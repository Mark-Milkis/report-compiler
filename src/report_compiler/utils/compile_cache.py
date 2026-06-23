"""
Persistent cache of compiled sub-document PDFs.

Compiling a nested DOCX insert is the most expensive step in a run: each one
invokes Microsoft Word (or LibreOffice) to render a PDF. When a large report
fails late in the pipeline, re-running it would normally re-compile every
appendix from scratch. This cache stores each compiled sub-document PDF keyed
on a content signature of its source (and that source's dependencies), so a
re-run reuses unchanged appendices instead of rebuilding them.

The cache is intentionally conservative: any condition that prevents computing a
reliable signature (unreadable file, parse error) yields a volatile key, so a
stale PDF is never served in place of changed content.
"""

import os
import shutil
import time
import hashlib
from typing import Optional

from ..core.config import Config
from .logging_config import get_file_logger


class CompileCache:
    """Content-addressed cache for compiled sub-document PDFs."""

    _READ_CHUNK = 1024 * 1024  # 1 MiB

    def __init__(self, cache_dir: str, enabled: bool = True):
        self.enabled = enabled
        self.cache_dir = cache_dir
        self.logger = get_file_logger()
        # Lazily imported to avoid a document<->utils import cycle.
        self._parser = None
        if self.enabled:
            try:
                os.makedirs(self.cache_dir, exist_ok=True)
                self._prune()
            except OSError as e:
                self.logger.warning(
                    "Could not initialise compile cache at %s (%s); caching disabled.",
                    self.cache_dir, e,
                )
                self.enabled = False

    # -- public API ---------------------------------------------------------

    def compute_key(self, docx_path: str) -> Optional[str]:
        """Compute a stable cache key for a source DOCX and its dependencies.

        Returns None if caching is disabled. The key changes whenever the source
        document, or any file it references (transitively, for nested DOCX
        inserts), changes.
        """
        if not self.enabled:
            return None
        return self._signature(docx_path, set())

    def get(self, key: Optional[str]) -> Optional[str]:
        """Return the path to a cached PDF for ``key``, or None on a miss."""
        if not self.enabled or not key:
            return None
        path = self._path_for(key)
        if os.path.exists(path):
            # Touch so frequently reused entries survive TTL pruning.
            try:
                os.utime(path, None)
            except OSError:
                pass
            return path
        return None

    def put(self, key: Optional[str], pdf_path: str) -> Optional[str]:
        """Copy a freshly compiled PDF into the cache and return its cached path.

        Writes to a temporary name and atomically renames so a concurrent run
        never observes a half-written cache entry. Failures are non-fatal.
        """
        if not self.enabled or not key:
            return None
        if not os.path.exists(pdf_path):
            return None
        dest = self._path_for(key)
        tmp = f"{dest}.{os.getpid()}.{int(time.time() * 1000)}.tmp"
        try:
            shutil.copy2(pdf_path, tmp)
            os.replace(tmp, dest)
            return dest
        except OSError as e:
            self.logger.warning("Could not write cache entry for %s (%s).",
                                os.path.basename(pdf_path), e)
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass
            return None

    # -- internals ----------------------------------------------------------

    def _path_for(self, key: str) -> str:
        return os.path.join(self.cache_dir, f"{key}.pdf")

    def _get_parser(self):
        if self._parser is None:
            from ..document.placeholder_parser import PlaceholderParser
            self._parser = PlaceholderParser()
        return self._parser

    def _hash_file(self, path: str, hasher) -> None:
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(self._READ_CHUNK), b""):
                hasher.update(chunk)

    def _signature(self, docx_path: str, seen: set) -> str:
        """Recursively hash a DOCX and the files it references.

        The document's own bytes capture its text (and therefore every
        placeholder's parameters, e.g. page ranges). Referenced files are folded
        in by content (nested DOCX) or by size+mtime (PDFs/images), so a changed
        dependency invalidates the parent even when the parent text is untouched.
        """
        docx_path = os.path.abspath(docx_path)
        hasher = hashlib.sha256()

        if docx_path in seen:
            # Cycle guard: contribute a stable marker and stop recursing.
            hasher.update(b"<cycle>")
            return hasher.hexdigest()
        seen.add(docx_path)

        try:
            self._hash_file(docx_path, hasher)
        except OSError:
            # Cannot read the source -> never serve a cached hit for it.
            return self._volatile_key()

        base_dir = os.path.dirname(docx_path)
        try:
            placeholders = self._get_parser().find_all_placeholders(docx_path)
        except Exception:
            return self._volatile_key()

        deps = placeholders.get("table", []) + placeholders.get("paragraph", [])
        parts = []
        for p in deps:
            file_path = p.get("file_path")
            if not file_path:
                continue
            dep_abs = os.path.abspath(os.path.join(base_dir, file_path))
            if dep_abs.lower().endswith(".docx") and os.path.exists(dep_abs):
                parts.append(self._signature(dep_abs, seen))
            else:
                try:
                    st = os.stat(dep_abs)
                    parts.append(f"{dep_abs}:{st.st_size}:{st.st_mtime_ns}")
                except OSError:
                    parts.append(f"{dep_abs}:missing")

        for part in sorted(parts):
            hasher.update(part.encode("utf-8", "ignore"))
        return hasher.hexdigest()

    @staticmethod
    def _volatile_key() -> str:
        """A never-repeating key, guaranteeing a cache miss."""
        return "volatile-" + os.urandom(16).hex()

    def _prune(self) -> None:
        """Delete cache entries not accessed within the configured TTL."""
        ttl_seconds = Config.CACHE_TTL_DAYS * 24 * 3600
        cutoff = time.time() - ttl_seconds
        try:
            entries = os.listdir(self.cache_dir)
        except OSError:
            return
        for name in entries:
            if not name.endswith(".pdf"):
                continue
            path = os.path.join(self.cache_dir, name)
            try:
                if os.path.getmtime(path) < cutoff:
                    os.remove(path)
            except OSError:
                pass
