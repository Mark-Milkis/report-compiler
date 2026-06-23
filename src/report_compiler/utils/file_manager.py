"""
File management utilities for temporary files and cleanup.
"""

import hashlib
import os
import time
from pathlib import Path
from typing import List, Optional
from ..core.config import Config
from .logging_config import get_file_logger


class FileManager:
    """Manages temporary files and cleanup operations."""

    def __init__(self, keep_temp: bool = False, work_dir: Optional[str] = None):
        """
        Args:
            keep_temp: Whether to keep temporary files for debugging.
            work_dir: Directory in which to place temporary files. When provided,
                all temp files are created here (instead of alongside the source
                document), which keeps cloud-synced source folders free of churn
                and avoids OneDrive sync/lock issues. The directory is created if
                needed and removed on cleanup if empty.
        """
        self.keep_temp = keep_temp
        self.temp_files: List[str] = []
        self.timestamp = int(time.time() * 1000)  # Millisecond timestamp
        self.logger = get_file_logger()
        self.work_dir = work_dir
        if self.work_dir:
            try:
                os.makedirs(self.work_dir, exist_ok=True)
            except OSError as e:
                # Fall back to source-adjacent temp files rather than failing the run.
                self.logger.warning(
                    "Could not create temp work directory %s (%s); "
                    "falling back to source-adjacent temp files.",
                    self.work_dir, e,
                )
                self.work_dir = None

    def generate_temp_path(self, base_path: str, suffix: str = "") -> str:
        """
        Generate a temporary file path based on the base path.

        Args:
            base_path: Original file path
            suffix: Additional suffix for the temp file

        Returns:
            Path to temporary file
        """
        base_name = os.path.splitext(os.path.basename(base_path))[0]
        ext = os.path.splitext(base_path)[1]

        if self.work_dir:
            base_dir = self.work_dir
            # When several source files from different directories share a base
            # name, a single flat work directory would collide. Disambiguate with
            # a short hash of the source's directory so the temp names stay unique.
            digest = hashlib.sha1(
                os.path.dirname(os.path.abspath(base_path)).encode("utf-8", "ignore")
            ).hexdigest()[:8]
            name_core = f"{base_name}_{digest}"
        else:
            base_dir = os.path.dirname(base_path)
            name_core = base_name

        if suffix:
            temp_name = f"{Config.TEMP_FILE_PREFIX}{name_core}_{suffix}_{self.timestamp}{ext}"
        else:
            temp_name = f"{Config.TEMP_FILE_PREFIX}{name_core}_{self.timestamp}{ext}"

        temp_path = os.path.join(base_dir, temp_name)
        self.temp_files.append(temp_path)
        return temp_path
    
    def retarget_temp_path(self, old_path: str, new_path: str) -> str:
        """
        Replace a previously generated temp path with a corrected one.

        Keeps the cleanup tracking list consistent when a caller needs to adjust a
        path returned by ``generate_temp_path`` (for example to force a ``.pdf``
        extension) so the file that is actually written still gets cleaned up.

        Args:
            old_path: The path previously returned by ``generate_temp_path``.
            new_path: The corrected path to track instead.

        Returns:
            The new path.
        """
        try:
            index = self.temp_files.index(old_path)
            self.temp_files[index] = new_path
        except ValueError:
            # Not tracked (e.g. caller passed an untracked path); just track the new one.
            self.temp_files.append(new_path)
        return new_path

    def create_temp_copy(self, source_path: str, suffix: str = "") -> str:
        """
        Create a temporary copy of a file.
        
        Args:
            source_path: Path to source file
            suffix: Additional suffix for the temp file
            
        Returns:
            Path to temporary copy        """
        import shutil
        
        temp_path = self.generate_temp_path(source_path, suffix)
        shutil.copy2(source_path, temp_path)
        return temp_path
    
    def cleanup(self) -> None:
        """Clean up all temporary files created by this manager."""
        if self.keep_temp:
            self.logger.info("Keeping temporary files for debugging...")
            for temp_file in self.temp_files:
                if os.path.exists(temp_file):
                    self.logger.info("  • %s", os.path.basename(temp_file))
            return
        
        self.logger.info("Cleaning up temporary files...")
        removed_count = 0
        
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                    self.logger.info("  ✓ Removed: %s", os.path.basename(temp_file))
                    removed_count += 1
            except Exception as e:
                self.logger.warning("  ⚠️ Could not remove %s: %s", os.path.basename(temp_file), e)
        
        if removed_count == 0 and len(self.temp_files) > 0:
            self.logger.info("  • No temporary files to clean up")

        self.temp_files.clear()

        # Remove the per-run work directory if we created one and it is now empty.
        if self.work_dir and os.path.isdir(self.work_dir):
            try:
                os.rmdir(self.work_dir)
            except OSError:
                # Not empty (e.g. unexpected leftovers) or in use; leave it in place.
                pass
    
    def __enter__(self):
        """Context manager entry."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit with automatic cleanup."""
        self.cleanup()
    
    @staticmethod
    def validate_path(file_path: str, must_exist: bool = False) -> Optional[str]:
        """
        Validate and normalize a file path.
        
        Args:
            file_path: Path to validate
            must_exist: Whether the file must exist
            
        Returns:
            Normalized absolute path, or None if invalid
        """
        try:
            path = Path(file_path).resolve()
            
            if must_exist and not path.exists():
                return None
            
            return str(path)
        except Exception:
            return None
    
    @staticmethod
    def ensure_directory_exists(file_path: str) -> bool:
        """
        Ensure the directory for a file path exists.
        
        Args:
            file_path: Path to file (directory will be created)
            
        Returns:
            True if directory exists or was created successfully
        """
        try:
            directory = os.path.dirname(file_path)
            if directory:
                os.makedirs(directory, exist_ok=True)
            return True
        except Exception:
            return False
    
    @staticmethod
    def get_file_size_mb(file_path: str) -> float:
        """
        Get file size in megabytes.
        
        Args:
            file_path: Path to file
            
        Returns:
            File size in MB, or 0 if file doesn't exist
        """
        try:
            return os.path.getsize(file_path) / (1024 * 1024)
        except Exception:
            return 0.0
    
    @staticmethod
    def is_file_locked(file_path: str) -> bool:
        """
        Check if a file is locked (in use by another process).
        
        Args:
            file_path: Path to file
            
        Returns:
            True if file appears to be locked
        """
        try:
            with open(file_path, 'r+b'):
                return False
        except (IOError, OSError):
            return True
    
    @staticmethod
    def copy_file(source_path: str, dest_path: str) -> bool:
        """
        Copy a file from source to destination.

        Args:
            source_path: Path to the source file.
            dest_path: Path to the destination file.

        Returns:
            True if copy was successful, False otherwise.
        """
        import shutil
        logger = get_file_logger()
        try:
            # Ensure destination directory exists
            FileManager.ensure_directory_exists(dest_path)
            shutil.copy2(source_path, dest_path)
            logger.debug(f"Successfully copied file from {source_path} to {dest_path}")
            return True
        except Exception as e:
            logger.error(f"Failed to copy file from {source_path} to {dest_path}: {e}")
            return False

    @staticmethod
    def move_file(source_path: str, dest_path: str) -> bool:
        """
        Move a file from source to destination.

        Args:
            source_path: Path to the source file.
            dest_path: Path to the destination file.

        Returns:
            True if move was successful, False otherwise.
        """
        import shutil
        logger = get_file_logger()
        try:
            # Ensure destination directory exists
            FileManager.ensure_directory_exists(dest_path)
            shutil.move(source_path, dest_path)
            logger.debug(f"Successfully moved file from {source_path} to {dest_path}")
            return True
        except Exception as e:
            logger.error(f"Failed to move file from {source_path} to {dest_path}: {e}")
            return False
