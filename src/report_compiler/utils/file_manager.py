"""
File management utilities for temporary files and cleanup.
"""

import os
import tempfile
import time
from pathlib import Path
from typing import List, Optional
from ..core.config import Config
from .logging_config import get_file_logger


class FileManager:
    """Manages temporary files and cleanup operations."""
    
    def __init__(self, keep_temp: bool = False):
        self.keep_temp = keep_temp
        self.temp_files: List[str] = []
        self.timestamp = int(time.time() * 1000)  # Millisecond timestamp
        self.logger = get_file_logger()
        self.logger = get_file_logger()
    
    def generate_temp_path(self, base_path: str, suffix: str = "") -> str:
        """
        Generate a temporary file path based on the base path.
        
        Args:
            base_path: Original file path
            suffix: Additional suffix for the temp file
            
        Returns:
            Path to temporary file
        """
        base_dir = os.path.dirname(base_path)
        base_name = os.path.splitext(os.path.basename(base_path))[0]
        ext = os.path.splitext(base_path)[1]
        
        if suffix:
            temp_name = f"{Config.TEMP_FILE_PREFIX}{base_name}_{suffix}_{self.timestamp}{ext}"
        else:
            temp_name = f"{Config.TEMP_FILE_PREFIX}{base_name}_{self.timestamp}{ext}"
        
        temp_path = os.path.join(base_dir, temp_name)
        self.temp_files.append(temp_path)
        return temp_path
    
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
        try:
            shutil.copy2(source_path, dest_path)
            return True
        except Exception:
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
        try:
            shutil.move(source_path, dest_path)
            return True
        except Exception:
            return False
