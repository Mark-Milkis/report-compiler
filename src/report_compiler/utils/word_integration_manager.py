"""
Word Integration Manager - Manages Word template installation, removal, and updates.
"""

import os
import platform
import shutil
from pathlib import Path
from typing import Optional, Tuple
from .logging_config import get_logger

class WordIntegrationManager:
    """Manages Word integration template installation and management."""
    
    TEMPLATE_FILENAME = "ReportCompilerTemplate.dotm"
    
    def __init__(self):
        self.logger = get_logger()
        self.platform = platform.system().lower()
        
    def get_word_startup_folder(self) -> Optional[Path]:
        """
        Get the Word startup folder path for the current platform.
        
        Returns:
            Path to Word startup folder, or None if not found/supported
        """
        if self.platform == "windows":
            # Windows: %APPDATA%\Microsoft\Word\STARTUP\
            appdata = os.environ.get("APPDATA")
            if appdata:
                startup_folder = Path(appdata) / "Microsoft" / "Word" / "STARTUP"
                return startup_folder
            else:
                self.logger.error("APPDATA environment variable not found")
                return None
        
        elif self.platform == "darwin":
            # macOS: ~/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/
            home = Path.home()
            startup_folder = home / "Library" / "Group Containers" / "UBF8T346G9.Office" / "User Content.localized" / "Startup.localized" / "Word"
            return startup_folder
            
        else:
            # Linux and other platforms - Word integration not typically supported
            self.logger.warning(f"Word integration is not supported on {self.platform}")
            return None
    
    def get_template_source_path(self) -> Path:
        """
        Get the path to the template file in the package.
        
        Returns:
            Path to the ReportCompilerTemplate.dotm file
        """
        # Get the package root directory
        package_root = Path(__file__).parent.parent.parent.parent
        template_path = package_root / "word_integration" / self.TEMPLATE_FILENAME
        return template_path
    
    def get_template_destination_path(self) -> Optional[Path]:
        """
        Get the full path where the template should be installed.
        
        Returns:
            Full path to template destination, or None if startup folder not found
        """
        startup_folder = self.get_word_startup_folder()
        if startup_folder:
            return startup_folder / self.TEMPLATE_FILENAME
        return None
    
    def is_template_installed(self) -> bool:
        """
        Check if the Word template is currently installed.
        
        Returns:
            True if template is installed, False otherwise
        """
        dest_path = self.get_template_destination_path()
        return dest_path is not None and dest_path.exists()
    
    def install_template(self) -> Tuple[bool, str]:
        """
        Install the Word template to the startup folder.
        
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            # Check if platform is supported
            startup_folder = self.get_word_startup_folder()
            if startup_folder is None:
                return False, f"Word integration is not supported on {self.platform}"
            
            # Get source template path
            source_path = self.get_template_source_path()
            if not source_path.exists():
                return False, f"Template file not found: {source_path}"
            
            # Get destination path
            dest_path = self.get_template_destination_path()
            if dest_path is None:
                return False, "Could not determine Word startup folder"
            
            # Check if already installed
            if dest_path.exists():
                return False, f"Template is already installed at: {dest_path}"
            
            # Create startup folder if it doesn't exist
            startup_folder.mkdir(parents=True, exist_ok=True)
            
            # Copy template file
            shutil.copy2(source_path, dest_path)
            
            self.logger.info(f"Successfully installed Word template to: {dest_path}")
            return True, f"Word template installed successfully to: {dest_path}"
            
        except PermissionError as e:
            error_msg = f"Permission denied when installing template: {e}"
            self.logger.error(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"Failed to install Word template: {e}"
            self.logger.error(error_msg)
            return False, error_msg
    
    def remove_template(self) -> Tuple[bool, str]:
        """
        Remove the Word template from the startup folder.
        
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            # Get destination path
            dest_path = self.get_template_destination_path()
            if dest_path is None:
                return False, f"Word integration is not supported on {self.platform}"
            
            # Check if template is installed
            if not dest_path.exists():
                return False, f"Template is not installed at: {dest_path}"
            
            # Remove template file
            dest_path.unlink()
            
            self.logger.info(f"Successfully removed Word template from: {dest_path}")
            return True, f"Word template removed successfully from: {dest_path}"
            
        except PermissionError as e:
            error_msg = f"Permission denied when removing template: {e}"
            self.logger.error(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"Failed to remove Word template: {e}"
            self.logger.error(error_msg)
            return False, error_msg
    
    def update_template(self) -> Tuple[bool, str]:
        """
        Update the Word template by replacing the existing one.
        
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            # Get source template path
            source_path = self.get_template_source_path()
            if not source_path.exists():
                return False, f"Template file not found: {source_path}"
            
            # Get destination path
            dest_path = self.get_template_destination_path()
            if dest_path is None:
                return False, f"Word integration is not supported on {self.platform}"
            
            # Check if template is currently installed
            if not dest_path.exists():
                return False, f"Template is not currently installed. Use 'install' command first."
            
            # Create backup of existing template
            backup_path = dest_path.with_suffix('.dotm.backup')
            shutil.copy2(dest_path, backup_path)
            
            try:
                # Update template file
                shutil.copy2(source_path, dest_path)
                
                # Remove backup since update was successful
                backup_path.unlink()
                
                self.logger.info(f"Successfully updated Word template at: {dest_path}")
                return True, f"Word template updated successfully at: {dest_path}"
                
            except Exception as e:
                # Restore from backup if update failed
                if backup_path.exists():
                    shutil.copy2(backup_path, dest_path)
                    backup_path.unlink()
                raise e
                
        except PermissionError as e:
            error_msg = f"Permission denied when updating template: {e}"
            self.logger.error(error_msg)
            return False, error_msg
        except Exception as e:
            error_msg = f"Failed to update Word template: {e}"
            self.logger.error(error_msg)
            return False, error_msg
    
    def get_status(self) -> dict:
        """
        Get the current status of Word integration.
        
        Returns:
            Dictionary with status information
        """
        status = {
            "platform": self.platform,
            "supported": self.get_word_startup_folder() is not None,
            "startup_folder": str(self.get_word_startup_folder()) if self.get_word_startup_folder() else None,
            "template_installed": self.is_template_installed(),
            "template_path": str(self.get_template_destination_path()) if self.get_template_destination_path() else None,
            "source_template_exists": self.get_template_source_path().exists(),
            "source_template_path": str(self.get_template_source_path())
        }
        return status