import urllib.request
import json
from pathlib import Path
import shutil
import os
import tempfile
import subprocess
import time 
import threading
from typing import Callable, Optional, List, Tuple, Dict, Any

try:
    import py7zr
    PY7ZR_AVAILABLE = True
    UnsupportedPy7zrMethodError = py7zr.exceptions.UnsupportedCompressionMethodError
except ImportError:
    PY7ZR_AVAILABLE = False
    class UnsupportedPy7zrMethodError(Exception): pass

from app_logger import log_debug, log_error, log_extended, log_essential, log_warning

class GitHubReleaseDownloader:
    def __init__(self,
                 repo_owner_slash_repo: str,
                 status_callback: Optional[Callable[[str], None]] = None,
                 progress_callback: Optional[Callable[[int], None]] = None,
                 completion_callback: Optional[Callable[[Optional[Path], Optional[Path]], None]] = None,
                 error_callback: Optional[Callable[[str], None]] = None
                 ):
        self.repo_url_fragment = repo_owner_slash_repo
        self.api_url = f"https://api.github.com/repos/{self.repo_url_fragment}/releases/latest"
        
        self._status_cb = status_callback
        self._progress_cb_percent = progress_callback
        self._completion_cb = completion_callback 
        self._error_cb = error_callback

        self.thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()

    def _status(self, msg: str):
        log_debug(f"[GH Downloader] Status: {msg}")
        if self._status_cb:
            self._status_cb(msg)

    def _progress_percent(self, percent: int):
        if self._progress_cb_percent:
            self._progress_cb_percent(min(max(0, percent), 100))
    
    def _notify_error(self, error_message: str):
        log_error(f"[GH Downloader] Error: {error_message}")
        if self._error_cb:
            self._error_cb(error_message)
        elif self._status_cb:
             self._status_cb(f"Error: {error_message}")

    def _notify_completion(self, result_exe_path: Optional[Path], temp_dir_path_used: Optional[Path]):
        if self._completion_cb:
            self._completion_cb(result_exe_path, temp_dir_path_used)

    def _get_latest_release_asset_info(self, asset_keyword: str, prefer_windows: bool = True) -> Optional[Tuple[str, str, int]]:
        self._status(f"Fetching latest release from '{self.repo_url_fragment}'...")
        try:
            repo_check_url = f"https://api.github.com/repos/{self.repo_url_fragment}"
            with urllib.request.urlopen(urllib.request.Request(repo_check_url, headers={'User-Agent': 'WhisperR-App'})) as response:
                if response.status != 200:
                    msg = f"Repository '{self.repo_url_fragment}' not found or inaccessible (HTTP {response.status})."
                    self._notify_error(msg)
                    return None
        except Exception as e:
            msg = f"Error verifying repository '{self.repo_url_fragment}': {e}"
            self._notify_error(msg)
            log_error(msg, exc_info=True) 
            return None

        try:
            req = urllib.request.Request(self.api_url, headers={'User-Agent': 'WhisperR-App'})
            with urllib.request.urlopen(req) as response:
                if response.status != 200:
                    msg = f"API request to '{self.api_url}' failed (HTTP {response.status})."
                    self._notify_error(msg)
                    return None
                release_data = json.loads(response.read().decode('utf-8'))
        except Exception as e:
            msg = f"Error fetching release info from '{self.api_url}': {e}"
            self._notify_error(msg)
            log_error(msg, exc_info=True)
            return None

        assets = release_data.get('assets')
        if not assets:
            msg = "No assets found in the latest release."
            self._notify_error(msg)
            log_error(f"No assets found in latest release data for '{self.repo_url_fragment}'. Release data: {release_data.get('name', 'N/A')}")
            return None

        log_debug(f"Available assets for release '{release_data.get('name', 'N/A')}': {[a['name'] for a in assets]}")
        # Collect all matching assets
        matching_assets = []
        for asset in assets:
            name_lower = asset['name'].lower()
            if asset_keyword.lower() in name_lower:
                if prefer_windows and 'windows' not in name_lower:
                    continue
                if not name_lower.endswith('.7z'):
                    continue
                matching_assets.append(asset)

        if not matching_assets:
            log_warning(f"Could not find a '.7z' asset matching '{asset_keyword}'. Looking for any asset matching '{asset_keyword}'.")
            for asset in assets:
                name_lower = asset['name'].lower()
                if asset_keyword.lower() in name_lower:
                    matching_assets.append(asset)
                    break

        # Sort matching assets by upload time (newest first)
        matching_assets.sort(key=lambda x: x['updated_at'], reverse=True)

        def _parse_version(name: str) -> List[int]:
            """Parse version string from filename into list of integers."""
            # Look for patterns like r192.3.4 or v1.2.3
            version_str = None
            if 'r' in name.lower():
                try:
                    version_str = name.lower().split('r')[1].split('_')[0]
                except IndexError:
                    return []
            elif 'v' in name.lower():
                try:
                    version_str = name.lower().split('v')[1].split('_')[0]
                except IndexError:
                    return []
            
            if not version_str:
                return []
            
            # Split into components and convert to integers
            try:
                return [int(part) for part in version_str.split('.')]
            except ValueError:
                return []

        def _compare_versions(v1: List[int], v2: List[int]) -> int:
            """Compare two version lists. Returns 1 if v1 > v2, -1 if v1 < v2, 0 if equal."""
            for i in range(max(len(v1), len(v2))):
                part1 = v1[i] if i < len(v1) else 0
                part2 = v2[i] if i < len(v2) else 0
                if part1 > part2:
                    return 1
                elif part1 < part2:
                    return -1
            return 0

        # Try to parse version numbers from filenames
        versioned_assets = []
        for asset in matching_assets:
            version_parts = _parse_version(asset['name'])
            versioned_assets.append((version_parts, asset))

        # Sort by version (newest first) then by upload time
        versioned_assets.sort(
            key=lambda x: (x[0], x[1]['updated_at']),
            reverse=True
        )
        found_asset = versioned_assets[0][1] if versioned_assets else None
        
        if found_asset:
            asset_url = found_asset['browser_download_url']
            asset_name = found_asset['name']
            asset_size = found_asset.get('size', -1)
            log_essential(f"Found asset: '{asset_name}' (Size: {asset_size} bytes) for keyword '{asset_keyword}'. URL: {asset_url}")
            self._status(f"Found asset: {asset_name}")
            return asset_url, asset_name, asset_size
        else:
            msg = f"No asset found matching keyword '{asset_keyword}' and ending with .7z (or any matching asset)."
            self._notify_error(msg)
            log_error(f"No asset found for '{self.repo_url_fragment}' matching keyword '{asset_keyword}'. Searched {len(assets)} assets.")
            return None

    def _download_file(self, url: str, download_to_path: Path, asset_name: str, asset_size: int) -> bool:
        self._status(f"Downloading '{asset_name}'...")
        self._progress_percent(0)
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'WhisperR-App'})
            with urllib.request.urlopen(req) as response, open(download_to_path, 'wb') as out_file:
                effective_size = asset_size if asset_size > 0 else int(response.getheader('Content-Length', -1))
                block_size = 8192 * 4 
                downloaded_bytes = 0
                while True:
                    if self._stop_event.is_set():
                        self._status("Download cancelled.")
                        return False
                    buffer = response.read(block_size)
                    if not buffer:
                        break
                    out_file.write(buffer)
                    downloaded_bytes += len(buffer)
                    if effective_size > 0:
                        percent = int(downloaded_bytes * 100 / effective_size)
                        self._progress_percent(percent)
            
            self._progress_percent(100)
            self._status(f"'{asset_name}' downloaded successfully.")
            log_extended(f"File '{asset_name}' downloaded to {download_to_path}")
            return True
        except Exception as e:
            msg = f"Download of '{asset_name}' failed: {e}"
            self._notify_error(msg)
            log_error(f"Failed to download '{url}' to '{download_to_path}': {e}", exc_info=True)
            if download_to_path.exists():
                try: download_to_path.unlink()
                except OSError: pass
            return False

    def _extract_archive(self, archive_path: Path, extract_to_dir: Path) -> bool:
        if not archive_path.name.lower().endswith('.7z'):
            msg = f"Unsupported archive type: {archive_path.name}. Only .7z is supported."
            self._notify_error(msg)
            return False

        self._status(f"Preparing to extract '{archive_path.name}' to '{extract_to_dir}'...")
        self._progress_percent(0) 

        # Robustly clean the target extraction directory before attempting to extract
        if extract_to_dir.exists() and extract_to_dir.is_dir():
            self._status(f"Cleaning existing target directory: {extract_to_dir}...")
            cleaned_successfully = False
            for attempt in range(3): # Try up to 3 times
                try:
                    # Check if directory is empty before attempting rmtree
                    if not any(extract_to_dir.iterdir()):
                        log_debug(f"Target directory '{extract_to_dir}' is already empty or does not exist (attempt {attempt+1}).")
                        cleaned_successfully = True
                        break # No need to rmtree if empty or gone
                    
                    shutil.rmtree(extract_to_dir)
                    log_debug(f"Removed existing extraction directory: {extract_to_dir} (attempt {attempt + 1})")
                    cleaned_successfully = True
                    break # Success
                except FileNotFoundError: # If rmtree was called on a non-existent dir due to race condition
                    log_debug(f"Target directory '{extract_to_dir}' was not found during cleanup attempt {attempt+1}, considering it cleaned.")
                    cleaned_successfully = True
                    break
                except PermissionError as e_perm:
                    log_warning(f"Attempt {attempt + 1} to remove '{extract_to_dir}' failed with PermissionError: {e_perm}. Retrying after short delay...")
                    if attempt < 2: # Don't sleep on the last attempt
                        time.sleep(0.5 + attempt * 0.5) # Increasing delay: 0.5s, 1s
                except Exception as e:
                    log_error(f"Attempt {attempt + 1} to remove '{extract_to_dir}' failed with unexpected error: {e}", exc_info=True)
                    if attempt < 2:
                        time.sleep(0.5 + attempt * 0.5)
            
            if not cleaned_successfully:
                msg = f"Critical: Could not clean existing target directory '{extract_to_dir}' after multiple attempts. Extraction cannot proceed safely."
                self._notify_error(msg)
                log_error(msg + " Please check for locked files or permissions in the target directory.")
                return False
        
        try:
            extract_to_dir.mkdir(parents=True, exist_ok=True) # exist_ok=True is fine if cleaned_successfully handled it
            self._status(f"Target directory '{extract_to_dir}' ready for extraction.")
        except Exception as e:
            msg = f"Error creating target directory '{extract_to_dir}': {e}"
            self._notify_error(msg)
            return False

        self._status(f"Extracting '{archive_path.name}'...") # Update status after prep
        extracted_successfully = False
        
        if PY7ZR_AVAILABLE:
            try:
                self._status("Attempting extraction with py7zr library...")
                with py7zr.SevenZipFile(archive_path, mode='r') as z:
                    z.extractall(path=extract_to_dir)
                extracted_successfully = True
                log_extended(f"Extracted '{archive_path.name}' using py7zr to '{extract_to_dir}'.")
            except UnsupportedPy7zrMethodError as bcj_error: 
                self._status(f"py7zr: Unsupported method (e.g., BCJ2). Will try system 7z. Error: {str(bcj_error)[:200]}") # Truncate long error
                log_warning(f"py7zr does not support a method in '{archive_path}': {bcj_error}. Falling back to system 7z.")
            except Exception as py7zr_error:
                self._status(f"py7zr extraction failed: {str(py7zr_error)[:200]}. Will try system 7z if available.")
                log_error(f"py7zr extraction for '{archive_path}' failed: {py7zr_error}", exc_info=True)
        
        if not extracted_successfully:
            try:
                self._status("Attempting extraction with system '7z' command...")
                # Check common Windows 7z.exe locations
                seven_zip_paths = [
                    "7z",  # Check PATH first
                    "C:\\Program Files\\7-Zip\\7z.exe",
                    "C:\\Program Files (x86)\\7-Zip\\7z.exe"
                ]
                
                seven_zip_exe = None
                for path in seven_zip_paths:
                    try:
                        subprocess.run([path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=2,
                                      creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
                        seven_zip_exe = path
                        break
                    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
                        continue
                
                if not seven_zip_exe:
                    msg = ("7-Zip not found. Please ensure 7-Zip is installed or provide the path to 7z.exe.\n\n"
                          "You can download 7-Zip from https://www.7-zip.org/")
                    self._notify_error(msg)
                    return False

                cmd = [seven_zip_exe, 'x', str(archive_path), f'-o{str(extract_to_dir)}', '-aoa', '-y']
                log_extended(f"Using 7z at: {seven_zip_exe}")
                log_debug(f"Running system 7z command: {' '.join(cmd)}")
                
                startupinfo = None
                if os.name == 'nt':
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE
                
                process = subprocess.run(cmd, check=True, capture_output=True, text=True, encoding='utf-8', errors='replace', startupinfo=startupinfo)
                
                if process.stdout: log_extended(f"7z extraction stdout: {process.stdout.strip()}")
                if process.stderr: log_warning(f"7z extraction stderr (normal for some info): {process.stderr.strip()}") 
                
                extracted_successfully = True
                log_extended(f"Extracted '{archive_path.name}' using system 7z to '{extract_to_dir}'.")
            except subprocess.CalledProcessError as e:
                err_output_stdout = e.stdout.strip() if e.stdout else ""
                err_output_stderr = e.stderr.strip() if e.stderr else ""
                full_err_output = (err_output_stdout + "\n" + err_output_stderr).strip()

                # More specific error message for the user
                if "Cannot delete output file" in full_err_output:
                    msg = f"System 7z: Failed to overwrite a file in the target directory (it might be locked). Details: {err_output_stderr}"
                elif "Sub items Errors: 1" in err_output_stdout: # 7zip often reports this on stdout for various errors
                     msg = f"System 7z: Extraction completed with errors. Details: {full_err_output}"
                else:
                    msg = f"System 7z extraction failed. Output: {full_err_output}"
                
                self._notify_error(msg[:300]) # Truncate very long CLI outputs for UI
                log_error(f"System 7z extraction for '{archive_path}' failed (return code {e.returncode}):\nSTDOUT: {e.stdout}\nSTDERR: {e.stderr}", exc_info=False)
            except FileNotFoundError:
                 msg = "System '7z' command not found. Cannot extract."
                 self._notify_error(msg)
            except Exception as sys_7z_error:
                msg = f"System 7z extraction failed with an unexpected error: {sys_7z_error}"
                self._notify_error(msg)
                log_error(f"System 7z extraction for '{archive_path}' encountered an error: {sys_7z_error}", exc_info=True)
        
        if extracted_successfully:
            self._status("Extraction completed successfully.")
            self._progress_percent(100) 
            return True
        else:
            log_error(f"Failed to extract '{archive_path.name}' to '{extract_to_dir}' after all attempts.")
            return False

    def _find_executable(self, base_dir: Path, executable_names: List[str]) -> Optional[Path]:
        self._status(f"Searching for executables ({', '.join(executable_names)}) in '{base_dir}'...")
        for root_str, _, files in os.walk(base_dir):
            root = Path(root_str) 
            for file_name in files:
                if file_name.lower() in [name.lower() for name in executable_names]:
                    exe_path = root / file_name
                    self._status(f"Found executable: {exe_path.name} at {exe_path.parent}")
                    log_debug(f"Found executable '{exe_path}' during search.")
                    return exe_path
        log_warning(f"Could not find any of {executable_names} in '{base_dir}' after extraction.")
        return None

    def _perform_download_and_extract(self,
                                      asset_keyword: str,
                                      target_extraction_dir: Path,
                                      executable_names: List[str],
                                      prefer_windows_in_asset_name: bool = True):
        exe_path_found: Optional[Path] = None
        temp_dir_obj_for_download: Optional[tempfile.TemporaryDirectory] = None 

        try:
            asset_info = self._get_latest_release_asset_info(asset_keyword, prefer_windows_in_asset_name)
            if not asset_info:
                self._notify_completion(None, None) 
                return

            download_url, asset_filename, asset_size = asset_info
            
            temp_dir_obj_for_download = tempfile.TemporaryDirectory(prefix="whisperr_dl_", suffix="_archive")
            temp_download_parent_path = Path(temp_dir_obj_for_download.name)
            downloaded_asset_filepath = temp_download_parent_path / asset_filename
            
            if not self._download_file(download_url, downloaded_asset_filepath, asset_filename, asset_size):
                self._notify_completion(None, Path(temp_dir_obj_for_download.name) if temp_dir_obj_for_download else None)
                return

            if self._stop_event.is_set(): return

            if not self._extract_archive(downloaded_asset_filepath, target_extraction_dir):
                self._notify_completion(None, Path(temp_dir_obj_for_download.name) if temp_dir_obj_for_download else None)
                return
            
            if self._stop_event.is_set(): return

            exe_path_found = self._find_executable(target_extraction_dir, executable_names)

            if exe_path_found:
                self._status(f"Process complete. Executable: {exe_path_found}")
                log_essential(f"Successfully installed '{asset_keyword}'. Executable found at '{exe_path_found}'.")
                self._notify_completion(exe_path_found, Path(temp_dir_obj_for_download.name) if temp_dir_obj_for_download else None)
            else:
                # This error message will be shown if _find_executable returns None
                self._notify_error(f"Executable not found in '{target_extraction_dir}' after extraction.")
                self._notify_completion(None, Path(temp_dir_obj_for_download.name) if temp_dir_obj_for_download else None)

        except Exception as e:
            error_message = f"Unexpected error in download worker: {e}"
            log_error(error_message, exc_info=True)
            self._notify_error(error_message)
            self._notify_completion(None, Path(temp_dir_obj_for_download.name) if temp_dir_obj_for_download else None)

    def download_extract_and_find_exe_threaded(self,
                                               asset_keyword: str,
                                               target_extraction_dir: Path,
                                               executable_names: List[str],
                                               prefer_windows_in_asset_name: bool = True):
        if self.thread and self.thread.is_alive():
            self._notify_error("Another download process is already running.")
            return

        self._stop_event.clear()
        self.thread = threading.Thread(
            target=self._perform_download_and_extract,
            args=(asset_keyword, target_extraction_dir, executable_names, prefer_windows_in_asset_name),
            daemon=True 
        )
        self.thread.start()

    def cancel_download(self): 
        if self.thread and self.thread.is_alive():
            self._stop_event.set()
            self._status("Cancellation requested...")
