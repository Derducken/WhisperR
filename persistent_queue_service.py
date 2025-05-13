import json
import threading
from pathlib import Path
from typing import List, Optional, Any
from app_logger import get_logger, log_essential, log_error, log_extended, log_debug

# Name of the file to store pending tasks
PERSISTENT_QUEUE_FILENAME = "pending_transcriptions.json"

class PersistentTaskQueue:
    def __init__(self, storage_directory: Path):
        self.storage_path = storage_directory / PERSISTENT_QUEUE_FILENAME
        self._lock = threading.Lock()
        self._pending_tasks: List[str] = []
        self._load_tasks()
        log_essential(f"PersistentTaskQueue initialized. Loaded {len(self._pending_tasks)} pending tasks from {self.storage_path}")

    def _load_tasks(self) -> None:
        """Loads tasks from the persistent storage file."""
        with self._lock:
            try:
                if self.storage_path.exists():
                    with open(self.storage_path, 'r', encoding='utf-8') as f:
                        tasks_on_disk = json.load(f)
                        if isinstance(tasks_on_disk, list):
                            self._pending_tasks = [str(task) for task in tasks_on_disk if isinstance(task, str)]
                            log_extended(f"Loaded {len(self._pending_tasks)} tasks from {self.storage_path}")
                        else:
                            log_error(f"Persistent queue file {self.storage_path} does not contain a list. Initializing empty queue.")
                            self._pending_tasks = []
                            self._save_tasks_nolock() # Save an empty list to fix format
                else:
                    log_extended(f"Persistent queue file {self.storage_path} not found. Initializing empty queue.")
                    self._pending_tasks = []
                    self._save_tasks_nolock() # Create the file with an empty list
            except json.JSONDecodeError:
                log_error(f"Error decoding JSON from {self.storage_path}. Initializing with empty queue and attempting to overwrite.", exc_info=True)
                self._pending_tasks = []
                self._save_tasks_nolock() # Attempt to save a valid empty list
            except Exception as e:
                log_error(f"Failed to load tasks from {self.storage_path}: {e}", exc_info=True)
                # Keep potentially in-memory loaded tasks if any, or default to empty
                if not hasattr(self, '_pending_tasks') or self._pending_tasks is None:
                    self._pending_tasks = []


    def _save_tasks_nolock(self) -> bool:
        """Saves the current list of tasks to the persistent storage file (without acquiring lock)."""
        try:
            # Ensure directory exists
            self.storage_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.storage_path, 'w', encoding='utf-8') as f:
                json.dump(self._pending_tasks, f, indent=2)
            log_debug(f"Saved {len(self._pending_tasks)} tasks to {self.storage_path}")
            return True
        except Exception as e:
            log_error(f"Failed to save tasks to {self.storage_path}: {e}", exc_info=True)
            return False

    def add_task(self, task_filepath: str) -> bool:
        """Adds a task (audio file path) to the queue and persists it."""
        if not task_filepath:
            log_warning("Attempted to add an empty task filepath to persistent queue.")
            return False
        
        task_path_str = str(task_filepath) # Ensure it's a string

        with self._lock:
            if task_path_str not in self._pending_tasks:
                self._pending_tasks.append(task_path_str)
                if self._save_tasks_nolock():
                    log_extended(f"Task '{task_path_str}' added to persistent queue. Total: {len(self._pending_tasks)}")
                    return True
                else:
                    # If save failed, roll back the addition from in-memory list
                    try:
                        self._pending_tasks.remove(task_path_str)
                    except ValueError:
                        pass # Should not happen if logic is correct
                    log_error(f"Failed to save after adding task '{task_path_str}'. Task not added.")
                    return False
            else:
                log_extended(f"Task '{task_path_str}' already in persistent queue. Not adding again.")
                return True # Consider it success if already there

    def mark_task_complete(self, task_filepath: str) -> bool:
        """Removes a task from the queue upon completion and persists the change."""
        if not task_filepath:
            log_warning("Attempted to mark an empty task filepath as complete in persistent queue.")
            return False

        task_path_str = str(task_filepath)

        with self._lock:
            if task_path_str in self._pending_tasks:
                original_tasks = list(self._pending_tasks) # For potential rollback
                self._pending_tasks.remove(task_path_str)
                if self._save_tasks_nolock():
                    log_extended(f"Task '{task_path_str}' marked complete and removed from persistent queue. Remaining: {len(self._pending_tasks)}")
                    return True
                else:
                    # If save failed, roll back the removal
                    self._pending_tasks = original_tasks
                    log_error(f"Failed to save after marking task '{task_path_str}' complete. Task not removed from persistent state.")
                    return False
            else:
                log_warning(f"Task '{task_path_str}' not found in persistent queue to mark complete.")
                return False # Or True if "not found" is acceptable as "not needing removal"

    def get_pending_tasks(self) -> List[str]:
        """Returns a copy of the current list of pending tasks."""
        with self._lock:
            return list(self._pending_tasks)

    def get_queue_size(self) -> int:
        """Returns the number of pending tasks."""
        with self._lock:
            return len(self._pending_tasks)

    def clear_all_tasks(self) -> bool:
        """Clears all tasks from the queue and persists the change."""
        with self._lock:
            if not self._pending_tasks:
                log_extended("Persistent queue is already empty. No action taken for clear_all_tasks.")
                return True
            
            original_tasks = list(self._pending_tasks)
            self._pending_tasks = []
            if self._save_tasks_nolock():
                log_essential(f"All {len(original_tasks)} tasks cleared from persistent queue.")
                return True
            else:
                self._pending_tasks = original_tasks # Rollback
                log_error("Failed to save after clearing all tasks. Queue not cleared from persistent state.")
                return False

if __name__ == '__main__':
    # Example Usage (for testing purposes)
    # Ensure app_logger is minimally functional or mock it if running standalone
    class MockLogger:
        def __init__(self, name): self.name = name
        def info(self, msg): print(f"INFO: {msg}")
        def error(self, msg, exc_info=None): print(f"ERROR: {msg}")
        def warning(self, msg): print(f"WARNING: {msg}")
        def debug(self, msg): print(f"DEBUG: {msg}")

    if 'app_logger' not in globals() or not hasattr(globals()['app_logger'], 'get_logger'):
        def get_logger(name="test"): return MockLogger(name)
        log_essential = print
        log_error = print
        log_extended = print
        log_debug = print
        log_warning = print


    test_dir = Path("./test_persistent_queue_data")
    test_dir.mkdir(exist_ok=True)

    pq = PersistentTaskQueue(test_dir)

    print(f"Initial queue size: {pq.get_queue_size()}")
    print(f"Initial tasks: {pq.get_pending_tasks()}")

    pq.add_task("/path/to/audio1.wav")
    pq.add_task("/path/to/audio2.wav")
    print(f"After adds: {pq.get_pending_tasks()}")

    pq.mark_task_complete("/path/to/audio1.wav")
    print(f"After complete: {pq.get_pending_tasks()}")

    pq.add_task("/path/to/audio3.wav")
    print(f"Tasks before clear: {pq.get_pending_tasks()}")
    
    # Test loading from existing file
    del pq
    print("\nRe-initializing queue to test loading:")
    pq2 = PersistentTaskQueue(test_dir)
    print(f"Loaded tasks: {pq2.get_pending_tasks()}")
    
    pq2.clear_all_tasks()
    print(f"Tasks after clear: {pq2.get_pending_tasks()}")

    # Clean up test file
    # (pq2.storage_path).unlink(missing_ok=True)
    # test_dir.rmdir()
    print(f"\nTest complete. Check content of {pq2.storage_path}")
