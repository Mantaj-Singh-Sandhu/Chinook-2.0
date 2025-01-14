import sys
import subprocess
from PyQt5.QtCore import QThread, pyqtSignal
import time

class CommandRunnerThread(QThread):
    progress = pyqtSignal(str)
    progress_bar_update = pyqtSignal(int)  # Signal to update the progress bar
    finished = pyqtSignal()  # Signal emitted when all tasks are complete

    def __init__(self, first_script, remaining_scripts, repetitions):
        super().__init__()
        self.first_script = first_script
        self.remaining_scripts = remaining_scripts
        self.repetitions = repetitions
        self.input_parameter = 1  # Initial parameter for input.py

    def run(self):
        try:
            total_tasks = 1 + (len(self.remaining_scripts) * self.repetitions) + self.repetitions
            completed_tasks = 0

            # Run the first script once
            self.progress.emit(f"Downloading {self.first_script}...")
            subprocess.run([sys.executable, self.first_script], check=True)
            completed_tasks += 1
            self.progress_bar_update.emit(int((completed_tasks / total_tasks) * 100))

            # Run the remaining scripts multiple times
            for i in range(self.repetitions):  # Run scripts 'repetitions' times
                # Run input.py with the current parameter
                self.progress.emit(f"Running input.py {self.input_parameter}...")
                time.sleep(1)
                subprocess.run([sys.executable, "input.py", str(self.input_parameter)], check=True)
                completed_tasks += 1
                self.progress_bar_update.emit(int((completed_tasks / total_tasks) * 100))

                # Increment the input parameter for the next repetition
                self.input_parameter += 1

                # Run all remaining scripts
                for script in self.remaining_scripts:
                    self.progress.emit(f"Running {script}...")
                    subprocess.run([sys.executable, script], check=True)
                    completed_tasks += 1
                    self.progress_bar_update.emit(int((completed_tasks / total_tasks) * 100))

            # All tasks completed
            self.progress.emit("All tasks completed successfully!")
            self.finished.emit()  # Notify completion

        except subprocess.CalledProcessError as e:
            self.progress.emit(f"Error: {e}")