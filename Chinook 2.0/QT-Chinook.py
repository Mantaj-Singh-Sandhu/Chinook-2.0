import sys
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QGridLayout, QPushButton,
    QLabel, QMessageBox, QProgressBar, QLineEdit, QGraphicsDropShadowEffect
)
from PyQt5.QtCore import Qt
from command_runner import CommandRunnerThread  # Import CommandRunnerThread

class ScriptRunnerApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Chinook")
        self.setGeometry(200, 200, 600, 400)

        # Main layout
        self.main_layout = QVBoxLayout()

        # Center-aligned message label
        self.info_label = QLabel("Enter the date, Click the Buttons, and Sit back.")
        self.info_label.setAlignment(Qt.AlignCenter)  # Center-align text
        self.info_label.setStyleSheet("font-size: 16px; font-weight: bold;")  # Larger, bold font
        self.main_layout.addWidget(self.info_label)

        # Input for the date
        self.date_input = QLineEdit()
        self.date_input.setPlaceholderText("Enter date (YYYY-MM-DD)")
        self.date_input.setStyleSheet("font-size: 14px;")  # Larger input font
        self.main_layout.addWidget(self.date_input)

        # Set the current date when the app starts
        self.set_current_date()

        # Grid layout for buttons
        self.grid_layout = QGridLayout()

        # Buttons for tasks
        self.task_repetitions = {
            "MOP": 17,
            "BUZ": 9,
            "DRN": 16,
            "TPA": 2,
            "SQA": 5,
        }

        row = 0
        col = 0
        for task_name in self.task_repetitions.keys():
            button = QPushButton(task_name)
            
            # Set the button style (without box-shadow for compatibility)
            button.setStyleSheet(
                """
                QPushButton {
                    font-size: 14px;
                    padding: 10px;
                    margin: 9px;
                    background-color: #0078D4;
                    color: white;
                    border-radius: 5px;
                    border: none;
                }
                QPushButton:pressed {
                    background-color: #005BB5;
                }
                """
            )
            
            # Apply shadow effect to the button
            shadow_effect = QGraphicsDropShadowEffect()
            shadow_effect.setOffset(2, 2)  # Set shadow offset
            shadow_effect.setBlurRadius(5)  # Set shadow blur radius
            shadow_effect.setColor(Qt.black)  # Set shadow color

            button.setGraphicsEffect(shadow_effect)

            # Connect the button's click event to handle the task
            button.clicked.connect(lambda checked, task=task_name: self.handle_file_action(task))
            
            # Add the button to the grid layout
            self.grid_layout.addWidget(button, row, col)
            col += 1
            if col > 1:  # Move to the next row after 2 columns
                col = 0
                row += 1

        self.main_layout.addLayout(self.grid_layout)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("height: 20px;")  # Increase height
        self.main_layout.addWidget(self.progress_bar)

        self.setLayout(self.main_layout)

        # Initialize task-related attributes
        self.first_script = "From_AWS.py"
        self.remaining_scripts = [
            "chinook.py",
            "fmi.py",
            "j1939_stage1.py",
            "j1939_stage2.py",
            "j1939_stage3.py",
            "CDL_stage1.py",
            "CDL_stage2.py",
            "CDL_stage3.py",
            "one.py"
        ]

        self.runner_thread = None
        self.selected_repetitions = 1

    def set_current_date(self):
        """Set the current date in the date input field and save it to date.txt."""
        current_date = datetime.today().strftime('%Y-%m-%d')
        self.date_input.setText(current_date)
        self.update_date_file(current_date)

    def update_date_file(self, date_text):
        """Write the date to the date.txt file."""
        try:
            with open("date.txt", "w") as file:
                file.write(date_text)
        except Exception as e:
            self.info_label.setText(f"Failed to update date: {e}")

    def validate_date(self, date_text):
        """Validate the entered date is in YYYY-MM-DD format."""
        try:
            datetime.strptime(date_text, '%Y-%m-%d')
            return True
        except ValueError:
            return False

    def handle_file_action(self, task_name):
        """Handle the selected task and update the date file."""
        try:
            # Capture the current value of date_input
            date_text = self.date_input.text().strip()
            if not date_text:
                raise ValueError("Date field is empty!")

            # Validate date format
            if not self.validate_date(date_text):
                self.info_label.setText("Invalid date format! Use YYYY-MM-DD.")
                return

            # Update the date.txt file with the entered date
            self.update_date_file(date_text)

            # Get repetitions for the task
            repetitions = self.task_repetitions.get(task_name, 1)
            self.selected_repetitions = repetitions

            # Update devices_list.txt with the content of the selected task's file
            filename = f"{task_name}.txt"
            with open(filename, "r") as input_file:
                content = input_file.read()

            with open("devices_list.txt", "w") as devices_file:
                devices_file.write(content)

            self.info_label.setText(f"{filename} content copied to devices_list.txt successfully!")
            self.run_tasks()  # Run the sequence after handling the file
        except Exception as e:
            self.info_label.setText(f"Failed to handle {task_name}: {e}")

    def toggle_ui(self, enable=True):
        """Enable or disable the input and buttons during task execution."""
        self.date_input.setEnabled(enable)
        for i in range(self.grid_layout.count()):
            widget = self.grid_layout.itemAt(i).widget()
            if isinstance(widget, QPushButton):
                widget.setEnabled(enable)

    def run_tasks(self):
        """Run the tasks using the runner thread."""
        self.info_label.setText("Starting task execution...")
        self.progress_bar.setValue(1)

        # Disable the UI elements during task execution
        self.toggle_ui(enable=False)

        self.runner_thread = CommandRunnerThread(
            self.first_script, self.remaining_scripts, self.selected_repetitions
        )
        self.runner_thread.progress.connect(self.update_status)
        self.runner_thread.progress_bar_update.connect(self.progress_bar.setValue)
        self.runner_thread.finished.connect(self.task_complete)
        self.runner_thread.start()

    def update_status(self, message):
        """Update the info label with the progress message."""
        self.info_label.setText(message)

    def task_complete(self):
        """Handle the task completion."""
        self.info_label.setText("All tasks completed successfully!")
        QMessageBox.information(self, "Success", "Report Files are ready in excel_outputs Folder")

        # Re-enable the UI elements after task completion
        self.toggle_ui(enable=True)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ScriptRunnerApp()
    window.show()
    sys.exit(app.exec_())
