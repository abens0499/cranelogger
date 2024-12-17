import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QTabWidget, 
                             QFormLayout, QLineEdit, QComboBox, QMessageBox, 
                             QDialog, QCompleter)
from PyQt5.QtCore import QTimer, QTime, Qt, QDateTime
from PyQt5.QtGui import QFont, QColor,QDoubleValidator
from datetime import datetime, timedelta

from excel_handler import ExcelHandler
from data_manager import DataManager

class SearchableComboBox(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        
        # Create a completer
        self.completer = QCompleter(self)
        self.completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchContains)
        self.setCompleter(self.completer)
        
        # Populate the completer
        self.addItems([
            'Maintenance', 
            'Operator Break', 
            'Equipment Issue', 
            'Material Shortage', 
            'Mechanical Failure',
            'Weather Conditions',
            'Waiting for Cargo',
            'Shift Change',
            'Other'
        ])

class CraneOperationSystem(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Initial list of operators
        self.initial_operators = [
            'John Doe', 
            'Jane Smith', 
            'Mike Johnson', 
            'Sarah Williams',
            'Alex Brown',
            'Emily Davis'
        ]
        
        # Initialize Excel and Data Management
        self.excel_handler = ExcelHandler()
        self.data_manager = DataManager(self.excel_handler)
        # Initialize crane_timer_labels before initUI
        self.crane_timer_labels = {
        1: None,
        2: None
        }
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Crane Operation Management System')
        self.setGeometry(100, 100, 800, 600)

        # Main widget and layout
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # Shift and Time Display
        self.shift_label = QLabel()
        self.time_label = QLabel()
        self.update_time_and_shift()
        
        time_layout = QHBoxLayout()
        time_layout.addWidget(self.shift_label)
        time_layout.addWidget(self.time_label)
        main_layout.addLayout(time_layout)

        # Timer to update time
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time_and_shift)
        self.timer.start(1000)  # Update every second

        # Tab Widget
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # Create Tabs
        self.cranes_tab = self.create_cranes_tab()
        self.barges_tab = self.create_barges_tab()
        self.generators_tab = self.create_generators_tab()
        self.ships_tab = self.create_ships_tab()

        self.tab_widget.addTab(self.cranes_tab, "Cranes")
        self.tab_widget.addTab(self.barges_tab, "Barges")
        self.tab_widget.addTab(self.generators_tab, "Generators")
        self.tab_widget.addTab(self.ships_tab, "Ships")

        # Crane Running Timers
        self.crane_timers = {
            1: QTimer(self),
            2: QTimer(self)
        }
        
        

    def update_time_and_shift(self):
        current_time = QTime.currentTime()
        self.time_label.setText(current_time.toString('hh:mm:ss'))
        
        current_shift = self.data_manager.get_current_shift()
        self.shift_label.setText(f'Current Shift: {current_shift}')

    def create_cranes_tab(self):
        crane_widget = QWidget()
        crane_layout = QHBoxLayout()
        
        # Function to create crane box
        def create_crane_box(crane_number):
            crane_box = QVBoxLayout()
            crane_label = QLabel(f'Crane {crane_number}')
            crane_status = QLabel('Stopped')
            crane_timer = QLabel('00:00:00')
            self.crane_timer_labels[crane_number] = crane_timer
            crane_operator = QLabel('No Operator')
            
            # Start/Stop Button
            crane_start_stop = QPushButton('Start')
            crane_start_stop.setStyleSheet("""
                QPushButton { 
                    background-color: green; 
                    color: white; 
                    font-weight: bold; 
                }
            """)
            
            crane_start_stop.clicked.connect(
                lambda checked, cn=crane_number, 
                css=crane_start_stop, 
                cs=crane_status, 
                ct=crane_timer, 
                co=crane_operator: 
                self.toggle_crane(cn, css, cs, ct, co)
            )
            
            # Assign Operator Button
            crane_assign_operator = QPushButton('Assign Operator')
            crane_assign_operator.clicked.connect(
                lambda checked, cn=crane_number, 
                co=crane_operator: 
                self.assign_operator(cn, co)
            )
            
            crane_box.addWidget(crane_label)
            crane_box.addWidget(crane_status)
            crane_box.addWidget(crane_timer)
            crane_box.addWidget(crane_operator)
            crane_box.addWidget(crane_start_stop)
            crane_box.addWidget(crane_assign_operator)
            
            return crane_box

        # Crane 1
        crane1_box = create_crane_box(1)
        
        # Crane 2
        crane2_box = create_crane_box(2)
        
        # Stop Both Cranes
        stop_both_button = QPushButton('Stop Both Cranes')
        stop_both_button.clicked.connect(self.stop_both_cranes)
        
        crane_layout.addLayout(crane1_box)
        crane_layout.addLayout(crane2_box)
        
        crane_widget.setLayout(crane_layout)
        crane_widget.layout().addWidget(stop_both_button)
        
        return crane_widget
    
    def toggle_crane(self, crane_number, button, status_label, timer_label, operator_label):
        if not self.data_manager.is_crane_running(crane_number):
            # Start crane
            operator = operator_label.text()
            if operator == 'No Operator':
                QMessageBox.warning(self, 'Error', 'Please assign an operator first.')
                return
            
            # Start crane
            self.data_manager.start_crane(crane_number, operator)
            button.setText('Stop')
            button.setStyleSheet("""
                QPushButton { 
                    background-color: red; 
                    color: white; 
                    font-weight: bold; 
                }
            """)
            status_label.setText('Running')
            
            # Start timer
            self.start_crane_timer(crane_number, timer_label)
        else:
            # Stop crane - show idle reason dialog
            result = self.show_idle_reason_dialog(crane_number)
            if result:
                button.setText('Start')
                button.setStyleSheet("""
                    QPushButton { 
                        background-color: green; 
                        color: white; 
                        font-weight: bold; 
                    }
                """)
                status_label.setText('Stopped')
                
                # Stop timer
                self.stop_crane_timer(crane_number)
                timer_label.setText('00:00:00')

    def start_crane_timer(self, crane_number, timer_label):
        def update_timer():
            elapsed_time = self.data_manager.get_crane_elapsed_time(crane_number)
            hours, remainder = divmod(elapsed_time.seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            timer_label.setText(f'{hours:02d}:{minutes:02d}:{seconds:02d}')
        
        # Create and start timer
        self.crane_timers[crane_number].timeout.connect(lambda: update_timer())
        self.crane_timers[crane_number].start(1000)  # Update every second

    def stop_crane_timer(self, crane_number):
        # Stop and disconnect the timer
        self.crane_timers[crane_number].stop()
        try:
            self.crane_timers[crane_number].timeout.disconnect()
        except TypeError:
            pass  # No connections to disconnect

    def show_idle_reason_dialog(self, crane_number):
        dialog = QDialog(self)
        dialog.setWindowTitle('Select Idle Reason')
        layout = QVBoxLayout()
        
        # Use custom searchable combo box
        idle_dropdown = SearchableComboBox()
        
        custom_reason = QLineEdit()
        custom_reason.setPlaceholderText('Enter custom reason if "Other" selected')
        custom_reason.setVisible(False)
        
        idle_dropdown.currentTextChanged.connect(
            lambda text: custom_reason.setVisible(text == 'Other')
        )
        
        layout.addWidget(QLabel('Select Idle Reason:'))
        layout.addWidget(idle_dropdown)
        layout.addWidget(custom_reason)
        
        buttons = QHBoxLayout()
        submit_btn = QPushButton('Submit')
        cancel_btn = QPushButton('Cancel')
        buttons.addWidget(submit_btn)
        buttons.addWidget(cancel_btn)
        
        layout.addLayout(buttons)
        dialog.setLayout(layout)
        
        submit_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)
        
        if dialog.exec_() == QDialog.Accepted:
            idle_reason = idle_dropdown.currentText()
            if idle_reason == 'Other':
                idle_reason = custom_reason.text() or 'Unspecified'
            
            # Stop both cranes
            for cn in [1, 2]:
                result = self.data_manager.stop_crane(cn, idle_reason)
            
            return True
        return False

    def stop_both_cranes(self):
        # Reuse the existing idle reason dialog for stopping both cranes
        result = self.show_idle_reason_dialog(1)  # Passes 1 as a placeholder
        
        # Reset UI for both cranes
        for crane_number in [1, 2]:
            # Reset button
            button = self.cranes_tab.findChild(QPushButton, f'crane{crane_number}_start_stop')
            if button:
                button.setText('Start')
                button.setStyleSheet("""
                    QPushButton { 
                        background-color: green; 
                        color: white; 
                        font-weight: bold; 
                    }
                """)
            
            # Reset status
            status_label = self.cranes_tab.findChild(QLabel, f'crane{crane_number}_status')
            if status_label:
                status_label.setText('Stopped')
            
            # Reset timer
            timer_label = self.crane_timer_labels[crane_number]
            if timer_label:
                timer_label.setText('00:00:00')
            
            # Stop timer
            self.stop_crane_timer(crane_number)

    def assign_operator(self, crane_number, operator_label):
        # Get available operators
        available_operators = self.data_manager.get_available_operators(self.initial_operators)
        
        if not available_operators:
            QMessageBox.warning(self, 'No Available Operators', 'All operators have been assigned.')
            return None

        dialog = QDialog(self)
        dialog.setWindowTitle(f'Assign Operator to Crane {crane_number}')
        layout = QVBoxLayout()
        
        # Operator Dropdown
        operators = QComboBox()
        operators.addItems(available_operators)
        
        submit_btn = QPushButton('Assign')
        submit_btn.clicked.connect(dialog.accept)
        
        layout.addWidget(QLabel(f'Select Operator for Crane {crane_number}:'))
        layout.addWidget(operators)
        layout.addWidget(submit_btn)
        
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            operator = operators.currentText()
            operator_label.setText(operator)
            return operator
        return None

    def create_barges_tab(self):
        barge_widget = QWidget()
        barge_layout = QFormLayout()
        
        # Barge Name/ID
        barge_name = QLineEdit()
        
        # Tons Loaded
        tons_loaded = QLineEdit()
        tons_loaded.setPlaceholderText('Enter tons loaded')
        tons_loaded.setValidator(QDoubleValidator())
        
        # Start/Stop Time (automatically captured)
        start_time_label = QLabel(datetime.now().strftime('%H:%M:%S'))
        stop_time_label = QLabel('Not Stopped')
        
        start_btn = QPushButton('Start Barge')
        stop_btn = QPushButton('Stop Barge')
        
        def start_barge():
            start_time_label.setText(datetime.now().strftime('%H:%M:%S'))
            start_btn.setEnabled(False)
            stop_btn.setEnabled(True)
        
        def stop_barge():
            if not barge_name.text() or not tons_loaded.text():
                QMessageBox.warning(self, 'Error', 'Please fill Barge Name and Tons Loaded')
                return
            
            stop_time = datetime.now()
            stop_time_label.setText(stop_time.strftime('%H:%M:%S'))
            
            # Log barge data
            try:
                self.excel_handler.log_barge_data(
                    barge_name.text(), 
                    datetime.now(),  # start time
                    stop_time,  # stop time
                    float(tons_loaded.text())  # tons loaded
                )
                QMessageBox.information(self, 'Success', 'Barge data logged successfully')
                
                # Reset UI
                barge_name.clear()
                tons_loaded.clear()
                start_time_label.setText(datetime.now().strftime('%H:%M:%S'))
                stop_time_label.setText('Not Stopped')
                start_btn.setEnabled(True)
                stop_btn.setEnabled(False)
            except ValueError:
                QMessageBox.warning(self, 'Error', 'Invalid tons loaded. Use a numeric value.')
        
        start_btn.clicked.connect(start_barge)
        stop_btn.clicked.connect(stop_barge)
        stop_btn.setEnabled(False)
        
        barge_layout.addRow('Barge Name/ID:', barge_name)
        barge_layout.addRow('Tons Loaded:', tons_loaded)
        barge_layout.addRow('Start Time:', start_time_label)
        barge_layout.addRow('Stop Time:', stop_time_label)
        barge_layout.addRow(start_btn, stop_btn)
        
        barge_widget.setLayout(barge_layout)
        return barge_widget

    def create_generators_tab(self):
        generator_widget = QWidget()
        generator_layout = QVBoxLayout()
        
        # Generators with start/stop functionality
        for i in range(1, 4):
            generator_box = QFormLayout()
            
            # Generator ID
            generator_id = QLabel(f'Generator {i}')
            
            # Timer Label
            timer_label = QLabel('00:00:00')
            
            # Status Label
            status_label = QLabel('Stopped')
            
            # Start/Stop Button
            start_stop_btn = QPushButton('Start')
            start_stop_btn.setStyleSheet("""
                QPushButton { 
                    background-color: green; 
                    color: white; 
                    font-weight: bold; 
                }
            """)
            
            # Timer for this generator
            generator_timer = QTimer(self)
            generator_start_time = None
            
            def create_generator_toggle_handler(gen_num, btn, status_lbl, timer_lbl, gen_timer):
                def toggle_generator():
                    nonlocal generator_start_time
                    
                    if btn.text() == 'Start':
                        # Start generator
                        generator_start_time = datetime.now()
                        btn.setText('Stop')
                        btn.setStyleSheet("""
                            QPushButton { 
                                background-color: red; 
                                color: white; 
                                font-weight: bold; 
                            }
                        """)
                        status_lbl.setText('Running')
                        
                        # Start timer
                        def update_timer():
                            if generator_start_time:
                                elapsed_time = datetime.now() - generator_start_time
                                hours, remainder = divmod(elapsed_time.seconds, 3600)
                                minutes, seconds = divmod(remainder, 60)
                                timer_lbl.setText(f'{hours:02d}:{minutes:02d}:{seconds:02d}')
                        
                        gen_timer.timeout.connect(update_timer)
                        gen_timer.start(1000)  # Update every second
                    else:
                        # Stop generator
                        stop_time = datetime.now()
                        
                        # Log generator data
                        self.excel_handler.log_generator_data(
                            f'Generator {gen_num}', 
                            generator_start_time, 
                            stop_time
                        )
                        
                        btn.setText('Start')
                        btn.setStyleSheet("""
                            QPushButton { 
                                background-color: green; 
                                color: white; 
                                font-weight: bold; 
                            }
                        """)
                        status_lbl.setText('Stopped')
                        timer_lbl.setText('00:00:00')
                        
                        # Stop timer
                        gen_timer.stop()
                        try:
                            gen_timer.timeout.disconnect()
                        except TypeError:
                            pass
                        
                        generator_start_time = None
                
                return toggle_generator
            
            # Bind the handler
            start_stop_btn.clicked.connect(
                create_generator_toggle_handler(i, start_stop_btn, status_label, timer_label, generator_timer)
            )
            
            generator_box.addRow('Generator:', generator_id)
            generator_box.addRow('Status:', status_label)
            generator_box.addRow('Running Time:', timer_label)
            generator_box.addRow(start_stop_btn)
            
            generator_layout.addLayout(generator_box)
        
        generator_widget.setLayout(generator_layout)
        return generator_widget

    def create_ships_tab(self):
        ships_widget = QWidget()
        ships_layout = QFormLayout()
        
        # Ship Name
        ship_name = QLineEdit()
        
        # Quantity to Load/Unload
        quantity = QLineEdit()
        quantity.setPlaceholderText('Enter quantity')
        
        # Number of Hatches
        hatches = QLineEdit()
        hatches.setPlaceholderText('Number of hatches')
        
        # Start/Finished Ship Buttons
        start_time_label = QLabel('Not Started')
        finished_time_label = QLabel('Not Finished')
        
        start_btn = QPushButton('Start Ship')
        finished_btn = QPushButton('Finish Ship')
        finished_btn.setEnabled(False)
        
        def start_ship():
            if not ship_name.text():
                QMessageBox.warning(self, 'Error', 'Please enter ship name')
                return
            
            start_time = datetime.now()
            start_time_label.setText(start_time.strftime('%H:%M:%S'))
            start_btn.setEnabled(False)
            finished_btn.setEnabled(True)
        
        def finish_ship():
            if not ship_name.text() or not quantity.text() or not hatches.text():
                QMessageBox.warning(self, 'Error', 'Please fill all fields')
                return
            
            finished_time = datetime.now()
            finished_time_label.setText(finished_time.strftime('%H:%M:%S'))
            
            try:
                # Log ship data
                self.excel_handler.log_ship_data(
                    ship_name.text(), 
                    datetime.now(),  # start time 
                    finished_time,  # finished time
                    float(quantity.text()),  # quantity
                    int(hatches.text())  # hatches
                )
                
                QMessageBox.information(self, 'Success', 'Ship data logged successfully')
                
                # Reset UI
                ship_name.clear()
                quantity.clear()
                hatches.clear()
                start_time_label.setText('Not Started')
                finished_time_label.setText('Not Finished')
                start_btn.setEnabled(True)
                finished_btn.setEnabled(False)
            except ValueError:
                QMessageBox.warning(self, 'Error', 'Invalid quantity or hatches. Use numeric values.')
        
        start_btn.clicked.connect(start_ship)
        finished_btn.clicked.connect(finish_ship)
        
        ships_layout.addRow('Ship Name:', ship_name)
        ships_layout.addRow('Quantity:', quantity)
        ships_layout.addRow('Number of Hatches:', hatches)
        ships_layout.addRow('Start Time:', start_time_label)
        ships_layout.addRow('Finished Time:', finished_time_label)
        ships_layout.addRow(start_btn, finished_btn)
        
        ships_widget.setLayout(ships_layout)
        return ships_widget

# Explicitly export the CraneOperationSystem class
__all__ = ['CraneOperationSystem']