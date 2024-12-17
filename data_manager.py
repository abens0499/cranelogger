import pandas as pd
from datetime import datetime, time, timedelta

class DataManager:
    def __init__(self, excel_handler):
        self.excel_handler = excel_handler
        
        # Crane tracking
        self.crane_states = {
            1: {
                'running': False,
                'start_time': None,
                'operator': None,
                'elapsed_time': timedelta()
            },
            2: {
                'running': False,
                'start_time': None,
                'operator': None,
                'elapsed_time': timedelta()
            }
        }

        # Shift definitions
        self.shifts = [
            {'name': 'Shift 1', 'start': time(7, 30), 'end': time(14, 0)},
            {'name': 'Shift 2', 'start': time(14, 0), 'end': time(22, 0)},
            {'name': 'Shift 3', 'start': time(22, 0), 'end': time(6, 0)}
        ]

        # Track used operators
        self.used_operators = set()

    def get_current_shift(self):
        now = datetime.now().time()
        for shift in self.shifts:
            if shift['name'] == 'Shift 3':
                if now >= shift['start'] or now < time(6, 0):
                    return shift['name']
            else:
                if shift['start'] <= now < shift['end']:
                    return shift['name']
        return 'Unknown Shift'

    def start_crane(self, crane_number, operator):
        if crane_number not in [1, 2]:
            raise ValueError("Invalid crane number")
        
        self.crane_states[crane_number]['running'] = True
        self.crane_states[crane_number]['start_time'] = datetime.now()
        self.crane_states[crane_number]['operator'] = operator
        
        # Mark operator as used
        self.used_operators.add(operator)

    def stop_crane(self, crane_number, idle_reason):
        if crane_number not in [1, 2]:
            raise ValueError("Invalid crane number")
        
        if not self.crane_states[crane_number]['running']:
            return None

        start_time = self.crane_states[crane_number]['start_time']
        stop_time = datetime.now()
        operator = self.crane_states[crane_number]['operator']

        # Log data to Excel
        self.excel_handler.log_crane_data(
            crane_number, 
            operator, 
            start_time, 
            stop_time, 
            idle_reason
        )

        # Reset crane state
        self.crane_states[crane_number]['running'] = False
        self.crane_states[crane_number]['start_time'] = None
        self.crane_states[crane_number]['operator'] = None

        return start_time, stop_time

    def get_crane_elapsed_time(self, crane_number):
        if not self.crane_states[crane_number]['running']:
            return timedelta()
        
        start_time = self.crane_states[crane_number]['start_time']
        return datetime.now() - start_time

    def is_crane_running(self, crane_number):
        return self.crane_states[crane_number]['running']

    def get_available_operators(self, initial_operators):
        # Return operators not yet used
        return [op for op in initial_operators if op not in self.used_operators]

    def reset_crane_timer(self, crane_number):
        self.crane_states[crane_number]['start_time'] = datetime.now()