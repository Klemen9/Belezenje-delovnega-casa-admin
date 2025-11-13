# Beleženje delovnega časa Admin, program spremljanje podatkov o delovnem času.
# Copyright © 2025  Klemen Ambrožič
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
# E-mail: klemen.ambrozic@oscg.si

import sys
import os
from datetime import datetime, timedelta
import sqlite3
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                            QTableWidget, QTableWidgetItem, QComboBox, 
                            QTabWidget, QCalendarWidget, QMessageBox, 
                            QFileDialog, QCheckBox, QGroupBox, QDialog,
                            QTableWidget, QHeaderView, QStyle, QRadioButton,
                            QSpinBox, QMenu, QScrollArea, QGridLayout)
from PyQt6.QtGui import QAction, QTextCharFormat, QBrush, QColor, QFont, QPen, QIcon
from PyQt6.QtCore import QSize
from PyQt6.QtCore import Qt, QDate, QTimer, QThread, pyqtSignal
import pandas as pd
from smb.SMBConnection import SMBConnection
import configparser
import keyboard
import threading
import io
import json
import re

class DateRangeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Izberi obdobje")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # Start date
        start_layout = QHBoxLayout()
        start_layout.addWidget(QLabel("Začetni datum:"))
        self.start_calendar = QCalendarWidget()
        start_layout.addWidget(self.start_calendar)
        layout.addLayout(start_layout)
        
        # End date
        end_layout = QHBoxLayout()
        end_layout.addWidget(QLabel("Končni datum:"))
        self.end_calendar = QCalendarWidget()
        end_layout.addWidget(self.end_calendar)
        layout.addLayout(end_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("Potrdi")
        cancel_button = QPushButton("Prekliči")
        
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
    
    def get_dates(self):
        return self.start_calendar.selectedDate().toPyDate(), self.end_calendar.selectedDate().toPyDate()

class ArchiveDialog(QDialog):
    def __init__(self, worker_name, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Arhiviraj delavca - {worker_name}")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # Worker name display
        worker_label = QLabel(f"Delavec: {worker_name}")
        worker_label.setStyleSheet("font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(worker_label)
        
        # Start date
        start_layout = QHBoxLayout()
        start_layout.addWidget(QLabel("Začetni datum:"))
        self.start_calendar = QCalendarWidget()
        start_layout.addWidget(self.start_calendar)
        layout.addLayout(start_layout)
        
        # End date
        end_layout = QHBoxLayout()
        end_layout.addWidget(QLabel("Končni datum:"))
        self.end_calendar = QCalendarWidget()
        end_layout.addWidget(self.end_calendar)
        layout.addLayout(end_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("Arhiviraj")
        cancel_button = QPushButton("Prekliči")
        
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
    
    def get_dates(self):
        return self.start_calendar.selectedDate().toPyDate(), self.end_calendar.selectedDate().toPyDate()

class DeleteTimestampsDialog(QDialog):
    def __init__(self, worker_name, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Izbriši evidence delavca - {worker_name}")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # Worker name display
        worker_label = QLabel(f"Delavec: {worker_name}")
        worker_label.setStyleSheet("font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(worker_label)
        
        # Start date
        start_layout = QHBoxLayout()
        start_layout.addWidget(QLabel("Začetni datum:"))
        self.start_calendar = QCalendarWidget()
        start_layout.addWidget(self.start_calendar)
        layout.addLayout(start_layout)
        
        # End date
        end_layout = QHBoxLayout()
        end_layout.addWidget(QLabel("Končni datum:"))
        self.end_calendar = QCalendarWidget()
        end_layout.addWidget(self.end_calendar)
        layout.addLayout(end_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("Izbriši")
        cancel_button = QPushButton("Prekliči")
        
        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
    
    def get_dates(self):
        return self.start_calendar.selectedDate().toPyDate(), self.end_calendar.selectedDate().toPyDate()

class ResultsDialog(QDialog):
    def __init__(self, data, summary=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Rezultati")
        self.setMinimumSize(1000, 600)  # Increased minimum size
        
        layout = QVBoxLayout(self)
        
        # Add summary if provided
        if summary:
            if isinstance(summary, dict):
                # Convert dictionary to formatted string
                summary_text = "\n".join([f"{key}: {value}" for key, value in summary.items()])
            else:
                summary_text = str(summary)
                
            self.summary_label = QLabel(summary_text)
            self.summary_label.setStyleSheet("font-weight: bold; margin-bottom: 10px;")
            layout.addWidget(self.summary_label)
        
        # Results table
        self.table = QTableWidget()
        self.table.setColumnCount(len(data.columns))
        self.table.setHorizontalHeaderLabels(data.columns)
        self.table.setRowCount(len(data))
        
        # Set column widths based on content
        for i, column in enumerate(data.columns):
            # Set minimum width based on header text
            header_width = len(column) * 10  # Approximate width based on character count
            self.table.setColumnWidth(i, max(header_width, 100))  # Minimum width of 100 pixels
        
        # Fill the table with data
        for i, row in data.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                self.table.setItem(i, j, item)
        
        # Make the table stretch to fill the dialog
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        layout.addWidget(self.table)
        
        # Export button
        export_button = QPushButton("Izvozi v CSV")
        export_button.clicked.connect(self.export_to_csv)
        layout.addWidget(export_button)

    def export_to_csv(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Shrani CSV datoteko",
            "",
            "CSV Files (*.csv)"
        )
        
        if file_path:
            try:
                # Get data from table
                data = []
                headers = []
                
                # Get headers
                for col in range(self.table.columnCount()):
                    header = self.table.horizontalHeaderItem(col).text()
                    headers.append(header)
                
                # Get data
                for row in range(self.table.rowCount()):
                    row_data = []
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        row_data.append(item.text() if item else "")
                    data.append(row_data)
                
                # Create DataFrame
                df = pd.DataFrame(data, columns=headers)
                
                # Write to CSV with semicolon separator, including summary at the top
                with open(file_path, 'w', encoding='utf-8') as f:
                    # Write summary if present
                    if hasattr(self, 'summary_label') and self.summary_label.text().strip():
                        summary_lines = self.summary_label.text().strip().split('\n')
                        for line in summary_lines:
                            f.write(f"# {line}\n")
                        f.write('\n')  # Extra newline after summary
                    df.to_csv(f, sep=';', index=False, encoding='utf-8')
                
                QMessageBox.information(self, "Uspeh", "Podatki so bili uspešno izvoženi.")
            except Exception as e:
                QMessageBox.critical(self, "Napaka", f"Napaka pri izvozu podatkov: {str(e)}")
                print(f"Error in export_to_csv: {str(e)}")

class SearchEmployeeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Poišči delavca glede na ID kartice")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # Card ID input
        card_layout = QHBoxLayout()
        card_layout.addWidget(QLabel("ID kartice:"))
        self.card_input = QLineEdit()
        self.card_input.setMaxLength(14)
        self.card_input.setPlaceholderText("Vnesite 14-mestno šestnajstiško število")
        card_layout.addWidget(self.card_input)
        layout.addLayout(card_layout)
        
        # Search button
        search_button = QPushButton("Išči")
        search_button.clicked.connect(self.search_employee)
        layout.addWidget(search_button)
        
        # Result label
        self.result_label = QLabel()
        self.result_label.setWordWrap(True)
        layout.addWidget(self.result_label)

    def search_employee(self):
        card_id = self.card_input.text().lower()
        
        # Validate hexadecimal format
        if len(card_id) != 14 or not all(c in '0123456789abcdef' for c in card_id):
            self.result_label.setText("Neveljavna številka kartice! Vnesite 14-mestno šestnajstiško število.")
            return
        
        try:
            # Find employee in database
            self.parent().cursor.execute("""
                SELECT e.name, e.daily_hours, g.name as group_name
                FROM employees e
                LEFT JOIN groups g ON e.group_id = g.id
                WHERE LOWER(e.card_id) = ?
            """, (card_id,))
            
            employee = self.parent().cursor.fetchone()
            
            if employee:
                name, daily_hours, group_name = employee
                message = f"Ime delavca: {name}\n"
                message += f"Dnevni delovni čas: {daily_hours} ur\n"
                message += f"Skupina: {group_name if group_name else 'Ni dodeljena'}"
                self.result_label.setText(message)
            else:
                self.result_label.setText("Kartica ni bila najdena v bazi podatkov.")
                
        except Exception as e:
            self.result_label.setText(f"Napaka pri iskanju delavca: {str(e)}")

class CalendarDialog(QDialog):
    def __init__(self, parent=None, card_id=None):
        super().__init__(parent)
        self.setWindowTitle("Koledar delovnega časa")
        self.setMinimumSize(800, 600)
        self.card_id = card_id
        self.conn = parent.conn
        self.cursor = parent.cursor
        
        # Initialize month_data dictionary and month cache
        self.month_data = {}
        self.month_cache = {}  # Cache for storing month data
        self.special_day_formats = {}  # Store special day formatting
        self.current_month = None  # Track current month
        
        # Get worker's name
        self.cursor.execute("SELECT name FROM employees WHERE card_id = ?", (self.card_id,))
        self.worker_name = self.cursor.fetchone()[0]
        
        layout = QVBoxLayout(self)
        
        # Add worker's name label
        self.name_label = QLabel(f"Delavec: {self.worker_name}")
        self.name_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(self.name_label)
        
        # Create calendar widget
        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(True)
        self.calendar.clicked.connect(self.on_date_clicked)
        self.calendar.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.calendar.customContextMenuRequested.connect(self.show_context_menu)
        
        # Connect currentPageChanged signal to update display
        self.calendar.currentPageChanged.connect(self.on_month_changed)
        
        # Initialize selected dates set
        self.selected_dates = set()
        self.last_clicked_date = None
        
        layout.addWidget(self.calendar)
        
        # Add month navigation
        nav_layout = QHBoxLayout()
        prev_month = QPushButton("Prejšni mesec")
        next_month = QPushButton("Naslednji mesec")
        refresh_btn = QPushButton("Osveži")
        prev_month.clicked.connect(self.prev_month)
        next_month.clicked.connect(self.next_month)
        refresh_btn.clicked.connect(self.refresh_current_month)
        nav_layout.addWidget(prev_month)
        nav_layout.addWidget(next_month)
        nav_layout.addWidget(refresh_btn)
        layout.addLayout(nav_layout)
        
        # Add year selection
        year_layout = QHBoxLayout()
        year_layout.addWidget(QLabel("Leto:"))
        self.year_spin = QSpinBox()
        self.year_spin.setRange(2000, 2100)
        self.year_spin.setValue(QDate.currentDate().year())
        self.year_spin.valueChanged.connect(self.on_year_changed)
        year_layout.addWidget(self.year_spin)
        layout.addLayout(year_layout)
        
        # Add day details
        self.day_details = QLabel()
        self.day_details.setWordWrap(True)
        layout.addWidget(self.day_details)
        
        # Load data for current month
        self.load_month_data()
        
    def refresh_current_month(self):
        """Clear cache and refresh the current month data"""
        # Clear the cache for the current month
        if self.current_month:
            month_key = self.current_month
            if month_key in self.month_cache:
                del self.month_cache[month_key]
        
        # Reload the current month data
        self.load_month_data()
        
    def clear_all_cache(self):
        """Clear all cached month data"""
        self.month_cache.clear()
        self.month_data = {}
        self.special_day_formats = {}
        
    def on_date_clicked(self, date):
        """Handle date selection"""
        # Get keyboard modifiers
        modifiers = QApplication.keyboardModifiers()
        
        if modifiers & Qt.KeyboardModifier.ShiftModifier and self.last_clicked_date is not None:
            # Handle range selection
            start_date = min(self.last_clicked_date, date)
            end_date = max(self.last_clicked_date, date)
            
            # Clear current selection
            self.selected_dates.clear()
            
            # Add all dates in the range
            current_date = start_date
            while current_date <= end_date:
                self.selected_dates.add(current_date)
                current_date = current_date.addDays(1)
        else:
            # Single click - clear previous selection and select only this date
            self.selected_dates.clear()
            self.selected_dates.add(date)
        
        # Store the last clicked date
        self.last_clicked_date = date
        
        # Clear all formatting first
        self.calendar.setDateTextFormat(QDate(), QTextCharFormat())
        
        # Update calendar appearance
        self.update_calendar_appearance()
        
        # Show details for the clicked date
        self.show_date_details(date)
        
    def show_context_menu(self, pos):
        """Show context menu for the selected dates"""
        menu = QMenu(self)
        
        if not self.selected_dates:
            return
            
        # Check if any of the selected dates already have special days
        has_special_days = False
        for date in self.selected_dates:
            self.cursor.execute("""
                SELECT type FROM special_days 
                WHERE card_id = ? AND date = ?
            """, (self.card_id, date.toPyDate()))
            if self.cursor.fetchone():
                has_special_days = True
                break
        
        # Add worktime edit option for single day selection
        if len(self.selected_dates) == 1:
            worktime_action = menu.addAction("Spremeni vrednosti za izbran dan")
            worktime_action.triggered.connect(lambda: self.edit_worktime_for_day(list(self.selected_dates)[0]))
            menu.addSeparator()
        
        if has_special_days:
            remove_action = menu.addAction("Odstrani dodane vrednosti za izbrane dneve")
            remove_action.triggered.connect(lambda: self.remove_special_days(list(self.selected_dates)))
        else:
            sick_leave_action = menu.addAction("Dodaj bolniški stalež za izbrane dneve")
            vacation_action = menu.addAction("Dodaj dopust za izbrane dneve")
            sick_leave_action.triggered.connect(lambda: self.add_special_days(list(self.selected_dates), 'sick_leave'))
            vacation_action.triggered.connect(lambda: self.add_special_days(list(self.selected_dates), 'vacation'))
        
        menu.exec(self.calendar.mapToGlobal(pos))
        
    def on_month_changed(self, year, month):
        """Handle month change in calendar"""
        # Clear previous selection
        self.selected_dates.clear()
        self.last_clicked_date = None
        
        # Set the selected date to the first day of the new month
        new_date = QDate(year, month, 1)
        self.calendar.setSelectedDate(new_date)
        
        # Update the display
        self.load_month_data()
        
        # Show details for the first day of the month
        self.show_date_details(new_date)
        
        # Force calendar update
        self.calendar.updateCells()
        self.calendar.repaint()
        
        # Update year spinbox to match the new year
        self.year_spin.setValue(year)
        
    def prev_month(self):
        """Navigate to previous month"""
        current_date = self.calendar.selectedDate()
        new_date = current_date.addMonths(-1)
        self.calendar.setSelectedDate(new_date)
        self.on_month_changed(new_date.year(), new_date.month())
        
    def next_month(self):
        """Navigate to next month"""
        current_date = self.calendar.selectedDate()
        new_date = current_date.addMonths(1)
        self.calendar.setSelectedDate(new_date)
        self.on_month_changed(new_date.year(), new_date.month())
        
    def on_year_changed(self, year):
        """Handle year change"""
        current_date = self.calendar.selectedDate()
        new_date = QDate(year, current_date.month(), 1)
        self.calendar.setSelectedDate(new_date)
        self.on_month_changed(year, current_date.month())
        
    def show_date_details(self, date):
        """Show details for the selected date, including all arrivals and departures"""
        py_date = date.toPyDate()
        details = ""
        # Show summary if available
        if py_date in self.month_data:
            data = self.month_data[py_date]
            details += f"Delovne ure: {data['hours']}\n"
            details += f"Status: {data['status']}\n"
            details += f"Prihod na delo: {data['arrival']}\n"
            details += f"Izhod iz dela: {data['departure']}"
            # Add special day status if it exists
            if data['special']:
                if data['special'] == 'sick_leave':
                    details += "\n\nBolniški stalež"
                else:  # vacation
                    details += "\n\nDopust"
        else:
            # Check if this is a special day that's not in month_data
            self.cursor.execute("""
                SELECT type FROM special_days
                WHERE card_id = ? AND date = ?
            """, (self.card_id, py_date))
            special_day = self.cursor.fetchone()
            if special_day:
                if special_day[0] == 'sick_leave':
                    self.day_details.setText("Bolniški stalež")
                else:  # vacation
                    self.day_details.setText("Dopust")
                return
            else:
                self.day_details.setText("Ni podatkov za ta dan")
                return

        # Fetch and show all arrivals and departures for this day
        try:
            # Read only this day's data from SMB
            parent = self.parent()
            if parent is not None:
                # Use parent's read_smb_files to get raw data for this day
                df = parent.read_smb_files(py_date, py_date)
                if df is not None and not df.empty:
                    # Filter for this card_id and date
                    df = df[(df['CardID'] == self.card_id) & (df['Timestamp'].dt.date == py_date)]
                    df = df.sort_values('Timestamp')
                    if not df.empty:
                        details += "\n\nVsi prihodi in odhodi za ta dan:\n"
                        for _, row in df.iterrows():
                            time_str = row['Timestamp'].strftime('%H:%M:%S')
                            if row['Status'] == 'Prihod na delo':
                                details += f"Prihod na delo: {time_str}\n"
                            elif row['Status'] == 'Izhod iz dela':
                                details += f"Izhod iz dela: {time_str}\n"
        except Exception as e:
            details += f"\nNapaka pri pridobivanju vseh prihodov/odhodov: {str(e)}"
        self.day_details.setText(details)
        
    def add_special_days(self, dates, day_type):
        """Add special days to the database for multiple dates"""
        try:
            # Check if any dates are in the future
            today = datetime.now().date()
            future_dates = [date for date in dates if date.toPyDate() > today]
            
            if future_dates:
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("Opozorilo")
                msg_box.setText(f"Ali ste prepričani, da želite dodati {'bolniški stalež' if day_type == 'sick_leave' else 'dopust'} za prihodnje datume?")
                msg_box.setIcon(QMessageBox.Icon.Question)
                
                # Add custom buttons with Slovenian text
                da_button = msg_box.addButton("Da", QMessageBox.ButtonRole.YesRole)
                ne_button = msg_box.addButton("Ne", QMessageBox.ButtonRole.NoRole)
                
                # Set Ne as default button
                msg_box.setDefaultButton(ne_button)
                
                msg_box.exec()
                
                if msg_box.clickedButton() == ne_button:
                    return
            
            # Process all dates
            for date in dates:
                py_date = date.toPyDate()
                
                # First remove any existing special day for this date
                self.cursor.execute("""
                    DELETE FROM special_days
                    WHERE card_id = ? AND date = ?
                """, (self.card_id, py_date))
                
                # Then add the new special day
                self.cursor.execute("""
                    INSERT INTO special_days (card_id, date, type)
                    VALUES (?, ?, ?)
                """, (self.card_id, py_date, day_type))
                
                # Create and store the special day format
                format = QTextCharFormat()
                if day_type == 'sick_leave':
                    color = QColor(139, 69, 19)  # Brown
                    text = "BS"
                else:  # vacation
                    color = QColor(0, 0, 255)  # Blue
                    text = "D"
                format.setBackground(QBrush(color))
                format.setForeground(QBrush(Qt.GlobalColor.black))
                format.setToolTip(text)
                self.special_day_formats[py_date] = format
                
                # Apply the format immediately
                self.calendar.setDateTextFormat(date, format)
            
            self.conn.commit()
            
            # Save shared data to SMB
            self.parent().save_shared_data_with_retry()
            
            # Update the month cache and reload the current month
            current_date = self.calendar.selectedDate()
            month_key = (current_date.year(), current_date.month())
            if month_key in self.month_cache:
                del self.month_cache[month_key]  # Clear the cache for this month
            
            # Reload the month data to update the display
            self.load_month_data()
            
            # Force the calendar to update its appearance
            self.calendar.updateCells()
            self.calendar.repaint()
            
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri dodajanju posebnih dni: {str(e)}")
            
    def remove_special_days(self, dates):
        """Remove special days from the database for multiple dates"""
        try:
            # First, remove from database
            for date in dates:
                py_date = date.toPyDate()
                self.cursor.execute("""
                    DELETE FROM special_days
                    WHERE card_id = ? AND date = ?
                """, (self.card_id, py_date))
                
                # Remove from special_day_formats
                if py_date in self.special_day_formats:
                    del self.special_day_formats[py_date]
            
            self.conn.commit()
            
            # Save shared data to SMB
            self.parent().save_shared_data_with_retry()
            
            # Clear the month cache and reload data
            current_date = self.calendar.selectedDate()
            month_key = (current_date.year(), current_date.month())
            if month_key in self.month_cache:
                del self.month_cache[month_key]
            
            # Clear selection before reloading data
            self.selected_dates.clear()
            self.last_clicked_date = None
            
            # Clear all formatting
            self.calendar.setDateTextFormat(QDate(), QTextCharFormat())
            
            # Reload the month data
            self.load_month_data()
            
            # Force calendar update
            self.calendar.updateCells()
            self.calendar.repaint()
            
            # Show details for the current date
            self.show_date_details(current_date)
            
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri odstranjevanju posebnih dni: {str(e)}")
            
    def apply_month_formatting(self):
        """Apply formatting from cached month data"""
        for date, data in self.month_data.items():
            format = QTextCharFormat()
            
            if data['special']:
                if data['special'] == 'sick_leave':
                    color = QColor(139, 69, 19)  # Brown
                    text = "BS"
                else:  # vacation
                    color = QColor(0, 0, 255)  # Blue
                    text = "D"
                format.setBackground(QBrush(color))
                format.setForeground(QBrush(Qt.GlobalColor.black))
                format.setToolTip(text)
                # Store special day format
                self.special_day_formats[date] = format
            else:
                if data['status'] == 'Nepopolni podatki':
                    color = QColor(255, 0, 0)  # Red
                elif data['status'].startswith('Manjko ur'):
                    color = QColor(255, 165, 0)  # Orange
                else:
                    color = QColor(0, 255, 0)  # Green
                format.setBackground(QBrush(color))
                format.setForeground(QBrush(Qt.GlobalColor.black))
            
            qdate = QDate(date.year, date.month, date.day)
            self.calendar.setDateTextFormat(qdate, format)
        
        # Force the calendar to update its appearance
        self.calendar.updateCells()
        self.calendar.repaint()

    def load_month_data(self):
        """Load attendance data for the current month"""
        try:
            # Get the first and last day of the current month
            current_date = self.calendar.selectedDate()
            first_day = QDate(current_date.year(), current_date.month(), 1)
            last_day = QDate(current_date.year(), current_date.month(), current_date.daysInMonth())
            
            # Store current month
            self.current_month = (current_date.year(), current_date.month())
            
            # Check if we have cached data for this month
            month_key = (current_date.year(), current_date.month())
            if month_key in self.month_cache:
                # Use cached data
                self.month_data = self.month_cache[month_key]
                self.apply_month_formatting()
                return
            
            # Get employee's daily hours
            self.cursor.execute("SELECT daily_hours FROM employees WHERE card_id = ?", (self.card_id,))
            daily_hours = self.cursor.fetchone()[0]
            is_flexible = daily_hours == -1
            
            # Get attendance data for the month
            result = self.parent().calculate_working_hours(self.card_id, first_day.toPyDate(), last_day.toPyDate())
            
            # Get special days for the month
            self.cursor.execute("""
                SELECT date, type FROM special_days
                WHERE card_id = ? AND date BETWEEN ? AND ?
            """, (self.card_id, first_day.toPyDate(), last_day.toPyDate()))
            
            # Create a dictionary of special days
            special_days = {}
            for row in self.cursor.fetchall():
                date = row[0]
                if isinstance(date, str):
                    date = datetime.strptime(date, '%Y-%m-%d').date()
                special_days[date] = row[1]
            
            # Clear previous formatting
            self.calendar.setDateTextFormat(QDate(), QTextCharFormat())
            
            if result is not None:
                # Create a dictionary of dates and their status
                self.month_data = {}
                for _, row in result.iterrows():
                    date = row['Datum']
                    if isinstance(date, str):
                        date = datetime.strptime(date, '%Y-%m-%d').date()
                    hours = row['Delovne ure']
                    status = row['Status']
                    arrival = row['Prihod na delo']
                    departure = row['Izhod iz dela']
                    
                    # Create text format
                    format = QTextCharFormat()
                    
                    # Check if this is a special day
                    if date in special_days:
                        if special_days[date] == 'sick_leave':
                            color = QColor(139, 69, 19)  # Brown
                            text = "BS"
                        else:  # vacation
                            color = QColor(0, 0, 255)  # Blue
                            text = "D"
                        format.setBackground(QBrush(color))
                        format.setForeground(QBrush(Qt.GlobalColor.black))
                        format.setToolTip(text)
                        # Store special day format
                        self.special_day_formats[date] = format
                    else:
                        # Determine color based on status
                        if status == 'Nepopolni podatki':
                            color = QColor(255, 0, 0)  # Red
                        elif status.startswith('Manjko ur'):
                            color = QColor(255, 165, 0)  # Orange
                        else:
                            color = QColor(0, 255, 0)  # Green
                        format.setBackground(QBrush(color))
                        format.setForeground(QBrush(Qt.GlobalColor.black))
                    
                    # Set the format for this date
                    qdate = QDate(date.year, date.month, date.day)
                    self.calendar.setDateTextFormat(qdate, format)
                    
                    self.month_data[date] = {
                        'hours': hours,
                        'status': status,
                        'arrival': arrival,
                        'departure': departure,
                        'special': special_days.get(date)
                    }
                    
            # Also format any special days that might not be in the result
            for date, day_type in special_days.items():
                if date not in self.month_data:
                    format = QTextCharFormat()
                    if day_type == 'sick_leave':
                        color = QColor(139, 69, 19)  # Brown
                        text = "BS"
                    else:  # vacation
                        color = QColor(0, 0, 255)  # Blue
                        text = "D"
                    format.setBackground(QBrush(color))
                    format.setForeground(QBrush(Qt.GlobalColor.black))
                    format.setToolTip(text)
                    qdate = QDate(date.year, date.month, date.day)
                    self.calendar.setDateTextFormat(qdate, format)
                    # Store special day format
                    self.special_day_formats[date] = format
            
            # Cache the month data
            self.month_cache[month_key] = self.month_data.copy()
            
            # Force the calendar to update its appearance
            self.calendar.updateCells()
            self.calendar.repaint()
            
            # Show details for the first day of the month
            self.show_date_details(first_day)
            
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri nalaganju podatkov: {str(e)}")
            print(f"Error in load_month_data: {str(e)}")
            import traceback
            print(traceback.format_exc())
            
    def update_calendar_appearance(self):
        """Update the calendar appearance with both special days and selection"""
        # First, restore all special day formatting
        self.load_month_data()
        
        # Create selection format
        selection_format = QTextCharFormat()
        selection_format.setBackground(QBrush(QColor(200, 200, 200, 100)))  # Semi-transparent light gray
        selection_format.setForeground(QBrush(Qt.GlobalColor.black))
        
        # First apply all special day formatting
        for date, format in self.special_day_formats.items():
            qdate = QDate(date.year, date.month, date.day)
            self.calendar.setDateTextFormat(qdate, format)
        
        # Then apply selection formatting to selected dates
        for date in self.selected_dates:
            # Check if this is a special day
            py_date = date.toPyDate()
            if py_date in self.special_day_formats:
                # For special days, combine the formats
                special_format = self.special_day_formats[py_date]
                combined_format = QTextCharFormat(special_format)
                # Add selection highlight while preserving special day color
                combined_format.setBackground(selection_format.background())
                self.calendar.setDateTextFormat(date, combined_format)
            else:
                # For normal days, just apply selection format
                self.calendar.setDateTextFormat(date, selection_format)
        
        # Force the calendar to update its appearance
        self.calendar.updateCells()
        self.calendar.repaint()
    
    def edit_worktime_for_day(self, date):
        """Open worktime edit dialog for the selected day"""
        dialog = WorktimeEditDialog(self.parent(), self.card_id, date.toPyDate())
        result = dialog.exec()
        
        # If the dialog was accepted (data was modified), refresh the calendar
        if result == QDialog.DialogCode.Accepted:
            self.refresh_current_month()

class WorktimeEditDialog(QDialog):
    def __init__(self, parent=None, card_id=None, date=None):
        super().__init__(parent)
        self.main_window = parent
        self.card_id = card_id
        self.date = date
        self.worktime_entries = []
        self.data_modified = False  # Track if data was modified
        
        self.setWindowTitle(f"Spremeni vrednosti za {date.strftime('%d.%m.%Y')}")
        self.setMinimumWidth(600)
        self.setMinimumHeight(400)
        
        self.init_ui()
        self.load_worktime_data()
    
    def closeEvent(self, event):
        """Handle dialog close event"""
        if self.data_modified:
            self.accept()  # Set result to Accepted if data was modified
        else:
            self.reject()  # Set result to Rejected if no data was modified
        event.accept()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        title_label = QLabel(f"Delovni čas za {self.date.strftime('%d.%m.%Y')}")
        title_label.setStyleSheet("font-size: 14px; font-weight: bold; margin: 10px;")
        layout.addWidget(title_label)
        
        # Scroll area for worktime entries
        scroll_area = QScrollArea()
        scroll_widget = QWidget()
        self.entries_layout = QVBoxLayout(scroll_widget)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        layout.addWidget(scroll_area)
        
        # Buttons layout
        buttons_layout = QHBoxLayout()
        
        add_arrival_btn = QPushButton("Dodaj prihod")
        add_arrival_btn.clicked.connect(self.add_arrival)
        buttons_layout.addWidget(add_arrival_btn)
        
        add_departure_btn = QPushButton("Dodaj izhod")
        add_departure_btn.clicked.connect(self.add_departure)
        buttons_layout.addWidget(add_departure_btn)
        
        buttons_layout.addStretch()
        
        close_btn = QPushButton("Zapri")
        close_btn.clicked.connect(self.accept)
        buttons_layout.addWidget(close_btn)
        
        layout.addLayout(buttons_layout)
        
        # Add info label about saving
        info_label = QLabel("Vse spremembe so shranjene avtomatsko")
        info_label.setStyleSheet("color: #666; font-size: 10px;")
        layout.addWidget(info_label)
        self.setLayout(layout)
    
    def load_worktime_data(self):
        """Load worktime data for the selected day from SMB share"""
        try:
            # Get worktime data from main window's SMB files
            df = self.main_window.read_smb_files(self.date, self.date)
            
            if df is not None and not df.empty:
                # Filter data for this card_id
                card_data = df[df['CardID'] == str(self.card_id)]
                
                if not card_data.empty:
                    for _, row in card_data.iterrows():
                        timestamp = row['Timestamp']
                        status = row['Status']
                        
                        # Create entry widget
                        self.add_entry_widget(timestamp, status)
                else:
                    # No data for this card_id on this date
                    no_data_label = QLabel("Ni podatkov za ta dan. Lahko dodajate nove vrednosti z gumbi spodaj.")
                    no_data_label.setStyleSheet("color: #666; font-style: italic; padding: 10px;")
                    self.entries_layout.addWidget(no_data_label)
            else:
                # No data at all for this date
                no_data_label = QLabel("Ni podatkov za ta dan. Lahko dodajate nove vrednosti z gumbi spodaj.")
                no_data_label.setStyleSheet("color: #666; font-style: italic; padding: 10px;")
                self.entries_layout.addWidget(no_data_label)
        except Exception as e:
            # Show error but still allow adding new data
            error_label = QLabel(f"Napaka pri nalaganju podatkov: {str(e)}\nLahko dodajate nove vrednosti z gumbi spodaj.")
            error_label.setStyleSheet("color: #ff6b6b; font-style: italic; padding: 10px;")
            self.entries_layout.addWidget(error_label)
    
    def add_entry_widget(self, timestamp, status):
        """Add a worktime entry widget to the layout"""
        entry_widget = QWidget()
        entry_layout = QHBoxLayout(entry_widget)
        
        # Format timestamp for display
        try:
            if isinstance(timestamp, str):
                # If it's already a full datetime string, extract time part
                if ' ' in timestamp:
                    # Full datetime format: "YYYY-MM-DD HH:MM:SS"
                    dt = pd.to_datetime(timestamp)
                    time_str = dt.strftime('%H:%M:%S')
                else:
                    # Just time format: "HH:MM:SS"
                    time_str = timestamp
            else:
                time_str = str(timestamp)
        except:
            time_str = str(timestamp)
        
        # Entry label
        entry_label = QLabel(f"{time_str}, {status}")
        entry_label.setStyleSheet("padding: 5px; border: 1px solid #ccc; background-color: #f9f9f9;")
        entry_layout.addWidget(entry_label)
        
        # Delete button
        delete_btn = QPushButton("Izbriši vrednost")
        delete_btn.setStyleSheet("background-color: #ff6b6b; color: white; padding: 5px 10px;")
        delete_btn.clicked.connect(lambda: self.delete_entry(entry_widget, timestamp, status))
        entry_layout.addWidget(delete_btn)
        
        self.entries_layout.addWidget(entry_widget)
        self.worktime_entries.append((timestamp, status, entry_widget))
    
    def delete_entry(self, widget, timestamp, status):
        """Delete a worktime entry"""
        try:
            # Check if the widget exists in the entries list
            entry_found = False
            for t, s, w in self.worktime_entries:
                if w == widget and str(t) == str(timestamp) and s == status:
                    entry_found = True
                    break
            
            if not entry_found:
                QMessageBox.warning(self, "Opozorilo", "Vrednost ni bila najdena v seznamu.")
                return
            
            # Try to remove from CSV file first
            try:
                self.delete_from_csv(timestamp, status)
            except Exception as e:
                print(f"Warning: Could not delete from CSV file: {str(e)}")
                # Continue with UI deletion even if CSV deletion fails
                # This handles cases where the file doesn't exist or can't be accessed
            
            # Mark data as modified
            self.data_modified = True
            
            # Remove from UI
            self.entries_layout.removeWidget(widget)
            widget.deleteLater()
            
            # Remove from entries list
            self.worktime_entries = [(t, s, w) for t, s, w in self.worktime_entries if w != widget]
            
            QMessageBox.information(self, "Uspeh", "Vrednost je bila uspešno izbrisana.")
        except Exception as e:
            error_message = str(e)
            if "NAPAKA DOVOLJENJ" in error_message:
                # Permission error - show user-friendly message
                QMessageBox.critical(self, "Napaka dovoljenj", 
                    "Aplikacija nima dovoljenja za pisanje v SMB mapo.\n\n"
                    "To pomeni, da lahko aplikacija prebere podatke, vendar ne more\n"
                    "posodobiti ali izbrisati obstoječih zapisov.\n\n"
                    "Rešitve:\n"
                    "1. Kontaktirajte sistemskega administratorja za nastavitev\n"
                    "   ustreznih dovoljenj za pisanje v SMB mapo\n"
                    "2. Preverite, ali imate pravice za spreminjanje datotek\n"
                    "   v omrežni mapi\n\n"
                    f"Tehnični opis napake:\n{error_message}")
            else:
                QMessageBox.critical(self, "Napaka", f"Napaka pri brisanju vrednosti: {error_message}")
    
    def delete_from_csv(self, timestamp, status):
        """Delete entry from CSV file in SMB share"""
        try:
            debug_info = []
            debug_info.append("=== DELETE FROM CSV DEBUG ===")
            debug_info.append(f"Original timestamp: {timestamp} (type: {type(timestamp)})")
            debug_info.append(f"Status: {status}")
            debug_info.append(f"Card ID: {self.card_id}")
            debug_info.append(f"Date: {self.date}")
            
            print(f"=== DELETE FROM CSV DEBUG ===")
            print(f"Original timestamp: {timestamp} (type: {type(timestamp)})")
            print(f"Status: {status}")
            print(f"Card ID: {self.card_id}")
            print(f"Date: {self.date}")
            
            # Try to read current CSV file
            df = None
            try:
                df = self.main_window.read_smb_files(self.date, self.date)
                debug_info.append(f"Loaded DataFrame shape: {df.shape if df is not None else 'None'}")
                print(f"Loaded DataFrame shape: {df.shape if df is not None else 'None'}")
                if df is not None and not df.empty:
                    debug_info.append(f"DataFrame columns: {df.columns.tolist()}")
                    debug_info.append(f"DataFrame content:\n{df}")
                    print(f"DataFrame columns: {df.columns.tolist()}")
                    print(f"DataFrame content:\n{df}")
            except Exception as e:
                error_msg = f"Could not read existing data for deletion: {str(e)}"
                debug_info.append(error_msg)
                print(error_msg)
                # If we can't read the file, there's nothing to delete
                # This is not an error - just means the file doesn't exist or is empty
                return
            
            if df is not None and not df.empty:
                # Create full timestamp for comparison (combine date with time)
                if ' ' not in str(timestamp):
                    # If timestamp is just time, combine with date
                    full_timestamp = f"{self.date.strftime('%Y-%m-%d')} {timestamp}"
                else:
                    # If timestamp is already full datetime, use as is
                    full_timestamp = str(timestamp)
                
                debug_info.append(f"Full timestamp for comparison: {full_timestamp}")
                print(f"Full timestamp for comparison: {full_timestamp}")
                
                # Check what timestamps are actually in the DataFrame
                debug_info.append("Timestamps in DataFrame:")
                print(f"Timestamps in DataFrame:")
                for i, ts in enumerate(df['Timestamp']):
                    debug_info.append(f"  {i}: {ts} (type: {type(ts)})")
                    print(f"  {i}: {ts} (type: {type(ts)})")
                
                # Check if the entry exists before trying to delete it
                card_matches = df['CardID'] == str(self.card_id)
                timestamp_matches = df['Timestamp'].astype(str) == full_timestamp
                status_matches = df['Status'] == status
                
                debug_info.append(f"Card matches: {card_matches.sum()}")
                debug_info.append(f"Timestamp matches: {timestamp_matches.sum()}")
                debug_info.append(f"Status matches: {status_matches.sum()}")
                print(f"Card matches: {card_matches.sum()}")
                print(f"Timestamp matches: {timestamp_matches.sum()}")
                print(f"Status matches: {status_matches.sum()}")
                
                entry_exists = (card_matches & timestamp_matches & status_matches).any()
                debug_info.append(f"Entry exists: {entry_exists}")
                print(f"Entry exists: {entry_exists}")
                
                if entry_exists:
                    debug_info.append("Entry found, proceeding with deletion...")
                    print("Entry found, proceeding with deletion...")
                    # Filter out the entry to delete
                    df_filtered = df[~(card_matches & timestamp_matches & status_matches)]
                    debug_info.append(f"Filtered DataFrame shape: {df_filtered.shape}")
                    print(f"Filtered DataFrame shape: {df_filtered.shape}")
                    
                    # Update CSV file only if there are still entries left
                    if not df_filtered.empty:
                        debug_info.append("Updating CSV file with remaining entries...")
                        print("Updating CSV file with remaining entries...")
                        try:
                            self.update_csv_file(df_filtered)
                            debug_info.append("CSV file updated successfully!")
                        except Exception as e:
                            debug_info.append(f"ERROR updating CSV file: {str(e)}")
                            raise
                    else:
                        debug_info.append("No entries left, deleting CSV file...")
                        print("No entries left, deleting CSV file...")
                        # If no entries left, delete the file
                        try:
                            self.delete_csv_file()
                            debug_info.append("CSV file deleted successfully!")
                        except Exception as e:
                            debug_info.append(f"ERROR deleting CSV file: {str(e)}")
                            raise
                    debug_info.append("Deletion completed successfully!")
                    print("Deletion completed successfully!")
                else:
                    debug_info.append(f"Entry not found for deletion: {full_timestamp}, {status}")
                    print(f"Entry not found for deletion: {full_timestamp}, {status}")
                    # Try alternative timestamp formats
                    debug_info.append("Trying alternative timestamp formats...")
                    print("Trying alternative timestamp formats...")
                    for alt_format in [timestamp, str(timestamp), full_timestamp]:
                        alt_matches = df['Timestamp'].astype(str) == alt_format
                        if (card_matches & alt_matches & status_matches).any():
                            debug_info.append(f"Found match with format: {alt_format}")
                            print(f"Found match with format: {alt_format}")
                            # Use this format for deletion
                            df_filtered = df[~(card_matches & alt_matches & status_matches)]
                            if not df_filtered.empty:
                                self.update_csv_file(df_filtered)
                            else:
                                self.delete_csv_file()
                            debug_info.append("Deletion completed with alternative format!")
                            print("Deletion completed with alternative format!")
                            return
                    debug_info.append("No matches found with any timestamp format")
                    print("No matches found with any timestamp format")
            else:
                # No data to delete
                debug_info.append("No data found to delete")
                print("No data found to delete")
            
            # Debug info is available in console output if needed
            
        except Exception as e:
            error_msg = f"Error in delete_from_csv: {str(e)}"
            print(error_msg)
            pass  # Error details are in console output
            raise Exception(f"Napaka pri posodabljanju CSV datoteke: {str(e)}")
    
    def delete_csv_file(self):
        """Delete the CSV file from SMB share"""
        try:
            import configparser
            from smb.SMBConnection import SMBConnection
            
            # Load configuration
            config = configparser.ConfigParser()
            config.read('config.ini')
            
            if 'SMB' not in config:
                raise ValueError("Nastavitve SMB niso bile najdene.")
            
            smb_path = config['SMB'].get('path', '')
            if not smb_path.startswith('\\\\'):
                raise ValueError("Neveljavna SMB pot")
            
            # Parse SMB path more safely
            path_parts = smb_path.split('\\')
            if len(path_parts) < 4:
                raise ValueError("Neveljavna SMB pot - manjkajo deli poti")
            
            server_name = path_parts[2]
            share_name = path_parts[3]
            username = config['SMB'].get('username', '')
            password = config['SMB'].get('password', '')
            
            # Connect to SMB share
            conn = SMBConnection(username, password, "CLIENT", server_name, use_ntlm_v2=True)
            connected = conn.connect(server_name, 445)
            
            if not connected:
                raise ConnectionError(f"Ni mogoče vzpostaviti povezave s strežnikom {server_name}")
            
            # Delete the file
            filename = f"time_records_{self.date.strftime('%Y%m%d')}.csv"
            try:
                conn.deleteFiles(share_name, filename)
                print(f"File {filename} deleted successfully")
            except Exception as e:
                print(f"Could not delete file {filename}: {str(e)}")
                # This is not necessarily an error - file might not exist
            finally:
                conn.close()
                
        except Exception as e:
            print(f"Error in delete_csv_file: {str(e)}")
            # Don't raise exception here as this is not critical
    
    def update_csv_file(self, df):
        """Update the CSV file in SMB share"""
        try:
            import configparser
            import io
            import tempfile
            import os
            from smb.SMBConnection import SMBConnection
            
            print(f"Updating CSV file for date: {self.date}")
            print(f"DataFrame shape: {df.shape}")
            print(f"DataFrame content:\n{df}")
            
            # Load configuration
            config = configparser.ConfigParser()
            config.read('config.ini')
            
            if 'SMB' not in config:
                raise ValueError("Nastavitve SMB niso bile najdene.")
            
            smb_path = config['SMB'].get('path', '')
            if not smb_path.startswith('\\\\'):
                raise ValueError("Neveljavna SMB pot")
            
            # Parse SMB path more safely
            path_parts = smb_path.split('\\')
            if len(path_parts) < 4:
                raise ValueError("Neveljavna SMB pot - manjkajo deli poti")
            
            server_name = path_parts[2]
            share_name = path_parts[3]
            username = config['SMB'].get('username', '')
            password = config['SMB'].get('password', '')
            
            print(f"Connecting to SMB: {server_name}/{share_name}")
            print(f"Username: {username}")
            print(f"Path: {smb_path}")
            
            # Connect to SMB share
            conn = SMBConnection(username, password, "CLIENT", server_name, use_ntlm_v2=True)
            connected = conn.connect(server_name, 445)
            
            if not connected:
                raise ConnectionError(f"Ni mogoče vzpostaviti povezave s strežnikom {server_name}")
            
            # Create CSV content
            csv_content = df.to_csv(index=False, header=False)
            print(f"CSV content:\n{csv_content}")
            
            # Save to SMB share
            filename = f"time_records_{self.date.strftime('%Y%m%d')}.csv"
            print(f"Saving to file: {filename}")
            
            # Try multiple approaches to handle permission issues
            success = False
            error_messages = []
            
            # Method 1: Direct file storage
            try:
                file_obj = io.BytesIO(csv_content.encode('utf-8'))
                conn.storeFile(share_name, filename, file_obj)
                print("File saved successfully using direct method!")
                success = True
            except Exception as e:
                error_msg = f"Direct method failed: {str(e)}"
                print(error_msg)
                error_messages.append(error_msg)
            
            # Method 2: Try deleting existing file first, then create new one
            if not success:
                try:
                    print("Trying to delete existing file first...")
                    try:
                        conn.deleteFiles(share_name, filename)
                        print("Existing file deleted successfully")
                    except:
                        print("No existing file to delete or deletion failed")
                    
                    # Now try to create new file
                    file_obj = io.BytesIO(csv_content.encode('utf-8'))
                    conn.storeFile(share_name, filename, file_obj)
                    print("File saved successfully using delete-first method!")
                    success = True
                except Exception as e:
                    error_msg = f"Delete-first method failed: {str(e)}"
                    print(error_msg)
                    error_messages.append(error_msg)
            
            # Method 3: Try with temporary filename then rename
            if not success:
                try:
                    temp_filename = f"temp_{filename}"
                    print(f"Trying with temporary filename: {temp_filename}")
                    
                    file_obj = io.BytesIO(csv_content.encode('utf-8'))
                    conn.storeFile(share_name, temp_filename, file_obj)
                    print("Temporary file created successfully")
                    
                    # Try to delete original and rename temp
                    try:
                        conn.deleteFiles(share_name, filename)
                    except:
                        pass  # Ignore if original doesn't exist
                    
                    # Note: SMB doesn't have a direct rename, so we'll keep the temp file
                    # and try to delete it, then create the final file
                    try:
                        temp_file_obj = io.BytesIO()
                        conn.retrieveFile(share_name, temp_filename, temp_file_obj)
                        temp_file_obj.seek(0)
                        
                        final_file_obj = io.BytesIO(temp_file_obj.read())
                        conn.storeFile(share_name, filename, final_file_obj)
                        conn.deleteFiles(share_name, temp_filename)
                        print("File saved successfully using temporary file method!")
                        success = True
                    except Exception as rename_e:
                        # Keep the temp file as fallback
                        print(f"Rename failed, keeping temporary file: {rename_e}")
                        success = True  # At least we have the data in temp file
                        
                except Exception as e:
                    error_msg = f"Temporary file method failed: {str(e)}"
                    print(error_msg)
                    error_messages.append(error_msg)
            
            conn.close()
            
            if not success:
                # All methods failed
                all_errors = "; ".join(error_messages)
                raise Exception(f"Vse metode shranjevanja so neuspešne: {all_errors}")
            else:
                print("CSV file updated successfully!")
                
        except Exception as e:
            print(f"Error in update_csv_file: {str(e)}")
            # Check if this is a permission error
            if "0xC0000022" in str(e) or "ACCESS_DENIED" in str(e).upper() or "Unable to open file" in str(e):
                # Try to save locally as backup
                try:
                    import os
                    backup_dir = "backup_csv"
                    if not os.path.exists(backup_dir):
                        os.makedirs(backup_dir)
                    
                    local_filename = os.path.join(backup_dir, f"time_records_{self.date.strftime('%Y%m%d')}.csv")
                    csv_content = df.to_csv(index=False, header=False)
                    
                    with open(local_filename, 'w', encoding='utf-8') as f:
                        f.write(csv_content)
                    
                    print(f"Data saved locally as backup: {local_filename}")
                    
                    raise Exception(f"NAPAKA DOVOLJENJ: Aplikacija nima dovoljenja za pisanje v SMB mapo. Podatki so bili shranjeni lokalno kot varnostna kopija v {local_filename}. Prosimo kontaktirajte sistemskega administratorja za nastavitev ustreznih dovoljenj. Tehnični opis: {str(e)}")
                except Exception as backup_e:
                    print(f"Backup save also failed: {backup_e}")
                    raise Exception(f"NAPAKA DOVOLJENJ: Aplikacija nima dovoljenja za pisanje v SMB mapo. Varnostna kopija ni mogla biti shranjena. Prosimo kontaktirajte sistemskega administratorja za nastavitev ustreznih dovoljenj. Tehnični opis: {str(e)}")
            else:
                raise Exception(f"Napaka pri posodabljanju CSV datoteke: {str(e)}")
    
    def add_arrival(self):
        """Open dialog to add arrival time"""
        dialog = TimeInputDialog(self, "Napišite čas prihoda", "Prihod na delo")
        if dialog.exec() == QDialog.DialogCode.Accepted:
            time_str = dialog.get_time()
            self.add_entry_to_csv(time_str, "Prihod na delo")
            # Create full timestamp for display
            full_timestamp = f"{self.date.strftime('%Y-%m-%d')} {time_str}"
            self.add_entry_widget(full_timestamp, "Prihod na delo")
    
    def add_departure(self):
        """Open dialog to add departure time"""
        dialog = TimeInputDialog(self, "Napišite čas izhoda", "Izhod iz dela")
        if dialog.exec() == QDialog.DialogCode.Accepted:
            time_str = dialog.get_time()
            self.add_entry_to_csv(time_str, "Izhod iz dela")
            # Create full timestamp for display
            full_timestamp = f"{self.date.strftime('%Y-%m-%d')} {time_str}"
            self.add_entry_widget(full_timestamp, "Izhod iz dela")
    
    def add_entry_to_csv(self, time_str, status):
        """Add new entry to CSV file"""
        try:
            # Combine date with time to create full datetime string
            full_timestamp = f"{self.date.strftime('%Y-%m-%d')} {time_str}"
            
            # Create new entry
            new_entry = pd.DataFrame({
                'CardID': [str(self.card_id)],
                'Timestamp': [full_timestamp],
                'Status': [status]
            })
            
            # Try to read existing data
            existing_df = None
            try:
                existing_df = self.main_window.read_smb_files(self.date, self.date)
            except Exception as e:
                print(f"Could not read existing data: {str(e)}")
                # Continue with creating new file
            
            if existing_df is not None and not existing_df.empty:
                # Append new entry to existing data
                updated_df = pd.concat([existing_df, new_entry], ignore_index=True)
            else:
                # Create new file with just this entry
                updated_df = new_entry
            
            # Update CSV file
            self.update_csv_file(updated_df)
            
            # Mark data as modified
            self.data_modified = True
            
            QMessageBox.information(self, "Uspeh", "Vrednost je bila uspešno dodana.")
            
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri dodajanju vrednosti: {str(e)}")

class TimeInputDialog(QDialog):
    def __init__(self, parent=None, title="Vnesite čas", status=""):
        super().__init__(parent)
        self.status = status
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(300)
        
        layout = QVBoxLayout()
        
        # Instructions
        instruction_label = QLabel("Vnesite čas v formatu HH:MM:SS")
        layout.addWidget(instruction_label)
        
        # Time input layout
        time_layout = QHBoxLayout()
        
        self.hour_input = QLineEdit()
        self.hour_input.setPlaceholderText("HH")
        self.hour_input.setMaxLength(2)
        self.hour_input.textChanged.connect(self.validate_input)
        time_layout.addWidget(self.hour_input)
        
        time_layout.addWidget(QLabel(":"))
        
        self.minute_input = QLineEdit()
        self.minute_input.setPlaceholderText("MM")
        self.minute_input.setMaxLength(2)
        self.minute_input.textChanged.connect(self.validate_input)
        time_layout.addWidget(self.minute_input)
        
        time_layout.addWidget(QLabel(":"))
        
        self.second_input = QLineEdit()
        self.second_input.setPlaceholderText("SS")
        self.second_input.setMaxLength(2)
        self.second_input.textChanged.connect(self.validate_input)
        time_layout.addWidget(self.second_input)
        
        layout.addLayout(time_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        save_btn = QPushButton("Shrani")
        save_btn.clicked.connect(self.save_time)
        button_layout.addWidget(save_btn)
        
        cancel_btn = QPushButton("Prekliči")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
    
    def validate_input(self):
        """Validate time input fields"""
        # Auto-advance to next field when current is complete
        sender = self.sender()
        if len(sender.text()) == 2:
            if sender == self.hour_input:
                self.minute_input.setFocus()
            elif sender == self.minute_input:
                self.second_input.setFocus()
    
    def save_time(self):
        """Save the entered time"""
        try:
            hour = self.hour_input.text().zfill(2)
            minute = self.minute_input.text().zfill(2)
            second = self.second_input.text().zfill(2)
            
            # Validate time
            if not (0 <= int(hour) <= 23):
                raise ValueError("Neveljavna ura")
            if not (0 <= int(minute) <= 59):
                raise ValueError("Neveljavne minute")
            if not (0 <= int(second) <= 59):
                raise ValueError("Neveljavne sekunde")
            
            time_str = f"{hour}:{minute}:{second}"
            self.time_result = time_str
            self.accept()
            
        except ValueError as e:
            QMessageBox.warning(self, "Napaka", str(e))
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri vnosu časa: {str(e)}")
    
    def get_time(self):
        """Get the entered time"""
        return getattr(self, 'time_result', None)

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Nastavitve")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # SMB Path
        smb_path_layout = QHBoxLayout()
        smb_path_layout.addWidget(QLabel("SMB pot:"))
        self.smb_path_input = QLineEdit()
        smb_path_layout.addWidget(self.smb_path_input)
        layout.addLayout(smb_path_layout)
        
        # Authentication checkbox
        self.auth_checkbox = QCheckBox("Zahtevaj uporabniško ime in geslo")
        self.auth_checkbox.setChecked(True)
        self.auth_checkbox.stateChanged.connect(self.toggle_auth_fields)
        layout.addWidget(self.auth_checkbox)
        
        # Username
        username_layout = QHBoxLayout()
        username_layout.addWidget(QLabel("Uporabniško ime:"))
        self.username_input = QLineEdit()
        username_layout.addWidget(self.username_input)
        layout.addLayout(username_layout)
        
        # Password
        password_layout = QHBoxLayout()
        password_layout.addWidget(QLabel("Geslo:"))
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        password_layout.addWidget(self.password_input)
        layout.addLayout(password_layout)
        
        # Test connection button
        test_button = QPushButton("Preveri povezavo")
        test_button.clicked.connect(self.test_connection)
        layout.addWidget(test_button)
        
        # Save settings button
        self.save_settings_button = QPushButton("Shrani nastavitve")
        self.save_settings_button.clicked.connect(self.save_settings)
        layout.addWidget(self.save_settings_button)
        
        # Status label
        self.connection_status = QLabel("")
        layout.addWidget(self.connection_status)
        
        # Load existing configuration
        self.load_configuration()
        
    def toggle_auth_fields(self, state):
        """Toggle visibility of authentication fields"""
        self.username_input.setEnabled(state == Qt.CheckState.Checked.value)
        self.password_input.setEnabled(state == Qt.CheckState.Checked.value)
        
    def test_connection(self):
        """Test connection to SMB share"""
        try:
            smb_path = self.smb_path_input.text()
            if not smb_path.startswith('\\\\'):
                raise ValueError("Neveljavna SMB pot")
            
            server_name = smb_path.split('\\')[2]
            share_name = smb_path.split('\\')[3]
            
            if self.auth_checkbox.isChecked():
                username = self.username_input.text()
                password = self.password_input.text()
                if not username or not password:
                    raise ValueError("Vnesite uporabniško ime in geslo")
            else:
                username = "guest"
                password = ""
            
            # Try to connect using pysmb
            conn = SMBConnection(username, password, "CLIENT", server_name, use_ntlm_v2=True)
            connected = conn.connect(server_name, 445)
            
            if not connected:
                raise ConnectionError("Ni mogoče vzpostaviti povezave s strežnikom")
            
            # Try to list files
            files = conn.listPath(share_name, '/')
            
            self.connection_status.setText("Povezava uspešna!")
            self.connection_status.setStyleSheet("color: green")
            
        except Exception as e:
            self.connection_status.setText(f"Napaka pri povezavi: {str(e)}")
            self.connection_status.setStyleSheet("color: red")
        finally:
            if 'conn' in locals():
                conn.close()
                
    def save_settings(self):
        """Save SMB settings to configuration file"""
        try:
            config = configparser.ConfigParser()
            config['SMB'] = {
                'path': self.smb_path_input.text(),
                'username': self.username_input.text(),
                'password': self.password_input.text()
            }
            
            with open('config.ini', 'w') as configfile:
                config.write(configfile)
            
            QMessageBox.information(self, "Uspeh", "Nastavitve so bile uspešno shranjene!")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri shranjevanju nastavitev: {str(e)}")
            print(f"Error in save_settings: {str(e)}")
            
    def load_configuration(self):
        """Load configuration from file"""
        try:
            config = configparser.ConfigParser()
            if os.path.exists('config.ini'):
                config.read('config.ini')
                if 'SMB' in config:
                    self.smb_path_input.setText(config['SMB'].get('path', ''))
                    self.username_input.setText(config['SMB'].get('username', ''))
                    self.password_input.setText(config['SMB'].get('password', ''))
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri branju nastavitev: {str(e)}")
            print(f"Error in load_configuration: {str(e)}")

class ManualDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("O programu - Beleženje delovnega časa Admin v1.1.1")
        self.setMinimumSize(1200, 800)
        self.setModal(True)
        self.chapters = {}
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Create scroll area for content
        self.scroll_area = QScrollArea()
        scroll_widget = QWidget()
        content_layout = QVBoxLayout(scroll_widget)
        
        # Title
        title_label = QLabel("O programu")
        title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin: 10px; color: #2c3e50;")
        content_layout.addWidget(title_label)
        
        # Version info
        version_label = QLabel("Verzija: 1.1.1")
        version_label.setStyleSheet("font-size: 12px; color: #7f8c8d; margin-bottom: 20px;")
        content_layout.addWidget(version_label)
        
        # Table of contents with clickable links
        toc_label = QLabel("Kazalo:")
        toc_label.setStyleSheet("font-size: 14px; font-weight: bold; margin: 10px 0; color: #34495e;")
        content_layout.addWidget(toc_label)
        
        # Create clickable table of contents
        toc_widget = QWidget()
        toc_layout = QVBoxLayout(toc_widget)
        toc_layout.setContentsMargins(20, 0, 0, 0)
        
        chapters = [
            ("1. Ikone gumbov v programu", "toolbar_icons"),
            ("2. O programu", "about")
        ]
        
        for chapter_name, chapter_id in chapters:
            chapter_btn = QPushButton(chapter_name)
            chapter_btn.setStyleSheet("""
                QPushButton {
                    text-align: left;
                    padding: 5px 10px;
                    border: none;
                    background-color: transparent;
                    color: #3498db;
                    text-decoration: underline;
                }
                QPushButton:hover {
                    background-color: #ecf0f1;
                    color: #2980b9;
                }
            """)
            chapter_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            chapter_btn.clicked.connect(lambda checked, cid=chapter_id: self.scroll_to_chapter(cid))
            toc_layout.addWidget(chapter_btn)
            self.chapters[chapter_id] = chapter_btn
        
        content_layout.addWidget(toc_widget)
        
        # Store content layout reference for navigation
        self.content_layout = content_layout
        
        # Content sections with detailed explanations
        self.add_chapter_with_icons(content_layout, "toolbar_icons", "1. Ikone gumbov v programu", """
V programu so naslednji gumbi z ikonami. Vsak gumb ima svojo specifično funkcijo:

**Lokacije ikon:**
• Orodna vrstica - glavni gumbi v zgornji vrstici
• Glavno okno - gumbi za upravljanje delavcev  ob posameznem delavcu


        """)

        
        self.add_chapter_simple(content_layout, "about", "2. O programu", """
Beleženje delovnega časa Admin, program za ogled in spreminjanje podatkov o delovnem času.

Copyright © 2025  Klemen Ambrožič

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.

E-mail: klemen.ambrozic@oscg.si
        """)
        
        
        # Footer
        footer_label = QLabel("© 2025 Beleženje delovnega časa Admin v1.1.1")
        footer_label.setStyleSheet("font-size: 10px; color: #95a5a6; margin-top: 20px; text-align: center;")
        content_layout.addWidget(footer_label)
        
        self.scroll_area.setWidget(scroll_widget)
        self.scroll_area.setWidgetResizable(True)
        layout.addWidget(self.scroll_area)
        
        # Close button
        close_btn = QPushButton("Zapri")
        close_btn.clicked.connect(self.accept)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)
    
    def add_chapter(self, layout, chapter_id, title, content, icon_name):
        """Add a chapter to the manual with icon and navigation support"""
        # Section container
        section_widget = QWidget()
        section_layout = QHBoxLayout(section_widget)
        section_layout.setContentsMargins(0, 0, 0, 0)
        
        # Icon
        icon_label = QLabel()
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), icon_name)
        if os.path.exists(icon_path):
            try:
                icon = QIcon(icon_path)
                if not icon.isNull():
                    pixmap = icon.pixmap(32, 32)
                    icon_label.setPixmap(pixmap)
                else:
                    icon_label.setText("📋")
                    icon_label.setStyleSheet("font-size: 24px;")
            except Exception as e:
                icon_label.setText("📋")
                icon_label.setStyleSheet("font-size: 24px;")
        else:
            icon_label.setText("📋")
            icon_label.setStyleSheet("font-size: 24px;")
        
        icon_label.setFixedSize(40, 40)
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        section_layout.addWidget(icon_label)
        
        # Content
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(10, 0, 0, 0)
        
        # Section title
        title_label = QLabel(title)
        title_label.setStyleSheet("font-size: 14px; font-weight: bold; margin: 20px 0 10px 0; color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 5px;")
        content_layout.addWidget(title_label)
        
        # Section content
        content_label = QLabel(content.strip())
        content_label.setStyleSheet("font-size: 12px; margin: 0 0 15px 0; line-height: 1.6; color: #34495e;")
        content_label.setWordWrap(True)
        content_layout.addWidget(content_label)
        
        section_layout.addWidget(content_widget)
        layout.addWidget(section_widget)
        
        # Store reference for navigation
        self.chapters[chapter_id] = section_widget
    
    def scroll_to_chapter(self, chapter_id):
        """Scroll to a specific chapter"""
        if chapter_id in self.chapters:
            # Scroll to the chapter widget
            self.scroll_area.ensureWidgetVisible(self.chapters[chapter_id])
    
    def add_section_with_icon(self, layout, title, content, icon_name):
        """Add a section to the manual with icon (legacy method)"""
        self.add_chapter(layout, f"section_{len(self.chapters)}", title, content, icon_name)
    
    def add_chapter_simple(self, layout, chapter_id, title, content):
        """Add a chapter to the manual without icon"""
        # Section container
        section_widget = QWidget()
        section_layout = QVBoxLayout(section_widget)
        section_layout.setContentsMargins(0, 0, 0, 0)
        
        # Section title
        title_label = QLabel(title)
        title_label.setStyleSheet("font-size: 14px; font-weight: bold; margin: 20px 0 10px 0; color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 5px;")
        section_layout.addWidget(title_label)
        
        # Section content
        content_label = QLabel(content.strip())
        content_label.setStyleSheet("font-size: 12px; margin: 0 0 15px 0; line-height: 1.6; color: #34495e;")
        content_label.setWordWrap(True)
        section_layout.addWidget(content_label)
        
        layout.addWidget(section_widget)
        
        # Store reference for navigation
        self.chapters[chapter_id] = section_widget
    
    def add_chapter_with_icons(self, layout, chapter_id, title, content):
        """Add a chapter with toolbar button icons"""
        # Section container
        section_widget = QWidget()
        section_layout = QVBoxLayout(section_widget)
        section_layout.setContentsMargins(0, 0, 0, 0)
        
        # Section title
        title_label = QLabel(title)
        title_label.setStyleSheet("font-size: 14px; font-weight: bold; margin: 20px 0 10px 0; color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 5px;")
        section_layout.addWidget(title_label)
        
        # Section content
        content_label = QLabel(content.strip())
        content_label.setStyleSheet("font-size: 12px; margin: 0 0 15px 0; line-height: 1.6; color: #34495e;")
        content_label.setWordWrap(True)
        section_layout.addWidget(content_label)
        
        # Add toolbar button icons with descriptions
        self.add_toolbar_buttons(section_layout)
        
        layout.addWidget(section_widget)
        
        # Store reference for navigation
        self.chapters[chapter_id] = section_widget
    
    def add_toolbar_buttons(self, layout):
        """Add toolbar button icons with descriptions"""
        # Create a grid layout for buttons
        buttons_widget = QWidget()
        buttons_layout = QGridLayout(buttons_widget)
        buttons_layout.setSpacing(15)
        buttons_layout.setContentsMargins(10, 10, 10, 10)
        
        # Define all unique buttons with their icons and descriptions
        all_buttons = [
            # Toolbar buttons
            ("search.png", "Poišči delavca", "Orodna vrstica: Iskanje delavcev po ID-ju ali imenu\nKratice: Ctrl+F"),
            ("export.png", "Izvozi vse podatke", "Orodna vrstica: Izvoz vseh podatkov v JSON format\nKratice: Ctrl+E"),
            ("import.png", "Uvozi vse podatke", "Orodna vrstica: Uvoz podatkov iz JSON datoteke"),
            ("settings.png", "Nastavitve / Spremeni ID", "Orodna vrstica: Konfiguracija programa\nKontekstni meni: Spreminjanje ID-ja kartice"),
            ("refresh.png", "Osveži podatke", "Orodna vrstica: Ročna sinhronizacija z SMB shrambo\nKratice: F5"),
            ("manual.png", "O programu", "Orodna vrstica: Prikaz informacij o programu\nKratice: F1"),
            # Unique tablet/context menu buttons
            ("calendar.png", "Koledar", "Glavno okno: Prikaz koledarskega dialoga za izbiranje datuma"),
            ("archive.png", "Arhiviraj delavca", "Glavno okno / Kontekstni meni: Arhiviranje podatkov delavca"),
            ("lostcard.png", "Izgubljena kartica", "Glavno okno / Kontekstni meni: Označitev izgubljene kartice"),
            ("trashcan.png", "Izbriši delavca", "Glavno okno / Kontekstni meni: Brisanje delavca iz sistema")
        ]
        
        # Calculate grid dimensions (4 columns)
        cols = 4
        rows = (len(all_buttons) + cols - 1) // cols
        
        for i, (icon_name, button_name, description) in enumerate(all_buttons):
            row = i // cols
            col = i % cols
            
            # Create button container
            button_container = QWidget()
            button_container.setFixedSize(140, 110)
            button_layout = QVBoxLayout(button_container)
            button_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            button_layout.setSpacing(5)
            
            # Add icon
            icon_label = QLabel()
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), icon_name)
            if os.path.exists(icon_path):
                try:
                    icon = QIcon(icon_path)
                    if not icon.isNull():
                        pixmap = icon.pixmap(40, 40)
                        icon_label.setPixmap(pixmap)
                    else:
                        icon_label.setText("📋")
                        icon_label.setStyleSheet("font-size: 28px;")
                except Exception as e:
                    icon_label.setText("📋")
                    icon_label.setStyleSheet("font-size: 28px;")
            else:
                icon_label.setText("📋")
                icon_label.setStyleSheet("font-size: 28px;")
            
            icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            button_layout.addWidget(icon_label)
            
            # Add button name
            name_label = QLabel(button_name)
            name_label.setStyleSheet("font-size: 9px; font-weight: bold; text-align: center; color: #2c3e50;")
            name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            name_label.setWordWrap(True)
            button_layout.addWidget(name_label)
            
            # Add description
            desc_label = QLabel(description)
            desc_label.setStyleSheet("font-size: 8px; text-align: center; color: #7f8c8d;")
            desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            desc_label.setWordWrap(True)
            button_layout.addWidget(desc_label)
            
            buttons_layout.addWidget(button_container, row, col)
        
        layout.addWidget(buttons_widget)
    
    def add_section(self, layout, title, content):
        """Add a section to the manual (legacy method for compatibility)"""
        self.add_section_with_icon(layout, title, content, "manual.png")

class SMBUpdateWorker(QThread):
    """Worker thread for checking SMB updates without blocking the GUI"""
    update_available = pyqtSignal(int)  # Signal with version number when update is loaded
    update_complete = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_window = parent
        self.running = True
    
    def run(self):
        """Check for updates in background thread"""
        if not self.running or not self.parent_window:
            return
            
        try:
            # Get current version from SMB (this is the blocking operation)
            current_version = self.parent_window.get_data_version()
            
            # Check if update is needed
            if current_version > self.parent_window.last_known_version:
                # Load the updated data in background thread
                print(f"Background: Loading version {current_version}")
                if self.parent_window.load_shared_data_from_smb():
                    # Signal that update has been loaded successfully
                    self.update_available.emit(current_version)
                else:
                    print(f"Background: Failed to load version {current_version}")
            else:
                # No update needed, just update the known version
                self.parent_window.last_known_version = current_version
                
        except Exception as e:
            print(f"Error in SMB update worker: {str(e)}")
        finally:
            self.update_complete.emit()
    
    def stop(self):
        """Stop the worker thread"""
        self.running = False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Beleženje delovnega časa Admin")
        self.setMinimumSize(1200, 800)
        
        # Set window icon using absolute path
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon.ico')
        if os.path.exists(icon_path):
            try:
                icon = QIcon(icon_path)
                if not icon.isNull():
                    self.setWindowIcon(icon)
                else:
                    print(f"Warning: icon.ico could not be loaded as icon at {icon_path}")
            except Exception as e:
                print(f"Error loading icon.ico: {str(e)}")
        else:
            print(f"Warning: icon.ico not found at {icon_path}")
        
        # Initialize database
        self.init_database()
        
        # Create toolbar
        self.toolbar = self.addToolBar("Toolbar")
        self.toolbar.setMovable(False)
        
        # Create search action with custom icon
        self.search_action = QAction("Poišči delavca", self)
        search_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'search.png')
        if os.path.exists(search_icon_path):
            try:
                icon = QIcon(search_icon_path)
                if not icon.isNull():
                    self.search_action.setIcon(icon)
                else:
                    print(f"Warning: search.png could not be loaded as icon at {search_icon_path}")
            except Exception as e:
                print(f"Error loading search.png: {str(e)}")
        else:
            print(f"Warning: search.png not found at {search_icon_path}")
        self.search_action.triggered.connect(self.show_search_dialog)
        self.search_action.setToolTip("Poišči delavca po ID-ju")  # Updated tooltip
        self.toolbar.addAction(self.search_action)
        
        # Create export configuration action with custom icon
        self.export_config_action = QAction("Izvozi vse podatke", self)
        export_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'export.png')
        if os.path.exists(export_icon_path):
            try:
                icon = QIcon(export_icon_path)
                if not icon.isNull():
                    self.export_config_action.setIcon(icon)
                else:
                    print(f"Warning: export.png could not be loaded as icon at {export_icon_path}")
            except Exception as e:
                print(f"Error loading export.png: {str(e)}")
        else:
            print(f"Warning: export.png not found at {export_icon_path}")
        self.export_config_action.triggered.connect(self.export_configuration)
        self.toolbar.addAction(self.export_config_action)
        
        # Create import configuration action with custom icon
        self.import_config_action = QAction("Uvozi vse podatke", self)
        import_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'import.png')
        if os.path.exists(import_icon_path):
            try:
                icon = QIcon(import_icon_path)
                if not icon.isNull():
                    self.import_config_action.setIcon(icon)
                else:
                    print(f"Warning: import.png could not be loaded as icon at {import_icon_path}")
            except Exception as e:
                print(f"Error loading import.png: {str(e)}")
        else:
            print(f"Warning: import.png not found at {import_icon_path}")
        self.import_config_action.triggered.connect(self.import_configuration)
        self.toolbar.addAction(self.import_config_action)
        
        # Create settings action with custom icon
        self.settings_action = QAction("Nastavitve", self)
        settings_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'settings.png')
        if os.path.exists(settings_icon_path):
            try:
                icon = QIcon(settings_icon_path)
                if not icon.isNull():
                    self.settings_action.setIcon(icon)
                else:
                    print(f"Warning: settings.png could not be loaded as icon at {settings_icon_path}")
            except Exception as e:
                print(f"Error loading settings.png: {str(e)}")
        else:
            print(f"Warning: settings.png not found at {settings_icon_path}")
        self.settings_action.triggered.connect(self.show_settings)
        self.toolbar.addAction(self.settings_action)
        
        # Create refresh action with custom icon
        self.refresh_action = QAction("Osveži podatke", self)
        refresh_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'refresh.png')
        if os.path.exists(refresh_icon_path):
            try:
                icon = QIcon(refresh_icon_path)
                if not icon.isNull():
                    self.refresh_action.setIcon(icon)
                else:
                    print(f"Warning: refresh.png could not be loaded as icon at {refresh_icon_path}")
            except Exception as e:
                print(f"Error loading refresh.png: {str(e)}")
        else:
            print(f"Warning: refresh.png not found at {refresh_icon_path}")
        self.refresh_action.triggered.connect(self.manual_refresh)
        self.refresh_action.setToolTip("Osveži podatke iz SMB shrambe")
        self.toolbar.addAction(self.refresh_action)
        
        # Create manual action with custom icon
        self.manual_action = QAction("Navodila", self)
        manual_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'manual.png')
        if os.path.exists(manual_icon_path):
            try:
                icon = QIcon(manual_icon_path)
                if not icon.isNull():
                    self.manual_action.setIcon(icon)
                else:
                    print(f"Warning: manual.png could not be loaded as icon at {manual_icon_path}")
            except Exception as e:
                print(f"Error loading manual.png: {str(e)}")
        else:
            print(f"Warning: manual.png not found at {manual_icon_path}")
        self.manual_action.triggered.connect(self.show_manual)
        self.manual_action.setToolTip("O programu")
        self.toolbar.addAction(self.manual_action)
        
        # Create main widget and layout
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        self.layout = QVBoxLayout(self.main_widget)
        
        # Create tab widget
        self.tabs = QTabWidget()
        self.layout.addWidget(self.tabs)
        
        # Create tabs in desired order
        self.main_tab = QWidget()
        self.groups_tab = QWidget()
        self.group_calc_tab = QWidget()
        
        # Add tabs in desired order
        self.tabs.addTab(self.main_tab, "Glavno okno")
        self.tabs.addTab(self.groups_tab, "Skupine")
        self.tabs.addTab(self.group_calc_tab, "Naredi skupni izračun")
        
        # Initialize all tabs
        self.init_main_tab()
        self.init_groups_tab()
        self.init_group_calc_tab()
        
        # Initialize multi-user synchronization
        self.last_known_version = 0
        self.data_version = 0
        
        # Initialize worker thread for background SMB updates
        self.update_worker = None
        self.update_in_progress = False
        
        # Add timer for periodic update checks
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.check_for_updates)
        self.update_timer.start(30000)  # Check every 30 seconds
        
        # Load shared data on startup
        self.load_shared_data_from_smb()

    def show_settings(self):
        """Show settings dialog"""
        dialog = SettingsDialog(self)
        dialog.exec()
    
    def show_manual(self):
        """Show manual dialog"""
        dialog = ManualDialog(self)
        dialog.exec()
    
    def manual_refresh(self):
        """Manually refresh data from SMB share"""
        try:
            if self.load_shared_data_from_smb():
                self.update_employee_table()
                self.update_groups_list()
                QMessageBox.information(self, "Uspeh", "Podatki so bili uspešno osveženi!")
            else:
                QMessageBox.warning(self, "Napaka", "Ni mogoče naložiti podatkov iz SMB shrambe.")
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri osveževanju podatkov: {str(e)}")

    def init_database(self):
        """Initialize SQLite database for storing employee and group data"""
        try:
            self.conn = sqlite3.connect('employee_data.db')
            self.cursor = self.conn.cursor()
            
            # Create employees table (remove sp_enabled)
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS employees (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    card_id TEXT UNIQUE NOT NULL,
                    daily_hours REAL NOT NULL,
                    group_id INTEGER,
                    FOREIGN KEY (group_id) REFERENCES groups(id)
                )
            ''')
            
            # Remove sp_enabled column if it exists (SQLite does not support DROP COLUMN directly, so skip for now)
            # If you want to fully remove it, you would need to migrate the table.
            
            # Create groups table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS groups (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL
                )
            ''')
            
            # Create special days table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS special_days (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    card_id TEXT NOT NULL,
                    date DATE NOT NULL,
                    type TEXT NOT NULL,
                    FOREIGN KEY (card_id) REFERENCES employees(card_id),
                    UNIQUE(card_id, date)
                )
            ''')
            
            self.conn.commit()
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri inicializaciji baze: {str(e)}")
            print(f"Error in init_database: {str(e)}")

    def init_main_tab(self):
        """Initialize the main tab with employee table and buttons"""
        layout = QVBoxLayout(self.main_tab)
        
        # Add search box above the employee table
        search_layout = QHBoxLayout()
        self.name_search_box = QLineEdit()
        self.name_search_box.setPlaceholderText("Poišči po imenu")
        search_layout.addWidget(self.name_search_box)
        layout.addLayout(search_layout)
        
        # Create table for employee data (remove S.P. column)
        self.employee_table = QTableWidget()
        self.employee_table.setColumnCount(6)
        self.employee_table.setHorizontalHeaderLabels([
            "Ime delavca", "ID kartice", "Dnevni delovni čas", "Skupina", "Akcije", "SMB"
        ])
        # Set column widths
        self.employee_table.setColumnWidth(0, 200)  # Name
        self.employee_table.setColumnWidth(1, 150)  # Card ID
        self.employee_table.setColumnWidth(2, 150)  # Daily hours
        self.employee_table.setColumnWidth(3, 150)  # Group
        self.employee_table.setColumnWidth(4, 500)  # Actions (wider for full button text)
        self.employee_table.setColumnWidth(5, 50)   # SMB
        
        layout.addWidget(self.employee_table)
        
        # Add buttons for each row
        self.add_employee_button = QPushButton("Dodaj delavca")
        self.add_employee_button.clicked.connect(self.add_employee_row)
        layout.addWidget(self.add_employee_button)
        
        # Load existing employees
        self.update_employee_table()

    def init_groups_tab(self):
        """Initialize the groups management tab"""
        layout = QVBoxLayout(self.groups_tab)
        
        # Groups list
        self.groups_list = QTableWidget()
        self.groups_list.setColumnCount(2)
        self.groups_list.setHorizontalHeaderLabels(["Ime skupine", "Akcije"])
        layout.addWidget(self.groups_list)
        
        # Add group section
        add_group_layout = QHBoxLayout()
        self.new_group_input = QLineEdit()
        self.add_group_button = QPushButton("Dodaj skupino")
        self.add_group_button.clicked.connect(self.add_group)
        
        add_group_layout.addWidget(QLabel("Napiši ime nove skupine:"))
        add_group_layout.addWidget(self.new_group_input)
        add_group_layout.addWidget(self.add_group_button)
        layout.addLayout(add_group_layout)
        
        # Show existing groups
        self.update_groups_list()

    def update_groups_list(self):
        """Update the groups list display"""
        try:
            self.groups_list.setRowCount(0)
            self.cursor.execute("SELECT id, name FROM groups ORDER BY name")
            groups = self.cursor.fetchall()
            
            if groups:
                for group_id, name in groups:
                    row_position = self.groups_list.rowCount()
                    self.groups_list.insertRow(row_position)
                    self.groups_list.setItem(row_position, 0, QTableWidgetItem(name))
                    
                    delete_button = QPushButton("Izbriši")
                    delete_button.clicked.connect(lambda checked, gid=group_id: self.delete_group(gid))
                    self.groups_list.setCellWidget(row_position, 1, delete_button)
            
            # Also update the group combo box in the calculation tab
            if hasattr(self, 'group_combo'):
                self.group_combo.clear()
                self.group_combo.addItem("Vsi zaposleni", None)
                for group_id, name in groups:
                    self.group_combo.addItem(name, group_id)
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri posodabljanju seznama skupin: {str(e)}")
            print(f"Error in update_groups_list: {str(e)}")

    def delete_group(self, group_id):
        """Delete a group"""
        reply = QMessageBox.question(self, "Potrditev", 
                                   "Ali ste prepričani, da želite izbrisati to skupino?",
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            self.cursor.execute("DELETE FROM groups WHERE id = ?", (group_id,))
            self.conn.commit()
            self.update_groups_list()
            
            # Save shared data to SMB
            self.save_shared_data_with_retry()

    def show_calendar_dialog(self, callback):
        """Show calendar dialog and call callback with selected dates"""
        dialog = DateRangeDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            start_date, end_date = dialog.get_dates()
            callback(start_date, end_date)

    def calculate_overtime(self, card_id, daily_hours):
        """Calculate overtime hours for an employee"""
        def on_dates_selected(start_date, end_date):
            result = self.calculate_working_hours(card_id, start_date, end_date)
            if result is not None:
                # Check if employee has flexible hours
                if daily_hours == -1:
                    QMessageBox.information(self, "Opozorilo", "Delavec ima gibljivi delovni čas. Nadure niso izračunane.")
                    return
                # Calculate overtime (only positive values, excluding special days)
                result['Nadure'] = result.apply(
                    lambda x: max(0, x['Delovne ure'] - daily_hours) if x['Status'] not in ['Bolniški stalež', 'Dopust'] else 0,
                    axis=1
                )
                result['Status'] = result.apply(
                    lambda x: x['Status'] if x['Status'] in ['Bolniški stalež', 'Dopust'] else ('Nadure' if x['Nadure'] > 0 else 'Normalno'),
                    axis=1
                )
                # Calculate total overtime for the period
                total_overtime = result['Nadure'].sum()
                summary = f"Skupaj nadur za izbrano obdobje: {total_overtime:.2f} ur"
                dialog = ResultsDialog(result, summary)
                dialog.exec()
        self.show_calendar_dialog(on_dates_selected)

    def calculate_shortage(self, card_id, daily_hours):
        """Calculate shortage hours for an employee"""
        def on_dates_selected(start_date, end_date):
            result = self.calculate_working_hours(card_id, start_date, end_date)
            if result is not None:
                # Check if employee has flexible hours
                if daily_hours == -1:
                    QMessageBox.information(self, "Opozorilo", "Delavec ima gibljivi delovni čas. Manjko ur ni izračunano.")
                    return
                # Calculate shortage (only positive values, excluding special days)
                result['Manjko ur'] = result.apply(
                    lambda x: max(0, daily_hours - x['Delovne ure']) if x['Status'] not in ['Bolniški stalež', 'Dopust'] else 0,
                    axis=1
                )
                result['Status'] = result.apply(
                    lambda x: x['Status'] if x['Status'] in ['Bolniški stalež', 'Dopust'] else ('Manjko ur' if x['Manjko ur'] > 0 else 'Normalno'),
                    axis=1
                )
                # Correctly calculate total undertime for the period (unrounded sum)
                undertime = 0.0
                for idx, row in result.iterrows():
                    status = row['Status']
                    if isinstance(status, str) and status.startswith('Manjko ur'):
                        undertime += daily_hours - row['Delovne ure']
                summary = f"Skupaj manjko ur za izbrano obdobje: {undertime:.2f} ur"
                dialog = ResultsDialog(result, summary)
                dialog.exec()
        self.show_calendar_dialog(on_dates_selected)

    def init_group_calc_tab(self):
        """Initialize the group calculation tab"""
        layout = QVBoxLayout(self.group_calc_tab)
        
        # Group selection
        group_layout = QHBoxLayout()
        group_layout.addWidget(QLabel("Naredi izračun za:"))
        self.group_combo = QComboBox()
        
        # Add "All employees" option
        self.group_combo.addItem("Vsi zaposleni", None)
        
        # Add groups
        try:
            self.cursor.execute("SELECT id, name FROM groups ORDER BY name")
            groups = self.cursor.fetchall()
            if groups:
                for group_id, name in groups:
                    self.group_combo.addItem(name, group_id)
                    print(f"Added group: {name} with ID: {group_id}")  # Debug print
            else:
                print("No groups found in database")  # Debug print
        except Exception as e:
            print(f"Error loading groups: {str(e)}")  # Debug print
            QMessageBox.warning(self, "Napaka", f"Napaka pri nalaganju skupin: {str(e)}")
        
        group_layout.addWidget(self.group_combo)
        layout.addLayout(group_layout)
        
        # Calculation options
        options_group = QGroupBox("V izračun vključi:")
        options_layout = QVBoxLayout()
        
        self.calc_work_hours = QCheckBox("Izračunaj delovni čas za določeno obdobje")
        self.calc_overtime = QCheckBox("Izračunaj nadure")
        self.calc_shortage = QCheckBox("Izračunaj manjko ur")
        
        options_layout.addWidget(self.calc_work_hours)
        options_layout.addWidget(self.calc_overtime)
        options_layout.addWidget(self.calc_shortage)
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # Calculate button
        self.calculate_button = QPushButton("Izračunaj")
        self.calculate_button.clicked.connect(self.calculate_group_hours)
        layout.addWidget(self.calculate_button)

    def calculate_group_hours(self):
        """Calculate working hours for all employees in a group"""
        def on_dates_selected(start_date, end_date):
            group_id = self.group_combo.currentData()
            
            # Get all employees (either from specific group or all employees)
            if group_id is None:  # "Vsi zaposleni" is selected
                self.cursor.execute("""
                    SELECT card_id, name, daily_hours 
                    FROM employees 
                    ORDER BY name
                """)
            else:
                self.cursor.execute("""
                    SELECT card_id, name, daily_hours 
                    FROM employees 
                    WHERE group_id = ?
                    ORDER BY name
                """, (group_id,))
            
            employees = self.cursor.fetchall()
            
            if not employees:
                QMessageBox.warning(self, "Opozorilo", "Ni delavcev za izračun.")
                return
            
            all_results = []
            for card_id, name, daily_hours in employees:
                result = self.calculate_working_hours(card_id, start_date, end_date)
                if result is not None:
                    # Calculate summary statistics
                    total_days = len(result)
                    total_hours = result['Delovne ure'].sum()
                    
                    # Count days with different statuses
                    shortage_days = len(result[result['Status'].str.startswith('Manjko ur')])
                    overtime_days = len(result[result['Status'].str.startswith('Nadure')])
                    
                    # Get special days from database
                    self.cursor.execute("""
                        SELECT type, COUNT(*) as count
                        FROM special_days
                        WHERE card_id = ? AND date BETWEEN ? AND ?
                        GROUP BY type
                    """, (card_id, start_date, end_date))
                    
                    special_days = self.cursor.fetchall()
                    sick_leave_days = 0
                    vacation_days = 0
                    
                    for day_type, count in special_days:
                        if day_type == 'sick_leave':
                            sick_leave_days = count
                        elif day_type == 'vacation':
                            vacation_days = count
                    
                    # Calculate overtime and shortage hours if checkboxes are selected
                    overtime_hours = 0
                    shortage_hours = 0
                    
                    if self.calc_overtime.isChecked():
                        overtime_hours = result[result['Status'].str.startswith('Nadure')]['Delovne ure'].sum() - (daily_hours * overtime_days)
                    
                    if self.calc_shortage.isChecked():
                        shortage_hours = (daily_hours * shortage_days) - result[result['Status'].str.startswith('Manjko ur')]['Delovne ure'].sum()
                    
                    # Calculate net hours (Skupaj) if at least one is checked
                    net_hours = None
                    if self.calc_overtime.isChecked() or self.calc_shortage.isChecked():
                        net_hours = round(overtime_hours - shortage_hours, 2)
                    
                    # Create summary row
                    summary_row = {
                        'Številka kartice': card_id,
                        'Ime delavca': name,
                        'Dnevni delovni čas': 'Gibljiv' if daily_hours == -1 else str(daily_hours),
                        'Skupno delovnih dni': total_days,
                        'Skupno delovnih ur': round(total_hours, 2),
                        'Dnevi z manjko ur': shortage_days,
                        'Dnevi z nadurami': overtime_days,
                        'Bolniški stalež - št. dni': sick_leave_days,
                        'Dopust - št. dni': vacation_days
                    }
                    
                    # Add optional columns if checkboxes are selected
                    if self.calc_overtime.isChecked():
                        summary_row['Nadure'] = round(overtime_hours, 2)
                    
                    if self.calc_shortage.isChecked():
                        summary_row['Manjko ur'] = round(shortage_hours, 2)
                    
                    if net_hours is not None:
                        summary_row['Neto ur'] = net_hours
                    
                    all_results.append(summary_row)
            
            if not all_results:
                QMessageBox.warning(self, "Opozorilo", "Ni podatkov za izračun.")
                return
            
            # Create DataFrame from results
            result_df = pd.DataFrame(all_results)
            
            # Show results without summary
            dialog = ResultsDialog(result_df)
            dialog.exec()
        
        self.show_calendar_dialog(on_dates_selected)

    def update_employee_table(self):
        """Update the employee table with current data from the database"""
        try:
            # Clear the table
            self.employee_table.setRowCount(0)
            
            # Get all employees with their group names (remove sp_enabled)
            self.cursor.execute("""
                SELECT e.id, e.name, e.card_id, e.daily_hours, g.id as group_id, g.name as group_name
                FROM employees e
                LEFT JOIN groups g ON e.group_id = g.id
                ORDER BY e.name
            """)
            employees = self.cursor.fetchall()
            
            # Get all available groups
            self.cursor.execute("SELECT id, name FROM groups ORDER BY name")
            groups = self.cursor.fetchall()
            
            for employee in employees:
                row_position = self.employee_table.rowCount()
                self.employee_table.insertRow(row_position)
                
                # Always set a QTableWidgetItem for the name column
                name_item = QTableWidgetItem(employee[1])
                self.employee_table.setItem(row_position, 0, name_item)
                
                # Add employee data for card ID
                self.employee_table.setItem(row_position, 1, QTableWidgetItem(employee[2]))  # Card ID
                
                # Create and add daily hours combo box
                hours_combo = QComboBox()
                hours_combo.addItem("Gibljivi delovni čas", -1)
                for hours in range(1, 13):  # 1 to 12 hours
                    hours_combo.addItem(f"{hours} ur", hours)
                
                # Set current hours
                current_hours = employee[3]
                if current_hours == -1:
                    hours_combo.setCurrentText("Gibljivi delovni čas")
                else:
                    hours_combo.setCurrentText(f"{int(current_hours)} ur")
                
                # Connect combo box change to update database
                def on_hours_changed(combo, emp_id):
                    new_hours = combo.currentData()
                    try:
                        self.cursor.execute("UPDATE employees SET daily_hours = ? WHERE id = ?", 
                                          (new_hours, emp_id))
                        self.conn.commit()
                        
                        # Save shared data to SMB
                        self.save_shared_data_with_retry()
                    except Exception as e:
                        QMessageBox.critical(self, "Napaka", f"Napaka pri posodabljanju delovnega časa: {str(e)}")
                
                hours_combo.currentIndexChanged.connect(
                    lambda index, combo=hours_combo, emp_id=employee[0]: on_hours_changed(combo, emp_id)
                )
                
                self.employee_table.setCellWidget(row_position, 2, hours_combo)
                
                # Create and add group combo box
                group_combo = QComboBox()
                group_combo.addItem("", None)  # Empty option for no group
                for group_id, group_name in groups:
                    group_combo.addItem(group_name, group_id)
                
                # Set current group if exists
                if employee[4]:  # if group_id exists
                    index = group_combo.findData(employee[4])
                    if index >= 0:
                        group_combo.setCurrentIndex(index)
                
                def on_group_changed(combo, emp_id):
                    new_group_id = combo.currentData()
                    try:
                        self.cursor.execute("UPDATE employees SET group_id = ? WHERE id = ?", 
                                          (new_group_id, emp_id))
                        self.conn.commit()
                        
                        # Save shared data to SMB
                        self.save_shared_data_with_retry()
                    except Exception as e:
                        QMessageBox.critical(self, "Napaka", f"Napaka pri posodabljanju skupine: {str(e)}")
                group_combo.currentIndexChanged.connect(
                    lambda index, combo=group_combo, emp_id=employee[0]: on_group_changed(combo, emp_id)
                )
                self.employee_table.setCellWidget(row_position, 3, group_combo)
                
                # Add action buttons
                button_widget = QWidget()
                button_layout = QHBoxLayout(button_widget)
                button_layout.setContentsMargins(0, 0, 0, 0)
                button_layout.setSpacing(2)
                
                calendar_button = QPushButton()
                calendar_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'calendar.png')
                # Add error handling for icon loading
                if os.path.exists(calendar_icon_path):
                    try:
                        icon = QIcon(calendar_icon_path)
                        if not icon.isNull():
                            calendar_button.setIcon(icon)
                        else:
                            print(f"Warning: calendar.png could not be loaded as icon at {calendar_icon_path}")
                            # Fallback to text
                            calendar_button.setText("📅")
                    except Exception as e:
                        print(f"Error loading calendar.png: {str(e)}")
                        # Fallback to text
                        calendar_button.setText("📅")
                else:
                    print(f"Warning: calendar.png not found at {calendar_icon_path}")
                    # Fallback to text
                    calendar_button.setText("📅")
                calendar_button.clicked.connect(lambda checked, cid=employee[2]: self.show_calendar(cid))
                button_layout.addWidget(calendar_button)
                
                # Add lost card button
                lost_card_button = QPushButton()
                lost_card_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'lostcard.png')
                # Add error handling for icon loading
                if os.path.exists(lost_card_icon_path):
                    try:
                        icon = QIcon(lost_card_icon_path)
                        if not icon.isNull():
                            lost_card_button.setIcon(icon)
                        else:
                            print(f"Warning: lostcard.png could not be loaded as icon at {lost_card_icon_path}")
                            # Fallback to text
                            lost_card_button.setText("ID")
                    except Exception as e:
                        print(f"Error loading lostcard.png: {str(e)}")
                        # Fallback to text
                        lost_card_button.setText("ID")
                else:
                    print(f"Warning: lostcard.png not found at {lost_card_icon_path}")
                    # Fallback to text
                    lost_card_button.setText("ID")
                lost_card_button.setToolTip("Spremeni ID")
                lost_card_button.clicked.connect(lambda checked, cid=employee[2], name=employee[1]: self.show_change_card_id_dialog(cid, name))
                button_layout.addWidget(lost_card_button)
                
                # Add archive button
                archive_button = QPushButton()
                archive_icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'archive.png')
                # Add error handling for icon loading
                if os.path.exists(archive_icon_path):
                    try:
                        icon = QIcon(archive_icon_path)
                        if not icon.isNull():
                            archive_button.setIcon(icon)
                        else:
                            print(f"Warning: archive.png could not be loaded as icon at {archive_icon_path}")
                            # Fallback to text
                            archive_button.setText("📦")
                    except Exception as e:
                        print(f"Error loading archive.png: {str(e)}")
                        # Fallback to text
                        archive_button.setText("📦")
                else:
                    print(f"Warning: archive.png not found at {archive_icon_path}")
                    # Fallback to text
                    archive_button.setText("📦")
                archive_button.setToolTip("Arhiviraj delavca")
                archive_button.clicked.connect(lambda checked, cid=employee[2], name=employee[1]: self.archive_worker_data(cid, name))
                button_layout.addWidget(archive_button)
                
                calc_button = QPushButton("Izračunaj delovni čas")
                overtime_button = QPushButton("Pokaži nadure")
                shortage_button = QPushButton("Pokaži manjko ur")
                delete_button = QPushButton("Izbriši delavca")
                
                # Add trashcan button for deleting timestamps from SMB
                trashcan_button = QPushButton()
                trashcan_icon_path = os.path.join(os.path.dirname(__file__), "trashcan.png")
                if os.path.exists(trashcan_icon_path):
                    try:
                        trashcan_icon = QIcon(trashcan_icon_path)
                        trashcan_button.setIcon(trashcan_icon)
                        trashcan_button.setIconSize(QSize(20, 20))
                    except Exception as e:
                        print(f"Error loading trashcan.png: {str(e)}")
                        # Fallback to text
                        trashcan_button.setText("🗑️")
                else:
                    print(f"Warning: trashcan.png not found at {trashcan_icon_path}")
                    # Fallback to text
                    trashcan_button.setText("🗑️")
                trashcan_button.setToolTip("Izbriši delavca iz shrambe")
                trashcan_button.clicked.connect(lambda checked, cid=employee[2], name=employee[1]: self.delete_worker_timestamps(cid, name))
                
                calc_button.clicked.connect(lambda checked, cid=employee[2]: self.calculate_employee_hours(cid))
                overtime_button.clicked.connect(lambda checked, cid=employee[2], hours=employee[3]: self.calculate_overtime(cid, hours))
                shortage_button.clicked.connect(lambda checked, cid=employee[2], hours=employee[3]: self.calculate_shortage(cid, hours))
                delete_button.clicked.connect(lambda checked, emp_id=employee[0], emp_name=employee[1]: self.delete_employee(emp_id, emp_name))
                
                button_layout.addWidget(calc_button)
                button_layout.addWidget(overtime_button)
                button_layout.addWidget(shortage_button)
                button_layout.addWidget(delete_button)
                
                self.employee_table.setCellWidget(row_position, 4, button_widget)
                
                # Add trashcan button to SMB column
                self.employee_table.setCellWidget(row_position, 5, trashcan_button)

                # Set tooltip for all cells in this row to the worker's name
                for col in range(self.employee_table.columnCount()):
                    item = self.employee_table.item(row_position, col)
                    if item is None:
                        item = QTableWidgetItem()
                        self.employee_table.setItem(row_position, col, item)
                    item.setToolTip(employee[1])
            
            # Reconnect the search box signal to ensure it works after table update
            try:
                self.name_search_box.returnPressed.disconnect()
            except Exception:
                pass
            self.name_search_box.returnPressed.connect(self.search_employee_by_name)
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri posodabljanju tabele: {str(e)}")
            print(f"Error in update_employee_table: {str(e)}")


    def delete_employee(self, employee_id, employee_name):
        """Delete an employee from the database"""
        reply = QMessageBox.question(
            self,
            "Potrditev brisanja",
            f"Ali ste prepričani, da želite izbrisati delavca {employee_name}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                self.cursor.execute("DELETE FROM employees WHERE id = ?", (employee_id,))
                self.conn.commit()
                self.update_employee_table()
                
                # Update worker_id.csv file
                self.update_worker_id_file()
                
                # Save shared data to SMB
                self.save_shared_data_with_retry()
                
                QMessageBox.information(self, "Uspeh", f"Delavec {employee_name} je bil uspešno izbrisan.")
            except Exception as e:
                QMessageBox.critical(self, "Napaka", f"Napaka pri brisanju delavca: {str(e)}")

    def add_group(self):
        """Add a new group"""
        group_name = self.new_group_input.text()
        if group_name:
            try:
                self.cursor.execute("INSERT INTO groups (name) VALUES (?)", (group_name,))
                self.conn.commit()
                self.new_group_input.clear()
                self.update_groups_list()
                
                # Save shared data to SMB
                self.save_shared_data_with_retry()
                
                QMessageBox.information(self, "Uspeh", "Skupina je bila uspešno dodana!")
            except sqlite3.IntegrityError:
                QMessageBox.warning(self, "Napaka", "Skupina s tem imenom že obstaja!")
            except Exception as e:
                QMessageBox.critical(self, "Napaka", f"Napaka pri dodajanju skupine: {str(e)}")

    def add_employee_row(self):
        """Add a new row to the employee table"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Dodaj novega delavca")
        layout = QVBoxLayout(dialog)
        
        # Name input
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("Ime delavca:"))
        name_input = QLineEdit()
        name_layout.addWidget(name_input)
        layout.addLayout(name_layout)
        
        # Card ID input
        card_layout = QHBoxLayout()
        card_layout.addWidget(QLabel("ID kartice:"))
        card_input = QLineEdit()
        card_input.setMaxLength(14)
        card_layout.addWidget(card_input)
        layout.addLayout(card_layout)
        
        # Working hours input
        hours_layout = QVBoxLayout()
        hours_group = QGroupBox("Delovni čas")
        
        # Fixed hours option
        fixed_hours_layout = QHBoxLayout()
        self.fixed_hours_radio = QRadioButton("Fiksni delovni čas")
        self.fixed_hours_radio.setChecked(True)
        self.hours_input = QLineEdit()
        self.hours_input.setText("8.0")
        self.hours_input.setEnabled(True)
        fixed_hours_layout.addWidget(self.fixed_hours_radio)
        fixed_hours_layout.addWidget(self.hours_input)
        fixed_hours_layout.addWidget(QLabel("ur/dan"))
        hours_layout.addLayout(fixed_hours_layout)
        
        # Flexible hours option
        flexible_hours_radio = QRadioButton("Gibljivi delovni čas")
        hours_layout.addWidget(flexible_hours_radio)
        
        # Connect radio buttons
        def on_radio_changed():
            self.hours_input.setEnabled(self.fixed_hours_radio.isChecked())
        
        self.fixed_hours_radio.toggled.connect(on_radio_changed)
        flexible_hours_radio.toggled.connect(on_radio_changed)
        
        hours_group.setLayout(hours_layout)
        layout.addWidget(hours_group)
        
        # Group selection
        group_layout = QHBoxLayout()
        group_layout.addWidget(QLabel("Skupina:"))
        group_combo = QComboBox()
        self.cursor.execute("SELECT id, name FROM groups")
        for group_id, name in self.cursor.fetchall():
            group_combo.addItem(name, group_id)
        group_layout.addWidget(group_combo)
        layout.addLayout(group_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        save_button = QPushButton("Shrani")
        cancel_button = QPushButton("Prekliči")
        
        def save_employee():
            try:
                name = name_input.text()
                if not name:
                    QMessageBox.warning(dialog, "Napaka", "Vnesite ime delavca!")
                    return
                
                card_id = card_input.text().lower()  # Convert to lowercase
                if not card_id:
                    QMessageBox.warning(dialog, "Napaka", "Vnesite številko kartice!")
                    return
                
                # Validate hexadecimal format
                if len(card_id) != 14 or not all(c in '0123456789abcdef' for c in card_id):
                    QMessageBox.warning(dialog, "Napaka", "Neveljavna številka kartice! Vnesite 14-mestno šestnajstiško število.")
                    return
                
                # Handle working hours
                if self.fixed_hours_radio.isChecked():
                    try:
                        hours = float(self.hours_input.text())
                        if hours <= 0:
                            raise ValueError
                    except ValueError:
                        QMessageBox.warning(dialog, "Napaka", "Neveljavna vrednost za dnevni delovni čas!")
                        return
                else:
                    hours = -1  # Special value for flexible hours
                
                group_id = group_combo.currentData()
                
                self.cursor.execute("""
                    INSERT INTO employees (name, card_id, daily_hours, group_id)
                    VALUES (?, ?, ?, ?)
                """, (name, card_id, hours, group_id))
                self.conn.commit()
                
                # Update the employee table
                self.update_employee_table()
                
                # Update worker_id.csv file
                self.update_worker_id_file()
                
                # Save shared data to SMB
                self.save_shared_data_with_retry()
                
                dialog.accept()
                QMessageBox.information(self, "Uspeh", "Delavec je bil uspešno dodan!")
            except sqlite3.IntegrityError:
                QMessageBox.warning(dialog, "Napaka", "Delavec s to kartico že obstaja!")
            except Exception as e:
                QMessageBox.critical(dialog, "Napaka", f"Napaka pri shranjevanju: {str(e)}")
        save_button.clicked.connect(save_employee)
        cancel_button.clicked.connect(dialog.reject)
        
        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        dialog.exec()

    def calculate_employee_hours(self, card_id):
        """Calculate working hours for an employee"""
        try:
            def on_dates_selected(start_date, end_date):
                try:
                    result = self.calculate_working_hours(card_id, start_date, end_date)
                    if result is not None:
                        # Calculate total hours
                        total_hours = result['Delovne ure'].sum()
                        total_days = len(result)
                        normal_days = len(result[result['Status'] == 'Normalno'])
                        shortage_days = len(result[result['Status'] == 'Manjko ur'])
                        incomplete_days = len(result[result['Status'] == 'Nepopolni podatki'])
                        # Get daily_hours for this employee
                        self.cursor.execute("SELECT daily_hours FROM employees WHERE card_id = ?", (card_id,))
                        daily_hours = self.cursor.fetchone()[0]
                        # Correctly calculate total overtime (nadure) for the period (unrounded sum)
                        overtime = 0.0
                        undertime = 0.0
                        for idx, row in result.iterrows():
                            status = row['Status']
                            if isinstance(status, str) and status.startswith('Nadure'):
                                overtime += row['Delovne ure'] - daily_hours
                            if isinstance(status, str) and status.startswith('Manjko ur'):
                                undertime += daily_hours - row['Delovne ure']
                        neto_ur = overtime - undertime
                        # Add 'Neto ur' column to the result DataFrame
                        def calc_neto_ur(row):
                            status = row['Status']
                            if isinstance(status, str) and status.startswith('Nadure'):
                                return round(row['Delovne ure'] - daily_hours, 2)
                            elif isinstance(status, str) and status.startswith('Manjko ur'):
                                return round(row['Delovne ure'] - daily_hours, 2)
                            else:
                                return 0.0
                        result['Neto ur'] = result.apply(calc_neto_ur, axis=1)
                        # Create a summary string
                        summary = f"Skupni seštevek ur za izbrano obdobje:\n"
                        summary += f"Skupno število delovnih dni: {total_days}\n"
                        summary += f"Skupno število delovnih ur: {total_hours:.2f}\n"
                        summary += f"Število normalnih dni: {normal_days}\n"
                        summary += f"Število dni z manjko ur: {shortage_days}\n"
                        summary += f"Število dni z nepopolnimi podatki: {incomplete_days}"
                        summary += f"\nSkupaj nadur: {overtime:.2f} ur"
                        summary += f"\nSkupaj manjko ur: {undertime:.2f} ur"
                        summary += f"\nNeto ur: {neto_ur:.2f} ur"
                        # Create and show the results dialog
                        dialog = ResultsDialog(result, summary)
                        dialog.exec()
                except Exception as e:
                    QMessageBox.critical(self, "Napaka", f"Napaka pri izračunu: {str(e)}")
                    print(f"Error in on_dates_selected: {str(e)}")
            self.show_calendar_dialog(on_dates_selected)
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri začetku izračuna: {str(e)}")
            print(f"Error in calculate_employee_hours: {str(e)}")

    def show_search_dialog(self):
        dialog = SearchEmployeeDialog(self)
        dialog.exec()

    def show_calendar(self, card_id):
        """Show the calendar dialog for a specific employee"""
        dialog = CalendarDialog(self, card_id)
        dialog.exec()

    def show_change_card_id_dialog(self, card_id, worker_name):
        """Show dialog to change worker's card ID"""
        dialog = ChangeCardIDDialog(worker_name, card_id, self)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.new_card_id:
            self.change_worker_card_id(card_id, dialog.new_card_id, worker_name)

    def export_configuration(self):
        """Export all data (employees, groups, special_days) to a JSON file - excludes settings"""
        try:
            # Get save file path
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Izvozi vse podatke",
                "",
                "JSON Files (*.json)"
            )
            if not file_path:
                return

            config = {}
            
            # Export groups
            self.cursor.execute("SELECT id, name FROM groups")
            groups = [{"id": row[0], "name": row[1]} for row in self.cursor.fetchall()]
            config["groups"] = groups
            
            # Export employees
            self.cursor.execute("""
                SELECT name, card_id, daily_hours, group_id
                FROM employees
            """)
            employees = [
                {
                    "name": row[0],
                    "card_id": row[1],
                    "daily_hours": row[2],
                    "group_id": row[3]
                }
                for row in self.cursor.fetchall()
            ]
            config["employees"] = employees
            
            # Export special_days
            self.cursor.execute("""
                SELECT card_id, date, type
                FROM special_days
            """)
            special_days = [
                {
                    "card_id": row[0],
                    "date": row[1],
                    "type": row[2]
                }
                for row in self.cursor.fetchall()
            ]
            config["special_days"] = special_days

            # Write to file
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)

            QMessageBox.information(self, "Uspeh", "Vsi podatki so bili uspešno izvoženi!")
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri izvozu podatkov: {str(e)}")
            print(f"Error in export_configuration: {str(e)}")

    def import_configuration(self):
        """Import all data (employees, groups, special_days) from a JSON file - excludes settings"""
        try:
            # Get file path
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Uvozi vse podatke",
                "",
                "JSON Files (*.json)"
            )
            
            if not file_path:
                return
            
            # Read configuration
            with open(file_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # Confirm import
            reply = QMessageBox.question(
                self,
                "Potrditev uvoza",
                "Uvoz vseh podatkov bo prepisal obstoječe podatke (zaposleni, skupine, posebni dnevi). Ali želite nadaljevati?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.No:
                return
            
            # Start transaction
            self.cursor.execute("BEGIN TRANSACTION")
            
            try:
                # Clear existing data
                self.cursor.execute("DELETE FROM special_days")
                self.cursor.execute("DELETE FROM employees")
                self.cursor.execute("DELETE FROM groups")
                
                # Import groups
                if "groups" in config:
                    for group in config["groups"]:
                        self.cursor.execute(
                            "INSERT INTO groups (id, name) VALUES (?, ?)",
                            (group["id"], group["name"])
                        )
                
                # Import employees
                if "employees" in config:
                    for employee in config["employees"]:
                        self.cursor.execute("""
                            INSERT INTO employees (name, card_id, daily_hours, group_id)
                            VALUES (?, ?, ?, ?)
                        """, (
                            employee["name"],
                            employee["card_id"],
                            employee["daily_hours"],
                            employee["group_id"]
                        ))
                
                # Import special_days if present
                if "special_days" in config:
                    special_days_data = [
                        (sd["card_id"], sd["date"], sd["type"]) for sd in config["special_days"]
                    ]
                    if special_days_data:
                        self.cursor.executemany(
                            "INSERT INTO special_days (card_id, date, type) VALUES (?, ?, ?)",
                            special_days_data
                        )
                
                # Commit transaction
                self.conn.commit()
                
                # Update UI
                self.update_groups_list()
                self.update_employee_table()
                
                # Update worker_id.csv file
                self.update_worker_id_file()
                
                QMessageBox.information(self, "Uspeh", "Vsi podatki so bili uspešno uvoženi!")
                
            except Exception as e:
                # Rollback on error
                self.conn.rollback()
                raise e
            
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri uvozu konfiguracije: {str(e)}")
            print(f"Error in import_configuration: {str(e)}")

    def read_smb_files(self, start_date, end_date):
        """Read CSV files from SMB share for the given date range"""
        try:
            # Get settings from config file
            config = configparser.ConfigParser()
            if not os.path.exists('config.ini'):
                raise ValueError("Nastavitve niso bile najdene. Najprej nastavite SMB povezavo.")
            
            config.read('config.ini')
            if 'SMB' not in config:
                raise ValueError("Nastavitve SMB niso bile najdene. Najprej nastavite SMB povezavo.")
            
            smb_path = config['SMB'].get('path', '')
            if not smb_path.startswith('\\\\'):
                raise ValueError("Neveljavna SMB pot")
            
            server_name = smb_path.split('\\')[2]
            share_name = smb_path.split('\\')[3]
            
            # Get authentication settings
            username = config['SMB'].get('username', '')
            password = config['SMB'].get('password', '')
            
            # Connect to SMB share
            conn = SMBConnection(username, password, "CLIENT", server_name, use_ntlm_v2=True)
            connected = conn.connect(server_name, 445)
            
            if not connected:
                raise ConnectionError("Ni mogoče vzpostaviti povezave s strežnikom")
            
            # Get list of files
            files = []
            current_date = start_date
            while current_date <= end_date:
                filename = f"time_records_{current_date.strftime('%Y%m%d')}.csv"
                files.append(filename)
                current_date += timedelta(days=1)
            
            # Read and process files
            all_data = []
            for file in files:
                try:
                    # Read file content
                    file_obj = io.BytesIO()
                    conn.retrieveFile(share_name, file, file_obj)
                    file_obj.seek(0)
                    
                    # Convert to DataFrame with column names
                    df = pd.read_csv(file_obj, names=['CardID', 'Timestamp', 'Status'])
                    if not df.empty:
                        # Parse timestamps - they might be in different formats
                        # Try full datetime first, then just time format
                        try:
                            df['Timestamp'] = pd.to_datetime(df['Timestamp'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
                        except:
                            # If that fails, try just time format (HH:MM:SS)
                            df['Timestamp'] = pd.to_datetime(df['Timestamp'], format='%H:%M:%S', errors='coerce')
                        print(f"Successfully read file {file} with {len(df)} rows")
                        all_data.append(df)
                except Exception as e:
                    print(f"Napaka pri branju datoteke {file}: {str(e)}")
                    continue
            
            if not all_data:
                # Return empty DataFrame instead of raising error
                # This allows the WorktimeEditDialog to open even when no files exist
                return pd.DataFrame(columns=['CardID', 'Timestamp', 'Status'])
            
            combined_data = pd.concat(all_data, ignore_index=True)
            print(f"Combined data shape: {combined_data.shape}")
            return combined_data
            
        except Exception as e:
            print(f"Error in read_smb_files: {str(e)}")
            # Don't show critical error dialog for missing files - just return empty DataFrame
            # This allows the WorktimeEditDialog to open even when SMB connection fails
            return pd.DataFrame(columns=['CardID', 'Timestamp', 'Status'])
        finally:
            if 'conn' in locals():
                conn.close()

    def update_worker_id_file(self):
        """Update the worker_id.csv file in the SMB share with current employee data"""
        try:
            # Get settings from config file
            config = configparser.ConfigParser()
            if not os.path.exists('config.ini'):
                raise ValueError("Nastavitve niso bile najdene. Najprej nastavite SMB povezavo.")
            
            config.read('config.ini')
            if 'SMB' not in config:
                raise ValueError("Nastavitve SMB niso bile najdene. Najprej nastavite SMB povezavo.")
            
            smb_path = config['SMB'].get('path', '')
            if not smb_path.startswith('\\'):
                raise ValueError("Neveljavna SMB pot")
            
            server_name = smb_path.split('\\')[2]
            share_name = smb_path.split('\\')[3]
            
            # Get authentication settings
            username = config['SMB'].get('username', '')
            password = config['SMB'].get('password', '')
            
            # Connect to SMB share
            conn = SMBConnection(username, password, "CLIENT", server_name, use_ntlm_v2=True)
            connected = conn.connect(server_name, 445)
            
            if not connected:
                raise ConnectionError("Ni mogoče vzpostaviti povezave s strežnikom")
            
            # Get all employees (remove S.P. logic)
            self.cursor.execute("SELECT name, card_id FROM employees ORDER BY name")
            employees = self.cursor.fetchall()
            
            # Create CSV lines
            lines = []
            for name, card_id in employees:
                line = f"{name};{card_id}"
                lines.append(line)
            csv_content = "\n".join(lines)
            
            # Save to SMB share
            file_obj = io.BytesIO(csv_content.encode('utf-8'))
            conn.storeFile(share_name, "worker_id.csv", file_obj)
            
            print("Successfully updated worker_id.csv file")
            
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri posodabljanju datoteke worker_id.csv: {str(e)}")
            print(f"Error in update_worker_id_file: {str(e)}")
        finally:
            if 'conn' in locals():
                conn.close()

    def update_card_id_in_time_records(self, old_card_id, new_card_id):
        """Update card ID in all time_records CSV files"""
        try:
            # Get settings from config file
            config = configparser.ConfigParser()
            if not os.path.exists('config.ini'):
                raise ValueError("Nastavitve niso bile najdene. Najprej nastavite SMB povezavo.")
            
            config.read('config.ini')
            if 'SMB' not in config:
                raise ValueError("Nastavitve SMB niso bile najdene. Najprej nastavite SMB povezavo.")
            
            smb_path = config['SMB'].get('path', '')
            if not smb_path.startswith('\\\\'):
                raise ValueError("Neveljavna SMB pot")
            
            server_name = smb_path.split('\\')[2]
            share_name = smb_path.split('\\')[3]
            
            # Get authentication settings
            username = config['SMB'].get('username', '')
            password = config['SMB'].get('password', '')
            
            # Connect to SMB share
            conn = SMBConnection(username, password, "CLIENT", server_name, use_ntlm_v2=True)
            connected = conn.connect(server_name, 445)
            
            if not connected:
                raise ConnectionError("Ni mogoče vzpostaviti povezave s strežnikom")
            
            # Get list of all time_records files
            try:
                files = conn.listPath(share_name, '/')
                time_record_files = [f.filename for f in files if f.filename.startswith('time_records_') and f.filename.endswith('.csv')]
            except Exception as e:
                print(f"Error listing files: {str(e)}")
                time_record_files = []
            
            updated_files = 0
            for filename in time_record_files:
                try:
                    # Read file content
                    file_obj = io.BytesIO()
                    conn.retrieveFile(share_name, filename, file_obj)
                    file_obj.seek(0)
                    
                    # Read CSV content
                    content = file_obj.read().decode('utf-8')
                    
                    # Replace old card ID with new one
                    if old_card_id in content:
                        content = content.replace(old_card_id, new_card_id)
                        
                        # Write back to file
                        file_obj = io.BytesIO(content.encode('utf-8'))
                        conn.storeFile(share_name, filename, file_obj)
                        updated_files += 1
                        print(f"Updated {filename}")
                    
                except Exception as e:
                    print(f"Error updating {filename}: {str(e)}")
                    continue
            
            print(f"Successfully updated {updated_files} time_records files")
            
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri posodabljanju time_records datotek: {str(e)}")
            print(f"Error in update_card_id_in_time_records: {str(e)}")
        finally:
            if 'conn' in locals():
                conn.close()

    def change_worker_card_id(self, old_card_id, new_card_id, worker_name):
        """Change worker's card ID in database and all related files"""
        try:
            # Update database
            self.cursor.execute("UPDATE employees SET card_id = ? WHERE card_id = ?", (new_card_id, old_card_id))
            self.conn.commit()
            
            # Update special_days table
            self.cursor.execute("UPDATE special_days SET card_id = ? WHERE card_id = ?", (new_card_id, old_card_id))
            self.conn.commit()
            
            # Update worker_id.csv file
            self.update_worker_id_file()
            
            # Update all time_records files
            self.update_card_id_in_time_records(old_card_id, new_card_id)
            
            # Update the employee table display
            self.update_employee_table()
            
            QMessageBox.information(self, "Uspeh", f"ID kartice za delavca {worker_name} je bil uspešno spremenjen!")
            
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri spreminjanju ID-ja kartice: {str(e)}")
            print(f"Error in change_worker_card_id: {str(e)}")

    def calculate_working_hours(self, card_id, start_date, end_date):
        """Calculate working hours for a specific card ID in the given date range"""
        try:
            data = self.read_smb_files(start_date, end_date)
            if data is None:
                return None
            
            # Filter data for the specific card
            card_data = data[data['CardID'] == card_id].copy()
            
            # Don't show warning if the date is in the future
            if card_data.empty and end_date > datetime.now().date():
                return pd.DataFrame(columns=['Datum', 'Prihod na delo', 'Izhod iz dela', 'Delovne ure', 'Status'])
            
            if card_data.empty:
                QMessageBox.warning(self, "Opozorilo", "Ni podatkov za izbranega delavca v tem obdobju.")
                return None
            
            # Get employee's daily hours
            self.cursor.execute("SELECT daily_hours FROM employees WHERE card_id = ?", (card_id,))
            daily_hours = self.cursor.fetchone()[0]
            is_flexible = daily_hours == -1
            
            # Get special days for the date range
            self.cursor.execute("""
                SELECT date, type 
                FROM special_days 
                WHERE card_id = ? 
                AND date BETWEEN ? AND ?
            """, (card_id, start_date, end_date))
            special_days = {}
            for row in self.cursor.fetchall():
                date = row[0]
                if isinstance(date, str):
                    date = datetime.strptime(date, '%Y-%m-%d').date()
                # Translate special day types to Slovenian
                day_type = row[1]
                if day_type == 'vacation':
                    day_type = 'Dopust'
                elif day_type == 'sick_leave':
                    day_type = 'Bolniški stalež'
                special_days[date] = day_type
            
            # Convert timestamp to datetime
            card_data['Timestamp'] = pd.to_datetime(card_data['Timestamp'])
            
            # Group by date and calculate hours
            daily_hours_list = []
            for date, group in card_data.groupby(card_data['Timestamp'].dt.date):
                # Sort by timestamp
                group = group.sort_values('Timestamp')
                
                # Initialize variables for this day
                total_hours = 0
                current_entry = None
                first_entry = None
                last_exit = None
                
                # Process each record
                for _, row in group.iterrows():
                    if row['Status'] == 'Prihod na delo':
                        if current_entry is None:  # Only set if not already in work
                            current_entry = row['Timestamp']
                            if first_entry is None:  # Store first entry of the day
                                first_entry = row['Timestamp']
                    elif row['Status'] == 'Izhod iz dela':
                        if current_entry is not None:  # Only calculate if there was an entry
                            hours = (row['Timestamp'] - current_entry).total_seconds() / 3600
                            total_hours += hours
                            current_entry = None
                            last_exit = row['Timestamp']  # Update last exit time
                
                # Check if this is a special day
                if date in special_days:
                    status = special_days[date]
                    total_hours = 0  # Special days are counted as non-working days
                else:
                    # Determine status for normal days
                    if total_hours == 0:
                        status = 'Nepopolni podatki'
                    elif is_flexible:
                        status = 'Gibljivi delovni čas'
                    elif total_hours > daily_hours:
                        status = f'Nadure ({round(total_hours - daily_hours, 2)} ur)'
                    elif total_hours == daily_hours:
                        status = 'Normalno'
                    else:
                        status = f'Manjko ur ({round(daily_hours - total_hours, 2)} ur)'
                
                daily_hours_list.append({
                    'Datum': date,
                    'Prihod na delo': first_entry.strftime('%H:%M:%S') if first_entry else '',
                    'Izhod iz dela': last_exit.strftime('%H:%M:%S') if last_exit else '',
                    'Delovne ure': round(total_hours, 2),
                    'Status': status
                })
            
            # Add special days that don't have any records
            for date, day_type in special_days.items():
                if date not in [d['Datum'] for d in daily_hours_list]:
                    daily_hours_list.append({
                        'Datum': date,
                        'Prihod na delo': '',
                        'Izhod iz dela': '',
                        'Delovne ure': 0,
                        'Status': day_type
                    })
            
            result_df = pd.DataFrame(daily_hours_list)
            if result_df.empty:
                QMessageBox.warning(self, "Opozorilo", "Ni podatkov za izračun.")
                return None
                
            # Sort by date and ensure proper chronological order
            result_df['Datum'] = pd.to_datetime(result_df['Datum'])
            result_df = result_df.sort_values('Datum')
            result_df['Datum'] = result_df['Datum'].dt.date
            return result_df
            
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri izračunu delovnih ur: {str(e)}")
            print(f"Error in calculate_working_hours: {str(e)}")
            import traceback
            print(traceback.format_exc())
            return None

    def search_employee_by_name(self):
        """Highlight/select all rows where the worker's name contains the search text (case-insensitive)"""
        search_text = self.name_search_box.text().strip().lower()
        if not search_text:
            return
        self.employee_table.clearSelection()
        for row in range(self.employee_table.rowCount()):
            name_item = self.employee_table.item(row, 0)
            if name_item and search_text in name_item.text().lower():
                self.employee_table.selectRow(row)

    def archive_worker_data(self, card_id, worker_name):
        """Archive worker data by collecting timestamps from SMB CSV files"""
        try:
            # Show archive dialog for date selection
            dialog = ArchiveDialog(worker_name, self)
            if dialog.exec() != QDialog.DialogCode.Accepted:
                return
            
            start_date, end_date = dialog.get_dates()
            
            # Validate date range
            if start_date > end_date:
                QMessageBox.warning(self, "Napaka", "Začetni datum mora biti pred končnim datumom.")
                return
            
            # Show folder selection dialog
            folder_path = QFileDialog.getExistingDirectory(
                self, 
                "Izberi mapo za shranjevanje arhiviranih podatkov",
                "",
                QFileDialog.Option.ShowDirsOnly
            )
            
            if not folder_path:
                return
            
            # Collect timestamps for the worker
            QMessageBox.information(self, "Arhiviranje", "Začnem z arhiviranjem podatkov...")
            
            # Read SMB files for the date range
            all_data = self.read_smb_files(start_date, end_date)
            
            if all_data.empty:
                QMessageBox.warning(self, "Napaka", "Ni podatkov za izbrano obdobje.")
                return
            
            # Filter data for the specific worker
            worker_data = all_data[all_data['CardID'] == card_id].copy()
            
            if worker_data.empty:
                QMessageBox.warning(self, "Napaka", f"Ni podatkov za delavca {worker_name} v izbranem obdobju.")
                return
            
            # Sort by timestamp
            worker_data = worker_data.sort_values('Timestamp')
            
            # Create filename with worker name and date range
            safe_worker_name = "".join(c for c in worker_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            filename = f"arhiv_{safe_worker_name}_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.csv"
            file_path = os.path.join(folder_path, filename)
            
            # Export to CSV
            worker_data.to_csv(file_path, index=False, encoding='utf-8')
            
            QMessageBox.information(
                self, 
                "Uspešno arhiviranje", 
                f"Podatki za delavca {worker_name} so bili uspešno arhivirani v datoteko:\n{file_path}\n\nŠtevilo zapisov: {len(worker_data)}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri arhiviranju podatkov: {str(e)}")
            print(f"Error in archive_worker_data: {str(e)}")
            import traceback
            print(traceback.format_exc())

    def delete_worker_timestamps(self, card_id, worker_name):
        """Delete worker timestamps from SMB CSV files for a specific period"""
        try:
            # Show delete dialog for date selection
            dialog = DeleteTimestampsDialog(worker_name, self)
            if dialog.exec() != QDialog.DialogCode.Accepted:
                return
            
            start_date, end_date = dialog.get_dates()
            
            # Validate date range
            if start_date > end_date:
                QMessageBox.warning(self, "Napaka", "Začetni datum mora biti pred končnim datumom.")
                return
            
            # Show confirmation dialog
            reply = QMessageBox.question(
                self, 
                "Potrditev brisanja", 
                "Ali ste prepričani, da želite izbrisati evidence delavca iz smb shrambe?\n\nPriporočeno je arhiviranje.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply != QMessageBox.StandardButton.Yes:
                return
            
            # Read SMB files for the date range
            all_data = self.read_smb_files(start_date, end_date)
            
            if all_data.empty:
                QMessageBox.warning(self, "Napaka", "Ni podatkov za izbrano obdobje.")
                return
            
            # Filter data for the specific worker
            worker_data = all_data[all_data['CardID'] == card_id].copy()
            
            if worker_data.empty:
                QMessageBox.warning(self, "Napaka", f"Ni podatkov za delavca {worker_name} v izbranem obdobju.")
                return
            
            # Delete timestamps from SMB files
            deleted_count = self.delete_timestamps_from_smb(card_id, start_date, end_date)
            
            QMessageBox.information(
                self, 
                "Uspešno brisanje", 
                f"Evidence delavca {worker_name} so bile uspešno izbrisane iz SMB shrambe za obdobje {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}.\n\nŠtevilo izbrisanih zapisov: {deleted_count}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri brisanju evidence: {str(e)}")
            print(f"Error in delete_worker_timestamps: {str(e)}")
            import traceback
            print(traceback.format_exc())

    def delete_timestamps_from_smb(self, card_id, start_date, end_date):
        """Delete timestamps for specific worker from SMB CSV files using the same method as individual deletion"""
        try:
            print(f"Starting SMB deletion for card_id: {card_id}, period: {start_date} to {end_date}")
            
            # Process each day in the date range
            current_date = start_date
            files_processed = 0
            records_deleted = 0
            
            while current_date <= end_date:
                print(f"Processing date: {current_date}")
                
                try:
                    # Use the same method as individual deletion - read the data first
                    df = self.read_smb_files(current_date, current_date)
                    
                    if df is not None and not df.empty:
                        print(f"Loaded {len(df)} records for {current_date}")
                        print(f"DataFrame columns: {df.columns.tolist()}")
                        
                        if 'CardID' in df.columns:
                            # Check what CardID values exist
                            unique_card_ids = df['CardID'].unique()
                            print(f"Unique CardIDs in file: {unique_card_ids}")
                            print(f"Looking for card_id: {card_id} (type: {type(card_id)})")
                            
                            # Filter out the worker's records
                            original_count = len(df)
                            df_filtered = df[df['CardID'].astype(str) != str(card_id)]
                            new_count = len(df_filtered)
                            deleted_count = original_count - new_count
                            
                            print(f"Original records: {original_count}, After filtering: {new_count}, Deleted: {deleted_count}")
                            
                            if deleted_count > 0:
                                # Use the same update method as individual deletion
                                if not df_filtered.empty:
                                    print(f"Updating CSV file with remaining entries...")
                                    self.update_csv_file_for_date(df_filtered, current_date)
                                else:
                                    print(f"No entries left, deleting CSV file...")
                                    self.delete_csv_file_for_date(current_date)
                                
                                files_processed += 1
                                records_deleted += deleted_count
                                print(f"Successfully processed {current_date}: removed {deleted_count} records")
                            else:
                                print(f"No records found for card_id {card_id} on {current_date}")
                        else:
                            print(f"File for {current_date} is missing CardID column")
                    else:
                        print(f"No data found for {current_date}")
                
                except Exception as e:
                    print(f"Error processing {current_date}: {str(e)}")
                    import traceback
                    print(traceback.format_exc())
                    # Continue with other dates even if one fails
                
                current_date += timedelta(days=1)
            
            print(f"SMB deletion completed. Files processed: {files_processed}, Records deleted: {records_deleted}")
            return records_deleted
            
        except Exception as e:
            print(f"Error in delete_timestamps_from_smb: {str(e)}")
            import traceback
            print(traceback.format_exc())
            raise

    def update_csv_file_for_date(self, df, date):
        """Update CSV file for specific date using the same method as individual deletion"""
        try:
            import configparser
            import io
            import tempfile
            import os
            from smb.SMBConnection import SMBConnection
            
            print(f"Updating CSV file for date: {date}")
            print(f"DataFrame shape: {df.shape}")
            
            # Load configuration
            config = configparser.ConfigParser()
            config.read('config.ini')
            
            if 'SMB' not in config:
                raise ValueError("Nastavitve SMB niso bile najdene.")
            
            smb_path = config['SMB'].get('path', '')
            if not smb_path.startswith('\\\\'):
                raise ValueError("Neveljavna SMB pot")
            
            # Parse SMB path more safely
            path_parts = smb_path.split('\\')
            if len(path_parts) < 4:
                raise ValueError("Neveljavna SMB pot - manjkajo deli poti")
            
            server_name = path_parts[2]
            share_name = path_parts[3]
            username = config['SMB'].get('username', '')
            password = config['SMB'].get('password', '')
            
            # Connect to SMB share
            conn = SMBConnection(username, password, "CLIENT", server_name, use_ntlm_v2=True)
            connected = conn.connect(server_name, 445)
            
            if not connected:
                raise ConnectionError(f"Ni mogoče vzpostaviti povezave s strežnikom {server_name}")
            
            # Create CSV content (same format as individual deletion)
            csv_content = df.to_csv(index=False, header=False)
            
            # Save to SMB share
            filename = f"time_records_{date.strftime('%Y%m%d')}.csv"
            
            # Try multiple approaches to handle permission issues (same as individual deletion)
            success = False
            
            # Method 1: Direct file storage
            try:
                file_obj = io.BytesIO(csv_content.encode('utf-8'))
                conn.storeFile(share_name, filename, file_obj)
                print("File saved successfully using direct method!")
                success = True
            except Exception as e:
                print(f"Direct method failed: {str(e)}")
            
            # Method 2: Try deleting existing file first, then create new one
            if not success:
                try:
                    print("Trying to delete existing file first...")
                    try:
                        conn.deleteFiles(share_name, filename)
                        print("Existing file deleted successfully")
                    except:
                        print("No existing file to delete or deletion failed")
                    
                    # Now try to create new file
                    file_obj = io.BytesIO(csv_content.encode('utf-8'))
                    conn.storeFile(share_name, filename, file_obj)
                    print("File saved successfully using delete-first method!")
                    success = True
                except Exception as e:
                    print(f"Delete-first method failed: {str(e)}")
            
            # Method 3: Try with temporary filename
            if not success:
                try:
                    temp_filename = f"temp_{filename}"
                    print(f"Trying with temporary filename: {temp_filename}")
                    
                    file_obj = io.BytesIO(csv_content.encode('utf-8'))
                    conn.storeFile(share_name, temp_filename, file_obj)
                    print("Temporary file created successfully")
                    
                    # Try to delete original and rename temp
                    try:
                        conn.deleteFiles(share_name, filename)
                    except:
                        pass  # Ignore if original doesn't exist
                    
                    # Note: SMB doesn't have a direct rename, so we'll keep the temp file
                    # and try to delete it, then create the final file
                    try:
                        temp_file_obj = io.BytesIO()
                        conn.retrieveFile(share_name, temp_filename, temp_file_obj)
                        temp_file_obj.seek(0)
                        
                        final_file_obj = io.BytesIO(temp_file_obj.read())
                        conn.storeFile(share_name, filename, final_file_obj)
                        conn.deleteFiles(share_name, temp_filename)
                        print("File saved successfully using temporary file method!")
                        success = True
                    except Exception as rename_e:
                        # Keep the temp file as fallback
                        print(f"Rename failed, keeping temporary file: {rename_e}")
                        success = True  # Consider this a success since temp file exists
                except Exception as e:
                    print(f"Temporary file method failed: {str(e)}")
            
            conn.close()
            
            if not success:
                # Try to save locally as backup
                try:
                    local_filename = f"backup_{filename}"
                    df.to_csv(local_filename, index=False, header=False)
                    raise Exception(f"NAPAKA DOVOLJENJ: Aplikacija nima dovoljenja za pisanje v SMB mapo. Podatki so bili shranjeni lokalno kot varnostna kopija v {local_filename}. Prosimo kontaktirajte sistemskega administratorja za nastavitev ustreznih dovoljenj.")
                except Exception as backup_e:
                    print(f"Backup save also failed: {backup_e}")
                    raise Exception(f"NAPAKA DOVOLJENJ: Aplikacija nima dovoljenja za pisanje v SMB mapo. Varnostna kopija ni mogla biti shranjena. Prosimo kontaktirajte sistemskega administratorja za nastavitev ustreznih dovoljenj.")
                
        except Exception as e:
            print(f"Error in update_csv_file_for_date: {str(e)}")
            raise

    def delete_csv_file_for_date(self, date):
        """Delete CSV file for specific date using the same method as individual deletion"""
        try:
            import configparser
            from smb.SMBConnection import SMBConnection
            
            # Load configuration
            config = configparser.ConfigParser()
            config.read('config.ini')
            
            if 'SMB' not in config:
                raise ValueError("Nastavitve SMB niso bile najdene.")
            
            smb_path = config['SMB'].get('path', '')
            if not smb_path.startswith('\\\\'):
                raise ValueError("Neveljavna SMB pot")
            
            # Parse SMB path more safely
            path_parts = smb_path.split('\\')
            if len(path_parts) < 4:
                raise ValueError("Neveljavna SMB pot - manjkajo deli poti")
            
            server_name = path_parts[2]
            share_name = path_parts[3]
            username = config['SMB'].get('username', '')
            password = config['SMB'].get('password', '')
            
            # Connect to SMB share
            conn = SMBConnection(username, password, "CLIENT", server_name, use_ntlm_v2=True)
            connected = conn.connect(server_name, 445)
            
            if not connected:
                raise ConnectionError(f"Ni mogoče vzpostaviti povezave s strežnikom {server_name}")
            
            # Delete the file
            filename = f"time_records_{date.strftime('%Y%m%d')}.csv"
            try:
                conn.deleteFiles(share_name, filename)
                print(f"File {filename} deleted successfully")
            except Exception as e:
                print(f"Could not delete file {filename}: {str(e)}")
                # This is not necessarily an error - file might not exist
            finally:
                conn.close()
                
        except Exception as e:
            print(f"Error in delete_csv_file_for_date: {str(e)}")
            raise

    def get_smb_connection(self):
        """Get SMB connection with error handling"""
        try:
            config = configparser.ConfigParser()
            if not os.path.exists('config.ini'):
                return None, None
            
            config.read('config.ini')
            if 'SMB' not in config:
                return None, None
            
            smb_path = config['SMB'].get('path', '')
            if not smb_path.startswith('\\\\'):
                return None, None
            
            server_name = smb_path.split('\\')[2]
            share_name = smb_path.split('\\')[3]
            username = config['SMB'].get('username', '')
            password = config['SMB'].get('password', '')
            
            conn = SMBConnection(username, password, "CLIENT", server_name, use_ntlm_v2=True)
            connected = conn.connect(server_name, 445)
            
            if connected:
                return conn, share_name
            else:
                return None, None
                
        except Exception as e:
            print(f"Error creating SMB connection: {str(e)}")
            return None, None

    def get_data_version(self):
        """Get current data version for conflict detection"""
        try:
            conn, share_name = self.get_smb_connection()
            if not conn:
                return 1
            
            file_obj = io.BytesIO()
            try:
                conn.retrieveFile(share_name, "data_version.txt", file_obj)
                file_obj.seek(0)
                version = int(file_obj.read().decode('utf-8').strip())
            except:
                version = 1
            
            conn.close()
            return version
            
        except Exception as e:
            print(f"Error getting data version: {str(e)}")
            return 1

    def save_shared_data_to_smb(self):
        """Save all shared data to SMB share in JSON format"""
        try:
            conn, share_name = self.get_smb_connection()
            if not conn:
                print("No SMB connection available")
                return False
            
            # Get all data from local database
            self.cursor.execute("SELECT * FROM employees")
            employees = [dict(zip([col[0] for col in self.cursor.description], row)) 
                        for row in self.cursor.fetchall()]
            
            self.cursor.execute("SELECT * FROM groups")
            groups = [dict(zip([col[0] for col in self.cursor.description], row)) 
                     for row in self.cursor.fetchall()]
            
            self.cursor.execute("SELECT * FROM special_days")
            special_days = [dict(zip([col[0] for col in self.cursor.description], row)) 
                           for row in self.cursor.fetchall()]
            
            # Create shared data structure
            shared_data = {
                "employees": employees,
                "groups": groups,
                "special_days": special_days,
                "last_updated": datetime.now().isoformat(),
                "version": self.data_version + 1
            }
            
            # Save to SMB as JSON
            json_data = json.dumps(shared_data, ensure_ascii=False, indent=2)
            file_obj = io.BytesIO(json_data.encode('utf-8'))
            conn.storeFile(share_name, "shared_data.json", file_obj)
            
            # Update version file
            version_data = str(shared_data["version"])
            version_obj = io.BytesIO(version_data.encode('utf-8'))
            conn.storeFile(share_name, "data_version.txt", version_obj)
            
            # Update local version
            self.data_version = shared_data["version"]
            
            conn.close()
            print(f"Successfully saved shared data (version {self.data_version})")
            return True
            
        except Exception as e:
            print(f"Error saving shared data: {str(e)}")
            return False

    def load_shared_data_from_smb(self):
        """Load all shared data from SMB share"""
        try:
            conn, share_name = self.get_smb_connection()
            if not conn:
                print("No SMB connection available for loading shared data")
                return False
            
            # Load shared data
            file_obj = io.BytesIO()
            try:
                conn.retrieveFile(share_name, "shared_data.json", file_obj)
                file_obj.seek(0)
                shared_data = json.loads(file_obj.read().decode('utf-8'))
            except:
                print("No shared data file found, using local data only")
                conn.close()
                return False
            
            # Clear local database
            self.cursor.execute("DELETE FROM special_days")
            self.cursor.execute("DELETE FROM employees")
            self.cursor.execute("DELETE FROM groups")
            
            # Insert groups first (due to foreign key constraints)
            for group in shared_data.get("groups", []):
                self.cursor.execute("""
                    INSERT OR REPLACE INTO groups (id, name) VALUES (?, ?)
                """, (group["id"], group["name"]))
            
            # Insert employees
            for employee in shared_data.get("employees", []):
                self.cursor.execute("""
                    INSERT OR REPLACE INTO employees (id, name, card_id, daily_hours, group_id) 
                    VALUES (?, ?, ?, ?, ?)
                """, (employee["id"], employee["name"], employee["card_id"], 
                      employee["daily_hours"], employee["group_id"]))
            
            # Insert special days
            for special_day in shared_data.get("special_days", []):
                self.cursor.execute("""
                    INSERT OR REPLACE INTO special_days (id, card_id, date, type) 
                    VALUES (?, ?, ?, ?)
                """, (special_day["id"], special_day["card_id"], 
                      special_day["date"], special_day["type"]))
            
            self.conn.commit()
            
            # Update local version
            self.data_version = shared_data.get("version", 1)
            self.last_known_version = self.data_version
            
            conn.close()
            print(f"Successfully loaded shared data (version {self.data_version})")
            return True
            
        except Exception as e:
            print(f"Error loading shared data: {str(e)}")
            return False

    def check_for_updates(self):
        """Check if shared data has been updated by other users (non-blocking)"""
        # Skip if an update check is already in progress
        if self.update_in_progress:
            print("Update check already in progress, skipping...")
            return
        
        # Skip if worker is still running
        if self.update_worker and self.update_worker.isRunning():
            print("Worker thread still running, skipping...")
            return
        
        try:
            # Create and start worker thread for background checking
            self.update_in_progress = True
            self.update_worker = SMBUpdateWorker(self)
            self.update_worker.update_available.connect(self.handle_update_available)
            self.update_worker.update_complete.connect(self.handle_update_complete)
            self.update_worker.start()
            
        except Exception as e:
            print(f"Error starting update check: {str(e)}")
            self.update_in_progress = False
    
    def handle_update_available(self, version):
        """Handle when an update has been loaded (runs on main thread)"""
        try:
            # Data has already been loaded by the worker thread
            # Just update the UI elements
            print(f"Main thread: Updating UI for version {version}")
            self.update_employee_table()
            self.update_groups_list()
            print(f"Main thread: UI update completed for version {version}")
        except Exception as e:
            print(f"Error updating UI after data load: {str(e)}")
    
    def handle_update_complete(self):
        """Handle when update check is complete"""
        self.update_in_progress = False

    def save_shared_data_with_retry(self, max_retries=3):
        """Save shared data with retry logic"""
        for attempt in range(max_retries):
            try:
                if self.save_shared_data_to_smb():
                    return True
            except Exception as e:
                print(f"Save attempt {attempt + 1} failed: {str(e)}")
                if attempt == max_retries - 1:
                    QMessageBox.critical(self, "Napaka", 
                        f"Ni mogoče shraniti podatkov po {max_retries} poskusih. "
                        "Preverite povezavo z SMB strežnikom.")
                    return False
                import time
                time.sleep(1)  # Wait before retry
        return False
    
    def closeEvent(self, event):
        """Clean up worker thread when closing the application"""
        try:
            # Stop the update timer
            if hasattr(self, 'update_timer'):
                self.update_timer.stop()
            
            # Stop and wait for the worker thread to finish
            if hasattr(self, 'update_worker') and self.update_worker:
                if self.update_worker.isRunning():
                    self.update_worker.stop()
                    self.update_worker.wait(2000)  # Wait up to 2 seconds
            
            # Close database connection
            if hasattr(self, 'conn'):
                self.conn.close()
        except Exception as e:
            print(f"Error during cleanup: {str(e)}")
        finally:
            event.accept()

class ChangeCardIDDialog(QDialog):
    def __init__(self, worker_name, old_card_id, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Spremeni ID kartice")
        self.setMinimumWidth(400)
        self.old_card_id = old_card_id
        self.new_card_id = None
        
        layout = QVBoxLayout(self)
        
        # Worker info
        info_label = QLabel(f"Delavec: {worker_name}")
        info_label.setStyleSheet("font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(info_label)
        
        # Current card ID
        current_label = QLabel(f"Trenutni ID kartice: {old_card_id}")
        layout.addWidget(current_label)
        
        # Instructions
        instruction_label = QLabel("Zamenjaj obstoječi ID kartice z naslednjim:")
        layout.addWidget(instruction_label)
        
        # New card ID input
        card_layout = QHBoxLayout()
        card_layout.addWidget(QLabel("Novi ID kartice:"))
        self.new_card_input = QLineEdit()
        self.new_card_input.setMaxLength(14)
        self.new_card_input.setPlaceholderText("Vnesite 14-mestno šestnajstiško število")
        card_layout.addWidget(self.new_card_input)
        layout.addLayout(card_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        change_button = QPushButton("Spremeni")
        cancel_button = QPushButton("Prekliči")
        
        change_button.clicked.connect(self.change_card_id)
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(change_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        # Set focus to input field
        self.new_card_input.setFocus()
        
    def change_card_id(self):
        """Validate and accept the new card ID"""
        new_id = self.new_card_input.text().strip().lower()
        
        # Validate input
        if not new_id:
            QMessageBox.warning(self, "Napaka", "Vnesite novi ID kartice!")
            return
        
        # Validate hexadecimal format
        if len(new_id) != 14 or not all(c in '0123456789abcdef' for c in new_id):
            QMessageBox.warning(self, "Napaka", "Neveljavna številka kartice! Vnesite 14-mestno šestnajstiško število.")
            return
        
        # Check if new ID already exists
        try:
            self.parent().cursor.execute("SELECT name FROM employees WHERE card_id = ? AND card_id != ?", 
                                       (new_id, self.old_card_id))
            if self.parent().cursor.fetchone():
                QMessageBox.warning(self, "Napaka", "Kartica s tem ID-jem že obstaja!")
                return
        except Exception as e:
            QMessageBox.critical(self, "Napaka", f"Napaka pri preverjanju ID-ja: {str(e)}")
            return
        
        # Confirm change
        reply = QMessageBox.question(
            self,
            "Potrditev",
            "Ali ste prepričani, da želite spremeniti ID kartice?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.new_card_id = new_id
            self.accept()
    
    def keyPressEvent(self, event):
        """Handle keyboard shortcuts"""
        if event.key() == Qt.Key.Key_F1:
            self.show_manual()
        elif event.key() == Qt.Key.Key_F5:
            self.manual_refresh()
        elif event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.key() == Qt.Key.Key_F:
                self.show_search_dialog()
            elif event.key() == Qt.Key.Key_S:
                self.save_configuration()
            elif event.key() == Qt.Key.Key_E:
                self.export_all_data()
        else:
            super().keyPressEvent(event)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
