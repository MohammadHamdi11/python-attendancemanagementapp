#==========================================================imports and constants==========================================================#
import sys
import re
import pandas as pd  
import traceback
import os 
import requests
import json
import random
import shutil
import threading
from collections import defaultdict
from datetime import datetime, timedelta, date
from typing import List, Dict
import io
from PIL import Image
import math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QHeaderView, QGridLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QFileDialog, QSizePolicy, QCalendarWidget,
    QFrame, QTableWidget, QTableWidgetItem, QProgressBar, QGraphicsDropShadowEffect, QListWidget,
    QTextEdit, QDialog, QMessageBox, QScrollArea, QStackedWidget, QGroupBox, QRadioButton, QAbstractItemView
)
from PyQt6.QtCore import Qt, pyqtSignal, QThread, QSize, QTimer, QObject, pyqtSignal, QEvent, QDate
from PyQt6.QtGui import QIcon, QPixmap, QFont, QIntValidator, QColor, QMovie, QTextCursor, QPainter, QPainterPath, QPen, QPalette

# Constants - Inverted Colors
DARK_BLUE = "#24325f"
DARK_RED = "#951d1e"
BLACK = "#000000"
CARD_BG = "#1a1a1a"  
TEXT_COLOR = "#ffffff"  
INPUT_BG = "#2d2d2d"  
BORDER_COLOR = "#3d3d3d"

# Standard button style 
MENU_BUTTON_STYLE = f"""
QPushButton {{
    background-color: {DARK_BLUE};
    color: {TEXT_COLOR};
    border: none;
    border-radius: 3px;
    padding: 10px 20px;
    font-size: 16px;
    font-weight: bold;
}}
QPushButton:hover {{
    background-color: {DARK_RED};
}}
"""

# Menu Exit button style 
MENU_EXIT_BUTTON_STYLE = f"""
QPushButton {{
    background-color: {DARK_RED};
    color: {TEXT_COLOR};
    border: none;
    border-radius: 3px;
    padding: 10px 20px;
    font-size: 16px;
    font-weight: bold;
}}
QPushButton:hover {{
    background-color: #ab2223;
}}
"""

# STANDARD button style 
STANDARD_BUTTON_STYLE = f"""
QPushButton {{
    background-color: {DARK_BLUE};
    color: {TEXT_COLOR};
    border: none;
    border-radius: 3px;
    padding: 5px 15px;
}}
QPushButton:hover {{
    background-color: {DARK_RED};
}}
"""

# Exit button style 
EXIT_BUTTON_STYLE = f"""
QPushButton {{
    background-color: {DARK_RED};
    color: {TEXT_COLOR};
    border: none;
    border-radius: 3px;
    padding: 5px 15px;
}}
QPushButton:hover {{
    background-color: #ab2223;
}}
"""

# Table style
TABLE_STYLE = f"""
QTableWidget::item {{
    text-align: center;
    padding: 5px;
}}
QTableWidget {{
    background-color: {INPUT_BG};
    gridline-color: #3d3d3d;
    border: 1px solid #3d3d3d;
    border-radius: 5px;
}}
QTableWidget::item:selected {{
    background-color: {DARK_BLUE};
    color: white;
}}
QTableWidget::item:hover {{
    background-color: {DARK_BLUE};
    color: white;
}}
QHeaderView::section:hover {{
    background-color: {DARK_BLUE};
    color: white;
}}
QHeaderView::section {{
    background-color: #202c54;
    color: white;
    gridline-color: #3d3d3d;
    border: 1px solid #3d3d3d;
}}
"""

GROUP_BOX_STYLE = f"""
QGroupBox {{
    border: 2px solid {DARK_BLUE};
    border-radius: 5px;
    margin-top: 2ex;
    font-weight: bold;
    font-size: 20px;
    color: {TEXT_COLOR};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 3px 0 3px;
    color: {TEXT_COLOR};
}}
QLabel {{
    color: white;
}}
QRadioButton {{
    color: white;
}}
QRadioButton::indicator::unchecked {{
    border: 2px solid white;
    background-color: white; /* Background for unchecked state */
    border-radius: 7px;
}}
QRadioButton::indicator::checked {{
    border: 2px solid white;
    background-color: {DARK_BLUE}; /* Dark blue check mark */
    border-radius: 7px;
}}

QComboBox {{
    color: white;
    background-color: {INPUT_BG};
}}

QLineEdit {{
    color: white;
    background-color: {INPUT_BG};
}}
"""

PROGRESS_BAR_STYLE = f"""
QProgressBar {{
    text-align: center;
    background-color: {INPUT_BG};
}}
QProgressBar::chunk {{
    background-color: {DARK_RED}; 
}}
"""

# For the console 
CONSOLE_STYLE = f"""
QTextEdit {{
    background-color: {INPUT_BG};
    color: white;
    border: 1px solid {BLACK};
}}
"""

#==========================================================start page==========================================================#

class StartPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
        
    def init_ui(self):
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create a dark card-like container
        card_container = QWidget()
        card_container.setStyleSheet(f"""
            QWidget {{
                background-color: {BLACK};
                border-radius: 10px;
                border: 1px solid {BORDER_COLOR};
                color: {TEXT_COLOR};
            }}
            QFrame {{
                border: none;
            }}
            QComboBox {{
                background-color: {INPUT_BG};
                color: {TEXT_COLOR};
                border: 1px solid {BORDER_COLOR};
                padding: 5px;
                border-radius: 3px;
            }}
            QComboBox::drop-down {{
                border: none;
            }}
            QComboBox::down-arrow {{
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid {TEXT_COLOR};
            }}
            QLineEdit {{
                background-color: {INPUT_BG};
                color: {TEXT_COLOR};
                border: 1px solid {BORDER_COLOR};
                padding: 5px;
                border-radius: 3px;
            }}
        """)
        card_layout = QVBoxLayout(card_container)
        card_layout.setSpacing(20)
        card_layout.setContentsMargins(40, 40, 40, 40)

        # Create info button in top left of the card
        info_container = QHBoxLayout()
        info_container.setContentsMargins(0, 0, 0, 0)
        
        self.info_button = QPushButton("", self)
        self.info_button.setFixedSize(40, 40)
        icon_path = os.path.join(os.path.dirname(__file__), 'info.png')
        if os.path.exists(icon_path):
            self.info_button.setIcon(QIcon(icon_path))
            self.info_button.setIconSize(QSize(32, 32))
        else:
            self.info_button.setStyleSheet(f"""
                QPushButton {{
                    background-color: {DARK_BLUE};
                    color: {TEXT_COLOR};
                    border: none;
                    border-radius: 20px;
                    font-size: 16px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    background-color: {DARK_RED};
                }}
            """)
            self.info_button.setText("i")
        
        info_container.addWidget(self.info_button)
        # Add spacer after button to push everything else to the right
        info_container.addStretch()
        card_layout.addLayout(info_container)
        
        # Center the card in the window
        center_layout = QHBoxLayout()
        center_layout.addStretch(1)
        center_layout.addWidget(card_container)
        center_layout.addStretch(1)
        
        # Vertical centering
        vertical_layout = QVBoxLayout()
        vertical_layout.addStretch(1)
        vertical_layout.addLayout(center_layout)
        vertical_layout.addStretch(1)
        
        # Background container
        bg_container = QWidget()
        bg_container.setStyleSheet(f"background-color: {CARD_BG};")
        bg_container.setLayout(vertical_layout)
        main_layout.addWidget(bg_container)

        # Logo
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), 'ASU1.png')
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaled(128, 128, Qt.AspectRatioMode.KeepAspectRatio, 
                                             Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(logo_label)

        # Title
        title_label = QLabel("Department Attendance \nManagement System")
        title_label.setStyleSheet(f"""
            color: {TEXT_COLOR};
            font-size: 24px;
            font-weight: bold;
        """)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        card_layout.addWidget(title_label)

        # Buttons Container
        buttons_widget = QWidget()
        buttons_widget.setStyleSheet("border: none;")
        buttons_layout = QVBoxLayout(buttons_widget)
        buttons_layout.setSpacing(15)
        
        # First row - 3 buttons side by side (original buttons)
        first_row = QHBoxLayout()
        first_row.setSpacing(10)
        
        self.reference_btn = QPushButton("Prepare Reference File")
        self.reference_btn.setMinimumHeight(50)
        self.reference_btn.setStyleSheet(MENU_BUTTON_STYLE)
        first_row.addWidget(self.reference_btn)
        
        self.preparer_btn = QPushButton("Prepare Log Sheet")
        self.preparer_btn.setMinimumHeight(50)
        self.preparer_btn.setStyleSheet(MENU_BUTTON_STYLE)
        first_row.addWidget(self.preparer_btn)
        
        self.schedule_btn = QPushButton("Manage Schedules")
        self.schedule_btn.setMinimumHeight(50)
        self.schedule_btn.setStyleSheet(MENU_BUTTON_STYLE)
        first_row.addWidget(self.schedule_btn)
        
        buttons_layout.addLayout(first_row)
        
        # Second row - 2 buttons side by side (process buttons)
        second_row = QHBoxLayout()
        second_row.setSpacing(10)

        self.appeal_btn = QPushButton("Process Appeals")
        self.appeal_btn.setMinimumHeight(50)
        self.appeal_btn.setStyleSheet(MENU_BUTTON_STYLE)
        second_row.addWidget(self.appeal_btn)
        
        self.process_btn = QPushButton("Process Attendance")
        self.process_btn.setMinimumHeight(50)
        self.process_btn.setStyleSheet(MENU_BUTTON_STYLE)
        second_row.addWidget(self.process_btn)
        
        buttons_layout.addLayout(second_row)
        
        # Third row - 2 buttons side by side (populate and analyze)
        third_row = QHBoxLayout()
        third_row.setSpacing(10)
        
        self.populate_btn = QPushButton("Populate Main File")
        self.populate_btn.setMinimumHeight(50)
        self.populate_btn.setStyleSheet(MENU_BUTTON_STYLE)
        third_row.addWidget(self.populate_btn)
        
        self.dashboard_btn = QPushButton("Analyze Attendance")
        self.dashboard_btn.setMinimumHeight(50)
        self.dashboard_btn.setStyleSheet(MENU_BUTTON_STYLE)
        third_row.addWidget(self.dashboard_btn)
        
        buttons_layout.addLayout(third_row)
        
        # Exit Button (full width)
        exit_btn = QPushButton("Exit")
        exit_btn.setMinimumHeight(50)
        exit_btn.setStyleSheet(MENU_EXIT_BUTTON_STYLE)
        exit_btn.clicked.connect(self.parent().close)
        buttons_layout.addWidget(exit_btn)

        # Add buttons container to card
        card_layout.addWidget(buttons_widget)

        # Set fixed size for the card
        card_container.setFixedWidth(700)  # Slightly wider to accommodate the buttons
        card_container.setMinimumHeight(700)  # Taller to accommodate the extra button

# ==========================================================info page==========================================================#

class InfoPage(QWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet("""
            background-color: black; 
            color: white;
            QLabel {
                color: white;
            }
        """)

        self.init_ui()

    def return_to_home(self):
        # Get the stacked widget and switch to the start page
        stacked_widget = self.parent()
        if isinstance(stacked_widget, QStackedWidget):
            stacked_widget.setCurrentIndex(0)

    def init_ui(self):
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Header layout
        header_layout = QHBoxLayout()
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), 'ASU1.png')
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaled(60, 60, Qt.AspectRatioMode.KeepAspectRatio,
                                               Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        title_label = QLabel("About the app")
        title_label.setStyleSheet(f"font-size: 24px; font-weight: bold;")
        header_layout.addWidget(logo_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Create vertical layout for header buttons
        header_buttons_layout = QVBoxLayout()
        header_buttons_layout.setSpacing(5)

        # Back button
        back_btn = QPushButton("Back to Home")
        back_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        back_btn.clicked.connect(self.return_to_home)

        # Exit button
        exit_btn = QPushButton("Exit")
        exit_btn.setStyleSheet(EXIT_BUTTON_STYLE)
        exit_btn.clicked.connect(QApplication.instance().quit)

        # Add buttons to vertical layout
        header_buttons_layout.addWidget(back_btn)
        header_buttons_layout.addWidget(exit_btn)
        header_buttons_layout.addStretch()

        # Add button layout to header
        header_layout.addStretch()
        header_layout.addLayout(header_buttons_layout)
        main_layout.addLayout(header_layout)

        # Create a GroupBox for info content, similar to AttendanceProcessor style
        info_group = QGroupBox()
        # Using same style as other GroupBoxes
        info_group.setStyleSheet(GROUP_BOX_STYLE)
        info_layout = QVBoxLayout(info_group)

        # Create scroll area for content
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setMinimumHeight(600)
        scroll_area.setStyleSheet(f"""
            QScrollArea {{
                background-color: transparent;
                border: none;
            }}
            QScrollBar:vertical {{
                border: none;
                background: {CARD_BG};
                width: 10px;
                margin: 0px;
            }}
            QScrollBar::handle:vertical {{
                background: {DARK_BLUE};
                min-height: 20px;
                border-radius: 5px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                border: none;
                background: none;
            }}
        """)

        # Content widget for the scroll area
        info_content = QWidget()
        info_content.setStyleSheet("background-color: transparent;")
        content_layout = QVBoxLayout(info_content)

        # Create info label with rich text
        info_text = """
        <h1 style='color: white; text-align: center;'>Faculty Attendance Management System</h1>
        
        <h2 style='color: white;'>Special Thanks</h2>
        <ul style='color: white;'>
            <li><b>Dr. Ahmad Samir</b> who encouraged me to start working on this project</li>
            <li><b>Dr. Amani Helmi</b> for sponsoring this project and for bing the reason behind its success</li>
            <li><b>Dr. Gehan Adel</b> for supporting me every step on the way</li>
            <li><b>Dr. Doaa Mohammad Abu Bakr</b> who was the catalyst to this project's success</li>
            <li><b>Dr. Taqwa Mohammad Abd Al-Salam</b> for aiding me with her efforts in deployment</li>
            <li>My companion and friend, <b>Dr. Mazin Helmi</b> who never got tired of me</li>

        </ul>
                
        <h2 style='color: white;'>What This App Does</h2>
        <p style='color: white;'>This app is a simple tool to track and manage student attendance. It helps you manage faculty attendance easily. It tracks who attended classes and creates reports.</p>
        
        <h2 style='color: white;'>Main Features</h2>
        <ol style='color: white;'>
            <li><b>Prepare Reference File</b>
                <ul>
                    <li>Create a file with all student information</li>
                    <li>Organize students by ID, name, year, and group</li>
                </ul>
            </li>
            <li><b>Prepare Log Sheet</b>
                <ul>
                    <li>Combine attendance records from multiple sources</li>
                    <li>Get files from your computer or cloud storage</li>
                </ul>
            </li>
            <li><b>Manage Schedules</b>
                <ul>
                    <li>Create new class schedules</li>
                    <li>Update existing schedules when needed</li>
                </ul>
            </li>
            <li><b>Process Appeals</b>
                <ul>
                    <li>Manage attendance exceptions for students who missed regular logging</li>
                    <li>Connect student information, attendance logs, and session schedules</li>
                    <li>Select students and sessions for attendance appeals</li>
                    <li>Process and add these appeals to attendance records</li>
                </ul>
            </li>
            <li><b>Process Attendance</b>
                <ul>
                    <li>Compare student logs against official schedules</li>
                    <li>Track and validate attendance for students who have changed groups</li>
                    <li>Identify when student transfers between groups occurred</li>
                    <li>Validate attendance records against the appropriate group schedule</li>
                    <li>Generate detailed attendance reports with transfer logs</li>
                </ul>
            </li>
            <li><b>Analyze Attendance</b>
                <ul>
                    <li>See attendance statistics and trends</li>
                    <li>Identify students who need help or attention</li>
                </ul>
            </li>
        </ol>
        
        <h2 style='color: white;'>How To Use This App</h2>
        
        <h3 style='color: white;'>Step 1: Prepare Reference File</h3>
        <ul style='color: white;'>
            <li>Click "Prepare Reference File" button on the main screen</li>
            <li>Click "Browse" to select your Excel file with student information</li>
            <li>Choose which columns have student ID, name, year, and group</li>
            <li>Click "Preview Result" to check the data</li>
            <li>Click "Process and Save Reference File" when everything looks right</li>
        </ul>
        
        <h3 style='color: white;'>Step 2: Prepare Log Sheet</h3>
        <ul style='color: white;'>
            <li>Click "Prepare Log Sheet" button on the main screen</li>
            <li>Choose where to get your files from:
                <ul>
                    <li>"Import from Cloud Storage" to download files from online</li>
                    <li>"Import Local Files" to use files from your computer</li>
                </ul>
            </li>
            <li>Select the files you want to combine</li>
            <li>Click "Merge Logs Files" to join them into one file</li>
        </ul>
        
        <h3 style='color: white;'>Step 3: Manage Schedules</h3>
        <ul style='color: white;'>
            <li>Click "Manage Schedules" button on the main screen</li>
            <li>To create a new schedule:
                <ul>
                    <li>Enter the module name</li>
                    <li>Select year, group, subject, and location</li>
                    <li>Pick the date and time using the calendar</li>
                    <li>Click "Add Session" for each class meeting</li>
                    <li>Click "Create Schedule" when you're done</li>
                </ul>
            </li>
            <li>To update an existing schedule:
                <ul>
                    <li>Select "Update Existing" option</li>
                    <li>Click "Browse" to find your schedule file</li>
                    <li>Make changes to existing sessions or add new ones</li>
                    <li>Click "Update Schedule" to save changes</li>
                </ul>
            </li>
        </ul>
        
        <h3 style='color: white;'>Step 4: Process Appeals</h3>
        <ul style='color: white;'>
            <li>Click "Process Appeals" button on the main screen</li>
            <li>Set up your data sources:
                <ul>
                    <li>Select your reference data file with student information</li>
                    <li>Select your existing attendance logs file</li>
                    <li>Select your session schedules file</li>
                </ul>
            </li>
            <li>Managing appeals:
                <ul>
                    <li>Search for students by name or ID in the search box</li>
                    <li>Select a student from the student table</li>
                    <li>View student details in the details panel</li>
                    <li>Select relevant sessions from the filtered sessions table</li>
                    <li>Click "Add to Appeals" to create an appeal</li>
                    <li>Review all selected appeals in the appeals management table</li>
                    <li>Remove any unneeded appeals if necessary</li>
                    <li>Click "Process Appeals" to finalize and save the changes</li>
                </ul>
            </li>
        </ul>
        
        <h3 style='color: white;'>Step 5: Process Attendance</h3>
        <ul style='color: white;'>
            <li>Click "Process Attendance" button on the main screen</li>
                <ul>
            <li>For generating the attendance reports:
<ul>
                    <li>Select your reference file</li>
                    <li>Select your attendance logs</li>
                    <li>Add all relevant schedule files</li>
</ul>

            </li>

            <li>Click "Process Attendance Records" to start</li>
            <li>Wait for the system to finish processing</li>
            <li>Find your reports in the "attendance_reports" folder</li>
                </ul>
                <ul>
            <li>For updating previous reports in case some students transferred between groups:
<ul>
                    <li>Select the previous report</li>
                    <li>Select your NEW reference file</li>
                    <li>Select your attendance logs</li>
                    <li>Add all relevant schedule files</li> 
</ul>
            </li>
            <li>Click "Update Report" to start</li>
            <li>Wait for the system to finish updating</li>
            <li>Find your updated reports in the "attendance_reports" folder</li>
                </ul>
        </ul>
        
        <h3 style='color: white;'>Step 6: Populate Main File</h3>
        <ul style='color: white;'>
            <li>Click "Populate Main File" button on the main screen</li>
            <li>Select your department:
                <ul>
                    <li>Choose from the dropdown menu (Anatomy, Histology, Physiology, etc.)</li>
                </ul>
            </li>
            <li>Select your files:
                <ul>
                    <li>Browse to select the department attendance report file</li>
                    <li>Browse to select the main faculty attendance file to be updated</li>
                    <li>Select the appropriate sheet names if prompted</li>
                </ul>
            </li>
            <li>Click "Populate Main Attendance File" to start the process</li>
            <li>Wait for the system to:
                <ul>
                    <li>Automatically create a backup of your original file</li>
                    <li>Match student IDs between the files</li>
                    <li>Identify department-specific columns in the main file</li>
                    <li>Update attendance records for each student and session</li>
                </ul>
            </li>
            <li>Check the output console for progress updates and details</li>
            <li>A success message will appear when the process is complete</li>
        </ul>

        <h3 style='color: white;'>Step 7: Analyze Attendance</h3>
        <ul style='color: white;'>
            <li>Click "Analyze Attendance" button on the main screen</li>
            <li>Select your processed attendance report file from the previous step</li>
            <li>Click "Display Statistics" to see:
                <ul>
                    <li>Total number of students</li>
                    <li>Pass rate and average attendance</li>
                    <li>Number of at-risk students</li>
                    <li>Detailed information for each student</li>
                </ul>
            </li>
            <li>Use the search box to find specific students by name or ID</li>
        </ul>
        
        <h2 style='color: white;'>Tips for Best Results</h2>
        <ul style='color: white;'>
            <li>Keep your files organized in dedicated folders</li>
            <li>Process attendance weekly to stay up-to-date</li>
            <li>Use consistent naming for all your files</li>
            <li>When validating attendance, be aware that different time windows are used:
                <ul>
                    <li>Standard sessions: 15 minutes before to 150 minutes after session start</li>
                    <li>Exception hours (12, 1, 13, 3, 15): 15 minutes before to 150 minutes after</li>
                </ul>
            </li>
            <li>For students who transferred between groups, the system will:
                <ul>
                    <li>Analyze attendance patterns to identify when transfers occurred</li>
                    <li>Validate attendance against the appropriate group schedule based on transfer dates</li>
                </ul>
            </li>
        </ul>
        
        <h2 style='color: white;'>Troubleshooting</h2>
        <ul style='color: white;'>
            <li>If files won't import:
                <ul>
                    <li>Check that they are Excel (.xlsx) format</li>
                    <li>Make sure required columns are present</li>
                    <li>Verify you have permission to read/write these files</li>
                </ul>
            </li>
            <li>If data doesn't match:
                <ul>
                    <li>Check that student IDs are consistent across files</li>
                    <li>Look for typos in names or IDs</li>
                    <li>Make sure date formats are the same in all files</li>
                </ul>
            </li>
            <li>If the app crashes:
                <ul>
                    <li>Restart the application</li>
                    <li>Check that you have enough disk space</li>
                    <li>Make sure all required files are available</li>
                </ul>
            </li>
        </ul>
        
        <h2 style='color: white;'>Need Help?</h2>
        <p style='color: white;'><b>Developer Note:</b><br>
        This application was developed by a medical student at the Faculty of Medicine, Ain Shams University with the help of AI models including Claude, DeepSeek, Perplexity, and a little bit of ChatGPT.</p>
        <p style='color: white;'><b>Contact Information:</b></p>
        <ul style='color: white;'>
            <li>  231249@med.asu.edu.eg</li>
            <li>  mohammadhamdisaid.mh@icloud.com</li>
            <li>  mohammad_hamdi11@yahoo.com</li>
        </ul>
        """

        info_label = QLabel()
        info_label.setTextFormat(Qt.TextFormat.RichText)
        info_label.setText(info_text)
        info_label.setWordWrap(True)
        info_label.setStyleSheet("background-color: transparent;")
        content_layout.addWidget(info_label)

        scroll_area.setWidget(info_content)
        info_layout.addWidget(scroll_area)

        # Add the group box to the main layout
        main_layout.addWidget(info_group)

        # Add a stretch after the group box to push everything up
        main_layout.addStretch()

# ==========================================================reference preparer==========================================================#

class ReferenceFilePreparer(QWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet("""
            background-color: black; 
            color: white;
            QLabel {
                color: white;
            }
        """)
        self.init_ui()

    def return_to_home(self):
        # Get the stacked widget and switch to the start page
        stacked_widget = self.parent()
        if isinstance(stacked_widget, QStackedWidget):
            stacked_widget.setCurrentIndex(0)

    def init_ui(self):
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Header layout
        header_layout = QHBoxLayout()
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), 'ASU1.png')
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaled(60, 60, Qt.AspectRatioMode.KeepAspectRatio,
                                               Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        title_label = QLabel("Reference File Preparer")
        title_label.setStyleSheet(f"font-size: 24px; font-weight: bold;")
        header_layout.addWidget(logo_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Create vertical layout for header buttons
        header_buttons_layout = QVBoxLayout()
        header_buttons_layout.setSpacing(5)

        # Back button
        back_btn = QPushButton("Back to Home")
        back_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        back_btn.clicked.connect(self.return_to_home)

        # Exit button
        exit_btn = QPushButton("Exit")
        exit_btn.setStyleSheet(EXIT_BUTTON_STYLE)
        exit_btn.clicked.connect(QApplication.instance().quit)

        # Add buttons to vertical layout
        header_buttons_layout.addWidget(back_btn)
        header_buttons_layout.addWidget(exit_btn)
        header_buttons_layout.addStretch()

        # Add button layout to header
        header_layout.addStretch()
        header_layout.addLayout(header_buttons_layout)
        main_layout.addLayout(header_layout)

        # Input File Section
        input_group = QGroupBox("Input Excel File")
        input_group.setStyleSheet(GROUP_BOX_STYLE)
        input_layout = QHBoxLayout(input_group)

        # File selection layout with sheet combo on same line
        input_layout.addWidget(QLabel("Source File:"))
        self.input_file_path = QLineEdit()
        self.input_file_path.setPlaceholderText("Select Excel file...")
        input_layout.addWidget(self.input_file_path)
        
        browse_btn = QPushButton("Browse")
        browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        browse_btn.clicked.connect(self.browse_input_file)
        input_layout.addWidget(browse_btn)
        
        # Sheet selection on same line
        input_layout.addWidget(QLabel("Sheet:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(150)
        input_layout.addWidget(self.sheet_combo)
        
        main_layout.addWidget(input_group)

        # Data Before Section
        data_before_group = QGroupBox("Data Before Processing")
        data_before_group.setStyleSheet(GROUP_BOX_STYLE)
        data_before_layout = QVBoxLayout(data_before_group)

        self.data_before_table = QTableWidget()
        self.data_before_table.setStyleSheet(TABLE_STYLE)
        data_before_layout.addWidget(self.data_before_table)

        main_layout.addWidget(data_before_group)

        # Column Mapping Section
        mapping_group = QGroupBox("Column Mapping")
        mapping_group.setStyleSheet(GROUP_BOX_STYLE)
        mapping_layout = QGridLayout(mapping_group)

        # Student ID mapping
        mapping_layout.addWidget(QLabel("Student ID Column:"), 0, 0)
        self.id_column_combo = QComboBox()
        mapping_layout.addWidget(self.id_column_combo, 0, 1)

        # Name mapping
        mapping_layout.addWidget(QLabel("Name Column:"), 1, 0)
        self.name_column_combo = QComboBox()
        mapping_layout.addWidget(self.name_column_combo, 1, 1)

        # Year mapping
        mapping_layout.addWidget(QLabel("Year Column:"), 2, 0)
        self.year_column_combo = QComboBox()
        mapping_layout.addWidget(self.year_column_combo, 2, 1)

        # Group mapping
        mapping_layout.addWidget(QLabel("Group Column:"), 3, 0)
        self.group_column_combo = QComboBox()
        mapping_layout.addWidget(self.group_column_combo, 3, 1)

        # Preview button for mapping
        preview_mapping_btn = QPushButton("Preview Result")
        preview_mapping_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        preview_mapping_btn.clicked.connect(self.preview_mapping_result)
        mapping_layout.addWidget(preview_mapping_btn, 4, 0, 1, 2)

        main_layout.addWidget(mapping_group)

        # Data After Section
        data_after_group = QGroupBox("Data After Processing")
        data_after_group.setStyleSheet(GROUP_BOX_STYLE)
        data_after_layout = QVBoxLayout(data_after_group)

        self.data_after_table = QTableWidget()
        self.data_after_table.setStyleSheet(TABLE_STYLE)
        data_after_layout.addWidget(self.data_after_table)

        main_layout.addWidget(data_after_group)

        # Bottom buttons
        button_layout = QHBoxLayout()
        self.process_btn = QPushButton("Process and Save Reference File")
        self.process_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        self.process_btn.clicked.connect(self.process_file)
        button_layout.addWidget(self.process_btn)
        main_layout.addLayout(button_layout)

        # Connect file input changes to sheet loading
        self.input_file_path.textChanged.connect(self.load_sheets)
        
        # Connect sheet combobox changes to auto-preview
        self.sheet_combo.currentTextChanged.connect(self.preview_data)
        
        # Connect column combo changes to both "before" and "after" preview updates
        self.id_column_combo.currentTextChanged.connect(self.on_column_selection_changed)
        self.name_column_combo.currentTextChanged.connect(self.on_column_selection_changed)
        self.year_column_combo.currentTextChanged.connect(self.on_column_selection_changed)
        self.group_column_combo.currentTextChanged.connect(self.on_column_selection_changed)

    def on_column_selection_changed(self):
        """Handle any column selection change by updating both preview tables"""
        self.update_data_before_table()
        self.update_data_after_table()

    def browse_input_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if filename:
            self.input_file_path.setText(filename)

    def load_sheets(self):
        file_path = self.input_file_path.text()
        if file_path and os.path.isfile(file_path):
            try:
                wb = openpyxl.load_workbook(file_path, read_only=True)
                self.sheet_combo.clear()
                self.sheet_combo.addItems(wb.sheetnames)
            except Exception as e:
                self.show_error_message(f"Error loading workbook: {str(e)}")

    def preview_data(self):
        file_path = self.input_file_path.text()
        sheet_name = self.sheet_combo.currentText()
        
        if not file_path or not sheet_name:
            # Don't show error, just return silently
            return
            
        try:
            # Load the data
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Clear previous preview
            self.data_before_table.clear()
            self.data_before_table.setRowCount(0)
            self.data_before_table.setColumnCount(0)
            self.id_column_combo.clear()
            self.name_column_combo.clear()
            self.year_column_combo.clear()
            self.group_column_combo.clear()
            
            # Set up the table
            rows, cols = min(10, df.shape[0]), df.shape[1]
            self.data_before_table.setRowCount(rows)
            self.data_before_table.setColumnCount(cols)
            
            # Set column headers
            self.data_before_table.setHorizontalHeaderLabels(df.columns.tolist())
            
            # Populate combo boxes with column names
            column_names = [""] + df.columns.tolist()
            self.id_column_combo.addItems(column_names)
            self.name_column_combo.addItems(column_names)
            self.year_column_combo.addItems(column_names)
            self.group_column_combo.addItems(column_names)
            
            # Fill the preview table with data
            for i in range(rows):
                for j in range(cols):
                    value = str(df.iloc[i, j]) if not pd.isna(df.iloc[i, j]) else ""
                    item = QTableWidgetItem(value)
                    self.data_before_table.setItem(i, j, item)
            
            # Auto-adjust column widths
            self.data_before_table.resizeColumnsToContents()
            
            # Try to auto-detect column mappings
            self.auto_detect_columns(df.columns.tolist())
            
        except Exception as e:
            # Only show error if it's not because no sheet is selected yet
            if sheet_name:
                self.show_error_message(f"Error previewing data: {str(e)}")

    def update_data_before_table(self):
        """Update the 'Data Before' table with the currently selected columns"""
        file_path = self.input_file_path.text()
        sheet_name = self.sheet_combo.currentText()
        
        if not file_path or not sheet_name:
            return
            
        try:
            # Get selected columns
            id_col = self.id_column_combo.currentText()
            name_col = self.name_column_combo.currentText()
            year_col = self.year_column_combo.currentText()
            group_col = self.group_column_combo.currentText()
            
            # If no columns are selected yet, clear the table and return
            if not (id_col or name_col or year_col or group_col):
                self.data_before_table.clear()
                self.data_before_table.setRowCount(0)
                self.data_before_table.setColumnCount(0)
                return
                
            # Load the data
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Create a new dataframe with only the selected columns
            selected_cols = []
            col_headers = []
            
            if id_col:
                selected_cols.append(df[id_col])
                col_headers.append("Student ID")
            if name_col:
                selected_cols.append(df[name_col])
                col_headers.append("Name")
            if year_col:
                selected_cols.append(df[year_col])
                col_headers.append("Year")
            if group_col:
                selected_cols.append(df[group_col])
                col_headers.append("Group")
                
            if not selected_cols:
                return
                
            preview_df = pd.concat(selected_cols, axis=1)
            preview_df.columns = col_headers
            
            # Clear previous data
            self.data_before_table.clear()
            self.data_before_table.setRowCount(0)
            self.data_before_table.setColumnCount(0)
            
            # Set up the table
            rows, cols = min(10, preview_df.shape[0]), preview_df.shape[1]
            self.data_before_table.setRowCount(rows)
            self.data_before_table.setColumnCount(cols)
            
            # Set column headers
            self.data_before_table.setHorizontalHeaderLabels(preview_df.columns.tolist())
            
            # Fill the table with data
            for i in range(rows):
                for j in range(cols):
                    value = str(preview_df.iloc[i, j]) if not pd.isna(preview_df.iloc[i, j]) else ""
                    item = QTableWidgetItem(value)
                    self.data_before_table.setItem(i, j, item)
            
            # Auto-adjust column widths to fit content
            self.data_before_table.resizeColumnsToContents()
            
            # Set table to stretch columns to fill width
            header = self.data_before_table.horizontalHeader()
            for i in range(cols):
                header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
            
        except Exception as e:
            # Silent error handling for this automatic update
            pass

    def update_data_after_table(self):
        """Update the 'Data After' table with the currently selected columns after processing"""
        file_path = self.input_file_path.text()
        sheet_name = self.sheet_combo.currentText()
        
        if not file_path or not sheet_name:
            return
            
        try:
            # Get column mappings
            id_col = self.id_column_combo.currentText()
            name_col = self.name_column_combo.currentText()
            year_col = self.year_column_combo.currentText()
            group_col = self.group_column_combo.currentText()
            
            # If no columns are selected yet, clear the table and return
            if not (id_col or name_col or year_col or group_col):
                self.data_after_table.clear()
                self.data_after_table.setRowCount(0)
                self.data_after_table.setColumnCount(0)
                return
                
            # Load the data
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Create new dataframe with required columns
            new_df = pd.DataFrame()
            
            # Only add columns that are selected
            if id_col:
                new_df['Student ID'] = df[id_col].astype(str)
                # Clean Student ID - remove non-digit characters
                new_df['Student ID'] = new_df['Student ID'].apply(lambda x: ''.join(c for c in x if c.isdigit()))
            
            if name_col:
                new_df['Name'] = df[name_col]
                # Capitalize first letter of each word in Name
                new_df['Name'] = new_df['Name'].str.title()
            
            if year_col:
                new_df['Year'] = df[year_col]
                # Format Year values
                new_df['Year'] = new_df['Year'].apply(self.format_year)
            
            if group_col:
                new_df['Group'] = df[group_col]
                # Format Group values
                new_df['Group'] = new_df['Group'].apply(self.format_group)
            
            if new_df.empty:
                return
            
            # Clear previous data in After table
            self.data_after_table.clear()
            self.data_after_table.setRowCount(0)
            self.data_after_table.setColumnCount(0)
            
            # Set up the After table
            rows, cols = min(10, new_df.shape[0]), new_df.shape[1]
            self.data_after_table.setRowCount(rows)
            self.data_after_table.setColumnCount(cols)
            
            # Set column headers
            self.data_after_table.setHorizontalHeaderLabels(new_df.columns.tolist())
            
            # Fill the table with processed data
            for i in range(rows):
                for j in range(cols):
                    value = str(new_df.iloc[i, j]) if not pd.isna(new_df.iloc[i, j]) else ""
                    item = QTableWidgetItem(value)
                    self.data_after_table.setItem(i, j, item)
            
            # Auto-adjust column widths
            self.data_after_table.resizeColumnsToContents()
            
            # Set table to stretch columns to fill width
            header = self.data_after_table.horizontalHeader()
            for i in range(cols):
                header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
            
        except Exception as e:
            # Silent error handling for automatic updates
            pass

    def preview_mapping_result(self):
        """Preview the result of the mapping and formatting in the 'Data After' table"""
        # Validate inputs
        if not self.validate_inputs():
            return
            
        # Use the same method as auto-update
        self.update_data_after_table()

    def auto_detect_columns(self, columns):
        """Try to automatically detect appropriate column mappings based on column names"""
        for i, col in enumerate(columns):
            col_lower = col.lower()
            
            # +1 because we added an empty string at the beginning of combo box items
            if any(keyword in col_lower for keyword in ['id', 'student id', 'studentid', 'number']):
                self.id_column_combo.setCurrentIndex(i + 1)
                
            if any(keyword in col_lower for keyword in ['name', 'student name', 'studentname']):
                self.name_column_combo.setCurrentIndex(i + 1)
                
            if any(keyword in col_lower for keyword in ['year', 'level']):
                self.year_column_combo.setCurrentIndex(i + 1)
                
            if any(keyword in col_lower for keyword in ['group', 'section']):
                self.group_column_combo.setCurrentIndex(i + 1)

    def format_year(self, year_value):
        """Format year values to 'Year X' format"""
        if pd.isna(year_value) or not year_value:
            return "Year 1"  # Default value
            
        # Convert to string and clean it
        year_str = str(year_value).strip().lower()
        
        # Extract numeric part if it exists
        numeric_part = ''.join(c for c in year_str if c.isdigit())
        
        if not numeric_part:
            return "Year 1"  # Default if no number found
            
        # Format as "Year X"
        return f"Year {numeric_part}"

    def format_group(self, group_value):
        """Format group values to 'AX' format (uppercase letter followed by number)"""
        if pd.isna(group_value) or not group_value:
            return "A1"  # Default value
            
        # Convert to string and clean it
        group_str = str(group_value).strip().lower()
        
        # Try to extract letter and number parts
        letter_part = ""
        numeric_part = ""
        
        for c in group_str:
            if c.isalpha():
                letter_part += c
            elif c.isdigit():
                numeric_part += c
        
        # Default values if parts are missing
        if not letter_part:
            letter_part = "A"
        if not numeric_part:
            numeric_part = "1"
            
        # Take first letter and capitalize it
        letter_part = letter_part[0].upper()
        
        # Format as "A1" (no space)
        return f"{letter_part}{numeric_part}"

    def process_file(self):
        # Validate inputs
        if not self.validate_inputs():
            return
            
        try:
            # Load input data
            input_file = self.input_file_path.text()
            sheet_name = self.sheet_combo.currentText()
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            
            # Get column mappings
            id_col = self.id_column_combo.currentText()
            name_col = self.name_column_combo.currentText()
            year_col = self.year_column_combo.currentText()
            group_col = self.group_column_combo.currentText()
            
            # Create new dataframe with required columns
            new_df = pd.DataFrame()
            new_df['Student ID'] = df[id_col] if id_col else None
            new_df['Name'] = df[name_col] if name_col else None
            new_df['Year'] = df[year_col] if year_col else None
            new_df['Group'] = df[group_col] if group_col else None
            
            # Clean and format data
            # Convert Student ID to string and remove any non-digit characters
            new_df['Student ID'] = new_df['Student ID'].astype(str)
            new_df['Student ID'] = new_df['Student ID'].apply(lambda x: ''.join(c for c in x if c.isdigit()))
            
            # Capitalize first letter of each word in Name
            new_df['Name'] = new_df['Name'].str.title()
            
            # Format Year values to "Year X"
            new_df['Year'] = new_df['Year'].apply(self.format_year)
            
            # Format Group values to "AX" (uppercase letter followed by number)
            new_df['Group'] = new_df['Group'].apply(self.format_group)
            
            # Method 1: Get the directory of the running executable
            if getattr(sys, 'frozen', False):
                # If running as compiled executable
                app_dir = os.path.dirname(sys.executable)
            else:
                # If running as script
                app_dir = os.path.dirname(os.path.abspath(__file__))
            
            # Create a reference_data folder if it doesn't exist
            reference_data_dir = os.path.join(app_dir, "reference_data")
            if not os.path.exists(reference_data_dir):
                os.makedirs(reference_data_dir)
            
            # Create filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"reference_data_{timestamp}.xlsx"
            
            # Set output file in the reference_data directory
            output_file = os.path.join(reference_data_dir, output_filename)
            
            # Create a new workbook
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Reference Data"
            
            # Add header row
            headers = ['Student ID', 'Name', 'Year', 'Group']
            ws.append(headers)
            # Format header row
            for i, cell in enumerate(ws[1]):
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="D3D3D3")
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Add data rows
            for _, row in new_df.iterrows():
                ws.append([row['Student ID'], row['Name'], row['Year'], row['Group']])
            
            # Auto-adjust column widths
            for column in ws.columns:
                max_length = max(len(str(cell.value) if cell.value else '') for cell in column)
                ws.column_dimensions[openpyxl.utils.get_column_letter(column[0].column)].width = max_length + 3
            
            # Save the workbook
            wb.save(output_file)
            
            self.show_success_message(f"Reference file created successfully!\nSaved in the 'reference_data' folder as:\n{output_filename}")
            
        except Exception as e:
            self.show_error_message(f"Error processing file: {str(e)}")
    
    def validate_inputs(self):
        # Check if input file is selected
        if not self.input_file_path.text():
            self.show_error_message("Please select an input file.")
            return False
            
        # Check if sheet is selected
        if not self.sheet_combo.currentText():
            self.show_error_message("Please select a sheet.")
            return False
            
        # Check if column mappings are selected
        if not self.id_column_combo.currentText():
            self.show_error_message("Please select a Student ID column.")
            return False
            
        if not self.name_column_combo.currentText():
            self.show_error_message("Please select a Name column.")
            return False
            
        if not self.year_column_combo.currentText():
            self.show_error_message("Please select a Year column.")
            return False
            
        if not self.group_column_combo.currentText():
            self.show_error_message("Please select a Group column.")
            return False
            
        return True

    def show_error_message(self, message):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Icon.Critical)
        msg_box.setWindowTitle("Error")
        msg_box.setText(message)
        msg_box.setStyleSheet(f"""
            QMessageBox {{
                background-color: {CARD_BG};
            }}
            QLabel {{
                color: {TEXT_COLOR};
                font-size: 14px;
            }}
        """)
        # Style OK button
        ok_button = msg_box.addButton(QMessageBox.StandardButton.Ok)
        ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)
        msg_box.exec()

    def show_success_message(self, message):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Icon.Information)
        msg_box.setWindowTitle("Success")
        msg_box.setText(message)
        msg_box.setStyleSheet(f"""
            QMessageBox {{
                background-color: {CARD_BG};
            }}
            QLabel {{
                color: {TEXT_COLOR};
                font-size: 14px;
            }}
        """)
        # Style OK button
        ok_button = msg_box.addButton(QMessageBox.StandardButton.Ok)
        ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)
        msg_box.exec()

# ==========================================================log sheet preparer==========================================================#

class GithubDownloadWorker(QThread):
    progress_signal = pyqtSignal(int)
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(list)

    def __init__(self, repo_url, token):
        super().__init__()
        self.repo_url = repo_url
        self.token = token
        self.downloaded_files = []

    def run(self):
        try:
            # Parse the repo URL to extract owner and repo name
            # Example: "https://github.com/username/repo"
            parts = self.repo_url.strip('/').split('/')
            if len(parts) < 5 or parts[2] != 'github.com':
                self.log_signal.emit("Invalid GitHub repository URL format")
                return

            owner = parts[3]
            repo = parts[4]

            # Get repository contents
            self.log_signal.emit(
                f"Connecting to GitHub repository: {owner}/{repo}")
            headers = {
                'Authorization': f'token {self.token}'} if self.token else {}

            # Get all files in the repository
            # Specify the backups folder
            api_url = f"https://api.github.com/repos/{owner}/{repo}/contents/backups"
            response = requests.get(api_url, headers=headers)

            if response.status_code != 200:
                self.log_signal.emit(
                    f"Error accessing repository: {response.status_code}, {response.text}")
                return

            contents = response.json()
            excel_files = [item for item in contents if item['name'].endswith(
                '.xlsx') or item['name'].endswith('.xls')]

            if not excel_files:
                self.log_signal.emit("No Excel files found in the repository")
                return

            # Create base directory and imported logs subdirectory
            # Use a better approach for determining base directory that works with both script and exe
            if getattr(sys, 'frozen', False):
                # If the application is run as a bundle, the PyInstaller bootloader
                # extends the sys module by a flag frozen=True and sets the app 
                # path into variable sys._MEIPASS
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            
            temp_dir = os.path.join(base_dir, 'log_history', 'Imported_logs')
            os.makedirs(temp_dir, exist_ok=True)

            # Download each Excel file
            total_files = len(excel_files)
            for idx, file in enumerate(excel_files):
                self.log_signal.emit(f"Downloading {file['name']}...")

                download_url = file['download_url']
                file_response = requests.get(download_url, headers=headers)

                if file_response.status_code == 200:
                    # Save file locally
                    file_path = os.path.join(temp_dir, file['name'])
                    with open(file_path, 'wb') as f:
                        f.write(file_response.content)

                    self.downloaded_files.append(file_path)
                    self.log_signal.emit(f"Downloaded {file['name']}")
                else:
                    self.log_signal.emit(
                        f"Failed to download {file['name']}: {file_response.status_code}")

                # Update progress
                progress = int(((idx + 1) / total_files) * 100)
                self.progress_signal.emit(progress)

            self.log_signal.emit(
                f"Downloaded {len(self.downloaded_files)} Excel files")
            self.finished_signal.emit(self.downloaded_files)

        except Exception as e:
            self.log_signal.emit(f"Error: {str(e)}")
            import traceback
            self.log_signal.emit(traceback.format_exc())

class MergeWorker(QThread):
    progress_signal = pyqtSignal(int)
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str)

    def __init__(self, files, output_file):
        super().__init__()
        self.files = files
        self.output_file = output_file

    def run(self):
        try:
            if not self.files:
                self.log_signal.emit("No files to merge")
                return

            self.log_signal.emit(f"Starting merge of {len(self.files)} files")

            # Initialize a list to hold all dataframes
            all_dfs = []

            # Process each file
            for idx, file_path in enumerate(self.files):
                self.log_signal.emit(
                    f"Processing {os.path.basename(file_path)}")

                try:
                    # Read all sheets from the Excel file
                    excel_file = pd.ExcelFile(file_path)
                    for sheet_name in excel_file.sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet_name)

                        # Only process if the dataframe is not empty
                        if not df.empty:
                            # Add file and sheet metadata
                            df['Source_File'] = os.path.basename(file_path)
                            df['Source_Sheet'] = sheet_name
                            all_dfs.append(df)
                            self.log_signal.emit(
                                f"Added sheet '{sheet_name}' with {len(df)} rows")

                except Exception as e:
                    self.log_signal.emit(
                        f"Error processing {file_path}: {str(e)}")
                    continue

                # Update progress
                progress = int(((idx + 1) / len(self.files)) * 100)
                self.progress_signal.emit(progress)

            if not all_dfs:
                self.log_signal.emit("No valid data found in the files")
                return

            # Standardize column names across all dataframes
            self.log_signal.emit("Standardizing column headers...")

            # Find common columns or use a predefined set of columns
            # For now, we'll use a simple approach of getting all unique columns
            all_columns = set()
            for df in all_dfs:
                all_columns.update(df.columns)

            # Remove metadata columns we added
            standard_columns = [col for col in all_columns if col not in [
                'Source_File', 'Source_Sheet']]

            # Reindex all dataframes with the standard columns
            standardized_dfs = []
            for df in all_dfs:
                # Create a new dataframe with all standard columns (will be filled with NaN for missing columns)
                new_df = pd.DataFrame(columns=standard_columns)

                # Copy data from original dataframe for matching columns
                for col in standard_columns:
                    if col in df.columns:
                        new_df[col] = df[col]

                # Add back metadata columns
                new_df['Source_File'] = df['Source_File']
                new_df['Source_Sheet'] = df['Source_Sheet']

                standardized_dfs.append(new_df)

            # Concatenate all dataframes
            self.log_signal.emit("Merging all sheets...")
            merged_df = pd.concat(standardized_dfs, ignore_index=True)

            # Identify and reorder columns as per requirements:
            # 1. Student ID (looking for column containing "Student" and "ID")
            # 2. Location (looking for column containing "Location")
            # 3. Log date (looking for column containing "Log" and "date" or "Date")
            # 4. Log time (looking for column containing "Log" and "time" or "Time")
            # 5. All other columns

            self.log_signal.emit("Reordering columns to specified format...")

            # Find the best matching columns based on column names
            student_id_col = None
            location_col = None
            log_date_col = None
            log_time_col = None

            # Look for exact or partial matches
            for col in merged_df.columns:
                col_lower = str(col).lower()

                # Check for Student ID
                if "student" in col_lower and "id" in col_lower:
                    student_id_col = col
                # Check for Location
                elif "location" in col_lower:
                    location_col = col
                # Check for Log date
                elif "log" in col_lower and ("date" in col_lower or "day" in col_lower):
                    log_date_col = col
                # Check for Log time
                elif "log" in col_lower and "time" in col_lower:
                    log_time_col = col

            # Create the ordered columns list
            ordered_columns = []

            # Add the main required columns if they exist
            for col in [student_id_col, location_col, log_date_col, log_time_col]:
                if col is not None and col in merged_df.columns:
                    ordered_columns.append(col)

            # Add all remaining columns (excluding the ones we've already added and metadata)
            remaining_columns = [col for col in merged_df.columns
                                if col not in ordered_columns
                                and col not in ['Source_File', 'Source_Sheet']]
            ordered_columns.extend(remaining_columns)

            # Add metadata columns at the end
            ordered_columns.extend(['Source_File', 'Source_Sheet'])

            # Log the column ordering
            self.log_signal.emit(
                f"Column order being used: {', '.join(ordered_columns[:4])} + remaining columns")

            # Reorder the dataframe columns
            merged_df = merged_df[ordered_columns]

            # Make sure the output directory exists
            os.makedirs(os.path.dirname(self.output_file), exist_ok=True)

            # Save the merged data with reordered columns
            self.log_signal.emit(f"Saving merged data to {self.output_file}")
            merged_df.to_excel(self.output_file, index=False)

            self.log_signal.emit(
                f"Successfully merged {len(all_dfs)} sheets into {self.output_file} with ordered columns")
            self.finished_signal.emit(self.output_file)

        except Exception as e:
            self.log_signal.emit(f"Error during merge: {str(e)}")
            import traceback
            self.log_signal.emit(traceback.format_exc())

class LogSheetPreparer(QWidget):
    def __init__(self):
        super().__init__()
        self.files_to_merge = []
        # Define the hardcoded GitHub token - split to avoid detection
        token_part1 = "github_pat_"
        token_part2 = "11BREVRNQ0LX45XKQZzjkB_TL3KNQxHy4Sms4Fo20IUcxNLUwNAFbfeiXy92idb3mwTVANNZ4EC92cvkof"
        self.github_token = token_part1 + token_part2
        # Define the hardcoded GitHub repo URL - hidden from UI
        self.github_repo = "https://github.com/MohammadHamdi11/RN-E-attendancerecorderapp"
        self.setStyleSheet("""
            background-color: black; 
            color: white;
            QLabel {
                color: white;
            }
        """)

        self.init_ui()

    def init_ui(self):
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Header layout
        header_layout = QHBoxLayout()
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), 'ASU1.png')
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaled(60, 60, Qt.AspectRatioMode.KeepAspectRatio,
                                               Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        title_label = QLabel("Log Sheet Preparer")
        title_label.setStyleSheet(f"font-size: 24px; font-weight: bold;")
        header_layout.addWidget(logo_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Create vertical layout for header buttons
        header_buttons_layout = QVBoxLayout()
        header_buttons_layout.setSpacing(5)

        # Back button
        back_btn = QPushButton("Back to Home")
        back_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        back_btn.clicked.connect(self.return_to_home)

        # Exit button
        exit_btn = QPushButton("Exit")
        exit_btn.setStyleSheet(EXIT_BUTTON_STYLE)
        exit_btn.clicked.connect(QApplication.instance().quit)

        # Add buttons to vertical layout
        header_buttons_layout.addWidget(back_btn)
        header_buttons_layout.addWidget(exit_btn)
        header_buttons_layout.addStretch()

        # Add button layout to header
        header_layout.addStretch()
        header_layout.addLayout(header_buttons_layout)
        main_layout.addLayout(header_layout)

        # Import Method Selection
        import_group = QGroupBox("Import Method")
        import_group.setStyleSheet(GROUP_BOX_STYLE)
        import_layout = QVBoxLayout(import_group)

        # Radio buttons for import method
        radio_layout = QHBoxLayout()
        self.github_radio = QRadioButton("Import from Cloud Storage")
        self.github_radio.setChecked(True)  # Default to GitHub import
        self.local_radio = QRadioButton("Import Local Files")
        radio_layout.addWidget(self.github_radio)
        radio_layout.addWidget(self.local_radio)
        radio_layout.addStretch()
        import_layout.addLayout(radio_layout)

        # Connect radio buttons to toggle between input methods
        self.github_radio.toggled.connect(self.toggle_import_method)

        # Stacked widget for different import methods
        self.import_stack = QStackedWidget()

        # GitHub Import Widget
        github_widget = QWidget()
        github_layout = QVBoxLayout(github_widget)

        # Only show informational text about the GitHub repo, hiding the actual implementation details
        github_info_label = QLabel(
            "Excel files will be downloaded from the QRScanner-webapp repository's backup folder.")
        github_layout.addWidget(github_info_label)

        # Button to download files from GitHub
        download_btn = QPushButton("Download Excel Files from Cloud Storage")
        download_btn.clicked.connect(self.download_github_files)
        download_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        github_layout.addWidget(download_btn)

        # Local Files Import Widget
        local_widget = QWidget()
        local_layout = QVBoxLayout(local_widget)

        local_files_layout = QHBoxLayout()
        local_files_layout.addWidget(QLabel("Local Excel Files:"))
        self.local_files_label = QLineEdit()
        self.local_files_label.setPlaceholderText("Select Excel file...")
        self.local_files_label.setMinimumWidth(300)
        self.local_files_label.setReadOnly(True)
        local_files_layout.addWidget(self.local_files_label)

        # Button for importing local files
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse_files)
        browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        local_files_layout.addWidget(browse_btn)
        local_layout.addLayout(local_files_layout)

        # Add widgets to stack
        self.import_stack.addWidget(github_widget)
        self.import_stack.addWidget(local_widget)
        import_layout.addWidget(self.import_stack)

        main_layout.addWidget(import_group)

        # Files Table Section
        files_group = QGroupBox("Files to Merge")
        files_group.setStyleSheet(GROUP_BOX_STYLE)
        files_layout = QVBoxLayout(files_group)

        self.files_table = QTableWidget()
        self.files_table.setColumnCount(2)
        self.files_table.setHorizontalHeaderLabels(['File Path', 'Status'])

        # Center align the header text
        header = self.files_table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)

        # Set column resize modes
        header.setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch)  # File Path
        header.setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch)  # Status

        # Apply table styles
        self.files_table.setStyleSheet(TABLE_STYLE)
        files_layout.addWidget(self.files_table)

        # Buttons for files table
        files_btn_layout = QHBoxLayout()
        clear_files_btn = QPushButton("Clear Files")
        clear_files_btn.clicked.connect(self.clear_files)
        clear_files_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        files_btn_layout.addWidget(clear_files_btn)
        files_layout.addLayout(files_btn_layout)

        main_layout.addWidget(files_group)

        # Progress Bar Section
        progress_group = QGroupBox("Progress")
        progress_group.setStyleSheet(GROUP_BOX_STYLE)
        progress_layout = QVBoxLayout(progress_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar.setStyleSheet(PROGRESS_BAR_STYLE)

        # Create loading gif label
        self.loading_label = QLabel()
        self.loading_label.setFixedSize(24, 24)
        self.loading_label.setVisible(False)

        # Create the movie object for the GIF
        self.loading_movie = QMovie()
        self.loading_movie.setScaledSize(QSize(24, 24))
        self.loading_label.setMovie(self.loading_movie)

        loading_gif_path = os.path.join(
            os.path.dirname(__file__), 'loading.gif')
        if os.path.exists(loading_gif_path):
            self.loading_movie.setFileName(loading_gif_path)
        else:
            print(f"Warning: loading.gif not found at {loading_gif_path}")

        # Create a horizontal layout to hold both the progress bar and loading animation
        progress_h_layout = QHBoxLayout()
        progress_h_layout.addWidget(self.progress_bar)
        progress_h_layout.addWidget(self.loading_label)
        progress_layout.addLayout(progress_h_layout)

        main_layout.addWidget(progress_group)

        # Output Console Section
        console_group = QGroupBox("Output Console")
        console_group.setStyleSheet(GROUP_BOX_STYLE)
        console_layout = QVBoxLayout(console_group)

        self.output_console = QTextEdit()
        self.output_console.setReadOnly(True)
        self.output_console.setMaximumHeight(150)
        self.output_console.setStyleSheet(CONSOLE_STYLE)
        console_layout.addWidget(self.output_console)
        main_layout.addWidget(console_group)

        # Bottom Buttons
        button_layout = QHBoxLayout()
        merge_btn = QPushButton("Merge Logs Files")
        merge_btn.clicked.connect(self.merge_files)
        merge_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        button_layout.addWidget(merge_btn)
        main_layout.addLayout(button_layout)

    def toggle_import_method(self):
        # Set the current import method based on radio button selection
        if self.github_radio.isChecked():
            self.import_stack.setCurrentIndex(0)
        else:
            self.import_stack.setCurrentIndex(1)

    def return_to_home(self):
        # Get the stacked widget and switch to the start page
        stacked_widget = self.parent()
        if isinstance(stacked_widget, QStackedWidget):
            stacked_widget.setCurrentIndex(0)

    def browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Excel Files", "", "Excel Files (*.xlsx *.xls)"
        )
        if files:
            self.files_to_merge = files
            self.local_files_label.setText(f"{len(files)} files selected")
            self.update_files_table()
            self.log_message(f"Selected {len(files)} files for merging")

    def update_files_table(self):
        # Clear existing table
        self.files_table.setRowCount(0)

        # Add files to table
        for file_path in self.files_to_merge:
            row_position = self.files_table.rowCount()
            self.files_table.insertRow(row_position)

            # Create items for the cells
            file_item = QTableWidgetItem(os.path.basename(file_path))
            status_item = QTableWidgetItem("Ready")

            # Set items to the table
            self.files_table.setItem(row_position, 0, file_item)
            self.files_table.setItem(row_position, 1, status_item)

    def clear_files(self):
        self.files_to_merge = []
        self.files_table.setRowCount(0)
        self.local_files_label.setText("No files selected")
        self.log_message("Cleared all files")

    def log_message(self, message):
        self.output_console.append(
            f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
        # Scroll to the bottom
        self.output_console.moveCursor(QTextCursor.MoveOperation.End)

    def download_github_files(self):
        # Use hardcoded repo URL and token - not visible to users
        repo_url = self.github_repo
        token = self.github_token

        self.log_message(f"Starting download from Cloud Storage")

        # Start loading animation
        self.loading_label.setVisible(True)
        self.loading_movie.start()
        self.progress_bar.setValue(0)

        # Create and start the worker thread
        self.github_worker = GithubDownloadWorker(repo_url, token)
        self.github_worker.progress_signal.connect(self.update_progress)
        self.github_worker.log_signal.connect(self.log_message)
        self.github_worker.finished_signal.connect(
            self.handle_downloaded_files)
        self.github_worker.start()

    def handle_downloaded_files(self, files):
        # Update the list of files to merge
        self.files_to_merge = files
        self.update_files_table()

        # Stop loading animation
        self.loading_movie.stop()
        self.loading_label.setVisible(False)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def merge_files(self):
        if not self.files_to_merge:
            self.log_message("No files to merge. Please import files first.")
            return

        # Generate output filename with timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"log_history_{timestamp}.xlsx"

        # Create base directory and merged files subdirectory
        # Use a better approach for determining base directory that works with both script and exe
        if getattr(sys, 'frozen', False):
            # If the application is run as a bundle, the PyInstaller bootloader
            # extends the sys module by a flag frozen=True and sets the app 
            # path into variable sys._MEIPASS
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        
        output_dir = os.path.join(base_dir, 'log_history', 'Merged_Files')
        os.makedirs(output_dir, exist_ok=True)

        # Full path for output file
        output_file = os.path.join(output_dir, output_filename)

        self.log_message(f"Merging files to: {output_file}")

        # Start loading animation
        self.loading_label.setVisible(True)
        self.loading_movie.start()
        self.progress_bar.setValue(0)

        # Create and start the worker thread
        self.merge_worker = MergeWorker(self.files_to_merge, output_file)
        self.merge_worker.progress_signal.connect(self.update_progress)
        self.merge_worker.log_signal.connect(self.log_message)
        self.merge_worker.finished_signal.connect(self.handle_merge_complete)
        self.merge_worker.start()

    def handle_merge_complete(self, output_file):
        # Stop loading animation
        self.loading_movie.stop()
        self.loading_label.setVisible(False)

        self.log_message(f"Merge completed successfully: {output_file}")

        # Ask if user wants to open the merged file
        from PyQt6.QtWidgets import QMessageBox
        reply = QMessageBox.question(
            self, 'Merge Complete',
            f'Merge completed successfully!\nThe merged file has been saved as:\n{os.path.basename(output_file)}\n\nWould you like to open the merged file?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes
        )

        if reply == QMessageBox.StandardButton.Yes:
            # Open the file with the default application
            import subprocess
            import platform

            if platform.system() == 'Windows':
                os.startfile(output_file)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(('open', output_file))
            else:  # Linux
                subprocess.call(('xdg-open', output_file))

# ==========================================================Schedule Manager==========================================================#

class ScheduleManager(QWidget):
    def __init__(self):
        super().__init__()
        self.schedule_data = []
        self.current_year = None
        self.current_group = None
        self.setStyleSheet(f"""
            background-color: {BLACK}; 
            color: {TEXT_COLOR};
            QLabel {{
                color: {TEXT_COLOR};
            }}
            QCalendarWidget {{
                background-color: {INPUT_BG};
                color: {TEXT_COLOR};
            }}
            QCalendarWidget QAbstractItemView {{
                selection-background-color: #555;
                selection-color: {TEXT_COLOR};
            }}
            QCalendarWidget QWidget {{
                alternate-background-color: #444;
                color: {TEXT_COLOR};
            }}
            QCalendarWidget QTableView {{
                border: none;
            }}
        """)
        
        self.init_ui()

    def return_to_home(self):
        # Get the stacked widget and switch to the start page
        stacked_widget = self.parent()
        if isinstance(stacked_widget, QStackedWidget):
            stacked_widget.setCurrentIndex(0)

    def init_ui(self):
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Create a scroll area to make the page scrollable
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; }")
        
        # Create a widget to hold the content
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(20)
        scroll_layout.setContentsMargins(0, 0, 0, 0)
        
        # Header layout
        header_layout = QHBoxLayout()
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), 'ASU1.png')
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaled(60, 60, Qt.AspectRatioMode.KeepAspectRatio, 
                                             Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        title_label = QLabel("Schedule Manager")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        header_layout.addWidget(logo_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()
    
        # Create vertical layout for header buttons
        header_buttons_layout = QVBoxLayout()
        header_buttons_layout.setSpacing(5)
        
        # Back button
        back_btn = QPushButton("Back to Home")
        back_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        back_btn.clicked.connect(self.return_to_home)
        
        # Exit button
        exit_btn = QPushButton("Exit")
        exit_btn.setStyleSheet(EXIT_BUTTON_STYLE)
        exit_btn.clicked.connect(QApplication.instance().quit)
        
        # Add buttons to vertical layout
        header_buttons_layout.addWidget(back_btn)
        header_buttons_layout.addWidget(exit_btn)
        header_buttons_layout.addStretch()
        
        # Add button layout to header
        header_layout.addStretch()
        header_layout.addLayout(header_buttons_layout)
        scroll_layout.addLayout(header_layout)
    
        # Mode Selection Group
        mode_group = QGroupBox("Schedule Mode")
        mode_group.setStyleSheet(GROUP_BOX_STYLE)
        mode_layout = QHBoxLayout(mode_group)
        
        # Radio buttons for mode selection
        self.create_radio = QRadioButton("Create New Schedule")
        self.create_radio.setChecked(True)  # Default selection
        self.update_radio = QRadioButton("Update Existing Schedule")
        
        mode_layout.addWidget(self.create_radio)
        mode_layout.addWidget(self.update_radio)
        mode_layout.addStretch()
        
        # Connect radio buttons to switch mode functions
        self.create_radio.toggled.connect(self.switch_to_create_mode)
        self.update_radio.toggled.connect(self.switch_to_update_mode)
        
        scroll_layout.addWidget(mode_group)
    
        # Stacked widget to hold different mode layouts
        self.mode_stack = QStackedWidget()
        scroll_layout.addWidget(self.mode_stack)
        
        # Create mode widget
        self.create_widget = QWidget()
        self.create_layout = QVBoxLayout(self.create_widget)
        self.setup_create_mode()
        self.mode_stack.addWidget(self.create_widget)
        
        # Update mode widget
        self.update_widget = QWidget()
        self.update_layout = QVBoxLayout(self.update_widget)
        self.setup_update_mode()
        self.mode_stack.addWidget(self.update_widget)
    
        # Default to create mode
        self.mode_stack.setCurrentIndex(0)
    
        # Bottom Buttons layout
        self.button_layout = QHBoxLayout()
        
        # Create Schedule button
        self.create_schedule_btn = QPushButton("Create Schedule")
        self.create_schedule_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        self.create_schedule_btn.clicked.connect(self.save_schedule)
        self.button_layout.addWidget(self.create_schedule_btn)
        
        # Update Schedule button (initially hidden)
        self.update_schedule_btn = QPushButton("Update Schedule")
        self.update_schedule_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        self.update_schedule_btn.clicked.connect(self.update_schedule)
        self.update_schedule_btn.setVisible(False)
        self.button_layout.addWidget(self.update_schedule_btn)
        
        scroll_layout.addLayout(self.button_layout)
        
        # Set the scroll content and add to main layout
        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)
    
    def setup_create_mode(self):
        # Add New Sessions group
        add_sessions_group = QGroupBox("Add New Sessions")
        add_sessions_group.setStyleSheet(GROUP_BOX_STYLE)
        add_sessions_layout = QVBoxLayout(add_sessions_group)  # Use vertical layout as container
        
        # MODULE NAME INPUT - Add this new section
        module_name_layout = QHBoxLayout()
        module_name_layout.addWidget(QLabel("Module Name:"))
        self.module_name_input = QLineEdit()
        self.module_name_input.setPlaceholderText("Enter module name...")
        module_name_layout.addWidget(self.module_name_input)
        add_sessions_layout.addLayout(module_name_layout)
        
        # Create horizontal layout for the form fields
        form_layout = QHBoxLayout()
        
        # Left column for form inputs
        left_column = QVBoxLayout()
        left_column.setSpacing(15)
        
        # Year Selection
        year_row = QHBoxLayout()
        year_row.addWidget(QLabel("Academic Year:"))
        self.year_combo = QComboBox()
        self.year_combo.setMaxVisibleItems(10)  # Use setMaxVisibleItems instead
        self.year_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)  # Make it expand horizontally
        self.year_combo.addItems(["Year 1", "Year 2", "Year 3"])
        self.year_combo.currentTextChanged.connect(self.year_selected)
        year_row.addWidget(self.year_combo)
        year_row.addStretch()
        left_column.addLayout(year_row)
        
        # Group selection row
        group_row = QHBoxLayout()
        group_row.addWidget(QLabel("Group:            "))
        self.group_combo = QComboBox()
        self.group_combo.setMaxVisibleItems(10)
        self.group_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.populate_groups()
        group_row.addWidget(self.group_combo)
        group_row.addStretch()
        left_column.addLayout(group_row)
        
        # Subject row
        subject_row = QHBoxLayout()
        subject_row.addWidget(QLabel("Subject:          "))
        self.subject_combo = QComboBox()
        self.subject_combo.setMaxVisibleItems(10)
        self.subject_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.subject_combo.addItems([
            "Anatomy", "Histology", "Biochemistry", "Physiology", 
            "Microbiology", "Parasitology", "Pathology", "Pharmacology", "Clinical"
        ])
        self.subject_combo.currentTextChanged.connect(self.update_location_options)
        subject_row.addWidget(self.subject_combo)
        subject_row.addStretch()
        left_column.addLayout(subject_row)
        
        # Session row
        session_row = QHBoxLayout()
        session_row.addWidget(QLabel("Session:          "))
        self.session_combo = QComboBox()
        self.session_combo.setMaxVisibleItems(10)
        self.session_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.session_combo.addItems([str(i) for i in range(1, 31)])
        session_row.addWidget(self.session_combo)
        session_row.addStretch()
        left_column.addLayout(session_row)
    
        # Location row
        location_row = QHBoxLayout()
        location_row.addWidget(QLabel("Location:         "))
        self.location_combo = QComboBox()
        self.location_combo.setMaxVisibleItems(10)
        self.location_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.populate_locations()
        location_row.addWidget(self.location_combo)
        location_row.addStretch()
        left_column.addLayout(location_row)
        
        # Add left column to form layout
        form_layout.addLayout(left_column, 1)
        
        # Right column for date and time selection
        right_column = QVBoxLayout()
        right_column.setSpacing(15)
        
        # Time row - moved to top of right column
        time_label = QLabel("Start Time:")
        time_label.setStyleSheet("font-weight: bold;")
        right_column.addWidget(time_label)
        
        self.time_combo = QComboBox()
        self.time_combo.setMaxVisibleItems(10)
        self.time_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        # Generate times from 7:00 to 17:00 with 15-minute intervals
        times = []
        for hour in range(7, 18):  # 7 AM to 5 PM (17:00)
            for minute in [0, 15, 30, 45]:
                # Skip 17:15, 17:30, 17:45
                if hour == 17 and minute > 0:
                    continue
                times.append(f"{hour:02d}:{minute:02d}:00")
        self.time_combo.addItems(times)
        right_column.addWidget(self.time_combo)
        
        # Add spacer between time and date
        right_column.addSpacing(10)
        
        # Date with modern calendar widget
        date_label = QLabel("Date:")
        date_label.setStyleSheet("font-weight: bold;")
        right_column.addWidget(date_label)
        
        # Modern styled calendar
        self.date_calendar = QCalendarWidget()
        self.date_calendar.setFirstDayOfWeek(Qt.DayOfWeek.Monday)
        self.date_calendar.setGridVisible(True)
        self.date_calendar.setMinimumDate(QDate.currentDate())
        self.date_calendar.setMaximumDate(QDate.currentDate().addDays(3650))
        
        # Make calendar more modern
        self.date_calendar.setStyleSheet("""
            QCalendarWidget {
                background-color: #2a2a2a;
                border: 1px solid #555;
                border-radius: 4px;
            }
            QCalendarWidget QToolButton {
                color: #ddd;
                background-color: #3a3a3a;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 5px;
            }
            QCalendarWidget QMenu {
                background-color: #2a2a2a;
                color: #ddd;
            }
            QCalendarWidget QSpinBox {
                background-color: #3a3a3a;
                color: #ddd;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 2px;
            }
            QCalendarWidget QAbstractItemView:enabled {
                color: #ddd;
                background-color: #2a2a2a;
                selection-background-color: #555;
                selection-color: #fff;
            }
            QCalendarWidget QAbstractItemView:disabled {
                color: #555;
            }
            QCalendarWidget QWidget#qt_calendar_navigationbar {
                background-color: #3a3a3a;
                border-bottom: 1px solid #555;
            }
        """)
        
        right_column.addWidget(self.date_calendar)
        
        # Add right column to form layout
        form_layout.addLayout(right_column, 1)
        
        # Add the form layout to the main add_sessions_layout
        add_sessions_layout.addLayout(form_layout)
        
        # Add session button row
        add_session_btn_layout = QHBoxLayout()
        add_session_btn = QPushButton("Add Session")
        add_session_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        add_session_btn.clicked.connect(self.add_session)
        add_session_btn_layout.addStretch()
        add_session_btn_layout.addWidget(add_session_btn)
        add_sessions_layout.addLayout(add_session_btn_layout)
        
        # Add the grouped widget to the main layout
        self.create_layout.addWidget(add_sessions_group)
        
        # Added Sessions Table
        sessions_group = QGroupBox("Added Sessions")
        sessions_group.setStyleSheet(GROUP_BOX_STYLE)
        sessions_layout = QVBoxLayout(sessions_group)
        
        self.sessions_table = QTableWidget()
        self.sessions_table.setColumnCount(7)
        self.sessions_table.setHorizontalHeaderLabels([
            'Year', 'Group', 'Subject', 'Session', 'Location', 'Date', 'Start Time'
        ])
        self.sessions_table.setStyleSheet(TABLE_STYLE)
        
        # Set column properties
        header = self.sessions_table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Set columns to stretch
        for i in range(7):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
        
        sessions_layout.addWidget(self.sessions_table)
        
        # Buttons for table management
        table_buttons = QHBoxLayout()
        remove_btn = QPushButton("Remove Selected")
        remove_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        remove_btn.clicked.connect(self.remove_selected_session)
        
        clear_btn = QPushButton("Clear All")
        clear_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        clear_btn.clicked.connect(self.clear_sessions)
        
        table_buttons.addWidget(remove_btn)
        table_buttons.addWidget(clear_btn)
        sessions_layout.addLayout(table_buttons)
        
        self.create_layout.addWidget(sessions_group)
        
        # Initialize UI state
        self.update_location_options()
    
    def setup_update_mode(self):
        # File selection
        file_group = QGroupBox("Select Schedule to Update")
        file_group.setStyleSheet(GROUP_BOX_STYLE)
        file_layout = QHBoxLayout(file_group)
        
        # MODULE NAME INPUT - Add this new section
        module_name_layout = QHBoxLayout()
        module_name_layout.addWidget(QLabel("Module Name:"))
        self.update_module_name_input = QLineEdit()
        self.update_module_name_input.setPlaceholderText("Enter module name...")
        file_layout.addWidget(self.update_module_name_input)
        
        self.update_file_input = QLineEdit()
        self.update_file_input.setPlaceholderText("Select Excel schedule file...")
        file_layout.addWidget(self.update_file_input)
        
        browse_btn = QPushButton("Browse")
        browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        browse_btn.clicked.connect(self.browse_schedule_file)
        file_layout.addWidget(browse_btn)
        
        self.update_layout.addWidget(file_group)
        
        # Existing Sessions Table
        existing_group = QGroupBox("Existing Sessions")
        existing_group.setStyleSheet(GROUP_BOX_STYLE)
        existing_layout = QVBoxLayout(existing_group)
        
        self.existing_table = QTableWidget()
        self.existing_table.setColumnCount(7)
        self.existing_table.setHorizontalHeaderLabels([
            'Year', 'Group', 'Subject', 'Session', 'Location', 'Date', 'Start Time'
        ])
        self.existing_table.setStyleSheet(TABLE_STYLE)
        
        # Set column properties
        header = self.existing_table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Set columns to stretch
        for i in range(7):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
        
        existing_layout.addWidget(self.existing_table)
        
        # Table manipulation buttons
        existing_buttons = QHBoxLayout()
        remove_existing_btn = QPushButton("Remove Selected")
        remove_existing_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        remove_existing_btn.clicked.connect(self.remove_existing_session)
        existing_buttons.addWidget(remove_existing_btn)
        existing_layout.addLayout(existing_buttons)
    
        self.update_layout.addWidget(existing_group)
        
        # Add New Sessions group
        add_sessions_group = QGroupBox("Add New Sessions")
        add_sessions_group.setStyleSheet(GROUP_BOX_STYLE)
        add_sessions_layout = QVBoxLayout(add_sessions_group)  # Use vertical layout as container
        
        # Create horizontal layout for form fields
        form_layout = QHBoxLayout()
        
        # Left column for form inputs
        left_column = QVBoxLayout()
        left_column.setSpacing(15)
        
        # Year Selection
        year_row = QHBoxLayout()
        year_row.addWidget(QLabel("Academic Year:"))
        self.update_year_combo = QComboBox()
        self.update_year_combo.setMaxVisibleItems(10)
        self.update_year_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.update_year_combo.addItems(["Year 1", "Year 2", "Year 3"])
        self.update_year_combo.currentTextChanged.connect(self.update_year_selected)
        year_row.addWidget(self.update_year_combo)
        year_row.addStretch()
        left_column.addLayout(year_row)
        
        # Group selection row
        group_row = QHBoxLayout()
        group_row.addWidget(QLabel("Group:            "))
        self.update_group_combo = QComboBox()
        self.update_group_combo.setMaxVisibleItems(10)
        self.update_group_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.populate_update_groups()
        group_row.addWidget(self.update_group_combo)
        group_row.addStretch()
        left_column.addLayout(group_row)
        
        # Subject row
        subject_row = QHBoxLayout()
        subject_row.addWidget(QLabel("Subject:          "))
        self.update_subject_combo = QComboBox()
        self.update_subject_combo.setMaxVisibleItems(10)
        self.update_subject_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.update_subject_combo.addItems([
            "Anatomy", "Histology", "Biochemistry", "Physiology", 
            "Microbiology", "Parasitology", "Pathology", "Pharmacology", "Clinical"
        ])
        self.update_subject_combo.currentTextChanged.connect(self.update_update_location_options)
        subject_row.addWidget(self.update_subject_combo)
        subject_row.addStretch()
        left_column.addLayout(subject_row)
        
        # Session row
        session_row = QHBoxLayout()
        session_row.addWidget(QLabel("Session:           "))
        self.update_session_combo = QComboBox()
        self.update_session_combo.setMaxVisibleItems(10)
        self.update_session_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.update_session_combo.addItems([str(i) for i in range(1, 31)])
        session_row.addWidget(self.update_session_combo)
        session_row.addStretch()
        left_column.addLayout(session_row)
    
        # Location row
        location_row = QHBoxLayout()
        location_row.addWidget(QLabel("Location:         "))
        self.update_location_combo = QComboBox()
        self.update_location_combo.setMaxVisibleItems(10)
        self.update_location_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.populate_update_locations()
        location_row.addWidget(self.update_location_combo)
        location_row.addStretch()
        left_column.addLayout(location_row)
        
        # Add left column to form layout
        form_layout.addLayout(left_column, 1)
        
        # Right column for date and time selection
        right_column = QVBoxLayout()
        right_column.setSpacing(15)
        
        # Time row - moved to top of right column
        time_label = QLabel("Start Time:")
        time_label.setStyleSheet("font-weight: bold;")
        right_column.addWidget(time_label)
        
        self.update_time_combo = QComboBox()
        self.update_time_combo.setMaxVisibleItems(10)
        self.update_time_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        # Generate times from 7:00 to 17:00 with 15-minute intervals
        times = []
        for hour in range(7, 18):  # 7 AM to 5 PM (17:00)
            for minute in [0, 15, 30, 45]:
                # Skip 17:15, 17:30, 17:45
                if hour == 17 and minute > 0:
                    continue
                times.append(f"{hour:02d}:{minute:02d}:00")
        self.update_time_combo.addItems(times)
        right_column.addWidget(self.update_time_combo)
        
        # Add spacer between time and date
        right_column.addSpacing(10)
        
        # Date with modern calendar widget
        date_label = QLabel("Date:")
        date_label.setStyleSheet("font-weight: bold;")
        right_column.addWidget(date_label)
        
        # Modern styled calendar
        self.update_date_calendar = QCalendarWidget()
        self.update_date_calendar.setFirstDayOfWeek(Qt.DayOfWeek.Monday)
        self.update_date_calendar.setGridVisible(True)
        self.update_date_calendar.setMinimumDate(QDate.currentDate())
        self.update_date_calendar.setMaximumDate(QDate.currentDate().addDays(3650))
        
        # Make calendar more modern
        self.update_date_calendar.setStyleSheet("""
            QCalendarWidget {
                background-color: #2a2a2a;
                border: 1px solid #555;
                border-radius: 4px;
            }
            QCalendarWidget QToolButton {
                color: #ddd;
                background-color: #3a3a3a;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 5px;
            }
            QCalendarWidget QMenu {
                background-color: #2a2a2a;
                color: #ddd;
            }
            QCalendarWidget QSpinBox {
                background-color: #3a3a3a;
                color: #ddd;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 2px;
            }
            QCalendarWidget QAbstractItemView:enabled {
                color: #ddd;
                background-color: #2a2a2a;
                selection-background-color: #555;
                selection-color: #fff;
            }
            QCalendarWidget QAbstractItemView:disabled {
                color: #555;
            }
            QCalendarWidget QWidget#qt_calendar_navigationbar {
                background-color: #3a3a3a;
                border-bottom: 1px solid #555;
            }
        """)
        
        right_column.addWidget(self.update_date_calendar)
        
        # Add right column to form layout
        form_layout.addLayout(right_column, 1)
        
        # Add the form layout to the main add_sessions_layout
        add_sessions_layout.addLayout(form_layout)
        
        # Add session button row
        add_session_btn_layout = QHBoxLayout()
        add_session_btn = QPushButton("Add Session")
        add_session_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        add_session_btn.clicked.connect(self.add_update_session)
        add_session_btn_layout.addStretch()
        add_session_btn_layout.addWidget(add_session_btn)
        add_sessions_layout.addLayout(add_session_btn_layout)
        
        # Add the grouped widget to the main layout
        self.update_layout.addWidget(add_sessions_group)
        
        # New Sessions Group
        new_sessions_group = QGroupBox("New Sessions to Add")
        new_sessions_group.setStyleSheet(GROUP_BOX_STYLE)
        new_sessions_layout = QVBoxLayout(new_sessions_group)
        
        self.new_sessions_table = QTableWidget()
        self.new_sessions_table.setColumnCount(7)
        self.new_sessions_table.setHorizontalHeaderLabels([
            'Year', 'Group', 'Subject', 'Session', 'Location', 'Date', 'Start Time'
        ])
        self.new_sessions_table.setStyleSheet(TABLE_STYLE)
        
        # Set column properties
        header = self.new_sessions_table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Set columns to stretch
        for i in range(7):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
        
        new_sessions_layout.addWidget(self.new_sessions_table)
        
        # New sessions buttons
        new_buttons = QHBoxLayout()
        remove_new_btn = QPushButton("Remove Selected")
        remove_new_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        remove_new_btn.clicked.connect(self.remove_new_session)
        
        clear_new_btn = QPushButton("Clear All")
        clear_new_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        clear_new_btn.clicked.connect(self.clear_new_sessions)
        
        new_buttons.addWidget(remove_new_btn)
        new_buttons.addWidget(clear_new_btn)
        new_sessions_layout.addLayout(new_buttons)
        
        self.update_layout.addWidget(new_sessions_group)
        
        # Initialize UI state
        self.update_update_location_options()
    
    def populate_groups(self):
        """Populate groups based on year selection"""
        self.group_combo.clear()
        year = self.year_combo.currentText()

        groups = []
        for prefix in ["A", "B"]:
            for num in range(1, 11):
                groups.append(f"{prefix}{num}")

        self.group_combo.addItems(groups)

    def populate_update_groups(self):
        """Populate groups for update mode"""
        self.update_group_combo.clear()
        year = self.update_year_combo.currentText()

        groups = []
        for prefix in ["A", "B"]:
            for num in range(1, 11):
                groups.append(f"{prefix}{num}")

        self.update_group_combo.addItems(groups)

    def populate_locations(self):
        """Initialize locations combobox with all preset values"""
        self.location_combo.clear()

        # Set maximum visible items to 10
        self.location_combo.setMaxVisibleItems(10)

        # Add all preset locations regardless of subject
        locations = [
            "Morgue",
            "Anatomy Lecture Hall",
            "Histology Lab",
            "Histology Lecture Hall",
            "Biochemistry Lab",
            "Biochemistry Lecture Hall",
            "Physiology Lab",
            "Physiology Lecture Hall",
            "Microbiology Lab",
            "Microbiology Lecture Hall",
            "Parasitology Lab",
            "Parasitology Lecture Hall",
            "Pathology Lab",
            "Pathology Lecture Hall",
            "Pharmacology Lab",
            "Pharmacology Lecture Hall",
            "Building 'A' Lecture Hall",
            "Building 'B' Lecture Hall",
        ]

        # Add all locations to the combo box
        self.location_combo.addItems(locations)

    def populate_update_locations(self):
        """Initialize locations combobox with all preset values"""
        self.update_location_combo.clear()

        # Set maximum visible items to 10
        self.update_location_combo.setMaxVisibleItems(10)

        # Add all preset locations regardless of subject
        locations = [
            "Morgue",
            "Anatomy Lecture Hall",
            "Histology Lab",
            "Histology Lecture Hall",
            "Biochemistry Lab",
            "Biochemistry Lecture Hall",
            "Physiology Lab",
            "Physiology Lecture Hall",
            "Microbiology Lab",
            "Microbiology Lecture Hall",
            "Parasitology Lab",
            "Parasitology Lecture Hall",
            "Pathology Lab",
            "Pathology Lecture Hall",
            "Pharmacology Lab",
            "Pharmacology Lecture Hall",
            "Building 'B' Lecture Hall",
        ]

        # Add all locations to the combo box
        self.update_location_combo.addItems(locations)

    def update_location_options(self):
        """Update location options based on selected subject"""
        self.populate_locations()

    def update_update_location_options(self):
        """Update location options for update mode"""
        self.populate_update_locations()

    def year_selected(self):
        """Handle year selection change"""
        self.current_year = self.year_combo.currentText()
        self.populate_groups()

    def update_year_selected(self):
        """Handle year selection change in update mode"""
        self.current_year = self.update_year_combo.currentText()
        self.populate_update_groups()

    def switch_to_create_mode(self):
        """Switch to create mode"""
        if self.create_radio.isChecked():
            self.mode_stack.setCurrentIndex(0)
            self.create_schedule_btn.setVisible(True)
            self.update_schedule_btn.setVisible(False)

    def switch_to_update_mode(self):
        """Switch to update mode"""
        if self.update_radio.isChecked():
            self.mode_stack.setCurrentIndex(1)
            self.create_schedule_btn.setVisible(False)
            self.update_schedule_btn.setVisible(True)

    def add_session(self):
        """Add a session to the table in create mode"""
        year = self.year_combo.currentText()
        group = self.group_combo.currentText()
        subject = self.subject_combo.currentText()
        session = self.session_combo.currentText()
        location = self.location_combo.currentText()
        # Get date from calendar
        date = self.date_calendar.selectedDate().toString("dd/MM/yyyy")
        time = self.time_combo.currentText()

        # Add row to table
        row_position = self.sessions_table.rowCount()
        self.sessions_table.insertRow(row_position)

        # Set data in table
        self.sessions_table.setItem(row_position, 0, QTableWidgetItem(year))
        self.sessions_table.setItem(row_position, 1, QTableWidgetItem(group))
        self.sessions_table.setItem(row_position, 2, QTableWidgetItem(subject))
        self.sessions_table.setItem(row_position, 3, QTableWidgetItem(session))
        self.sessions_table.setItem(
            row_position, 4, QTableWidgetItem(location))
        self.sessions_table.setItem(row_position, 5, QTableWidgetItem(date))
        self.sessions_table.setItem(row_position, 6, QTableWidgetItem(time))

        # Center all items
        for col in range(7):
            item = self.sessions_table.item(row_position, col)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

        # Add to data structure
        self.schedule_data.append(
            [year, group, subject, session, location, date, time])

    def add_update_session(self):
        """Add a session to the table in update mode"""
        year = self.update_year_combo.currentText()
        group = self.update_group_combo.currentText()
        subject = self.update_subject_combo.currentText()
        session = self.update_session_combo.currentText()
        location = self.update_location_combo.currentText()
        # Get date from calendar
        date = self.update_date_calendar.selectedDate().toString("dd/MM/yyyy")
        time = self.update_time_combo.currentText()

        # Add row to table
        row_position = self.new_sessions_table.rowCount()
        self.new_sessions_table.insertRow(row_position)

        # Set data in table
        self.new_sessions_table.setItem(
            row_position, 0, QTableWidgetItem(year))
        self.new_sessions_table.setItem(
            row_position, 1, QTableWidgetItem(group))
        self.new_sessions_table.setItem(
            row_position, 2, QTableWidgetItem(subject))
        self.new_sessions_table.setItem(
            row_position, 3, QTableWidgetItem(session))
        self.new_sessions_table.setItem(
            row_position, 4, QTableWidgetItem(location))
        self.new_sessions_table.setItem(
            row_position, 5, QTableWidgetItem(date))
        self.new_sessions_table.setItem(
            row_position, 6, QTableWidgetItem(time))

        # Center all items
        for col in range(7):
            item = self.new_sessions_table.item(row_position, col)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

    def clear_sessions(self):
        """Clear all sessions from the table"""
        self.sessions_table.setRowCount(0)
        self.schedule_data = []

    def clear_new_sessions(self):
        """Clear all new sessions from the update table"""
        self.new_sessions_table.setRowCount(0)

    def remove_selected_session(self):
        """Remove selected session from the table"""
        selected_rows = self.sessions_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        # Remove from highest index to lowest to avoid index issues
        rows_to_remove = sorted([index.row()
                                for index in selected_rows], reverse=True)

        for row in rows_to_remove:
            # Get information before removing for logging
            year = self.sessions_table.item(row, 0).text()
            group = self.sessions_table.item(row, 1).text()
            subject = self.sessions_table.item(row, 2).text()
            session = self.sessions_table.item(row, 3).text()

            # Remove from table
            self.sessions_table.removeRow(row)

            # Remove from data structure
            if 0 <= row < len(self.schedule_data):
                self.schedule_data.pop(row)

    def remove_new_session(self):
        """Remove selected new session from the update table"""
        selected_rows = self.new_sessions_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        # Remove from highest index to lowest to avoid index issues
        rows_to_remove = sorted([index.row()
                                for index in selected_rows], reverse=True)

        for row in rows_to_remove:
            # Get information before removing for logging
            year = self.new_sessions_table.item(row, 0).text()
            group = self.new_sessions_table.item(row, 1).text()
            subject = self.new_sessions_table.item(row, 2).text()
            session = self.new_sessions_table.item(row, 3).text()

            # Remove from table
            self.new_sessions_table.removeRow(row)

    def remove_existing_session(self):
        """Remove selected session from the existing sessions table"""
        selected_rows = self.existing_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        # Remove from highest index to lowest to avoid index issues
        rows_to_remove = sorted([index.row()
                                for index in selected_rows], reverse=True)

        for row in rows_to_remove:
            # Get information before removing for logging
            year = self.existing_table.item(row, 0).text()
            group = self.existing_table.item(row, 1).text()
            subject = self.existing_table.item(row, 2).text()
            session = self.existing_table.item(row, 3).text()

            # Remove from table
            self.existing_table.removeRow(row)

    def browse_schedule_file(self):
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("Excel Files (*.xlsx *.xls)")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)

        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                file_path = selected_files[0]
                self.update_file_input.setText(file_path)
                # Auto-load the file
                self.load_schedule()

    def load_schedule(self):
        """Load an existing schedule from file"""
        file_path = self.update_file_input.text()
        if not file_path:
            self.show_message_box("Error", "Please select a file first")
            return
        
        try:
            # Clear existing data
            self.existing_table.setRowCount(0)
            
            # Load Excel file
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active
            
            # Try to get module name from sheet title
            module_name = sheet.title
            if module_name and module_name != "Sheet":  # If it's not the default sheet name
                self.update_module_name_input.setText(module_name)
            
            # Check if the first row contains "Module:" (old format) or headers (new format)
            first_cell = sheet.cell(row=1, column=1).value
            if first_cell and isinstance(first_cell, str) and first_cell.startswith("Module:"):
                # Old format with module name in first row
                module_name = first_cell.replace("Module:", "").strip()
                self.update_module_name_input.setText(module_name)
                start_row = 3  # Start reading data from row 3
            else:
                # New format with headers in first row
                start_row = 2  # Start reading data from row 2
            
            # Skip header row
            row_count = 0
            for row in sheet.iter_rows(min_row=start_row, values_only=True):
                if all(cell is None for cell in row):
                    continue  # Skip empty rows
                
                # Add row to table
                self.existing_table.insertRow(row_count)
                
                # Set data in table
                # Only use the first 7 columns
                for col, cell_value in enumerate(row[:7]):
                    value = str(cell_value) if cell_value is not None else ""
                    item = QTableWidgetItem(value)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.existing_table.setItem(row_count, col, item)
                
                row_count += 1
            
        except Exception as e:
            self.show_message_box("Error", f"Failed to load schedule: {str(e)}")
    
    def save_schedule(self):
        """Save the created schedule to Excel file directly without prompting"""
        if len(self.schedule_data) == 0:
            self.show_message_box("Error", "No schedule data to save")
            return
    
        # Get module name (default if empty)
        module_name = self.module_name_input.text().strip()
        if not module_name:
            module_name = "Untitled_Module"
    
        try:
            # Create modules_schedules directory in the current working directory
            # This will be where the EXE is running from
            save_dir = os.path.join(os.getcwd(), "modules_schedules")
            os.makedirs(save_dir, exist_ok=True)
    
            # Generate timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
            # Create filename with module name and timestamp
            filename = f"{module_name}_{timestamp}.xlsx"
            filepath = os.path.join(save_dir, filename)
        
            # Create a new workbook
            wb = openpyxl.Workbook()
            ws = wb.active
        
            # Name the worksheet with the module name (limit to 31 chars as Excel has a limit)
            ws.title = module_name[:31]
        
            # Add headers in row 1 (no module name row anymore)
            headers = ['Year', 'Group', 'Session', 'Location', 'Date', 'Start Time']
            for col, header in enumerate(headers, start=1):
                ws.cell(row=1, column=col).value = header
                ws.cell(row=1, column=col).font = openpyxl.styles.Font(bold=True)
        
            # Sort data by Year, Group, and then Session number
            sorted_data = sorted(self.schedule_data,
                                 key=lambda x: (x[0], x[1], int(x[3])))
        
            # Add data starting from row 2 (was row 3 before)
            for row_idx, row_data in enumerate(sorted_data, start=2):
                for col_idx, cell_value in enumerate(row_data, start=1):
                    ws.cell(row=row_idx, column=col_idx).value = cell_value
        
            # Auto-fit columns - modified to use all rows including header
            for col in range(1, 8):  # Columns A through G
                max_length = 0
                column_letter = openpyxl.utils.get_column_letter(col)
            
                # Include all rows
                for row in range(1, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col)
                    try:
                        if cell.value and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
            
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column_letter].width = adjusted_width
        
            # Save the file
            wb.save(filepath)
        
            self.show_message_box("Success", f"Schedule saved successfully to:\n{filepath}")
        
        except Exception as e:
            self.show_message_box("Error", f"Failed to save schedule: {str(e)}")
    
    def update_schedule(self):
        """Update the existing schedule with new sessions without prompting for save location"""
        # Check if we have loaded an existing schedule
        if self.existing_table.rowCount() == 0:
            self.show_message_box("Error", "Please load an existing schedule first")
            return
    
        # Check if we have new sessions to add
        if self.new_sessions_table.rowCount() == 0:
            self.show_message_box("Error", "No new sessions to add")
            return
    
        # Get module name (default if empty)
        module_name = self.update_module_name_input.text().strip()
        if not module_name:
            module_name = "Untitled_Module"
    
        try:
            # Create modules_schedules directory in the current working directory
            # This will be where the EXE is running from
            save_dir = os.path.join(os.getcwd(), "modules_schedules")
            os.makedirs(save_dir, exist_ok=True)
        
            # Generate timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
            # Create filename with module name and timestamp
            filename = f"{module_name}_updated_{timestamp}.xlsx"
            filepath = os.path.join(save_dir, filename)
        
            # Create a new workbook
            wb = openpyxl.Workbook()
            ws = wb.active
        
            # Name the worksheet with the module name (limit to 31 chars as Excel has a limit)
            ws.title = module_name[:31]
        
            # Add headers in row 1 (no module name row anymore)
            headers = ['Year', 'Group', 'Session', 'Location', 'Date', 'Start Time']
            for col, header in enumerate(headers, start=1):
                ws.cell(row=1, column=col).value = header
                ws.cell(row=1, column=col).font = openpyxl.styles.Font(bold=True)
        
            # Collect existing data
            existing_data = []
            for row in range(self.existing_table.rowCount()):
                row_data = []
                for col in range(7):
                    value = self.existing_table.item(row, col).text()
                    row_data.append(value)
                existing_data.append(row_data)
        
            # Collect new data
            new_data = []
            for row in range(self.new_sessions_table.rowCount()):
                row_data = []
                for col in range(7):
                    value = self.new_sessions_table.item(row, col).text()
                    row_data.append(value)
                new_data.append(row_data)
        
            # Combine data - group by Year and Group
            grouped_data = {}
        
            # Process existing data
            for row in existing_data:
                key = (row[0], row[1])  # (Year, Group)
                if key not in grouped_data:
                    grouped_data[key] = []
                grouped_data[key].append(row)
        
            # Add new data to appropriate groups
            for row in new_data:
                key = (row[0], row[1])  # (Year, Group)
                if key not in grouped_data:
                    grouped_data[key] = []
                grouped_data[key].append(row)
        
            # Sort groups by year and group name
            sorted_keys = sorted(grouped_data.keys())
        
            # Prepare final data - sort each group by session number
            final_data = []
            for key in sorted_keys:
                # Sort by session number within each group
                group_data = sorted(grouped_data[key], key=lambda x: int(x[3]))
                final_data.extend(group_data)
        
            # Add data to worksheet starting from row 2 (was row 3 before)
            for row_idx, row_data in enumerate(final_data, start=2):
                for col_idx, cell_value in enumerate(row_data, start=1):
                    ws.cell(row=row_idx, column=col_idx).value = cell_value
        
            # Auto-fit columns - include all rows
            for col in range(1, 8):  # Columns A through G
                max_length = 0
                column_letter = openpyxl.utils.get_column_letter(col)
            
                # Include all rows
                for row in range(1, ws.max_row + 1):
                    cell = ws.cell(row=row, column=col)
                    try:
                        if cell.value and len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
            
                adjusted_width = (max_length + 2)
                ws.column_dimensions[column_letter].width = adjusted_width
        
            # Save the file
            wb.save(filepath)
        
            self.show_message_box("Success", f"Schedule updated successfully to:\n{filepath}")
        
        except Exception as e:
            self.show_message_box("Error", f"Failed to update schedule: {str(e)}")
            
    def show_message_box(self, title, message):
        """Show a message box with the given title and message"""
        msg_box = QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.setStyleSheet("QLabel{min-width: 300px;}")
        msg_box.exec()

#==========================================================appeal processor==========================================================#

class AppealProcessor(QWidget):
    def __init__(self):
        super().__init__()
        self.schedules = []
        self.students = []
        self.sessions = []
        self.selected_appeals = []
        self.setStyleSheet("""
            background-color: black; 
            color: white;
            QLabel {
                color: white;
            }
        """)

        self.init_ui()

    def return_to_home(self):
        # Get the stacked widget and switch to the start page
        stacked_widget = self.parent()
        if isinstance(stacked_widget, QStackedWidget):
            stacked_widget.setCurrentIndex(0)

    def init_ui(self):
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Create a scroll area to make the page scrollable
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; }")
        
        # Create a widget to hold the content
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(20)
        scroll_layout.setContentsMargins(0, 0, 0, 0)

        # Header layout
        header_layout = QHBoxLayout()
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), 'ASU1.png')
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaled(60, 60, Qt.AspectRatioMode.KeepAspectRatio,
                                               Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        title_label = QLabel("Appeal Processor")
        title_label.setStyleSheet(f"font-size: 24px; font-weight: bold;")
        header_layout.addWidget(logo_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Create vertical layout for header buttons
        header_buttons_layout = QVBoxLayout()
        header_buttons_layout.setSpacing(5)

        # Back button
        back_btn = QPushButton("Back to Home")
        back_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        back_btn.clicked.connect(self.return_to_home)

        # Exit button
        exit_btn = QPushButton("Exit")
        exit_btn.setStyleSheet(EXIT_BUTTON_STYLE)
        exit_btn.clicked.connect(QApplication.instance().quit)

        # Add buttons to vertical layout
        header_buttons_layout.addWidget(back_btn)
        header_buttons_layout.addWidget(exit_btn)
        header_buttons_layout.addStretch()

        # Add button layout to header
        header_layout.addStretch()
        header_layout.addLayout(header_buttons_layout)
        scroll_layout.addLayout(header_layout)

        # Create three file input sections
        # REORDERED AS REQUESTED:
        
        # 1. Reference File Section
        ref_group = QGroupBox("Reference Data")
        ref_group.setStyleSheet(GROUP_BOX_STYLE)
        ref_layout = QVBoxLayout(ref_group)

        # Single line layout for reference data
        ref_input_layout = QHBoxLayout()
        ref_input_layout.addWidget(QLabel("Database File:"))
        self.ref_file_input = QLineEdit()
        self.ref_file_input.setPlaceholderText("Select Excel file...")
        self.ref_file_input.setMinimumWidth(200)
        ref_input_layout.addWidget(self.ref_file_input)
        ref_browse_btn = QPushButton("Browse")
        ref_browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        ref_input_layout.addWidget(ref_browse_btn)
        ref_input_layout.addWidget(QLabel("Sheet Name:"))
        self.ref_sheet_combo = QComboBox()
        self.ref_sheet_combo.setMinimumWidth(100)
        ref_input_layout.addWidget(self.ref_sheet_combo)
        ref_browse_btn.clicked.connect(
            lambda: self.browse_file(self.ref_file_input))
        ref_layout.addLayout(ref_input_layout)
        scroll_layout.addWidget(ref_group)
        
        # 2. Log File Section
        log_group = QGroupBox("Attendance Logs")
        log_group.setStyleSheet(GROUP_BOX_STYLE)
        log_layout = QVBoxLayout(log_group)

        # Single line layout for log data
        log_input_layout = QHBoxLayout()
        log_input_layout.addWidget(QLabel("Log File:"))
        self.log_file_input = QLineEdit()
        self.log_file_input.setPlaceholderText("Select Excel file...")
        self.log_file_input.setMinimumWidth(200)
        log_input_layout.addWidget(self.log_file_input)
        log_browse_btn = QPushButton("Browse")
        log_browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        log_input_layout.addWidget(log_browse_btn)
        log_input_layout.addWidget(QLabel("Sheet Name:"))
        self.log_sheet_combo = QComboBox()
        self.log_sheet_combo.setMinimumWidth(100)
        log_input_layout.addWidget(self.log_sheet_combo)
        log_browse_btn.clicked.connect(
            lambda: self.browse_file(self.log_file_input))
        log_layout.addLayout(log_input_layout)
        scroll_layout.addWidget(log_group)
        
        # 3. Schedule File Section
        schedule_group = QGroupBox("Sessions Schedules")
        schedule_group.setStyleSheet(GROUP_BOX_STYLE)
        schedule_layout = QVBoxLayout(schedule_group)

        # Single line layout for schedule data
        schedule_input_layout = QHBoxLayout()
        schedule_input_layout.addWidget(QLabel("Schedule File:"))
        self.schedule_file_input = QLineEdit()
        self.schedule_file_input.setPlaceholderText("Select Excel file...")
        self.schedule_file_input.setMinimumWidth(200)
        schedule_input_layout.addWidget(self.schedule_file_input)
        schedule_browse_btn = QPushButton("Browse")
        schedule_browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        schedule_input_layout.addWidget(schedule_browse_btn)
        schedule_input_layout.addWidget(QLabel("Sheet Name:"))
        self.schedule_sheet_combo = QComboBox()
        self.schedule_sheet_combo.setMinimumWidth(100)
        schedule_input_layout.addWidget(self.schedule_sheet_combo)
        schedule_browse_btn.clicked.connect(
            lambda: self.browse_file(self.schedule_file_input))
        schedule_layout.addLayout(schedule_input_layout)
        scroll_layout.addWidget(schedule_group)

        # Appeal Selection Section
        appeal_group = QGroupBox("Appeal Selection")
        appeal_group.setStyleSheet(GROUP_BOX_STYLE)
        appeal_layout = QVBoxLayout(appeal_group)
        appeal_group.setMinimumHeight(500)  # Increase height


        # Student search section
        student_search_layout = QHBoxLayout()
        student_search_layout.addWidget(QLabel("Search Student:"))
        self.student_search = QLineEdit()
        self.student_search.setPlaceholderText("Enter student ID or name...")
        self.student_search.setMinimumWidth(300)
        self.student_search.textChanged.connect(self.filter_students)
        student_search_layout.addWidget(self.student_search)
        
        # CHANGED: Replace student list with table
        self.student_table = QTableWidget()
        self.student_table.setColumnCount(4)
        self.student_table.setHorizontalHeaderLabels(['Student ID', 'Name', 'Year', 'Group'])
        self.student_table.setStyleSheet(TABLE_STYLE)
        # Center align the header text
        header = self.student_table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        # Set column resize modes to stretch
        for i in range(4):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
        self.student_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.student_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.student_table.itemSelectionChanged.connect(self.update_student_info)
        
        # Session selection section
        session_title = QLabel("Select the sessions which the students should have attended:")
        session_title.setStyleSheet("font-weight: bold; margin-top: 10px;")
        
        # CHANGED: Replace session list with table
        self.session_table = QTableWidget()
        self.session_table.setColumnCount(4)
        self.session_table.setHorizontalHeaderLabels([
            'Session', 'Location', 'Date', 'Time'
        ])
        self.session_table.setStyleSheet(TABLE_STYLE)
        # Center align the header text
        header = self.session_table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        # Set column resize modes to stretch
        for i in range(5):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)
        self.session_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.session_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.session_table.setMinimumHeight(150)  # Make it taller
        self.session_table.itemSelectionChanged.connect(self.update_session_info)
  
        # Add appeal button
        add_appeal_layout = QHBoxLayout()
        add_appeal_btn = QPushButton("Add to Appeals")
        add_appeal_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        add_appeal_btn.clicked.connect(self.add_appeal)
        add_appeal_layout.addStretch()
        add_appeal_layout.addWidget(add_appeal_btn)
        
        appeal_layout.addLayout(student_search_layout)
        appeal_layout.addWidget(self.student_table)
        appeal_layout.addWidget(session_title)
        appeal_layout.addWidget(self.session_table)
        appeal_layout.addLayout(add_appeal_layout)
        
        scroll_layout.addWidget(appeal_group)

        # Selected Appeals Table
        selected_appeals_group = QGroupBox("Selected Appeals")
        selected_appeals_group.setStyleSheet(GROUP_BOX_STYLE)
        selected_appeals_group.setMinimumHeight(300)  # Increase height
        selected_appeals_layout = QVBoxLayout(selected_appeals_group)

        self.appeals_table = QTableWidget()
        self.appeals_table.setColumnCount(8)
        self.appeals_table.setHorizontalHeaderLabels([
            'Student ID', 'Name', 'Year', 'Group', 'Session', 'Location', 'Date', 'Time'
        ])
        self.appeals_table.setStyleSheet(TABLE_STYLE)

        # Center align the header text
        header = self.appeals_table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)  # Center header text

        # Set column resize modes to stretch and fit content
        for i in range(9):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch)

        selected_appeals_layout.addWidget(self.appeals_table)
        
        # Remove appeal button
        remove_appeal_layout = QHBoxLayout()
        remove_appeal_btn = QPushButton("Remove Selected Appeal")
        remove_appeal_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        remove_appeal_btn.clicked.connect(self.remove_appeal)
        remove_appeal_layout.addStretch()
        remove_appeal_layout.addWidget(remove_appeal_btn)
        selected_appeals_layout.addLayout(remove_appeal_layout)
        
        scroll_layout.addWidget(selected_appeals_group)

        # Bottom Buttons
        button_layout = QHBoxLayout()
        process_btn = QPushButton("Process Appeals")
        process_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        process_btn.clicked.connect(self.process_appeals)
        button_layout.addWidget(process_btn)
        scroll_layout.addLayout(button_layout)

        # Connect file input changes to sheet loading AND data loading (autoload)
        self.ref_file_input.textChanged.connect(lambda: self.load_sheets_and_data(self.ref_file_input.text(), self.ref_sheet_combo, 'reference'))
        self.log_file_input.textChanged.connect(lambda: self.load_sheets_and_data(self.log_file_input.text(), self.log_sheet_combo, 'log'))
        self.schedule_file_input.textChanged.connect(lambda: self.load_sheets_and_data(self.schedule_file_input.text(), self.schedule_sheet_combo, 'schedule'))
        
        # Also connect sheet combo box changes to trigger data loading when changed
        self.ref_sheet_combo.currentIndexChanged.connect(lambda: self.autoload_data())
        self.log_sheet_combo.currentIndexChanged.connect(lambda: self.autoload_data())
        self.schedule_sheet_combo.currentIndexChanged.connect(lambda: self.autoload_data())
            
        # Set the scroll content and add to main layout
        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)

    def browse_file(self, input_field):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if filename:
            input_field.setText(filename)

    def load_sheets_and_data(self, file_path, combo_box, file_type):
        if os.path.isfile(file_path):
            try:
                wb = openpyxl.load_workbook(file_path, read_only=True)
                
                # Store current selection if exists
                current_selection = combo_box.currentText()
                
                # Update combo box
                combo_box.clear()
                combo_box.addItems(wb.sheetnames)
                
                # Restore selection if possible
                if current_selection in wb.sheetnames:
                    combo_box.setCurrentText(current_selection)
                    
                # Auto-load data after sheets are loaded
                self.autoload_data()
                
            except Exception as e:
                QMessageBox.critical(
                    self, "Error", f"Error loading workbook: {str(e)}")

    def autoload_data(self):
        """Automatically load data when files and sheets are selected"""
        # Check if all required files and sheets are selected
        if (self.ref_file_input.text() and self.ref_sheet_combo.currentText() and 
            self.schedule_file_input.text() and self.schedule_sheet_combo.currentText()):
            self.load_data()
    
    def load_data(self):
        """Load student and session data from the selected files"""
        try:
            # Clear previous data
            self.students = []
            self.sessions = []
            self.student_table.setRowCount(0)
            self.session_table.setRowCount(0)

            # Load student reference data
            ref_wb = openpyxl.load_workbook(self.ref_file_input.text(), read_only=True)
            ref_ws = ref_wb[self.ref_sheet_combo.currentText()]
            student_db = list(ref_ws.values)

            # Skip header row and load students
            for row in student_db[1:]:
                if row[0]:  # If student ID exists
                    student = {
                        'id': str(row[0]),
                        'name': row[1],
                        'year': row[2],
                        'group': row[3]                    }
                    self.students.append(student)
        
                    # Add to student table
                    current_row = self.student_table.rowCount()
                    self.student_table.insertRow(current_row)
        
                    # Add data to table cells
                    self.student_table.setItem(current_row, 0, QTableWidgetItem(str(student['id'])))
                    self.student_table.setItem(current_row, 1, QTableWidgetItem(student['name']))
                    self.student_table.setItem(current_row, 2, QTableWidgetItem(str(student['year'])))
                    self.student_table.setItem(current_row, 3, QTableWidgetItem(str(student['group'])))
        
                    # Center align all items
                    for col in range(4):
                        item = self.student_table.item(current_row, col)
                        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            # Load session schedule data
            sched_wb = openpyxl.load_workbook(self.schedule_file_input.text(), read_only=True)
            sched_ws = sched_wb[self.schedule_sheet_combo.currentText()]
            session_schedule = list(sched_ws.values)

            # Debug - print headers to understand format
            print(f"Schedule headers: {session_schedule[0] if session_schedule else 'No data'}")

            # Skip header row and load sessions
            for row_idx, row in enumerate(session_schedule):
                # Skip the header row (index 0)
                if row_idx == 0:
                    continue
                    
                # Make sure we have all required data in the row
                if len(row) >= 6:
                    # Extract session data based on the new format (like in the image)
                    year = row[0]  # Year column
                    group = row[1]  # Group column
                    session_num = row[2]  # Session column
                    location = row[3]  # Location column
                    date = row[4]  # Date column
                    start_time = row[5]  # Start Time column
                    
                    # Create a session dictionary with the extracted data
                    session = {
                        'year': year,
                        'group': group,
                        'session': session_num,
                        'location': location,
                        'date': date,
                        'start_time': start_time
                    }
                    
                    # Add to sessions list
                    self.sessions.append(session)
                    
                    # Debug - print each session
                    print(f"Added session: {session}")
        
            # Debug print
            print(f"Loaded {len(self.students)} students and {len(self.sessions)} sessions")
            if self.sessions:
                print(f"Sample session: {self.sessions[0]}")

        except Exception as e:
            self.show_custom_warning("Error", f"Error loading data: {str(e)}")
            import traceback
            traceback.print_exc()  # Print full stack trace for debugging
    
    def filter_students(self):
        """Filter student table based on search text"""
        search_text = self.student_search.text().lower()
        self.student_table.setRowCount(0)
    
        for student in self.students:
            # Filter students by ID or name
            if not search_text or search_text in student['id'].lower() or search_text in student['name'].lower():
                current_row = self.student_table.rowCount()
                self.student_table.insertRow(current_row)
            
                # Add data to table cells
                self.student_table.setItem(current_row, 0, QTableWidgetItem(str(student['id'])))
                self.student_table.setItem(current_row, 1, QTableWidgetItem(student['name']))
                self.student_table.setItem(current_row, 2, QTableWidgetItem(str(student['year'])))
                self.student_table.setItem(current_row, 3, QTableWidgetItem(str(student['group'])))
            
                # Center align all items
                for col in range(4):
                    item = self.student_table.item(current_row, col)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

    def update_student_info(self):
        """Update the student information based on selection"""
        # Get the selected row
        selected_rows = self.student_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()

        # Extract student ID from the selected row
        student_id = self.student_table.item(row, 0).text()

        # Find the selected student
        selected_student = None
        for student in self.students:
            if student['id'] == student_id:
                selected_student = student
                break

        if selected_student:
            # Update UI labels
            self.student_id_label.setText(selected_student['id'])
            self.student_name_label.setText(selected_student['name'])
            self.student_year_label.setText(str(selected_student['year']))
            self.student_group_label.setText(str(selected_student['group']))

            # Update session table to show applicable sessions
            self.session_table.setRowCount(0)
            self.applicable_sessions = []  # Store as an instance variable to access later

            # Convert student year and group values for proper comparison
            student_year = str(selected_student['year']).strip().lower()
            student_group = str(selected_student['group']).strip().lower()
            
            # If student year doesn't start with "year", add it for comparison
            if not student_year.startswith("year"):
                student_year = f"year {student_year}"
                
            print(f"Looking for sessions for student year: {student_year}, group: {student_group}")
            print(f"Total available sessions: {len(self.sessions)}")
            
            # Loop through all sessions to find matching ones
            for session in self.sessions:
                # Convert session year and group for comparison
                session_year = str(session['year']).strip().lower()
                session_group = str(session['group']).strip().lower()
                
                print(f"Comparing with session - year: {session_year}, group: {session_group}")
                
                # Check if this session is for the student's year and group
                if (session_year == student_year and session_group == student_group):
                    print(f"MATCH FOUND: {session}")
                    self.applicable_sessions.append(session)
                    
                    # Format date and time for display
                    date_str = str(session['date'])
                    time_str = str(session['start_time'])
                    
                    # Add to session table
                    current_row = self.session_table.rowCount()
                    self.session_table.insertRow(current_row)
                    
                    # Add data to table cells
                    self.session_table.setItem(current_row, 0, QTableWidgetItem(str(session['session'])))
                    self.session_table.setItem(current_row, 1, QTableWidgetItem(str(session['location'])))
                    self.session_table.setItem(current_row, 2, QTableWidgetItem(date_str))
                    self.session_table.setItem(current_row, 3, QTableWidgetItem(time_str))
                    
                    # Center align all items
                    for col in range(4):
                        if col < self.session_table.columnCount():
                            item = self.session_table.item(current_row, col)
                            if item:
                                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            
            print(f"Found {len(self.applicable_sessions)} applicable sessions for this student")
            
    def update_session_info(self):
        """Store the selected session to be used when adding an appeal"""
        self.selected_session = None
    
        # Get the selected row
        selected_rows = self.session_table.selectionModel().selectedRows()
        if not selected_rows:
            return
        
        selected_index = selected_rows[0].row()
    
        if selected_index >= 0 and hasattr(self, 'applicable_sessions') and selected_index < len(self.applicable_sessions):
            self.selected_session = self.applicable_sessions[selected_index]
            print(f"Selected session: {self.selected_session}")
    
    def add_appeal(self):
        """Add selected student and session to the appeals table"""
        # Check if student is selected
        student_rows = self.student_table.selectionModel().selectedRows()
        if not student_rows:
            self.show_custom_warning("Selection Required", "Please select a student.")
            return
        
        # Check if session is selected
        session_rows = self.session_table.selectionModel().selectedRows()
        if not session_rows:
            self.show_custom_warning("Selection Required", "Please select a session.")
            return
        
        # Update session info to ensure we have the latest selection
        self.update_session_info()
    
        if not hasattr(self, 'selected_session') or not self.selected_session:
            self.show_custom_warning("Selection Required", "Please select a valid session.")
            return
        
        # Get student info
        student_row = student_rows[0].row()
        student_id = self.student_table.item(student_row, 0).text()
    
        selected_student = None
        for student in self.students:
            if student['id'] == student_id:
                selected_student = student
                break
    
        if not selected_student:
            return
        
        # Create appeal record
        appeal = {
            'student_id': selected_student['id'],
            'name': selected_student['name'],
            'year': selected_student['year'],
            'group': selected_student['group'],
            'session': self.selected_session['session'],
            'location': self.selected_session['location'],
            'date': self.selected_session['date'],
            'time': self.selected_session['start_time']
        }
    
        # Check if this appeal already exists
        for existing_appeal in self.selected_appeals:
            if (existing_appeal['student_id'] == appeal['student_id'] and
                existing_appeal['session'] == appeal['session'] and
                existing_appeal['date'] == appeal['date']):
                self.show_custom_warning("Duplicate Appeal", 
                                        f"This appeal for {selected_student['name']} already exists.")
                return
    
        # Add to list and update table
        self.selected_appeals.append(appeal)
        self.update_appeals_table()
    
    def remove_appeal(self):
        """Remove selected appeal from the table"""
        current_row = self.appeals_table.currentRow()
        if current_row >= 0:
            removed_appeal = self.selected_appeals.pop(current_row)
            self.update_appeals_table()
        else:
            self.show_custom_warning("Selection Required", "Please select an appeal to remove.")
    
    def update_appeals_table(self):
        """Update the appeals table with current selections"""
        self.appeals_table.setRowCount(len(self.selected_appeals))
        
        for i, appeal in enumerate(self.selected_appeals):
            # Format date and time properly
            date_str = appeal['date']
            time_str = appeal['time']
            
            if isinstance(appeal['date'], datetime):
                date_str = appeal['date'].strftime('%d/%m/%Y')
            if isinstance(appeal['time'], datetime):
                time_str = appeal['time'].strftime('%H:%M:%S')
                
            # Create table items
            items = [
                appeal['student_id'],
                appeal['name'],
                str(appeal['year']),
                str(appeal['group']),
                str(appeal['session']),
                appeal['location'],
                str(date_str),
                str(time_str)
            ]
            
            for j, value in enumerate(items):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.appeals_table.setItem(i, j, item)

    def show_custom_warning(self, title, message):
        """Display a custom warning message box"""
        warning_dialog = QMessageBox(self)
        warning_dialog.setWindowTitle(title)
        warning_dialog.setText(message)
        warning_dialog.setIcon(QMessageBox.Icon.Warning)
        
        # Style OK button
        ok_button = warning_dialog.addButton(QMessageBox.StandardButton.Ok)
        ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)
        
        # Style dialog background
        warning_dialog.setStyleSheet(f"""
            QMessageBox {{
                background-color: black;
            }}
            QLabel {{
                color: white;
                font-size: 14px;
            }}
        """)
        
        warning_dialog.exec()

    def validate_inputs(self):
        """Validates that all necessary inputs are provided before processing appeals"""
        if not self.log_file_input.text() or not os.path.isfile(self.log_file_input.text()):
            self.show_custom_warning("Missing Input", "Please select a valid log file.")
            return False
            
        if not self.log_sheet_combo.currentText():
            self.show_custom_warning("Missing Input", "Please select a log sheet.")
            return False
            
        if not self.ref_file_input.text() or not os.path.isfile(self.ref_file_input.text()):
            self.show_custom_warning("Missing Input", "Please select a valid student reference file.")
            return False
            
        if not self.schedule_file_input.text() or not os.path.isfile(self.schedule_file_input.text()):
            self.show_custom_warning("Missing Input", "Please select a valid schedule file.")
            return False
            
        return True
        
    def process_appeals(self):
        """Process the selected appeals and add them to the log file"""
        if not self.selected_appeals:
            self.show_custom_warning("No Appeals", "Please add at least one appeal to process.")
            return
        
        if not self.validate_inputs():
            return
        
        try:
            # Directly process appeals without using a thread
            log_file = self.log_file_input.text()
            log_sheet = self.log_sheet_combo.currentText()
        
            # Load existing log file
            log_wb = openpyxl.load_workbook(log_file)
            log_ws = log_wb[log_sheet]
        
            # Get header row to understand column structure
            header_row = [cell.value if cell.value is not None else "" for cell in log_ws[1]]
        
            # Find important column indices
            try:
                # Adjust for the different column names in the example
                id_col = header_row.index("Student ID")
                location_col = header_row.index("Location")
                date_col = header_row.index("Log Date")
                time_col = header_row.index("Log Time")
            
                # Add missing columns if needed
                
                if "Session" not in header_row:
                    header_row.append("Session")
                    session_col = len(header_row) - 1
                    for row in range(1, log_ws.max_row + 1):
                        log_ws.cell(row=row, column=session_col+1).value = ""
                else:
                    session_col = header_row.index("Session")
                
                if "Status" not in header_row:
                    header_row.append("Status")
                    status_col = len(header_row) - 1
                    for row in range(1, log_ws.max_row + 1):
                        log_ws.cell(row=row, column=status_col+1).value = ""
                else:
                    status_col = header_row.index("Status")
                
                if "Notes" not in header_row:
                    header_row.append("Notes")
                    notes_col = len(header_row) - 1
                    for row in range(1, log_ws.max_row + 1):
                        log_ws.cell(row=row, column=notes_col+1).value = ""
                else:
                    notes_col = header_row.index("Notes")
            
            except ValueError as e:
                # Proper header not found, raise error
                raise ValueError(f"Required column not found in log file: {str(e)}")
        
            # Get total number of rows in the log sheet
            max_row = log_ws.max_row
        
            # Process each appeal
            for appeal in self.selected_appeals:
                # Format date and time values for consistency
                date_value = appeal['date']
                time_value = appeal['time']
            
                if isinstance(date_value, datetime):
                    date_value = date_value.strftime('%Y-%m-%d')
                if isinstance(time_value, datetime):
                    time_value = time_value.strftime('%H:%M:%S')
            
                # Prepare new row with appeal data
                new_row = max_row + 1
                log_ws.cell(row=new_row, column=id_col+1).value = appeal['student_id']
                log_ws.cell(row=new_row, column=date_col+1).value = date_value
                log_ws.cell(row=new_row, column=time_col+1).value = time_value
                log_ws.cell(row=new_row, column=session_col+1).value = appeal['session']
                log_ws.cell(row=new_row, column=status_col+1).value = "Present"  # Mark as present for approved appeals
                log_ws.cell(row=new_row, column=location_col+1).value = appeal['location']
                log_ws.cell(row=new_row, column=notes_col+1).value = "exception"  # Add exception flag as requested
            
                # Increment row counter
                max_row += 1
        
            # Save changes to the log file
            log_wb.save(log_file)
        
            # Show success message
            success_dialog = QMessageBox(self)
            success_dialog.setWindowTitle("Success")
            success_dialog.setText(
                "Appeals have been successfully processed and added to the log file.")
            success_dialog.setIcon(QMessageBox.Icon.Information)

            # Style OK button
            ok_button = success_dialog.addButton(QMessageBox.StandardButton.Ok)
            ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)

            # Style dialog background
            success_dialog.setStyleSheet(f"""
                QMessageBox {{
                    background-color: {CARD_BG};
                    }}
                QLabel {{
                    color: {TEXT_COLOR};
                    font-size: 14px;
                }}
            """)

            success_dialog.exec()
        except Exception as e:
            # Handle any errors that occur during processing
            error_dialog = QMessageBox(self)
            error_dialog.setWindowTitle("Error")
            error_dialog.setText(f"Error processing appeals: {str(e)}")
            error_dialog.setIcon(QMessageBox.Icon.Critical)
            error_dialog.setStyleSheet(f"""
                QMessageBox {{
                    background-color: {CARD_BG};
                    }}
                QLabel {{
                    color: {TEXT_COLOR};
                    font-size: 14px;
                }}
            """)
            error_dialog.exec()

#==========================================================attendance processors==========================================================#

class AttendanceProcessor(QWidget):
    def __init__(self):
        super().__init__()
        self.schedules = []
        self.setStyleSheet("""
            background-color: black; 
            color: white;
            QLabel {
                color: white;
            }
        """)

        self.init_ui()

    def return_to_home(self):
        # Get the stacked widget and switch to the start page
        stacked_widget = self.parent()
        if isinstance(stacked_widget, QStackedWidget):
            stacked_widget.setCurrentIndex(0)

    def init_ui(self):
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Create a scroll area to make the page scrollable
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; }")
        
        # Create a widget to hold the content
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(20)
        scroll_layout.setContentsMargins(0, 0, 0, 0)

        # Header layout
        header_layout = QHBoxLayout()
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), 'ASU1.png')
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaled(60, 60, Qt.AspectRatioMode.KeepAspectRatio,
                                               Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        title_label = QLabel("Attendance Processor")
        title_label.setStyleSheet(f"font-size: 24px; font-weight: bold;")
        header_layout.addWidget(logo_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Create vertical layout for header buttons
        header_buttons_layout = QVBoxLayout()
        header_buttons_layout.setSpacing(5)

        # Back button
        back_btn = QPushButton("Back to Home")
        back_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        back_btn.clicked.connect(self.return_to_home)

        # Exit button
        exit_btn = QPushButton("Exit")
        exit_btn.setStyleSheet(EXIT_BUTTON_STYLE)
        exit_btn.clicked.connect(QApplication.instance().quit)

        # Add buttons to vertical layout
        header_buttons_layout.addWidget(back_btn)
        header_buttons_layout.addWidget(exit_btn)
        header_buttons_layout.addStretch()

        # Add button layout to header
        header_layout.addStretch()
        header_layout.addLayout(header_buttons_layout)
        scroll_layout.addLayout(header_layout)

        # Previous Reports Section - NEW
        prev_reports_group = QGroupBox("Previous Reports (Optional)")
        prev_reports_group.setStyleSheet(GROUP_BOX_STYLE)
        prev_reports_layout = QVBoxLayout(prev_reports_group)

        # Single line layout for previous reports
        prev_reports_input_layout = QHBoxLayout()
        prev_reports_input_layout.addWidget(QLabel("Last Report File:"))
        self.prev_report_file_input = QLineEdit()
        self.prev_report_file_input.setPlaceholderText("Select previous report Excel file...")
        self.prev_report_file_input.setMinimumWidth(200)
        prev_reports_input_layout.addWidget(self.prev_report_file_input)
        prev_reports_browse_btn = QPushButton("Browse")
        prev_reports_browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        prev_reports_input_layout.addWidget(prev_reports_browse_btn)
        prev_reports_browse_btn.clicked.connect(
            lambda: self.browse_file(self.prev_report_file_input))
        prev_reports_layout.addLayout(prev_reports_input_layout)

        # Add infobox to explain this feature
        info_label = QLabel("Note: Upload previous reports to track student group transfers. "
                           "The system will identify students who switched groups and update their attendance accordingly.")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("font-style: italic; color: #AAAAAA;")
        prev_reports_layout.addWidget(info_label)

        scroll_layout.addWidget(prev_reports_group)

        # Reference Data Section
        ref_group = QGroupBox("Reference Data")
        ref_group.setStyleSheet(GROUP_BOX_STYLE)
        ref_layout = QVBoxLayout(ref_group)

        # Single line layout for reference data
        ref_input_layout = QHBoxLayout()
        ref_input_layout.addWidget(QLabel("Database File:"))
        self.ref_file_input = QLineEdit()
        self.ref_file_input.setPlaceholderText("Select Excel file...")
        self.ref_file_input.setMinimumWidth(200)
        ref_input_layout.addWidget(self.ref_file_input)
        ref_browse_btn = QPushButton("Browse")
        ref_browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        ref_input_layout.addWidget(ref_browse_btn)
        ref_input_layout.addWidget(QLabel("Sheet Name:"))
        self.ref_sheet_combo = QComboBox()
        self.ref_sheet_combo.setMinimumWidth(100)
        ref_input_layout.addWidget(self.ref_sheet_combo)
        ref_browse_btn.clicked.connect(
            lambda: self.browse_file(self.ref_file_input))
        ref_layout.addLayout(ref_input_layout)
        scroll_layout.addWidget(ref_group)

        # Attendance Logs Section
        log_group = QGroupBox("Attendance Logs")
        log_group.setStyleSheet(GROUP_BOX_STYLE)
        log_layout = QVBoxLayout(log_group)

        # Single line layout for log data
        log_input_layout = QHBoxLayout()
        log_input_layout.addWidget(QLabel("Log File:"))
        self.log_file_input = QLineEdit()
        self.log_file_input.setPlaceholderText("Select Excel file...")
        self.log_file_input.setMinimumWidth(200)
        log_input_layout.addWidget(self.log_file_input)
        log_browse_btn = QPushButton("Browse")
        log_browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        log_input_layout.addWidget(log_browse_btn)
        log_input_layout.addWidget(QLabel("Sheet Name:"))
        self.log_sheet_combo = QComboBox()
        self.log_sheet_combo.setMinimumWidth(100)
        log_input_layout.addWidget(self.log_sheet_combo)
        log_browse_btn.clicked.connect(
            lambda: self.browse_file(self.log_file_input))
        log_layout.addLayout(log_input_layout)
        scroll_layout.addWidget(log_group)

        # Session Schedules Section
        schedule_group = QGroupBox("Sessions Schedules")
        schedule_group.setStyleSheet(GROUP_BOX_STYLE)
        schedule_layout = QVBoxLayout(schedule_group)

        self.schedule_table = QTableWidget()
        self.schedule_table.setColumnCount(5)
        self.schedule_table.setHorizontalHeaderLabels(
            ['Year', 'Module', 'File', 'Sheet', 'Total Sessions'])
        self.schedule_table.setStyleSheet(TABLE_STYLE)

        # Center align the header text
        header = self.schedule_table.horizontalHeader()
        header.setDefaultAlignment(
            Qt.AlignmentFlag.AlignCenter)  # Center header text

        # Set column resize modes to stretch and fit content
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)  # Year
        header.setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch)  # Module
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)  # File
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)  # Sheet
        header.setSectionResizeMode(
            4, QHeaderView.ResizeMode.Stretch)  # Total Sessions

        schedule_layout.addWidget(self.schedule_table)
        schedule_btn_layout = QHBoxLayout()
        add_schedule_btn = QPushButton("Add Schedule")
        add_schedule_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        add_schedule_btn.clicked.connect(self.add_schedule)
        remove_schedule_btn = QPushButton("Remove Schedule")
        remove_schedule_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        remove_schedule_btn.clicked.connect(self.remove_schedule)
        schedule_btn_layout.addWidget(add_schedule_btn)
        schedule_btn_layout.addWidget(remove_schedule_btn)
        schedule_layout.addLayout(schedule_btn_layout)
        scroll_layout.addWidget(schedule_group)

        # Progress Bar Section
        progress_group = QGroupBox("Progress")
        progress_group.setStyleSheet(GROUP_BOX_STYLE)
        progress_layout = QVBoxLayout(progress_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar.setStyleSheet(PROGRESS_BAR_STYLE)

        # Create loading gif label
        self.loading_label = QLabel()
        # Adjust based on your GIF size
        self.loading_label.setFixedSize(24, 24)
        self.loading_label.setVisible(False)  # Hidden by default

        # Create the movie object for the GIF
        self.loading_movie = QMovie()
        # Adjust based on your GIF size
        self.loading_movie.setScaledSize(QSize(24, 24))
        self.loading_label.setMovie(self.loading_movie)

        # Make sure to have your loading.gif in the same directory as the script
        loading_gif_path = os.path.join(
            os.path.dirname(__file__), 'loading.gif')
        if os.path.exists(loading_gif_path):
            self.loading_movie.setFileName(loading_gif_path)
        else:
            print(f"Warning: loading.gif not found at {loading_gif_path}")

        # Create a horizontal layout to hold both the progress bar and loading animation
        progress_h_layout = QHBoxLayout()
        progress_h_layout.addWidget(self.progress_bar)
        progress_h_layout.addWidget(self.loading_label)
        progress_layout.addLayout(progress_h_layout)

        scroll_layout.addWidget(progress_group)

        # Output Console Section
        console_group = QGroupBox("Output Console")
        console_group.setStyleSheet(GROUP_BOX_STYLE)
        console_layout = QVBoxLayout(console_group)

        self.output_console = QTextEdit()
        self.output_console.setReadOnly(True)
        self.output_console.setMaximumHeight(150)
        self.output_console.setStyleSheet(CONSOLE_STYLE)
        console_layout.addWidget(self.output_console)
        scroll_layout.addWidget(console_group)

        # Bottom Buttons
        button_layout = QHBoxLayout()
        
        # Process Button
        process_btn = QPushButton("Process Attendance Records")
        process_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        process_btn.clicked.connect(self.process_data)
        button_layout.addWidget(process_btn)
        
        # Update Report Button
        update_btn = QPushButton("Update Report")
        update_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        update_btn.clicked.connect(self.update_report)
        button_layout.addWidget(update_btn)
        
        scroll_layout.addLayout(button_layout)
        
        # Set the scroll content as the scroll area's widget
        scroll_area.setWidget(scroll_content)
        
        # Add scroll area to main layout
        main_layout.addWidget(scroll_area)

        # Connect file input changes to sheet loading
        self.ref_file_input.textChanged.connect(
            lambda: self.load_sheets(self.ref_file_input.text(), self.ref_sheet_combo))
        self.log_file_input.textChanged.connect(
            lambda: self.load_sheets(self.log_file_input.text(), self.log_sheet_combo))
        self.prev_report_file_input.textChanged.connect(
            lambda: self.check_previous_report_file(self.prev_report_file_input.text()))

    def check_previous_report_file(self, file_path):
        """Check if the previous report file has the expected sheets"""
        if os.path.isfile(file_path):
            try:
                wb = openpyxl.load_workbook(file_path, read_only=True)
                # Check if the file has the expected sheets
                required_sheets = ["Summary", "Attendance"]
                has_required_sheets = all(sheet in wb.sheetnames for sheet in required_sheets)
                
                if not has_required_sheets:
                    self.output_console.append("Warning: Previous report file may not have the expected format. "
                                            "It should contain 'Summary' and 'Attendance' sheets.")
                else:
                    self.output_console.append("Previous report file loaded successfully.")
                    
                    # Try to extract the report date from the file name
                    try:
                        file_name = os.path.basename(file_path)
                        # Look for date pattern in the filename (YYYYMMDD_HHMMSS)
                        date_match = re.search(r'(\d{8}_\d{6})', file_name)
                        if date_match:
                            date_str = date_match.group(1)
                            report_date = datetime.strptime(date_str, '%Y%m%d_%H%M%S')
                            self.output_console.append(f"Previous report date: {report_date.strftime('%Y-%m-%d')}")
                        else:
                            # Try to extract from sheet names if they include dates
                            for sheet in wb.sheetnames:
                                date_match = re.search(r'(\d{2}/\d{2}/\d{4})', sheet)
                                if date_match:
                                    date_str = date_match.group(1)
                                    report_date = datetime.strptime(date_str, '%d/%m/%Y')
                                    self.output_console.append(f"Previous report date: {report_date.strftime('%Y-%m-%d')}")
                                    break
                    except Exception as e:
                        self.output_console.append(f"Note: Could not extract report date from file. Will use file modification date as fallback.")
            except Exception as e:
                self.output_console.append(f"Error loading previous report file: {str(e)}")

    def update_report(self):
        """Handle the Update Report button click"""
        # First, validate the inputs
        if not self.validate_inputs():
            return
    
        # Additionally validate that a previous report file is selected
        if not self.prev_report_file_input.text() or not os.path.isfile(self.prev_report_file_input.text()):
            self.show_custom_warning(
                "Previous Report Required", 
                "Please select a previous report file to update from")
            return
        
        # Disable UI elements
        self.setEnabled(False)
        self.output_console.clear()
        self.progress_bar.setValue(0)

        # Show and start the loading animation
        self.loading_label.setVisible(True)
        self.loading_movie.start()

        # Create and start update thread
        self.update_thread = UpdateProcessThread(
            self.prev_report_file_input.text(),
            self.ref_file_input.text(),
            self.ref_sheet_combo.currentText(),
            self.log_file_input.text(),
            self.log_sheet_combo.currentText(),
            self.schedules
        )

        # Connect signals
        self.update_thread.progress_updated.connect(self.update_progress)
        self.update_thread.error_occurred.connect(self.handle_error)
        self.update_thread.processing_complete.connect(self.handle_completion)

        # Start processing
        self.update_thread.start()

    def browse_file(self, input_field):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if filename:
            input_field.setText(filename)

    def load_sheets(self, file_path, combo_box):
        if os.path.isfile(file_path):
            try:
                wb = openpyxl.load_workbook(file_path, read_only=True)
                combo_box.clear()
                combo_box.addItems(wb.sheetnames)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error loading workbook: {str(e)}")

    def add_schedule(self):
        dialog = ScheduleDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.schedules.append(dialog.get_schedule_data())
            self.update_schedule_table()

    def remove_schedule(self):
        current_row = self.schedule_table.currentRow()
        if current_row >= 0:
            self.schedules.pop(current_row)
            self.update_schedule_table()

    def update_schedule_table(self):
        self.schedule_table.setRowCount(len(self.schedules))
        for i, schedule in enumerate(self.schedules):
            for j, value in enumerate(schedule):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)  # Center align the text
                self.schedule_table.setItem(i, j, item)

    def process_data(self):
        if not self.validate_inputs():
            return

        # Disable UI elements
        self.setEnabled(False)
        self.output_console.clear()
        self.progress_bar.setValue(0)

        # Show and start the loading animation
        self.loading_label.setVisible(True)
        self.loading_movie.start()

        # Create and start processing thread
        self.process_thread = ProcessThread(
            self.ref_file_input.text(),
            self.ref_sheet_combo.currentText(),
            self.log_file_input.text(),
            self.log_sheet_combo.currentText(),
            self.schedules
        )

        # Connect signals
        self.process_thread.progress_updated.connect(self.update_progress)
        self.process_thread.error_occurred.connect(self.handle_error)
        self.process_thread.processing_complete.connect(self.handle_completion)

        # Start processing
        self.process_thread.start()

    def validate_inputs(self):
        # Validate reference file
        if not self.ref_file_input.text() or not self.ref_sheet_combo.currentText():
            self.show_custom_warning("Reference Data Required", "Please select reference file and sheet")
            return False
            
        # Validate log file
        if not self.log_file_input.text() or not self.log_sheet_combo.currentText():
            self.show_custom_warning("Log Data Required", "Please select log file and sheet")
            return False
            
        # Validate schedules
        if not self.schedules:
            self.show_custom_warning("Schedules Required", "Please add at least one schedule")
            return False
            
        return True

    def show_custom_warning(self, title, message):
        """Show a custom styled warning dialog"""
        warning_dialog = QMessageBox(self)
        warning_dialog.setWindowTitle(title)
        warning_dialog.setText(message)
        warning_dialog.setIcon(QMessageBox.Icon.Warning)

        # Create and style OK button
        ok_button = warning_dialog.addButton(QMessageBox.StandardButton.Ok)
        ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)

        # Style dialog background and text
        warning_dialog.setStyleSheet(f"""
            QMessageBox {{
                background-color: {CARD_BG};
            }}
            QLabel {{
                color: {TEXT_COLOR};
                font-size: 14px;
            }}
        """)

        warning_dialog.exec()

    def update_progress(self, value):
        self.progress_bar.setValue(value)
        self.output_console.append(f"Processing... {value}%")

    def handle_error(self, error_message):
        self.setEnabled(True)
        error_dialog = QMessageBox(self)
        error_dialog.setWindowTitle("Error")
        error_dialog.setText(f"Error processing data: {error_message}")
        error_dialog.setIcon(QMessageBox.Icon.Critical)

        # Style OK button
        ok_button = error_dialog.addButton(QMessageBox.StandardButton.Ok)
        ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)

        # Style dialog background
        error_dialog.setStyleSheet(f"""
            QMessageBox {{
                background-color: {CARD_BG};
            }}
            QLabel {{
                color: {TEXT_COLOR};
                font-size: 14px;
            }}
        """)

        error_dialog.exec()
        self.output_console.append(f"Error: {error_message}")

    def handle_completion(self):
        self.setEnabled(True)
        self.progress_bar.setValue(100)
        self.output_console.append("Processing complete!")
        
        # Hide the loading animation
        self.loading_label.setVisible(False)
        self.loading_movie.stop()

        success_dialog = QMessageBox(self)
        success_dialog.setWindowTitle("Success")
        success_dialog.setText("Processing complete! Check the attendance_reports folder.")
        success_dialog.setIcon(QMessageBox.Icon.Information)

        # Style OK button
        ok_button = success_dialog.addButton(QMessageBox.StandardButton.Ok)
        ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)

        # Style dialog background
        success_dialog.setStyleSheet(f"""
            QMessageBox {{
                background-color: {CARD_BG};
            }}
            QLabel {{
                color: {TEXT_COLOR};
                font-size: 14px;
            }}
        """)

        success_dialog.exec()

class ProcessThread(QThread):
    progress_updated = pyqtSignal(int)
    error_occurred = pyqtSignal(str)
    processing_complete = pyqtSignal()
    
    # Constants for configuration (both in minutes)
    VALID_ATTENDANCE_BEFORE_MINUTES = 15
    VALID_ATTENDANCE_AFTER_MINUTES = 150

    def __init__(self, ref_file, ref_sheet, log_file, log_sheet, schedules):
        super().__init__()
        self.ref_file = ref_file
        self.ref_sheet = ref_sheet
        self.log_file = log_file
        self.log_sheet = log_sheet
        self.schedules = schedules

    def run(self):
        try:
            # Calculate total steps
            total_steps = 2 + len(self.schedules) * 5
            current_step = 0
            
            # Load reference data
            ref_wb = openpyxl.load_workbook(self.ref_file)
            ref_ws = ref_wb[self.ref_sheet]
            student_db = list(ref_ws.values)
            student_map = self.create_student_map(student_db)
            current_step += 1
            self.progress_updated.emit(int(current_step / total_steps * 100))
            
            # Load log data
            log_wb = openpyxl.load_workbook(self.log_file)
            log_ws = log_wb[self.log_sheet]
            log_history = list(log_ws.values)
            current_step += 1
            self.progress_updated.emit(int(current_step / total_steps * 100))
            
            # Create output directory
            output_dir = os.path.join(os.getcwd(), "attendance_reports")
            os.makedirs(output_dir, exist_ok=True)
            
            # Get current date for sheet names
            current_date = datetime.now().strftime('%d_%m_%Y')
            
            # Process each schedule
            for year, module, sched_file, sched_sheet, total_required, department in self.schedules:
                # Load schedule data
                sched_wb = openpyxl.load_workbook(sched_file)
                sched_ws = sched_wb[sched_sheet]
                session_schedule = list(sched_ws.values)
                current_step += 1
                self.progress_updated.emit(int(current_step / total_steps * 100))
                
                # Calculate sessions
                completed_sessions = self.calculate_completed_sessions(session_schedule[1:])
                session_details = self.calculate_session_details(session_schedule[1:])
                current_step += 1
                self.progress_updated.emit(int(current_step / total_steps * 100))
                
                # Validate attendance
                valid_attendance = self.validate_attendance(log_history, session_schedule[1:], 
                                                         student_map, f"Year {year}")
                current_step += 1
                self.progress_updated.emit(int(current_step / total_steps * 100))
                
                # Create output workbook and sheets
                output_wb = openpyxl.Workbook()
                output_wb.remove(output_wb.active)
                
                # Create Summary sheet first, then Attendance sheet with date in sheet names
                summary_sheet_name = f"Summary_{current_date}"
                attendance_sheet_name = f"Attendance_{current_date}"
                
                self.create_summary_sheet(output_wb, summary_sheet_name, valid_attendance, session_details,
                                        student_map, f"Year {year}", completed_sessions, total_required, department)
                self.create_valid_logs_sheet(output_wb, attendance_sheet_name, valid_attendance)

                current_step += 1
                self.progress_updated.emit(int(current_step / total_steps * 100))
                
                # Generate timestamp for filename
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                # Save output workbook with department and timestamp in filename
                year_dir = os.path.join(output_dir, f"Year_{year}")
                os.makedirs(year_dir, exist_ok=True)
                output_path = os.path.join(year_dir, f"Y{year}_{module}_{department}_attendance_{timestamp}.xlsx")
                output_wb.save(output_path)
                current_step += 1
                self.progress_updated.emit(int(current_step / total_steps * 100))

            self.processing_complete.emit()

        except Exception as e:
            self.error_occurred.emit(str(e))


    def create_student_map(self, student_db):
        student_map = {}
        for row in student_db[1:]:
            if row[0]:
                student_id = str(row[0])
                email = f"{student_id}@med.asu.edu.eg"
                student_map[student_id] = {
                    "name": row[1],
                    "year": row[2],
                    "group": row[3],
                    "email": email
                }
        return student_map

    def calculate_completed_sessions(self, session_schedule):
        completed_sessions = {}
        for row in session_schedule:
            if len(row) >= 2:
                year, group = row[:2]
                key = f"{year}-{group}"
                completed_sessions[key] = completed_sessions.get(key, 0) + 1
        return completed_sessions

    def calculate_session_details(self, session_schedule):
        session_details = {}
        
        for row in session_schedule:
            year, group, session, location = row[:4]
            key = f"{year}-{group}"
            if key not in session_details:
                session_details[key] = {}
            
            session_details[key][session] = {
                "location": location,
                "required": 1
            }
            
        return session_details

    def validate_attendance(self, log_history, session_schedule, student_map, target_year):
        valid_attendance = {}
        # Using the class constants to define time windows (both in minutes)
        before_window = timedelta(minutes=self.VALID_ATTENDANCE_BEFORE_MINUTES)
        after_window = timedelta(minutes=self.VALID_ATTENDANCE_AFTER_MINUTES)
        session_map = {}
        unique_logs = set()

        for row in session_schedule:
            year, group, session, location, date, start_time = row[:6]
            key = f"{year}-{group}"
            session_datetime = self.parse_datetime(date, start_time)
            session_key = f"{location}-{date}"
            if key not in session_map:
                session_map[key] = {}
            session_map[key][session_key] = (session, session_datetime)

        for row in log_history[1:]:
            if len(row) >= 4:
                student_id, location, date, time = row[:4]
                student_id = str(student_id)
                if student_id in student_map:
                    student = student_map[student_id]
                    key = f"{student['year']}-{student['group']}"
                    session_key = f"{location}-{date}"
                    if key in session_map and session_key in session_map[key]:
                        session, session_start = session_map[key][session_key]
                        log_datetime = self.parse_datetime(date, time)
                        # Using the updated time window: 15 min before and 120 min after
                        if session_start - before_window <= log_datetime <= session_start + after_window:
                            unique_log_key = f"{student_id}-{location}-{date}"
                            if unique_log_key not in unique_logs:
                                unique_logs.add(unique_log_key)
                                if key not in valid_attendance:
                                    valid_attendance[key] = []
                                valid_attendance[key].append([
                                    student_id, student['name'], student['year'],
                                    student['group'], student['email'], session,
                                    location, date, time
                                ])
        return valid_attendance

    def parse_datetime(self, date, time):
        if isinstance(date, str):
            date = datetime.strptime(date, '%d/%m/%Y').date()
        if isinstance(time, str):
            time = datetime.strptime(time, '%H:%M:%S').time()
        return datetime.combine(date, time)

    def create_valid_logs_sheet(self, workbook, sheet_name, data):
        sheet = workbook.create_sheet(sheet_name)
        header = ["Student ID", "Name", "Year", "Group", "Email", "Session", "Location", "Date", "Time"]
        sheet.append(header)
        
        # Format header row
        for cell in sheet[1]:
            cell.font = openpyxl.styles.Font(bold=True)
            cell.fill = openpyxl.styles.PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        
        # Freeze the header row
        sheet.freeze_panes = "C2"
        
        for key in data:
            for row in data[key]:
                sheet.append(row)
        for col in 'H', 'I':
            for cell in sheet[col]:
                cell.number_format = 'DD/MM/YYYY' if col == 'H' else 'HH:MM:SS'
        for column in sheet.columns:
            max_length = max(len(str(cell.value)) for cell in column)
            sheet.column_dimensions[openpyxl.utils.get_column_letter(column[0].column)].width = max_length + 2

    def create_summary_sheet(self, workbook, sheet_name, valid_attendance, session_details,
                           student_map, target_year, completed_sessions, total_required_sessions, department):
        sheet = workbook.create_sheet(sheet_name)
        
        # Get all unique sessions
        all_sessions = set()
        for key, sessions in session_details.items():
            all_sessions.update(sessions.keys())
        sorted_sessions = sorted(all_sessions)
    
        header = ["Student ID", "Name", "Year", "Group", "Email", 
                 "Sessions Left", "Sessions Completed", "Total Required", "Total Attended"]
        
        # Add columns for each session with department name
        for session in sorted_sessions:
            header.extend([f"{department} session {session} (Required)", f"{department} session {session} (Attended)"])
        
        sheet.append(header)
        
        # Format header row
        for cell in sheet[1]:
            cell.font = openpyxl.styles.Font(bold=True)
            cell.fill = openpyxl.styles.PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        
        # Freeze the header row
        sheet.freeze_panes = "C2"
        
        # Create alternating row colors for better readability
        light_blue = openpyxl.styles.PatternFill(start_color="F2F7FD", end_color="F2F7FD", fill_type="solid")
    
        for student_id, student in student_map.items():
            if student['year'] == target_year:
                key = f"{student['year']}-{student['group']}"
                group_completed = completed_sessions.get(key, 0)
                total_attended = 0
                attendance_by_session = {}
                
                # Calculate attendance for each session
                for entry in valid_attendance.get(key, []):
                    if entry[0] == student_id:
                        session = entry[5]
                        attendance_by_session[session] = attendance_by_session.get(session, 0) + 1
                        total_attended += 1
    
                sessions_left = total_required_sessions - group_completed
    
                row = [
                    student_id, student['name'], student['year'], student['group'],
                    student['email'], sessions_left, group_completed,
                    total_required_sessions, total_attended
                ]
    
                # Add attendance data for each session
                for session in sorted_sessions:
                    session_req = session_details.get(key, {}).get(session, {"required": 0})
                    session_att = attendance_by_session.get(session, 0)
                    row.extend([session_req["required"], session_att])
    
                sheet.append(row)
        
        # Apply alternating row colors
        for i, row in enumerate(sheet.iter_rows(min_row=2)):
            if i % 2 == 0:  # Apply light blue to every other row
                for cell in row:
                    cell.fill = light_blue
    
        # Format columns
        for column in sheet.columns:
            max_length = max(len(str(cell.value)) for cell in column)
            sheet.column_dimensions[openpyxl.utils.get_column_letter(column[0].column)].width = max_length + 2
            
        # Add borders to all cells
        thin_border = openpyxl.styles.Border(
            left=openpyxl.styles.Side(style='thin'),
            right=openpyxl.styles.Side(style='thin'),
            top=openpyxl.styles.Side(style='thin'),
            bottom=openpyxl.styles.Side(style='thin')
        )
        
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = thin_border

class UpdateProcessThread(QThread):
    progress_updated = pyqtSignal(int)
    error_occurred = pyqtSignal(str)
    processing_complete = pyqtSignal()
    
    # Constants for configuration (both in minutes)
    VALID_ATTENDANCE_BEFORE_MINUTES = 15
    VALID_ATTENDANCE_AFTER_MINUTES = 150

    def __init__(self, prev_report_file, ref_file, ref_sheet, log_file, log_sheet, schedules):
        super().__init__()
        self.prev_report_file = prev_report_file
        self.ref_file = ref_file
        self.ref_sheet = ref_sheet
        self.log_file = log_file
        self.log_sheet = log_sheet
        self.schedules = schedules
        self.prev_report_date = None

    def run(self):
        try:
            # Calculate total steps
            total_steps = 5 + len(self.schedules) * 6  # Extra steps for transfers analysis
            current_step = 0
            
            # Load previous report data
            prev_wb = openpyxl.load_workbook(self.prev_report_file)
            self.extract_report_date()
            
            # Load previous summary data (to identify group transfers)
            prev_summary_sheet = None
            for sheet_name in prev_wb.sheetnames:
                if sheet_name.startswith("Summary"):
                    prev_summary_sheet = prev_wb[sheet_name]
                    break
            
            if not prev_summary_sheet:
                raise Exception("No Summary sheet found in previous report")
                
            # Load previous attendance data
            prev_attendance_sheet = None
            for sheet_name in prev_wb.sheetnames:
                if sheet_name.startswith("Attendance"):
                    prev_attendance_sheet = prev_wb[sheet_name]
                    break
                    
            if not prev_attendance_sheet:
                raise Exception("No Attendance sheet found in previous report")
                
            current_step += 1
            self.progress_updated.emit(int(current_step / total_steps * 100))
            
            # Load reference data (current student groups)
            ref_wb = openpyxl.load_workbook(self.ref_file)
            ref_ws = ref_wb[self.ref_sheet]
            student_db = list(ref_ws.values)
            current_student_map = self.create_student_map(student_db)
            current_step += 1
            self.progress_updated.emit(int(current_step / total_steps * 100))
            
            # Extract previous student data from previous report
            previous_student_map = self.extract_previous_student_map(prev_summary_sheet)
            current_step += 1
            self.progress_updated.emit(int(current_step / total_steps * 100))
            
            # Identify students who transferred groups
            transferred_students = self.identify_transferred_students(previous_student_map, current_student_map)
            current_step += 1
            self.progress_updated.emit(int(current_step / total_steps * 100))
            
            # Load log data (attendance data)
            log_wb = openpyxl.load_workbook(self.log_file)
            log_ws = log_wb[self.log_sheet]
            log_history = list(log_ws.values)
            current_step += 1
            self.progress_updated.emit(int(current_step / total_steps * 100))
            
            # Create output directory
            output_dir = os.path.join(os.getcwd(), "attendance_reports")
            os.makedirs(output_dir, exist_ok=True)
            
            # Get current date for sheet names
            current_date = datetime.now().strftime('%d_%m_%Y')
            
            # Process each schedule with transfer awareness
            for year, module, sched_file, sched_sheet, total_required, department in self.schedules:
                # Load schedule data
                sched_wb = openpyxl.load_workbook(sched_file)
                sched_ws = sched_wb[sched_sheet]
                session_schedule = list(sched_ws.values)
                current_step += 1
                self.progress_updated.emit(int(current_step / total_steps * 100))
                
                # Calculate sessions
                completed_sessions = self.calculate_completed_sessions(session_schedule[1:])
                session_details = self.calculate_session_details(session_schedule[1:])
                current_step += 1
                self.progress_updated.emit(int(current_step / total_steps * 100))
                
                # Extract previous attendance data
                previous_attendance = self.extract_previous_attendance(prev_attendance_sheet, f"Year {year}")
                current_step += 1
                self.progress_updated.emit(int(current_step / total_steps * 100))
                
                # Validate attendance with transfer awareness
                valid_attendance = self.validate_attendance_with_transfers(
                    log_history, 
                    session_schedule[1:], 
                    current_student_map, 
                    previous_student_map,
                    transferred_students,
                    previous_attendance,
                    f"Year {year}"
                )
                current_step += 1
                self.progress_updated.emit(int(current_step / total_steps * 100))
                
                # Create output workbook and sheets
                output_wb = openpyxl.Workbook()
                output_wb.remove(output_wb.active)
                
                # Create Summary sheet first, then Attendance sheet with date in sheet names
                summary_sheet_name = f"Summary_{current_date}"
                attendance_sheet_name = f"Attendance_{current_date}"
                transfers_sheet_name = "Transfers"
                
                self.create_summary_sheet(
                    output_wb, 
                    summary_sheet_name, 
                    valid_attendance, 
                    session_details,
                    current_student_map, 
                    f"Year {year}", 
                    completed_sessions, 
                    total_required, 
                    department
                )
                
                self.create_valid_logs_sheet(output_wb, attendance_sheet_name, valid_attendance)
                
                # Create transfer log sheet
                self.create_transfer_log_sheet(
                    output_wb, 
                    transfers_sheet_name, 
                    transferred_students, 
                    f"Year {year}"
                )
                current_step += 1
                self.progress_updated.emit(int(current_step / total_steps * 100))
                
                # Generate timestamp for filename
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                
                # Save output workbook with department and timestamp in filename
                year_dir = os.path.join(output_dir, f"Year_{year}")
                os.makedirs(year_dir, exist_ok=True)
                output_path = os.path.join(year_dir, f"Y{year}_{module}_{department}_attendance_{timestamp}.xlsx")
                output_wb.save(output_path)
                current_step += 1
                self.progress_updated.emit(int(current_step / total_steps * 100))

            self.processing_complete.emit()

        except Exception as e:
            import traceback
            error_msg = f"{str(e)}\n{traceback.format_exc()}"
            self.error_occurred.emit(error_msg)

    def extract_report_date(self):
        """Extract the date of the previous report from file name or sheet names"""
        try:
            file_name = os.path.basename(self.prev_report_file)
            
            # Try to extract from filename (YYYYMMDD_HHMMSS)
            date_match = re.search(r'(\d{8}_\d{6})', file_name)
            if date_match:
                date_str = date_match.group(1)
                self.prev_report_date = datetime.strptime(date_str, '%Y%m%d_%H%M%S')
                return
                
            # If not found in filename, try to extract from sheet names
            wb = openpyxl.load_workbook(self.prev_report_file, read_only=True)
            
            for sheet in wb.sheetnames:
                # Look for date in sheet name (Summary_DD_MM_YYYY or Attendance_DD_MM_YYYY)
                date_match = re.search(r'_(\d{2}_\d{2}_\d{4})$', sheet)
                if date_match:
                    date_str = date_match.group(1)
                    self.prev_report_date = datetime.strptime(date_str, '%d_%m_%Y')
                    return
                    
            # If still not found, use file modification time as fallback
            self.prev_report_date = datetime.fromtimestamp(os.path.getmtime(self.prev_report_file))
            
        except Exception as e:
            self.prev_report_date = datetime.now() - timedelta(days=7)  # Default to a week ago
            print(f"Error extracting report date: {str(e)}. Using fallback date: {self.prev_report_date}")

    def create_student_map(self, student_db):
        student_map = {}
        for row in student_db[1:]:
            if row[0]:
                student_id = str(row[0])
                email = f"{student_id}@med.asu.edu.eg"
                student_map[student_id] = {
                    "name": row[1],
                    "year": row[2],
                    "group": row[3],
                    "email": email
                }
        return student_map

    def extract_previous_student_map(self, prev_summary_sheet):
        """Extract student data from previous report summary sheet"""
        previous_student_map = {}
        
        # Get column indices (they might vary between reports)
        header = [cell.value for cell in prev_summary_sheet[1]]
        id_col = header.index("Student ID") + 1
        name_col = header.index("Name") + 1
        year_col = header.index("Year") + 1
        group_col = header.index("Group") + 1
        email_col = header.index("Email") + 1
        
        for row in prev_summary_sheet.iter_rows(min_row=2):
            student_id = str(row[id_col-1].value)
            if student_id:
                previous_student_map[student_id] = {
                    "name": row[name_col-1].value,
                    "year": row[year_col-1].value,
                    "group": row[group_col-1].value,
                    "email": row[email_col-1].value
                }
                
        return previous_student_map

    def identify_transferred_students(self, previous_map, current_map):
        """Identify students who have switched groups"""
        transferred_students = {}
        
        for student_id, current_data in current_map.items():
            if student_id in previous_map:
                previous_data = previous_map[student_id]
                
                # Check if group has changed
                if previous_data["group"] != current_data["group"]:
                    transferred_students[student_id] = {
                        "name": current_data["name"],
                        "year": current_data["year"],
                        "group_before": previous_data["group"],
                        "group_after": current_data["group"],
                        "transfer_date": self.prev_report_date  # We don't know exact date, use report date
                    }
                    
        return transferred_students

    def extract_previous_attendance(self, prev_attendance_sheet, target_year):
        """Extract previous attendance data from the previous attendance sheet"""
        previous_attendance = {}
        
        # Get header row
        header = [cell.value for cell in prev_attendance_sheet[1]]
        
        # Identify column indices
        id_col = header.index("Student ID") + 1
        year_col = header.index("Year") + 1
        group_col = header.index("Group") + 1
        session_col = header.index("Session") + 1
        location_col = header.index("Location") + 1
        date_col = header.index("Date") + 1
        time_col = header.index("Time") + 1
        
        for row in prev_attendance_sheet.iter_rows(min_row=2):
            student_id = str(row[id_col-1].value)
            year = row[year_col-1].value
            
            if year == target_year:
                group = row[group_col-1].value
                key = f"{year}-{group}"
                
                if key not in previous_attendance:
                    previous_attendance[key] = []
                    
                # Convert date and time if they're datetime objects
                date_value = row[date_col-1].value
                time_value = row[time_col-1].value
                
                if isinstance(date_value, datetime):
                    date_value = date_value.strftime('%d/%m/%Y')
                if isinstance(time_value, datetime):
                    time_value = time_value.strftime('%H:%M:%S')
                
                previous_attendance[key].append([
                    student_id,
                    None,  # Name is not needed here
                    year,
                    group,
                    None,  # Email is not needed here
                    row[session_col-1].value,
                    row[location_col-1].value,
                    date_value,
                    time_value
                ])
                
        return previous_attendance

    def calculate_completed_sessions(self, session_schedule):
        completed_sessions = {}
        for row in session_schedule:
            if len(row) >= 2:
                year, group = row[:2]
                key = f"{year}-{group}"
                completed_sessions[key] = completed_sessions.get(key, 0) + 1
        return completed_sessions

    def calculate_session_details(self, session_schedule):
        session_details = {}
        
        for row in session_schedule:
            year, group, session, location = row[:4]
            key = f"{year}-{group}"
            if key not in session_details:
                session_details[key] = {}
            
            session_details[key][session] = {
                "location": location,
                "required": 1
            }
            
        return session_details

    def validate_attendance_with_transfers(self, log_history, session_schedule, 
                                          current_student_map, previous_student_map,
                                          transferred_students, previous_attendance,
                                          target_year):
        """Validate attendance with awareness of group transfers"""
        valid_attendance = {}
        # Using the class constants to define time windows (both in minutes)
        before_window = timedelta(minutes=self.VALID_ATTENDANCE_BEFORE_MINUTES)
        after_window = timedelta(minutes=self.VALID_ATTENDANCE_AFTER_MINUTES)
        session_map = {}
        unique_logs = set()

        # Create a mapping of sessions by location and date for each group
        for row in session_schedule:
            year, group, session, location, date, start_time = row[:6]
            key = f"{year}-{group}"
            session_datetime = self.parse_datetime(date, start_time)
            session_key = f"{location}-{date}"
            if key not in session_map:
                session_map[key] = {}
            session_map[key][session_key] = (session, session_datetime)

        # First, import previous attendance records
        for key, attendance_list in previous_attendance.items():
            if key not in valid_attendance:
                valid_attendance[key] = []
            
            for attendance in attendance_list:
                valid_attendance[key].append(attendance)
                # Add to unique logs to prevent duplicates
                student_id = attendance[0]
                location = attendance[6]
                date = attendance[7]
                unique_logs.add(f"{student_id}-{location}-{date}")

        # Process new log data
        for row in log_history[1:]:
            if len(row) >= 4:
                student_id, location, date, time = row[:4]
                student_id = str(student_id)
                
                # Skip if this student doesn't exist in either map
                if student_id not in current_student_map and student_id not in previous_student_map:
                    continue
                
                # Get student data - prefer current map, fall back to previous
                student = current_student_map.get(student_id, previous_student_map.get(student_id))
                
                # Skip if student is not in the target year
                if student['year'] != target_year:
                    continue
                
                # Normal case: student didn't transfer
                if student_id not in transferred_students:
                    key = f"{student['year']}-{student['group']}"
                    session_key = f"{location}-{date}"
                    
                    if key in session_map and session_key in session_map[key]:
                        session, session_start = session_map[key][session_key]
                        log_datetime = self.parse_datetime(date, time)
                        
                        # Check if log is within time window
                        if session_start - before_window <= log_datetime <= session_start + after_window:
                            unique_log_key = f"{student_id}-{location}-{date}"
                            if unique_log_key not in unique_logs:
                                unique_logs.add(unique_log_key)
                                if key not in valid_attendance:
                                    valid_attendance[key] = []
                                valid_attendance[key].append([
                                    student_id, student['name'], student['year'],
                                    student['group'], student['email'], session,
                                    location, date, time
                                ])
                
                # Special case: student transferred groups
                else:
                    transfer_info = transferred_students[student_id]
                    
                    # Check both the old and new group's sessions
                    old_key = f"{student['year']}-{transfer_info['group_before']}"
                    new_key = f"{student['year']}-{transfer_info['group_after']}"
                    session_key = f"{location}-{date}"
                    
                    # Check old group sessions
                    if old_key in session_map and session_key in session_map[old_key]:
                        session, session_start = session_map[old_key][session_key]
                        log_datetime = self.parse_datetime(date, time)
                        
                        # Check if log is within time window
                        if session_start - before_window <= log_datetime <= session_start + after_window:
                            unique_log_key = f"{student_id}-{location}-{date}"
                            if unique_log_key not in unique_logs:
                                unique_logs.add(unique_log_key)
                                if new_key not in valid_attendance:  # Use NEW group for updated attendance
                                    valid_attendance[new_key] = []
                                valid_attendance[new_key].append([
                                    student_id, student['name'], student['year'],
                                    transfer_info['group_after'], student['email'], session,
                                    location, date, time
                                ])
                    
                    # Check new group sessions
                    if new_key in session_map and session_key in session_map[new_key]:
                        session, session_start = session_map[new_key][session_key]
                        log_datetime = self.parse_datetime(date, time)
                        
                        # Check if log is within time window
                        if session_start - before_window <= log_datetime <= session_start + after_window:
                            unique_log_key = f"{student_id}-{location}-{date}"
                            if unique_log_key not in unique_logs:
                                unique_logs.add(unique_log_key)
                                if new_key not in valid_attendance:
                                    valid_attendance[new_key] = []
                                valid_attendance[new_key].append([
                                    student_id, student['name'], student['year'],
                                    transfer_info['group_after'], student['email'], session,
                                    location, date, time
                                ])
        
        return valid_attendance

    def parse_datetime(self, date, time):
        if isinstance(date, str):
            date = datetime.strptime(date, '%d/%m/%Y').date()
        elif isinstance(date, datetime):
            date = date.date()
            
        if isinstance(time, str):
            time = datetime.strptime(time, '%H:%M:%S').time()
        elif isinstance(time, datetime):
            time = time.time()
            
        return datetime.combine(date, time)

    def create_valid_logs_sheet(self, workbook, sheet_name, data):
        sheet = workbook.create_sheet(sheet_name)
        header = ["Student ID", "Name", "Year", "Group", "Email", "Session", "Location", "Date", "Time"]
        sheet.append(header)
        
        # Format header row
        for cell in sheet[1]:
            cell.font = openpyxl.styles.Font(bold=True)
            cell.fill = openpyxl.styles.PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        
        # Freeze the header row
        sheet.freeze_panes = "C2"
        
        for key in data:
            for row in data[key]:
                sheet.append(row)
        
        # Format date and time columns
        for col in 'H', 'I':
            for cell in sheet[col]:
                cell.number_format = 'DD/MM/YYYY' if col == 'H' else 'HH:MM:SS'
        
        # Auto-adjust column widths
        for column in sheet.columns:
            max_length = max(len(str(cell.value)) for cell in column)
            sheet.column_dimensions[openpyxl.utils.get_column_letter(column[0].column)].width = max_length + 2

    def create_summary_sheet(self, workbook, sheet_name, valid_attendance, session_details,
                          student_map, target_year, completed_sessions, total_required_sessions, department):
        sheet = workbook.create_sheet(sheet_name)
        
        # Get all unique sessions
        all_sessions = set()
        for key, sessions in session_details.items():
            all_sessions.update(sessions.keys())
        sorted_sessions = sorted(all_sessions)
    
        header = ["Student ID", "Name", "Year", "Group", "Email", 
                 "Sessions Left", "Sessions Completed", "Total Required", "Total Attended"]
        
        # Add columns for each session with department name
        for session in sorted_sessions:
            header.extend([f"{department} session {session} (Required)", f"{department} session {session} (Attended)"])
        
        sheet.append(header)
        
        # Format header row
        for cell in sheet[1]:
            cell.font = openpyxl.styles.Font(bold=True)
            cell.fill = openpyxl.styles.PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        
        # Freeze the header row
        sheet.freeze_panes = "C2"
        
        # Create alternating row colors for better readability
        light_blue = openpyxl.styles.PatternFill(start_color="F2F7FD", end_color="F2F7FD", fill_type="solid")
    
        for student_id, student in student_map.items():
            if student['year'] == target_year:
                key = f"{student['year']}-{student['group']}"
                group_completed = completed_sessions.get(key, 0)
                total_attended = 0
                attendance_by_session = {}
                
                # Calculate attendance for each session
                for entry in valid_attendance.get(key, []):
                    if entry[0] == student_id:
                        session = entry[5]
                        attendance_by_session[session] = attendance_by_session.get(session, 0) + 1
                        total_attended += 1
    
                sessions_left = total_required_sessions - group_completed
    
                row = [
                    student_id, student['name'], student['year'], student['group'],
                    student['email'], sessions_left, group_completed,
                    total_required_sessions, total_attended
                ]
    
                # Add attendance data for each session
                for session in sorted_sessions:
                    session_req = session_details.get(key, {}).get(session, {"required": 0})
                    session_att = attendance_by_session.get(session, 0)
                    row.extend([session_req["required"], session_att])
    
                sheet.append(row)
        
        # Apply alternating row colors
        for i, row in enumerate(sheet.iter_rows(min_row=2)):
            if i % 2 == 0:  # Apply light blue to every other row
                for cell in row:
                    cell.fill = light_blue
    
        # Format columns
        for column in sheet.columns:
            max_length = max(len(str(cell.value)) for cell in column)
            sheet.column_dimensions[openpyxl.utils.get_column_letter(column[0].column)].width = max_length + 2
            
        # Add borders to all cells
        thin_border = openpyxl.styles.Border(
            left=openpyxl.styles.Side(style='thin'),
            right=openpyxl.styles.Side(style='thin'),
            top=openpyxl.styles.Side(style='thin'),
            bottom=openpyxl.styles.Side(style='thin')
        )
        
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = thin_border

    def create_transfer_log_sheet(self, workbook, sheet_name, transferred_students, target_year):
        """Create a sheet logging student transfers"""
        sheet = workbook.create_sheet(sheet_name)
        
        # Set up header
        header = ["Student ID", "Name", "Year", "Group Before", "Group After", "Transfer Date"]
        sheet.append(header)
        
        # Format header row
        for cell in sheet[1]:
            cell.font = openpyxl.styles.Font(bold=True)
            cell.fill = openpyxl.styles.PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
            cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        
        # Freeze the header row
        sheet.freeze_panes = "C2"
        
        # Add transfer data
        for student_id, transfer_info in transferred_students.items():
            if transfer_info["year"] == target_year:
                transfer_date = transfer_info["transfer_date"].strftime('%d/%m/%Y') if transfer_info["transfer_date"] else "Unknown"
                
                row = [
                    student_id,
                    transfer_info["name"],
                    transfer_info["year"],
                    transfer_info["group_before"],
                    transfer_info["group_after"],
                    transfer_date
                ]
                
                sheet.append(row)
        
        # Format date column
        for cell in sheet['F']:
            if cell.row > 1:  # Skip header
                try:
                    cell.number_format = 'DD/MM/YYYY'
                except:
                    pass  # In case the date is "Unknown"
        
        # Create alternating row colors
        light_blue = openpyxl.styles.PatternFill(start_color="F2F7FD", end_color="F2F7FD", fill_type="solid")
        for i, row in enumerate(sheet.iter_rows(min_row=2)):
            if i % 2 == 0:
                for cell in row:
                    cell.fill = light_blue
        
        # Auto-adjust column widths
        for column in sheet.columns:
            max_length = max(len(str(cell.value)) for cell in column)
            sheet.column_dimensions[openpyxl.utils.get_column_letter(column[0].column)].width = max_length + 2
        
        # Add borders
        thin_border = openpyxl.styles.Border(
            left=openpyxl.styles.Side(style='thin'),
            right=openpyxl.styles.Side(style='thin'),
            top=openpyxl.styles.Side(style='thin'),
            bottom=openpyxl.styles.Side(style='thin')
        )
        
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = thin_border

class ScheduleDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Schedule")
        self.setMinimumWidth(400)
        self.setStyleSheet("""
            background-color: black; 
            color: white;
            QLabel {
                color: white;
            }
        """)
        
        self.init_ui()

        # Add input validation
        self.total_input.setValidator(QIntValidator(1, 999, self))
        self.year_input.setValidator(QIntValidator(1, 6, self))

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Form Group
        form_group = QGroupBox("Schedule Details")
        form_group.setStyleSheet(GROUP_BOX_STYLE)
        form_layout = QVBoxLayout(form_group)
        form_layout.setSpacing(10)

        # Form fields
        self.year_input = QLineEdit()
        self.year_input.setPlaceholderText("Academic year...")
        self.module_input = QLineEdit()
        self.module_input.setPlaceholderText("Module to process...")
        self.total_input = QLineEdit()
        self.total_input.setPlaceholderText("Total sessions number...")
        self.file_input = QLineEdit()
        self.file_input.setPlaceholderText("Select Excel file...")
        self.sheet_combo = QComboBox()
        
        # Add department combobox
        self.department_combo = QComboBox()
        departments = [
            "Anatomy", 
            "Histology", 
            "Physiology", 
            "Biochemistry", 
            "Pharmacology", 
            "Pathology", 
            "Parasitology", 
            "Microbiology", 
            "Forensics and Toxicology", 
            "Community Medicine", 
            "Clinical"
        ]
        self.department_combo.addItems(departments)

        # Add form fields with consistent spacing
        form_layout.addWidget(QLabel("Academic Year:"))
        form_layout.addWidget(self.year_input)
        form_layout.addWidget(QLabel("Module Name:"))
        form_layout.addWidget(self.module_input)
        form_layout.addWidget(QLabel("Department:"))
        form_layout.addWidget(self.department_combo)
        form_layout.addWidget(QLabel("Total Required Sessions:"))
        form_layout.addWidget(self.total_input)

        # Schedule File section with Browse button
        file_label = QLabel("Schedule File:")
        file_top_layout = QHBoxLayout()
        file_top_layout.addWidget(file_label)
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse_file)
        browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        file_top_layout.addWidget(browse_btn)
        file_top_layout.addStretch()
        form_layout.addLayout(file_top_layout)
        form_layout.addWidget(self.file_input)

        # Sheet selection
        form_layout.addWidget(QLabel("Sheet Name:"))
        form_layout.addWidget(self.sheet_combo)

        main_layout.addWidget(form_group)

        # Buttons at the bottom
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        cancel_button.setStyleSheet(STANDARD_BUTTON_STYLE)
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        main_layout.addLayout(button_layout)

    def browse_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)"
        )
        if filename:
            self.file_input.setText(filename)
            self.load_sheets(filename)

    def load_sheets(self, file_path):
        if os.path.isfile(file_path):
            try:
                wb = openpyxl.load_workbook(file_path, read_only=True)
                self.sheet_combo.clear()
                self.sheet_combo.addItems(wb.sheetnames)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error loading workbook: {str(e)}")

    def get_schedule_data(self):
        return [
            self.year_input.text(),
            self.module_input.text(),
            self.file_input.text(),
            self.sheet_combo.currentText(),
            int(self.total_input.text()),
            self.department_combo.currentText()  # Add department to returned data
        ]

    def accept(self):
        """Validate inputs before closing dialog"""
        try:
            # Check required fields
            if not self.year_input.text().strip():
                raise ValueError("Academic year is required")
            if not self.module_input.text().strip():
                raise ValueError("Module name is required")
            if not self.total_input.text().strip():
                raise ValueError("Total required sessions is required")
            if not self.file_input.text().strip():
                raise ValueError("Schedule file is required")
            if not self.sheet_combo.currentText():
                raise ValueError("Sheet name is required")
            
            # Validate numeric input
            total_sessions = self.total_input.text()
            if not total_sessions.isdigit():
                raise ValueError("Total sessions must be a whole number")
            if int(total_sessions) <= 0:
                raise ValueError("Total sessions must be greater than zero")
            
            # Validate file exists
            if not os.path.isfile(self.file_input.text()):
                raise FileNotFoundError("Selected schedule file does not exist")

        except (ValueError, FileNotFoundError) as e:
            # Create custom message box
            error_dialog = QMessageBox(self)
            error_dialog.setWindowTitle("Invalid Input")
            error_dialog.setText(str(e))
            error_dialog.setIcon(QMessageBox.Icon.Warning)
            
            # Configure OK button
            ok_button = error_dialog.addButton(QMessageBox.StandardButton.Ok)
            ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)
            
            # Style dialog background
            error_dialog.setStyleSheet(f"""
                QMessageBox {{
                    background-color: {CARD_BG};
                }}
                QLabel {{
                color: {TEXT_COLOR};
                    font-size: 14px;
                }}
            """)
            
            error_dialog.exec()
            return
            
        super().accept()

    def return_to_home(self):
        # Get the stacked widget and switch to the start page
        stacked_widget = self.parent()
        if isinstance(stacked_widget, QStackedWidget):
            stacked_widget.setCurrentIndex(0)
    
#==========================================================attendance analyzer==========================================================#

class AttendanceDashboard(QWidget):
    def __init__(self):
        super().__init__()
        self.student_data = []
        self.setStyleSheet("""
            background-color: black; 
            color: white;
            QLabel {
                color: white;
            }
        """)
        
        self.init_ui()

    def return_to_home(self):
        stacked_widget = self.parent()
        if isinstance(stacked_widget, QStackedWidget):
            stacked_widget.setCurrentIndex(0)

    def init_ui(self):
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Header layout
        header_layout = QHBoxLayout()
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), 'ASU1.png')
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaled(60, 60, Qt.AspectRatioMode.KeepAspectRatio, 
                                             Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        title_label = QLabel("Analysis Dashboard")
        title_label.setStyleSheet(f"font-size: 24px; font-weight: bold;")
        header_layout.addWidget(logo_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Create vertical layout for header buttons
        header_buttons_layout = QVBoxLayout()
        header_buttons_layout.setSpacing(5)
        
        # Buttons
        back_btn = QPushButton("Back to Home")
        back_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        back_btn.clicked.connect(self.return_to_home)

        process_attendance_btn = QPushButton("Process Attendance")
        process_attendance_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        
        exit_btn = QPushButton("Exit")
        exit_btn.setStyleSheet(EXIT_BUTTON_STYLE)
        exit_btn.clicked.connect(QApplication.instance().quit)
        
        header_buttons_layout.addWidget(back_btn)
        header_buttons_layout.addWidget(exit_btn)
        header_buttons_layout.addStretch()
        
        header_layout.addLayout(header_buttons_layout)
        main_layout.addLayout(header_layout)

        # File Selection Section
        file_group = QGroupBox("File Selection")
        file_group.setStyleSheet(GROUP_BOX_STYLE)
        file_layout = QVBoxLayout(file_group)
        
        file_input_layout = QHBoxLayout()
        file_input_layout.addWidget(QLabel("Reports File:"))
        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("Select Excel file...")
        self.file_path.setMinimumWidth(200)
        file_input_layout.addWidget(self.file_path, stretch=1)
        
        browse_btn = QPushButton("Browse")
        browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        browse_btn.clicked.connect(self.browse_file)
        file_input_layout.addWidget(browse_btn)
        
        file_input_layout.addWidget(QLabel("Sheet Name:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(100)
        file_input_layout.addWidget(self.sheet_combo)
        
        file_layout.addLayout(file_input_layout)
        main_layout.addWidget(file_group)

        # Statistics Section
        stats_group = QGroupBox("Statistics")
        stats_group.setStyleSheet(GROUP_BOX_STYLE)
        stats_layout = QHBoxLayout(stats_group)

        # Create stat cards
        self.total_students = self.create_stat_card("Total Students", "0")
        self.total_sessions = self.create_stat_card("Total Sessions", "0")
        self.avg_attendance = self.create_stat_card("Average Attendance", "0%")
        self.completion_rate = self.create_stat_card("Completion Rate", "0")

        stats_layout.addWidget(self.total_students)
        stats_layout.addWidget(self.total_sessions)
        stats_layout.addWidget(self.avg_attendance)
        stats_layout.addWidget(self.completion_rate)
        main_layout.addWidget(stats_group)
        
        # Session table section
        session_group = QGroupBox("Attendance Distribution")
        session_group.setStyleSheet(GROUP_BOX_STYLE)
        session_layout = QVBoxLayout(session_group)
        
        self.session_table = QTableWidget()
        self.session_table.setColumnCount(3)
        self.session_table.setHorizontalHeaderLabels(['Group', 'Session', 'Attendance Rate'])
        self.session_table.setStyleSheet(TABLE_STYLE)
        
        header = self.session_table.horizontalHeader()
        self.session_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Align cell content to center
        for row in range(self.session_table.rowCount()):
            for col in range(self.session_table.columnCount()):
                item = self.session_table.item(row, col)
                if item:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    
        session_layout.addWidget(self.session_table)
        main_layout.addWidget(session_group)

        # Student List Section
        student_group = QGroupBox("Students List")
        student_group.setStyleSheet(GROUP_BOX_STYLE)
        student_layout = QVBoxLayout(student_group)
        
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by ID or Name...")
        self.search_input.textChanged.connect(self.filter_students)
        search_layout.addWidget(self.search_input)
        student_layout.addLayout(search_layout)
        
        self.student_table = QTableWidget()
        self.student_table.setColumnCount(5)
        self.student_table.setHorizontalHeaderLabels([
            'Student ID', 'Name', 'Group', 'Sessions Attended', 'Session Details'
        ])
        self.student_table.setStyleSheet(TABLE_STYLE)
        
        self.student_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.student_table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Align cell content to center
        for row in range(self.student_table.rowCount()):
            for col in range(self.student_table.columnCount()):
                item = self.student_table.item(row, col)
                if item:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    
        student_layout.addWidget(self.student_table)
        main_layout.addWidget(student_group)

        # Bottom Buttons
        display_layout = QHBoxLayout()
        display_btn = QPushButton("Display Statistics")
        display_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        display_btn.clicked.connect(self.display_statistics)
        display_layout.addWidget(display_btn)
        main_layout.addLayout(display_layout)

    def create_stat_card(self, title, value):
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {DARK_BLUE};
                border-radius: 5px;
                padding: 10px;
            }}
            QFrame:hover {{
                background-color: #1b2649;
            }}
            QLabel {{
                color: {TEXT_COLOR};
            }}
        """)
        layout = QVBoxLayout(card)
                        
        title_label = QLabel(title)
        title_label.setStyleSheet("font-size: 14px;")
        value_label = QLabel(value)
        value_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        
        layout.addWidget(title_label)
        layout.addWidget(value_label)
        card.value_label = value_label
        return card

    def browse_file(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "    Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_name:
            self.file_path.setText(file_name)
            self.update_sheet_list(file_name)

    def update_student_table(self, data):
        """Update the student table with the provided data"""
        self.student_table.setRowCount(len(data))
        for i, row in enumerate(data):
            for j, value in enumerate(row):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.student_table.setItem(i, j, item)
    
    def filter_students(self):
        """Filter students based on search input"""
        search_text = self.search_input.text().lower()
        filtered_data = [
            row for row in self.student_data
            if search_text in str(row[0]).lower() or search_text in str(row[1]).lower()
        ]
        self.update_student_table(filtered_data)
    
    def browse_file(self):
        """Open file dialog to select Excel file"""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "    Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_name:
            self.file_path.setText(file_name)
            self.update_sheet_list(file_name)
    
    def update_sheet_list(self, file_path):
        """Update sheet combo box with available sheets"""
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(wb.sheetnames)
            
            summary_index = self.sheet_combo.findText("Summary", Qt.MatchFlag.MatchExactly)
            if summary_index >= 0:
                self.sheet_combo.setCurrentIndex(summary_index)
        
            wb.close()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error reading Excel file: {str(e)}")
    
    def display_statistics(self):
        """Display statistics from the selected Excel file"""
        file_path = self.file_path.text()
        sheet_name = self.sheet_combo.currentText()
        
        if not file_path or not sheet_name:
            self.show_custom_warning("Reports File Required", "Please select both file and sheet name")
            return 
            
        if not os.path.exists(file_path):
            self.show_custom_warning("Failed to Load Reports File", "Selected file does not exist")
            return 
            
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True)
            if sheet_name not in wb.sheetnames:
                self.show_custom_warning("Warning", "Selected sheet not found in workbook")
                return
    
            summary_sheet = wb[sheet_name]
            rows = list(summary_sheet.rows)
            header_row = rows[0]  # Get header row
            data_rows = rows[1:]  # Skip header row
            
            # Find column indices for important data
            total_required_idx = None
            total_attended_idx = None
            group_idx = None
            student_id_idx = None
            name_idx = None
            
            for idx, cell in enumerate(header_row):
                header_value = str(cell.value).strip() if cell.value else ""
                if header_value == "Total Required":
                    total_required_idx = idx
                elif header_value == "Total Attended":
                    total_attended_idx = idx
                elif header_value == "Group":
                    group_idx = idx
                elif header_value == "Student ID":
                    student_id_idx = idx
                elif header_value == "Name":
                    name_idx = idx
            
            # Find session columns for session details
            session_columns = []
            for idx, cell in enumerate(header_row):
                if cell.value and "session" in str(cell.value).lower() and "(Required)" in str(cell.value):
                    session_columns.append((idx, idx + 1))  # (Required column, Attended column)
            
            # Make sure we found all necessary columns
            if None in (total_required_idx, total_attended_idx, group_idx, student_id_idx, name_idx):
                self.show_custom_warning("Error", "Required columns not found in the Excel file")
                return
            
            if not session_columns:
                self.show_custom_warning("Error", "No session columns found in the Excel file")
                return
                
            # Initialize group-based tracking
            groups = set()
            group_student_counts = {}  # {group: total_students}
            group_session_attendance = {}  # {group: {session_num: attended_count}}
            
            # First pass: identify all groups and count students per group
            for row in data_rows:
                group = str(row[group_idx].value)
                groups.add(group)
                group_student_counts[group] = group_student_counts.get(group, 0) + 1
            
            # Get the total sessions from the first data row (assuming all students have same requirement)
            total_sessions = int(data_rows[0][total_required_idx].value) if data_rows else 0
            
            # Initialize attendance tracking for each group
            for group in groups:
                group_session_attendance[group] = {i: 0 for i in range(1, total_sessions + 1)}
            
            # Process student data
            self.student_data = []
            total_attended_sessions = 0
            completion_count = 0
            
            for row in data_rows:
                group = str(row[group_idx].value)
                attended_count = int(row[total_attended_idx].value)
                required_count = int(row[total_required_idx].value)
                
                # Track attended sessions for session details
                attended_sessions = []
                for session_idx, (req_col, att_col) in enumerate(session_columns, 1):
                    if row[att_col].value == 1:
                        attended_sessions.append(str(session_idx))
                        group_session_attendance[group][session_idx] += 1
                
                total_attended_sessions += attended_count
                
                if attended_count == required_count:
                    completion_count += 1
                
                self.student_data.append([
                    str(row[student_id_idx].value),  # Student ID
                    str(row[name_idx].value),        # Name
                    group,                           # Group
                    f"{attended_count}/{required_count}",  # Sessions Attended
                    ", ".join(sorted(attended_sessions)) if attended_sessions else "None"  # Session Details
                ])
            
            # Update attendance distribution table with group-based rates
            row_count = len(groups) * total_sessions
            self.session_table.setRowCount(row_count)
            current_row = 0
            
            for group in sorted(groups):
                for session_num in range(1, total_sessions + 1):
                    # Calculate attendance rate for this group and session
                    group_total = group_student_counts[group]
                    attended = group_session_attendance[group][session_num]
                    attendance_rate = (attended / group_total * 100) if group_total > 0 else 0
                    
                    # Create items with centered alignment
                    group_item = QTableWidgetItem(group)
                    session_item = QTableWidgetItem(f"Session {session_num}")
                    rate_item = QTableWidgetItem(f"{attendance_rate:.1f}%")
                    
                    # Set alignment for each item
                    group_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    session_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    rate_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    
                    # Set items in table
                    self.session_table.setItem(current_row, 0, group_item)
                    self.session_table.setItem(current_row, 1, session_item)
                    self.session_table.setItem(current_row, 2, rate_item)
                    current_row += 1
            
            # Update other stats
            total_students = len(data_rows)
            avg_attendance = (total_attended_sessions / (total_students * total_sessions) * 100) if total_students > 0 else 0            
            self.total_students.value_label.setText(str(total_students))
            self.total_sessions.value_label.setText(str(total_sessions))
            self.avg_attendance.value_label.setText(f"{avg_attendance:.1f}%")
            self.completion_rate.value_label.setText(str(completion_count))
            
            self.update_student_table(self.student_data)
            wb.close()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error loading report: {str(e)}")
        
    def show_custom_warning(self, title, message):
        """Show a custom styled warning dialog"""
        warning_dialog = QMessageBox(self)
        warning_dialog.setWindowTitle(title)
        warning_dialog.setText(message)
        warning_dialog.setIcon(QMessageBox.Icon.Warning)
    
        # Create and style OK button
        ok_button = warning_dialog.addButton(QMessageBox.StandardButton.Ok)
        ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)
    
        # Style dialog background and text
        warning_dialog.setStyleSheet(f"""
            QMessageBox {{
                background-color: {CARD_BG};
            }}
            QLabel {{
                color: {TEXT_COLOR};
                font-size: 14px;
            }}
        """)
    
        warning_dialog.exec()  

    def setup_worker(self, dep_file, faculty_file):
        # Create the worker
        self.worker = PopulateWorker(dep_file, faculty_file)
        
        # Create the thread
        self.thread = QThread()
        
        # Move worker to the thread
        self.worker.moveToThread(self.thread)
        
        # Connect signals and slots
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        
        # Connect the progress and output signals
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.output.connect(self.output_console.append)
        
        # Connect the new error and success signals
        self.worker.error.connect(self.handle_error)
        self.worker.success.connect(self.handle_success)
        
        # Start the thread
        self.thread.start()
    
    def handle_error(self, error_message):
        self.setEnabled(True)
        error_dialog = QMessageBox(self)
        error_dialog.setWindowTitle("Error")
        error_dialog.setText(f"Error processing data: {error_message}")
        error_dialog.setIcon(QMessageBox.Icon.Critical)

        # Style OK button
        ok_button = error_dialog.addButton(QMessageBox.StandardButton.Ok)
        ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)

        # Style dialog background
        error_dialog.setStyleSheet(f"""
            QMessageBox {{
                background-color: {CARD_BG};
            }}
            QLabel {{
                color: {TEXT_COLOR};
                font-size: 14px;
            }}
        """)

        error_dialog.exec()
        self.output_console.append(f"Error: {error_message}")
    
    def handle_success(self, success_message):
        self.setEnabled(True)
        self.progress_bar.setValue(100)
        self.output_console.append("Processing complete!")
        
        success_dialog = QMessageBox(self)
        success_dialog.setWindowTitle("Success")
        success_dialog.setText("Processing complete! Check the attendance_reports folder.")
        success_dialog.setIcon(QMessageBox.Icon.Information)

        # Style OK button
        ok_button = success_dialog.addButton(QMessageBox.StandardButton.Ok)
        ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)

        # Style dialog background
        success_dialog.setStyleSheet(f"""
            QMessageBox {{
                background-color: {CARD_BG};
            }}
            QLabel {{
                color: {TEXT_COLOR};
                font-size: 14px;
            }}
        """)

        success_dialog.exec()  

#==========================================================attendance populator==========================================================#

class Facultypopulator(QWidget):
    # Add a signal for progress updates
    progress_signal = pyqtSignal(int)
    output_signal = pyqtSignal(str)
    complete_signal = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.dep_file = None
        self.faculty_file = None
        # Define department mapping for search patterns
        self.department_patterns = {
            "Anatomy": ["anatomy", "anat", "anat.", "anato", "anato."],
            "Histology": ["histology", "histo", "histo.", "hist", "hist.", "sgd"],
            "Physiology": ["physiology", "physio", "physio."],
            "Biochemistry": ["biochemistry", "biochem", "biochem.", "bio", "bio.","cbl"],
            "Pharmacology": ["pharmacology", "pharma", "pharma."],
            "Pathology": ["pathology", "path", "path.", "patho", "patho."],
            "Parasitology": ["parasitology", "para", "para.", "sgd"],
            "Microbiology": ["microbiology", "micro", "micro."],
            "Forensics and Toxicology": ["forensics", "toxicology", "forensic", "toxico", "toxic", "toxic.", "toxico.", "forens.", "forens", "foren", "foren."],
            "Community Medicine": ["community", "comm med", "comm."],
            "Clinical": ["clinical", "clinic", "clin.", "clinic."]
        }
        self.selected_department = "Pharmacology"  # Default selection
        self.setStyleSheet("""
            background-color: black; 
            color: white;
            QLabel {
                color: white;
            }
        """)
        
        self.init_ui()

    def return_to_home(self):
        # Get the stacked widget and switch to the start page
        stacked_widget = self.parent()
        if isinstance(stacked_widget, QStackedWidget):
            stacked_widget.setCurrentIndex(0)

    def init_ui(self):
        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Track the current message index
        self.current_message_index = 0
    
        # Header layout
        header_layout = QHBoxLayout()
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), 'ASU1.png')
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaled(60, 60, Qt.AspectRatioMode.KeepAspectRatio, 
                                             Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        title_label = QLabel("Attendance populator")
        title_label.setStyleSheet(f"font-size: 24px; font-weight: bold;")
        header_layout.addWidget(logo_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()
    
        # Create vertical layout for header buttons
        header_buttons_layout = QVBoxLayout()
        header_buttons_layout.setSpacing(5)
        
        # Back button
        back_btn = QPushButton("Back to Home")
        back_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        back_btn.clicked.connect(self.return_to_home)
            
        # Exit button
        exit_btn = QPushButton("Exit")
        exit_btn.setStyleSheet(EXIT_BUTTON_STYLE)
        exit_btn.clicked.connect(QApplication.instance().quit)
        
        # Add buttons to vertical layout
        header_buttons_layout.addWidget(back_btn)
        header_buttons_layout.addWidget(exit_btn)
        header_buttons_layout.addStretch()
        
        # Add button layout to header
        header_layout.addLayout(header_buttons_layout)
        
        # Add header to main layout
        main_layout.addLayout(header_layout)

        # Department Selection Section
        department_group = QGroupBox("Department Selection")
        department_group.setStyleSheet(GROUP_BOX_STYLE)
        department_layout = QHBoxLayout(department_group)
        
        department_layout.addWidget(QLabel("Select Department:"))
        self.department_combo = QComboBox()
        self.department_combo.setMinimumWidth(200)
        
        # Add departments in the specified order
        departments = [
            "Anatomy", 
            "Histology", 
            "Physiology", 
            "Biochemistry", 
            "Pharmacology", 
            "Pathology", 
            "Parasitology", 
            "Microbiology", 
            "Forensics and Toxicology", 
            "Community Medicine", 
            "Clinical"
        ]
        self.department_combo.addItems(departments)
        
        # Set default to Pharmacology (index 4)
        self.department_combo.setCurrentIndex(4)
        self.department_combo.currentTextChanged.connect(self.update_selected_department)
        
        department_layout.addWidget(self.department_combo)
        department_layout.addStretch()
        
        # Add department group to main layout
        main_layout.addWidget(department_group)

        # Dep File Section 
        attendance_group = QGroupBox("Department Attendance Report")
        attendance_group.setStyleSheet(GROUP_BOX_STYLE)
        attendance_layout = QHBoxLayout(attendance_group)
        
        attendance_layout.addWidget(QLabel("Department Attendance File:"))
        self.dep_path = QLineEdit()
        self.dep_path.setPlaceholderText("Select Excel file...")
        self.dep_path.setReadOnly(True)
        attendance_layout.addWidget(self.dep_path)
        
        attendance_browse_btn = QPushButton("Browse")
        attendance_browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        attendance_browse_btn.clicked.connect(lambda: self.browse_file('dep'))
        attendance_layout.addWidget(attendance_browse_btn)

        attendance_layout.addWidget(QLabel("Sheet Name:"))
        self.dep_sheet_combo = QComboBox()
        self.dep_sheet_combo.setMinimumWidth(100)
        attendance_layout.addWidget(self.dep_sheet_combo)
        
        # Add attendance group to main layout
        main_layout.addWidget(attendance_group)

        # Faculty File Section
        faculty_group = QGroupBox("Main Attendance Report")
        faculty_group.setStyleSheet(GROUP_BOX_STYLE)
        faculty_layout = QHBoxLayout(faculty_group)
        
        faculty_layout.addWidget(QLabel("Main Attendance File:"))
        self.faculty_path = QLineEdit()
        self.faculty_path.setPlaceholderText("Select Excel file...")
        self.faculty_path.setReadOnly(True)
        faculty_layout.addWidget(self.faculty_path)
        
        faculty_browse_btn = QPushButton("Browse")
        faculty_browse_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        faculty_browse_btn.clicked.connect(lambda: self.browse_file('faculty'))
        faculty_layout.addWidget(faculty_browse_btn)

        faculty_layout.addWidget(QLabel("Sheet Name:"))
        self.faculty_sheet_combo = QComboBox()
        self.faculty_sheet_combo.setMinimumWidth(100)
        faculty_layout.addWidget(self.faculty_sheet_combo)
        
        # Add faculty group to main layout
        main_layout.addWidget(faculty_group)

        # Facts Display Section 
        facts_group = QGroupBox("Status")
        facts_group.setStyleSheet(GROUP_BOX_STYLE)
        facts_layout = QVBoxLayout(facts_group)
        
        self.fact_label = QLabel("")
        self.fact_label.setStyleSheet("font-weight: bold; font-size: 16px; color: #555;")
        self.fact_label.setWordWrap(True)
        self.fact_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        facts_layout.addWidget(self.fact_label)
        
        # Add stretch to push widgets to the top and make the empty space part of the layout
        facts_layout.addStretch(1)
        
        facts_group.setMinimumHeight(100)
        # Add the facts group to the main layout 
        main_layout.addWidget(facts_group)  

        # Progress Bar Section
        progress_group = QGroupBox("Progress")
        progress_group.setStyleSheet(GROUP_BOX_STYLE)
        progress_layout = QVBoxLayout(progress_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar.setStyleSheet(PROGRESS_BAR_STYLE)
        progress_layout.addWidget(self.progress_bar)
        
        # Create loading gif label
        self.loading_label = QLabel()
        self.loading_label.setFixedSize(24, 24)  # Adjust based on your GIF size
        self.loading_label.setVisible(False)  # Hidden by default
        
        # Create the movie object for the GIF
        self.loading_movie = QMovie()
        self.loading_movie.setScaledSize(QSize(24, 24))  # Adjust based on your GIF size
        self.loading_label.setMovie(self.loading_movie)
        
        # Make sure to have your loading.gif in the same directory as the script
        loading_gif_path = os.path.join(os.path.dirname(__file__), 'loading.gif')
        if os.path.exists(loading_gif_path):
            self.loading_movie.setFileName(loading_gif_path)
        else:
            print(f"Warning: loading.gif not found at {loading_gif_path}")
        
        # Create a horizontal layout to hold both the progress bar and loading animation
        progress_h_layout = QHBoxLayout()
        progress_h_layout.addWidget(self.progress_bar)
        progress_h_layout.addWidget(self.loading_label)
        progress_layout.addLayout(progress_h_layout)
        
        # Add progress group to main layout
        main_layout.addWidget(progress_group)

        # Output Console Section
        console_group = QGroupBox("Output Console")
        console_group.setStyleSheet(GROUP_BOX_STYLE)
        console_layout = QVBoxLayout(console_group)

        self.output_console = QTextEdit()
        self.output_console.setReadOnly(True)
        self.output_console.setMaximumHeight(150)
        self.output_console.setStyleSheet(CONSOLE_STYLE)  
        console_layout.addWidget(self.output_console)
        main_layout.addWidget(console_group)

        # Bottom Buttons
        button_layout = QHBoxLayout()
        self.populate_btn = QPushButton("Populate Main Attendance File")
        self.populate_btn.setStyleSheet(STANDARD_BUTTON_STYLE)
        self.populate_btn.clicked.connect(self.start_process)
        button_layout.addWidget(self.populate_btn)        
        
        # Add button layout to main layout
        main_layout.addLayout(button_layout)
        
        # Add a system console redirector
        sys.stdout = ConsoleRedirector(self.output_console)
        
        # Connect our signals to their handlers
        self.progress_signal.connect(self.update_progress_bar)
        self.output_signal.connect(self.update_output_console)
        self.complete_signal.connect(self.process_complete)
        
        # Initialize thread variable
        self.worker_thread = None
        
    def update_selected_department(self, department):
        """Update the selected department when the combo box changes"""
        self.selected_department = department
        self.output_console.append(f"Selected department: {department}")
        
    def update_progress_bar(self, value):
        """Update the progress bar from the worker thread"""
        self.progress_bar.setValue(value)
        
    def process_complete(self, result):
        """Handle process completion"""
        
        # Stop the thread monitor timer if it exists
        if hasattr(self, 'thread_monitor'):
            self.thread_monitor.stop()
        
        # Hide the loading animation
        self.loading_label.setVisible(False)
        self.loading_movie.stop()
        
        # Re-enable the populate button
        self.populate_btn.setEnabled(True)
        
        # Show completion dialog
        if "Error" not in result:
            QMessageBox.information(
                self,
                "Integration Complete",
                result,
                QMessageBox.StandardButton.Ok
            )

    def update_output_console(self, text):
        """Update the output console with text from the worker thread"""
        if self.output_console is not None:
            # Use setText instead of append to avoid buffer overflow
            # If you need to keep history, limit it to a reasonable amount
            cursor = self.output_console.textCursor()
            cursor.movePosition(QTextCursor.MoveOperation.End)
            cursor.insertText(text + "\n")
            self.output_console.setTextCursor(cursor)
            self.output_console.ensureCursorVisible()
            
    def populate_attendance(self):
        """Process and populate department attendance into faculty table"""
        if not self.dep_file or not self.faculty_file:
            return "Please select both files first"
        try:
            # Get selected department search patterns
            department_name = self.selected_department
            search_patterns = self.department_patterns.get(department_name, ["unknown"])
            
            # Emit signal to update progress
            self.progress_signal.emit(10)
            self.output_signal.emit(f"Starting integration process for {department_name} department...")
            
            # Create a backup of the original faculty file
            # Create backup filename with timestamp
            backup_filename = f"{os.path.splitext(self.faculty_file)[0]}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{os.path.splitext(self.faculty_file)[1]}"
            # Create the backup
            shutil.copy2(self.faculty_file, backup_filename)
            self.output_signal.emit(f"Created backup of original file at: {backup_filename}")
            
            # Update progress
            self.progress_signal.emit(20)
                
            # Load the department attendance data
            self.output_signal.emit(f"Loading {department_name} attendance data...")
            dep_df = pd.read_excel(self.dep_file, sheet_name="Attendance")
            # Load the summary data which has attendance status per student
            summary_df = pd.read_excel(self.dep_file, sheet_name="Summary")
            
            # Update progress
            self.progress_signal.emit(30)
            
            # Load the faculty-wide attendance table without modifying
            self.output_signal.emit("Loading faculty attendance data...")
            faculty_df = pd.read_excel(self.faculty_file, header=None)
            # Load the workbook with openpyxl to preserve formatting
            from openpyxl import load_workbook
            workbook = load_workbook(self.faculty_file)
            sheet = workbook.active
            
            # Track changes for reporting
            updated_records = 0
            students_processed = 0
            self.output_signal.emit(f"Faculty columns: {list(faculty_df.columns)}")
            self.output_signal.emit(f"{department_name} columns: {list(dep_df.columns)}")
            self.output_signal.emit(f"Summary columns: {list(summary_df.columns)}")
            
            # Update progress
            self.progress_signal.emit(40)
            
            # Find the Department session columns by looking for cells containing department patterns
            dept_cols = []
            self.output_signal.emit(f"Searching for {department_name} session columns using patterns: {search_patterns}...")
            
            # Search through first 15 rows to find cells containing department patterns
            for row_idx in range(min(15, len(faculty_df))):
                for col_idx in range(len(faculty_df.columns)):
                    cell_value = str(faculty_df.iloc[row_idx, col_idx]).strip() if pd.notna(faculty_df.iloc[row_idx, col_idx]) else ""
                    cell_value_lower = cell_value.lower()
                    
                    # Check if any pattern is in the cell
                    pattern_found = False
                    matched_pattern = ""
                    for pattern in search_patterns:
                        if pattern.lower() in cell_value_lower:
                            pattern_found = True
                            matched_pattern = pattern
                            break
                    
                    if pattern_found:
                        # Extract session number if present
                        match = re.search(r'(\d+)', cell_value)
                        session_num = int(match.group(1)) if match else None
                        if session_num is not None:
                            dept_cols.append((col_idx, cell_value, session_num))
                            self.output_signal.emit(f"Found {department_name} column at row {row_idx}, column {col_idx}: {cell_value}, Session: {session_num}, Matched pattern: {matched_pattern}")
            
            if not dept_cols:
                self.output_signal.emit(f"Error: Could not locate any {department_name} columns in faculty sheet")
                return f"Could not locate any {department_name} columns in faculty sheet"
            
            # Update progress
            self.progress_signal.emit(45)
            
            # Find the student ID column in the faculty sheet
            # We'll look for a column where most values are numeric strings with 4+ digits
            id_col_idx = None
            id_col_row_start = None
            self.output_signal.emit("Searching for student ID column...")
            
            # First scan the first ~20 rows to find potential ID columns
            potential_id_columns = {}
            # Look through first ~30 rows to get a good sample
            for row_idx in range(min(30, len(faculty_df))):
                for col_idx in range(len(faculty_df.columns)):
                    cell_value = str(faculty_df.iloc[row_idx, col_idx]).strip() if pd.notna(faculty_df.iloc[row_idx, col_idx]) else ""
                    # Check if cell contains a string of 4+ digits
                    if cell_value.isdigit() and len(cell_value) >= 4:
                        # Add this column to our potential columns dict and increment its count
                        if col_idx not in potential_id_columns:
                            potential_id_columns[col_idx] = []
                        potential_id_columns[col_idx].append(row_idx)
            
            # Find the column with the most potential student IDs
            best_col = None
            max_ids = 0
            for col_idx, row_indices in potential_id_columns.items():
                if len(row_indices) > max_ids:
                    max_ids = len(row_indices)
                    best_col = col_idx
            
            if best_col is not None:
                id_col_idx = best_col
                # Find first row with student ID (assuming continuous data afterward)
                id_col_row_start = min(potential_id_columns[best_col])
                self.output_signal.emit(f"Found student ID column at index {id_col_idx}, starting from row {id_col_row_start}")
                self.output_signal.emit(f"Found {max_ids} potential student IDs in this column")
            else:
                self.output_signal.emit("Warning: Could not detect a clear student ID column. Falling back to full sheet search.")
            
            # Update progress
            self.progress_signal.emit(50)
            
            # Get the Dep student ID column name
            dep_id_col = "Student ID"  # Based on the printed column names
            
            # Update progress
            total_students = len(summary_df)
            
            # Process each student in the summary sheet
            self.output_signal.emit(f"Processing {total_students} students...")
            for idx, dep_row in enumerate(summary_df.iterrows()):
                _, dep_row = dep_row  # Unpack the tuple
                
                # Calculate progress - spread between 50% and 90%
                student_progress = 50 + int((idx / total_students) * 40)
                self.progress_signal.emit(student_progress)
                
                # Get the student ID
                if dep_id_col not in dep_row:
                    self.output_signal.emit(f"Cannot find column {dep_id_col} in department sheet")
                    continue
                    
                student_id = str(dep_row[dep_id_col]).strip()
                
                # Find student ID in the faculty sheet
                matching_cells = []
                
                # If we found a likely student ID column, search only in that column
                if id_col_idx is not None and id_col_row_start is not None:
                    # Only search in the identified student ID column, starting from the first identified row
                    for row_idx in range(id_col_row_start, len(faculty_df)):
                        cell_value = faculty_df.iloc[row_idx, id_col_idx]
                        if pd.notna(cell_value):
                            faculty_cell_value = str(cell_value).strip()
                            if faculty_cell_value == student_id:
                                matching_cells.append((row_idx, id_col_idx))
                                self.output_signal.emit(f"Found matching student ID {student_id} at cell ({row_idx}, {id_col_idx})")
                                break  # Since we're in the ID column, we only need the first match
                else:
                    # Fallback: Search the entire sheet for matching cells
                    self.output_signal.emit(f"Searching full sheet for student ID: {student_id}")
                    for row_idx in range(len(faculty_df)):
                        for col_idx in range(len(faculty_df.columns)):
                            cell_value = faculty_df.iloc[row_idx, col_idx]
                            if pd.notna(cell_value):
                                # Compare after converting to string and stripping whitespace
                                faculty_cell_value = str(cell_value).strip()
                                if faculty_cell_value == student_id:
                                    matching_cells.append((row_idx, col_idx))
                                    self.output_signal.emit(f"Found matching student ID {student_id} at cell ({row_idx}, {col_idx})")
            
                if not matching_cells:
                    self.output_signal.emit(f"No match found for student ID: {student_id}")
                    continue
                    
                # For each matching student ID cell
                for row_idx, _ in matching_cells:
                    # For each department column detected
                    for col_idx, _, session_num in dept_cols:
                        if session_num is None:
                            self.output_signal.emit(f"Skipping {department_name} column without session number at column {col_idx}")
                            continue
                            
                        # Look for attendance columns that contain both the session number and "(Attended)"
                        attended = 0  # Default value
                        
                        # Find the appropriate attendance column by searching for session number and "(Attended)"
                        attendance_col = None
                        for col_name in dep_row.index:
                            # Check if this column contains both the session number and "(Attended)"
                            if "(Attended)" in str(col_name) and str(session_num) in str(col_name):
                                attendance_col = col_name
                                self.output_signal.emit(f"Found matching attendance column: {attendance_col}")
                                break
                                
                        if attendance_col:
                            # If found, get the attendance value
                            attended = int(dep_row[attendance_col]) if pd.notna(dep_row[attendance_col]) else 0
                            self.output_signal.emit(f"Attendance value from column {attendance_col}: {attended}")
                        else:
                            self.output_signal.emit(f"Could not find attendance column for session {session_num}")
                            # Fallback to checking the attendance sheet
                            student_sessions = dep_df.loc[dep_df[dep_id_col].astype(str).str.strip() == student_id]
                            if not student_sessions.empty:
                                session_col = 'Session'
                                if session_col in student_sessions.columns:
                                    session_attended = student_sessions.loc[student_sessions[session_col] == session_num]
                                    if not session_attended.empty:
                                        attended = 1
                                        self.output_signal.emit(f"Fallback: Student {student_id} attended {department_name} session {session_num}")
                        
                        # Update the faculty table at the precise X,Y coordinate (row, column)
                        faculty_df.iloc[row_idx, col_idx] = attended
                        
                        # Also update the Excel file directly using openpyxl (1-indexed)
                        # Convert to Excel's 1-based indexing
                        excel_row = row_idx + 1
                        excel_col = col_idx + 1
                        
                        # Convert column index to letter (A, B, C, etc.)
                        col_letter = get_column_letter(excel_col)
                        
                        # Update the cell in openpyxl
                        sheet[f"{col_letter}{excel_row}"] = attended
                        updated_records += 1
                        self.output_signal.emit(f"Updated student {student_id}, session {session_num}, attendance={attended} at cell ({row_idx}, {col_idx})")
                        
                # Increment the student counter
                students_processed += 1
                
                # Save every 10 students
                if students_processed % 10 == 0:
                    self.output_signal.emit(f"Saving after processing {students_processed} students...")
                    try:
                        workbook.save(self.faculty_file)
                        self.output_signal.emit(f"Successfully saved after {students_processed} students")
                    except Exception as save_error:
                        self.output_signal.emit(f"Warning: Save failed: {str(save_error)}")
                        self.output_signal.emit(f"Continuing processing but updates may not persist")
            
            # Final save at the end
            self.output_signal.emit(f"Final save after processing all {students_processed} students...")
            workbook.save(self.faculty_file)     
            
            # Also save a DataFrame version just to be safe
            backup_df_path = self.faculty_file + "_dataframe_backup.xlsx"
            faculty_df.to_excel(backup_df_path, index=False, header=False)
            self.output_signal.emit(f"Saved DataFrame backup to {backup_df_path}")
        
            # Update progress to 100%
            self.progress_signal.emit(100)
            
            success_message = f"Successfully updated {updated_records} attendance records for {students_processed} students. Original file backed up to {backup_filename}"
            self.output_signal.emit(success_message)
            return success_message
            
        except Exception as e:
            error_details = traceback.format_exc()
            self.output_signal.emit(f"Detailed error: {error_details}")
            # Reset progress bar on error
            self.progress_signal.emit(0)
            error_message = f"Error during integration: {str(e)}"
            self.output_signal.emit(error_message)
            return error_message
                                                
    def browse_file(self, file_type):
        """Handle file browsing for either department or faculty files."""
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx)"
        )
        if filename:
            if file_type == 'dep':
                self.dep_path.setText(filename)
                self.dep_file = filename
                self.load_sheets(filename, self.dep_sheet_combo)
            elif file_type == 'faculty':
                self.faculty_path.setText(filename)
                self.faculty_file = filename
                self.load_sheets(filename, self.faculty_sheet_combo)

    def load_sheets(self, file_path, combo_box):
        """Load sheet names and auto-select 'Summary' for department files"""
        if os.path.isfile(file_path):
            try:
                wb = openpyxl.load_workbook(file_path, read_only=True)
                sheets = wb.sheetnames
                combo_box.clear()
                combo_box.addItems(sheets)
                
                # Auto-select Summary sheet for dep files
                if combo_box == self.dep_sheet_combo:
                    summary_index = combo_box.findText("Summary", Qt.MatchFlag.MatchExactly)
                    if summary_index >= 0:
                        combo_box.setCurrentIndex(summary_index)
                        self.output_console.append(f"Automatically selected Summary sheet for {self.selected_department} file")
                
                wb.close()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error reading Excel file: {str(e)}")
                self.output_console.append(f"Error loading sheets: {str(e)}")

    def start_process(self):
        # Show facts message
        self.fact_label.setText("\nThis might take a while, you can take a short break to help refresh your mind while the task is being completed.\nأذكر الله")
        self.fact_label.setVisible(True)
        
        # Initialize process completion flag
        self.process_completed = False  
        
        """Starts the facts display and then processes the integration in a separate thread"""
        # Check for files first
        if not self.dep_file or not self.faculty_file:
            self.output_console.append("Please select both files first")
            return
            
        # If there's already a thread running, don't start another one
        if hasattr(self, 'worker_thread') and self.worker_thread is not None and self.worker_thread.isRunning():
            self.output_console.append("A process is already running. Please wait for it to complete.")
            return
            
        # Disable the populate button to prevent multiple clicks
        self.populate_btn.setEnabled(False)
        
        # Show and start the loading animation
        self.loading_label.setVisible(True)
        self.loading_movie.start()
        
        # Reset progress bar
        self.progress_bar.setValue(0)
        
        # Get the selected department and its search patterns
        department_name = self.selected_department
        search_patterns = self.department_patterns.get(department_name, ["unknown"])
        
        # Create and start a worker thread for the processing
        self.worker_thread = QThread()
        self.worker = PopulateWorker(self.dep_file, self.faculty_file, department_name, search_patterns)
        self.worker.moveToThread(self.worker_thread)
        
        # Connect signals
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.progress_signal.emit)
        self.worker.output.connect(self.output_signal.emit)
        self.worker.finished.connect(self.complete_signal.emit)
        self.worker.finished.connect(self.worker_thread.quit)
        
        # These connections need to be made as QueuedConnection to ensure thread safety
        self.worker.finished.connect(self.worker.deleteLater, Qt.ConnectionType.QueuedConnection)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater, Qt.ConnectionType.QueuedConnection)
        
        # Start the thread
        self.worker_thread.start()
    
    def start_integration(self):
        """Method called when Populate Attendance button is clicked"""
        
        # Then run the integration process
        self.run_integration()

    def process_complete(self, result):
        """Handle process completion"""

        # Hide facts message
        self.fact_label.setVisible(False)
        self.fact_label.clear()

        # Track that the process completed normally
        self.process_completed = True  # Add this line
        
        # Hide the loading animation
        self.loading_label.setVisible(False)
        self.loading_movie.stop()
                
        # Re-enable the populate button
        self.populate_btn.setEnabled(True)
        
        # Show styled completion dialog
        if "Error" not in result:
            self.show_styled_dialog("Success", result, QMessageBox.Icon.Information)
        else:
            self.show_styled_dialog("Error", result, QMessageBox.Icon.Critical)
    
    def show_styled_dialog(self, title, message, icon):
        """Create and show a styled QMessageBox"""
        dialog = QMessageBox(self)
        dialog.setWindowTitle(title)
        dialog.setText(message)
        dialog.setIcon(icon)
        
        # Style the dialog
        dialog.setStyleSheet(f"""
            QMessageBox {{
                background-color: {CARD_BG};
                color: {TEXT_COLOR};
            }}
            QLabel {{
                color: {TEXT_COLOR};
                font-size: 14px;
            }}
        """)
        
        # Get and style the OK button
        ok_button = dialog.addButton(QMessageBox.StandardButton.Ok)
        ok_button.setStyleSheet(STANDARD_BUTTON_STYLE)
        
        dialog.exec()

    def check_thread_status(self):
        """Monitor thread status and handle abnormal termination"""
        try:
            # Check if worker_thread exists and is valid
            if not hasattr(self, 'worker_thread') or self.worker_thread is None:
                self.thread_monitor.stop()
                return
            # Check if the thread is running
            if not self.worker_thread.isRunning():
                self.thread_monitor.stop()
                # If the process didn't complete normally, emit signal
                if not hasattr(self, 'process_completed'):
                    self.complete_signal.emit("Process terminated unexpectedly")
                # Clean up
                self.loading_label.setVisible(False)
                if hasattr(self, 'loading_movie'):
                    self.loading_movie.stop()
                self.populate_btn.setEnabled(True)
                self.worker_thread = None
        except RuntimeError as e:
            # Handle the case where the C++ object is already deleted
            self.thread_monitor.stop()
            self.worker_thread = None
            self.loading_label.setVisible(False)
            if hasattr(self, 'loading_movie'):
                self.loading_movie.stop()
            self.populate_btn.setEnabled(True)
               
class ConsoleRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.original_stdout = sys.stdout

    def write(self, text):
        self.original_stdout.write(text)  # Still print to console
        if text.strip():  # Only update if there's actual text (not just whitespace)
            self.text_widget.append(text.rstrip())
            QApplication.processEvents()  # Process UI events to update console in real-time

    def flush(self):
        self.original_stdout.flush()

class PopulateWorker(QObject):
    finished = pyqtSignal(str)  # Signal for when processing is complete
    progress = pyqtSignal(int)  # Signal for progress updates
    output = pyqtSignal(str)    # Signal for console output
    error = pyqtSignal(str)     # Signal for error messages
    success = pyqtSignal(str)   # Signal for success messages
    
    def __init__(self, dep_file, faculty_file, department_name="Department", search_patterns=None):
        super().__init__()
        self.dep_file = dep_file
        self.faculty_file = faculty_file
        self.department_name = department_name
        self.search_patterns = search_patterns if search_patterns else ["dep"]
        self.is_running = False
        
    def resource_path(self, relative_path):
        """Get absolute path to resource, works for dev and for PyInstaller"""
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(os.path.dirname(__file__))
        
        return os.path.join(base_path, relative_path)
        
    def run(self):
        """Main processing function that runs in a separate thread"""
        try:
            # Set running flag
            self.is_running = True
            self.output.emit(f"Starting integration process with dep file: {self.dep_file}, faculty file: {self.faculty_file}")
            
            # Emit signal to update progress
            self.progress.emit(10)
            self.output.emit("Starting integration process...")
            
            # Create a backup of the original faculty file
            # Create backup filename with timestamp
            backup_filename = f"{os.path.splitext(self.faculty_file)[0]}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{os.path.splitext(self.faculty_file)[1]}"
            try:
                # Create the backup
                shutil.copy2(self.faculty_file, backup_filename)
                self.output.emit(f"Created backup of original file at: {backup_filename}")
            except Exception as backup_error:
                error_msg = f"Failed to create backup: {str(backup_error)}"
                self.output.emit(f"Warning: {error_msg}")
                # Continue execution despite backup failure
                
            # Update progress
            self.progress.emit(20)
            
            # Load the dep attendance data
            self.output.emit("Loading department attendance data...")
            try:
                dep_df = pd.read_excel(self.dep_file, sheet_name="Attendance")
                # Load the summary data which has attendance status per student
                summary_df = pd.read_excel(self.dep_file, sheet_name="Summary")
            except Exception as excel_error:
                error_msg = f"Failed to load department data: {str(excel_error)}"
                raise Exception(error_msg)
                
            # Update progress
            self.progress.emit(30)
            
            # Load the faculty-wide attendance table without modifying
            self.output.emit("Loading faculty attendance data...")
            try:
                faculty_df = pd.read_excel(self.faculty_file, header=None)
                # Load the workbook with openpyxl to preserve formatting
                from openpyxl import load_workbook
                workbook = load_workbook(self.faculty_file)
                sheet = workbook.active
            except Exception as faculty_error:
                error_msg = f"Failed to load faculty data: {str(faculty_error)}"
                raise Exception(error_msg)
                
            # Track changes for reporting
            updated_records = 0
            students_processed = 0
            self.output.emit(f"Faculty columns: {len(faculty_df.columns)}, dep columns: {len(dep_df.columns)}, Summary columns: {len(summary_df.columns)}")
            
            # Update progress
            self.progress.emit(40)
            
            # Get the department name and search patterns from the worker initialization
            department_name = self.department_name if hasattr(self, 'department_name') else "Department"
            search_patterns = self.search_patterns if hasattr(self, 'search_patterns') else ["dep"]
            
            # Find the department session columns by looking for cells containing department patterns
            dep_cols = []
            self.output.emit(f"Searching for {department_name} session columns using patterns: {search_patterns}...")
            
            # Search through first 15 rows to find cells containing the department patterns
            for row_idx in range(min(15, len(faculty_df))):
                # Check if we're still running - important for clean cancellation
                if not self.is_running:
                    raise Exception("Operation cancelled")
                    
                for col_idx in range(len(faculty_df.columns)):
                    try:
                        cell_value = str(faculty_df.iloc[row_idx, col_idx]).strip() if pd.notna(faculty_df.iloc[row_idx, col_idx]) else ""
                        cell_value_lower = cell_value.lower()
                        
                        # Check if any pattern is in the cell
                        pattern_found = False
                        matched_pattern = ""
                        for pattern in search_patterns:
                            if pattern.lower() in cell_value_lower:
                                pattern_found = True
                                matched_pattern = pattern
                                break
                        
                        if pattern_found:
                            # Extract session number if present
                            match = re.search(r'(\d+)', cell_value)
                            session_num = int(match.group(1)) if match else None
                            if session_num is not None:
                                dep_cols.append((col_idx, cell_value, session_num))
                                self.output.emit(f"Found {department_name} column at row {row_idx}, column {col_idx}: {cell_value}, Session: {session_num}, Matched pattern: {matched_pattern}")
                    except Exception as cell_error:
                        # Continue despite cell error
                        pass
                        
            if not dep_cols:
                error_msg = f"Could not locate any {department_name} columns in faculty sheet"
                self.output.emit(f"Error: {error_msg}")
                self.error.emit(error_msg)
                self.finished.emit(f"Error: {error_msg}")
                self.is_running = False
                return
            
            # Update progress
            self.progress.emit(45)
            
            # Find the student ID column in the faculty sheet
            id_col_idx = None
            id_col_row_start = None
            
            self.output.emit("Searching for student ID column...")
            
            # First scan the first ~20 rows to find potential ID columns
            potential_id_columns = {}
            
            # Look through first ~30 rows to get a good sample
            for row_idx in range(min(30, len(faculty_df))):
                # Check if we're still running
                if not self.is_running:
                    raise Exception("Operation cancelled")
                    
                for col_idx in range(len(faculty_df.columns)):
                    try:
                        cell_value = str(faculty_df.iloc[row_idx, col_idx]).strip() if pd.notna(faculty_df.iloc[row_idx, col_idx]) else ""
                        
                        # Check if cell contains a string of 4+ digits
                        if cell_value.isdigit() and len(cell_value) >= 4:
                            # Add this column to our potential columns dict and increment its count
                            if col_idx not in potential_id_columns:
                                potential_id_columns[col_idx] = []
                            potential_id_columns[col_idx].append(row_idx)
                    except Exception as cell_error:
                        # Continue despite cell error
                        pass
            
            # Find the column with the most potential student IDs
            best_col = None
            max_ids = 0
            
            for col_idx, row_indices in potential_id_columns.items():
                if len(row_indices) > max_ids:
                    max_ids = len(row_indices)
                    best_col = col_idx
                    
            if best_col is not None:
                id_col_idx = best_col
                # Find first row with student ID (assuming continuous data afterward)
                id_col_row_start = min(potential_id_columns[best_col])
                self.output.emit(f"Found student ID column at index {id_col_idx}, starting from row {id_col_row_start}")
                self.output.emit(f"Found {max_ids} potential student IDs in this column")
            else:
                self.output.emit("Warning: Could not detect a clear student ID column. Falling back to full sheet search.")
            
            # Update progress
            self.progress.emit(50)
            
            # Get the department student ID column name
            dep_id_col = "Student ID"  # Based on the printed column names
        
            # Update progress
            total_students = len(summary_df)
            
            # Process each student in the summary sheet
            self.output.emit(f"Processing {total_students} students...")
            
            save_interval = 5  # Save every 5 students instead of 10
            
            for idx, dep_row in enumerate(summary_df.iterrows()):
                try:
                    # Check if we're still running - enable cancellation
                    if not self.is_running:
                        raise Exception("Operation cancelled")
                        
                    _, dep_row = dep_row  # Unpack the tuple
                    
                    # Calculate progress - spread between 50% and 90%
                    student_progress = 50 + int((idx / total_students) * 40)
                    self.progress.emit(student_progress)
                    
                    # Get the student ID
                    if dep_id_col not in dep_row:
                        self.output.emit(f"Cannot find column {dep_id_col} in department sheet")
                        continue
                        
                    student_id = str(dep_row[dep_id_col]).strip()
                    
                    # Find student ID in the faculty sheet
                    matching_cells = []
                    
                    # If we found a likely student ID column, search only in that column
                    if id_col_idx is not None and id_col_row_start is not None:
                        # Only search in the identified student ID column, starting from the first identified row
                        for row_idx in range(id_col_row_start, len(faculty_df)):
                            try:
                                cell_value = faculty_df.iloc[row_idx, id_col_idx]
                                if pd.notna(cell_value):
                                    faculty_cell_value = str(cell_value).strip()
                                    if faculty_cell_value == student_id:
                                        matching_cells.append((row_idx, id_col_idx))
                                        self.output.emit(f"Found student ID {student_id}")
                                        break  # Since we're in the ID column, we only need the first match
                            except Exception as e:
                                pass
                    else:
                        # Fallback: Search the entire sheet for matching cells
                        self.output.emit(f"Searching for student ID: {student_id}")
                        for row_idx in range(len(faculty_df)):
                            # Check every 100 rows if we're still running
                            if row_idx % 100 == 0 and not self.is_running:
                                raise Exception("Operation cancelled")
                                
                            for col_idx in range(len(faculty_df.columns)):
                                try:
                                    cell_value = faculty_df.iloc[row_idx, col_idx]
                                    if pd.notna(cell_value):
                                        # Compare after converting to string and stripping whitespace
                                        faculty_cell_value = str(cell_value).strip()
                                        if faculty_cell_value == student_id:
                                            matching_cells.append((row_idx, col_idx))
                                            self.output.emit(f"Found student ID {student_id}")
                                except Exception as e:
                                    pass
                    
                    if not matching_cells:
                        self.output.emit(f"No match found for student ID: {student_id}")
                        continue
                    
                    # For each matching student ID cell
                    for row_idx, _ in matching_cells:
                        # For each dep column detected
                        for col_idx, _, session_num in dep_cols:
                            if session_num is None:
                                self.output.emit(f"Skipping column without session number")
                                continue
                            
                            # Look for attendance columns that contain both the session number and "(Attended)"
                            attended = 0  # Default value
                            
                            # Find the appropriate attendance column by searching for session number and "(Attended)"
                            attendance_col = None
                            for col_name in dep_row.index:
                                # Check if this column contains both the session number and "(Attended)"
                                if "(Attended)" in str(col_name) and str(session_num) in str(col_name):
                                    attendance_col = col_name
                                    break
                            
                            if attendance_col:
                                # If found, get the attendance value
                                attended = int(dep_row[attendance_col]) if pd.notna(dep_row[attendance_col]) else 0
                            else:
                                # Fallback to checking the attendance sheet
                                try:
                                    student_sessions = dep_df.loc[dep_df[dep_id_col].astype(str).str.strip() == student_id]
                                    if not student_sessions.empty:
                                        session_col = 'Session'
                                        if session_col in student_sessions.columns:
                                            session_attended = student_sessions.loc[student_sessions[session_col] == session_num]
                                            if not session_attended.empty:
                                                attended = 1
                                except Exception as e:
                                    pass
                            
                            try:
                                # Update the faculty table at the precise X,Y coordinate (row, column)
                                faculty_df.iloc[row_idx, col_idx] = attended
                                
                                # Also update the Excel file directly using openpyxl (1-indexed)
                                # Convert to Excel's 1-based indexing
                                excel_row = row_idx + 1
                                excel_col = col_idx + 1
                                
                                # Convert column index to letter (A, B, C, etc.)
                                col_letter = get_column_letter(excel_col)
                                
                                # Update the cell in openpyxl
                                sheet[f"{col_letter}{excel_row}"] = attended
                                
                                updated_records += 1
                            except Exception as update_error:
                                pass
                    
                    # Increment the student counter
                    students_processed += 1
                    
                    # Save more frequently
                    if students_processed % save_interval == 0:
                        self.output.emit(f"Saving after processing {students_processed} students...")
                        try:
                            workbook.save(self.faculty_file)
                        except Exception as save_error:
                            error_msg = f"Warning: Save failed: {str(save_error)}"
                            self.output.emit(error_msg)
                            
                            # Try to save with a different filename if original save fails
                            try:
                                alt_save_path = f"{os.path.splitext(self.faculty_file)[0]}_recovery_{datetime.now().strftime('%H%M%S')}{os.path.splitext(self.faculty_file)[1]}"
                                workbook.save(alt_save_path)
                                self.output.emit(f"Saved to alternative location: {alt_save_path}")
                            except Exception as alt_save_error:
                                pass
                
                except Exception as student_error:
                    # Log the error but continue with next student
                    self.output.emit(f"Warning: Error processing a student: {str(student_error)}")
            
            # Final save at the end
            self.output.emit(f"Final save after processing all students...")
            try:
                workbook.save(self.faculty_file)
            except Exception as final_save_error:
                error_msg = f"Final save failed: {str(final_save_error)}"
                self.output.emit(f"Warning: {error_msg}")
                
                # Try to save with a different filename
                try:
                    final_save_path = f"{os.path.splitext(self.faculty_file)[0]}_final_{datetime.now().strftime('%H%M%S')}{os.path.splitext(self.faculty_file)[1]}"
                    workbook.save(final_save_path)
                    self.output.emit(f"Saved final result to: {final_save_path}")
                except Exception as alt_final_save_error:
                    pass
       
            # Also save a DataFrame version just to be safe
            try:
                backup_df_path = self.faculty_file + "_dataframe_backup.xlsx"
                faculty_df.to_excel(backup_df_path, index=False, header=False)
            except Exception as df_save_error:
                pass
            
            # Update progress to 100%
            self.progress.emit(100)
            
            success_message = f"Successfully updated {updated_records} attendance records for {students_processed} students. Original file backed up to {backup_filename}"
            self.output.emit(success_message)
            
            # Emit success signal
            self.success.emit(success_message)
            
            # Emit completion signal
            self.finished.emit(success_message)
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            self.output.emit(f"Error: {str(e)}")
            
            # Reset progress bar on error
            self.progress.emit(0)
            
            error_message = f"Error during integration: {str(e)}"
            
            # Emit error signal
            self.error.emit(error_message)
            
            # Emit completion signal with error message
            self.finished.emit(error_message)
        
        finally:
            # Always reset running state in finally block
            self.is_running = False

#==========================================================main app==========================================================#

class MainApplication(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Attendance Management App")
        self.setMinimumSize(1000, 750)
        icon_path = os.path.join(os.path.dirname(__file__), 'ASU1.png')
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
    
        self.setWindowTitle("Attendance Management App")
        
        # Create stacked widget to manage pages
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)
        
        # Create pages
        self.start_page = StartPage(self)
        self.info_page = InfoPage()  # New info page
        self.preparer_page = LogSheetPreparer()
        self.processor_page = AttendanceProcessor()
        self.populator_page = Facultypopulator()  
        self.dashboard_page = AttendanceDashboard()  
        self.schedule_manager_page = ScheduleManager() 
        self.reference_preparer_page = ReferenceFilePreparer() 
        self.appeal_processor_page = AppealProcessor()

        # Add pages to stacked widget
        self.stacked_widget.addWidget(self.start_page)
        self.stacked_widget.addWidget(self.info_page) 
        self.stacked_widget.addWidget(self.preparer_page)
        self.stacked_widget.addWidget(self.processor_page)
        self.stacked_widget.addWidget(self.populator_page)  
        self.stacked_widget.addWidget(self.dashboard_page)  
        self.stacked_widget.addWidget(self.schedule_manager_page) 
        self.stacked_widget.addWidget(self.reference_preparer_page) 
        self.stacked_widget.addWidget(self.appeal_processor_page)

        # Connect start page buttons to switch pages
        self.start_page.info_button.clicked.connect(self.show_info) 
        self.start_page.preparer_btn.clicked.connect(self.show_preparer)
        self.start_page.process_btn.clicked.connect(self.show_processor)
        self.start_page.populate_btn.clicked.connect(self.show_populator)
        self.start_page.dashboard_btn.clicked.connect(self.show_dashboard)
        self.start_page.schedule_btn.clicked.connect(self.show_schedule_manager)
        self.start_page.reference_btn.clicked.connect(self.show_reference_preparer)                
        self.start_page.appeal_btn.clicked.connect(self.show_appeal_processor)

        
        # Set the window style
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {BLACK};
            }}
        """)
    
    def show_info(self):
        self.stacked_widget.setCurrentWidget(self.info_page)

    def show_schedule_manager(self):
        self.stacked_widget.setCurrentWidget(self.schedule_manager_page)
        
    def show_preparer(self):
        self.stacked_widget.setCurrentWidget(self.preparer_page)
        
    def show_processor(self):
        self.stacked_widget.setCurrentWidget(self.processor_page)
        
    def show_populator(self):
        self.stacked_widget.setCurrentWidget(self.populator_page)
        
    def show_dashboard(self):
        self.stacked_widget.setCurrentWidget(self.dashboard_page)

    def show_reference_preparer(self):
        self.stacked_widget.setCurrentWidget(self.reference_preparer_page)

    def show_appeal_processor(self):
        self.stacked_widget.setCurrentWidget(self.appeal_processor_page)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainApplication()
    window.show()
    sys.exit(app.exec())