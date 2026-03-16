import sys
import os
import sqlite3
import json
import imaplib
import email
import smtplib
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import parsedate_to_datetime
import keyring
from cryptography.fernet import Fernet
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QTreeWidget, QTreeWidgetItem, QTextEdit, QLabel, QPushButton, QCalendarWidget, QFrame, QDialog, QFormLayout, QLineEdit, QDialogButtonBox, QMenuBar, QMenu, QMessageBox, QToolBar
from PyQt5.QtCore import Qt, QDate, QSize
from PyQt5.QtGui import QIcon, QFont, QTextCursor, QPixmap

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# --- CONFIGURATION ---
# Fill this in with the details from your client_secret.json file.
# This removes the need for users to have the file locally.
GOOGLE_CLIENT_CONFIG = {
    "installed": {
        "client_id": "PASTE_YOUR_CLIENT_ID_HERE",
        "client_secret": "PASTE_YOUR_CLIENT_SECRET_HERE",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": ["http://localhost"]
    }
}

class AccountSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Account Settings")
        self.setGeometry(100, 100, 400, 250)

        layout = QFormLayout()

        self.email_input = QLineEdit()
        layout.addRow("Email:", self.email_input)

        self.signin_button = QPushButton("Sign in with Google (Recommended)")
        self.signin_button.clicked.connect(self.signin)
        layout.addRow(self.signin_button)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        layout.addRow("App Password:", self.password_input)

        app_password_label = QLabel("For password-based login with Gmail, generate and use an App Password from your Google Account security settings.")
        app_password_label.setWordWrap(True)
        layout.addRow("", app_password_label)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

        self.setLayout(layout)

        self.credentials = None

    def signin(self):
        if "PASTE_YOUR" in GOOGLE_CLIENT_CONFIG["installed"]["client_id"]:
            QMessageBox.critical(self, "Setup Required",
                "You must paste your real Google Client ID and Secret into the 'OpenOutlook.py' file.\n\n"
                "Look for the GOOGLE_CLIENT_CONFIG dictionary near the top of the script."
            )
            return

        # Use embedded config instead of looking for a file
        flow = InstalledAppFlow.from_client_config(
            GOOGLE_CLIENT_CONFIG,
            scopes=["https://mail.google.com/"],
        )
        self.credentials = flow.run_local_server(port=0)
        
        # Use Gmail API to get the email address of the authenticated user
        service = build('gmail', 'v1', credentials=self.credentials)
        profile = service.users().getProfile(userId='me').execute()
        self.email_input.setText(profile['emailAddress'])


    def get_settings(self):
        auth_method = "oauth" if self.credentials else "password"
        settings = {
            "email": self.email_input.text(),
            "auth_method": auth_method,
            "imap_server": "imap.gmail.com",
            "imap_port": 993,
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 465
        }
        if auth_method == "oauth":
            settings["credentials"] = self.credentials
        else: # password
            settings["password"] = self.password_input.text()
        
        return settings

class SortableTreeWidgetItem(QTreeWidgetItem):
    def __lt__(self, other):
        column = self.treeWidget().sortColumn()
        try:
            # Column 2 is 'Received' date
            if column == 2:
                # Using a custom role for the datetime object
                date1 = self.data(0, Qt.UserRole + 1) 
                date2 = other.data(0, Qt.UserRole + 1)
                if date1 and date2:
                    return date1 < date2
            # Fallback to default string comparison for other columns
            return self.text(column).lower() < other.text(column).lower()
        except (TypeError, ValueError):
            return self.text(column).lower() < other.text(column).lower()

class ComposeWindow(QDialog):
    def __init__(self, parent=None, account=None, mode='new', original_msg=None):
        super().__init__(parent)
        self.parent_window = parent # To call send_email
        self.account = account
        self.mode = mode
        self.original_msg = original_msg

        self.setWindowTitle("Compose Email")
        self.setGeometry(150, 150, 700, 500)
        
        layout = QVBoxLayout()
        form_layout = QFormLayout()

        self.to_input = QLineEdit()
        self.cc_input = QLineEdit()
        self.subject_input = QLineEdit()
        
        form_layout.addRow("To:", self.to_input)
        form_layout.addRow("Cc:", self.cc_input)
        form_layout.addRow("Subject:", self.subject_input)
        
        layout.addLayout(form_layout)

        self.body_input = QTextEdit()
        layout.addWidget(self.body_input)

        self.buttons = QDialogButtonBox()
        self.buttons.addButton("Send", QDialogButtonBox.ButtonRole.AcceptRole)
        self.buttons.addButton(QDialogButtonBox.StandardButton.Cancel)
        self.buttons.accepted.connect(self.send)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)

        self.setLayout(layout)
        self.prepare_fields()

    def prepare_fields(self):
        if self.mode in ['reply', 'reply_all', 'forward'] and self.original_msg:
            if self.mode == 'reply':
                self.setWindowTitle(f"Re: {self.original_msg['subject']}")
                self.to_input.setText(self.original_msg.get('from', ''))
                self.subject_input.setText(f"Re: {self.original_msg['subject']}")
            elif self.mode == 'reply_all':
                self.setWindowTitle(f"Re: {self.original_msg['subject']}")
                self.to_input.setText(self.original_msg.get('from', ''))
                all_recipients = [self.original_msg.get('to', ''), self.original_msg.get('cc', '')]
                self.cc_input.setText(', '.join(filter(None, all_recipients)))
                self.subject_input.setText(f"Re: {self.original_msg['subject']}")
            elif self.mode == 'forward':
                self.setWindowTitle(f"Fw: {self.original_msg['subject']}")
                self.subject_input.setText(f"Fw: {self.original_msg['subject']}")

            quoted_body = f"\n\n---- Original Message ----\nFrom: {self.original_msg.get('from', '')}\nDate: {self.original_msg.get('date', '')}\nSubject: {self.original_msg.get('subject', '')}\nTo: {self.original_msg.get('to', '')}\n\n{self.original_msg.get('body', '')}"
            self.body_input.setText(quoted_body)
            self.body_input.moveCursor(QTextCursor.Start)

    def send(self):
        self.parent_window.send_email(self.account, self.to_input.text(), self.cc_input.text(), self.subject_input.text(), self.body_input.toPlainText())
        self.accept()

class OutlookLookalike(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OpenOutlook")
        self.setGeometry(100, 100, 800, 600)

        # Get the directory of the script
        self.script_dir = os.path.dirname(os.path.abspath(__file__))

        # Settings directory (Windows: %APPDATA%\OpenOutlook)
        if sys.platform == "win32":
            self.settings_dir = os.path.join(os.getenv("APPDATA"), "OpenOutlook")
        else:
            self.settings_dir = os.path.join(os.path.expanduser("~"), ".openoutlook")
        os.makedirs(self.settings_dir, exist_ok=True)
        self.settings_file = os.path.join(self.settings_dir, "settings.json")
        print(f"Settings file path: {self.settings_file}")  # Debug statement
        self.load_settings()
 
        # Encryption setup
        self.load_or_create_encryption_key()
 
        # SQLite database setup
        self.db_file = os.path.join(self.settings_dir, "emails.db")
        print(f"Database file path: {self.db_file}")  # Debug statement
        self.init_database()

        # App Icon
        if os.path.exists(os.path.join(self.script_dir, "icon.jpg")):
            self.setWindowIcon(QIcon(os.path.join(self.script_dir, "icon.jpg")))

        # Top Menu Bar
        menubar = QMenuBar(self)
        self.setMenuBar(menubar)

        # File Menu
        file_menu = QMenu("File", self)
        menubar.addMenu(file_menu)
        account_settings_action = file_menu.addAction("Account Settings")
        account_settings_action.triggered.connect(self.open_account_settings)

        # Edit Menu
        edit_menu = QMenu("Edit", self)
        menubar.addMenu(edit_menu)

        # View Menu
        view_menu = QMenu("View", self)
        menubar.addMenu(view_menu)

        # Tools Menu
        tools_menu = QMenu("Tools", self)
        menubar.addMenu(tools_menu)

        # Help Menu
        help_menu = QMenu("Help", self)
        menubar.addMenu(help_menu)
        about_action = help_menu.addAction("About OpenOutlook")
        about_action.triggered.connect(self.show_about_dialog)

        # Toolbar
        self.toolbar = QToolBar("Main Toolbar")
        self.addToolBar(self.toolbar)
        self.toolbar.setMovable(False)
        self.toolbar.setStyleSheet("QToolBar { background-color: #ECE9D8; border-bottom: 1px solid #C0C0C0; }")

        self.new_button = QPushButton("New")
        self.reply_button = QPushButton("Reply")
        self.reply_all_button = QPushButton("Reply to All")
        self.forward_button = QPushButton("Forward")
        
        self.new_button.clicked.connect(self.open_compose_window)
        self.reply_button.clicked.connect(lambda: self.open_compose_window(mode='reply'))
        self.reply_all_button.clicked.connect(lambda: self.open_compose_window(mode='reply_all'))
        self.forward_button.clicked.connect(lambda: self.open_compose_window(mode='forward'))

        self.toolbar.addWidget(self.new_button)
        self.toolbar.addWidget(self.reply_button)
        self.toolbar.addWidget(self.reply_all_button)
        self.toolbar.addWidget(self.forward_button)

        self.toolbar.addSeparator()

        self.print_button = QPushButton("Print")
        self.print_button.clicked.connect(self.print_email)
        self.toolbar.addWidget(self.print_button)

        self.delete_button = QPushButton("Delete")
        self.delete_button.clicked.connect(self.delete_email)
        self.toolbar.addWidget(self.delete_button)

        self.send_receive_button = QPushButton("Send/Receive")
        self.send_receive_button.clicked.connect(self.send_receive_all)
        self.toolbar.addWidget(self.send_receive_button)

        self.toolbar.addSeparator()

        self.find_button = QPushButton("Find")
        self.find_button.clicked.connect(self.open_find_dialog)
        self.toolbar.addWidget(self.find_button)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search Inbox")
        self.search_input.setMaximumWidth(200)
        self.search_input.returnPressed.connect(self.quick_search)
        self.toolbar.addWidget(self.search_input)

        # Main widget and layout
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        self.main_layout = QVBoxLayout(self.main_widget)
        self.main_layout.setContentsMargins(2, 2, 2, 2)
        self.main_layout.setSpacing(0)

        # Splitter for resizable panes
        self.splitter = QSplitter(Qt.Horizontal)
        self.splitter.setStyleSheet("QSplitter::handle { background-color: #808080; width: 1px; }")

        # Left pane: Folders and Buttons
        self.left_pane = QWidget()
        self.left_layout = QVBoxLayout(self.left_pane)
        self.left_pane.setStyleSheet("background-color: #DDE5ED; border: 1px solid #808080;")
        self.left_layout.setContentsMargins(2, 2, 2, 2)
        self.left_layout.setSpacing(2)

        # All Mail Folders
        self.all_folders = QTreeWidget()
        self.all_folders.setHeaderHidden(True)
        self.all_folders.setStyleSheet("""
            QTreeWidget {
                font: 8pt Tahoma;
                background-color: #DDE5ED;
                border: none;
            }
            QTreeWidget::item {
                padding: 0px;
                margin: 0px;
            }
            QTreeWidget::item:hover {
                background-color: #E6E6E6;
            }
            QTreeWidget::item:selected {
                background-color: #C0C0C0;
            }
        """)
        all_header = QTreeWidgetItem(self.all_folders, ["All Mail Folders"])
        all_header.setFlags(Qt.ItemIsEnabled)
        all_header.setFont(0, self.all_folders.font())
        all_header.setText(0, "All Mail Folders")
        folders = ["Drafts", "Inbox", "Junk E-mail", "Outbox", "Sent Items", "Deleted Items", "Sync Issues"]
        for folder in folders:
            item = QTreeWidgetItem(all_header, [folder])
            if folder == "Inbox":
                self.all_folders.setCurrentItem(item)  # Select "Inbox" by default
        self.all_folders.expandAll()
        self.left_layout.addWidget(self.all_folders)

        # Spacer to push buttons to the bottom
        self.left_layout.addStretch()

        # Button styling
        self.button_style = """
            QPushButton {
                font: 12pt Tahoma;
                font-weight: bold;
                background-color: #DDE5ED;
                border: none;
                text-align: left;
                padding: 5px 5px 5px 40px;  /* Increased left padding for larger icon */
                height: 30px;
            }
            QPushButton:hover {
                background-color: #C0C0C0;
            }
            QPushButton:pressed {
                background-color: #A0A0A0;
            }
        """
        self.selected_button_style = """
            QPushButton {
                font: 12pt Tahoma;
                font-weight: bold;
                background-color: #B0C4DE;
                border: none;
                text-align: left;
                padding: 5px 5px 5px 40px;  /* Increased left padding for larger icon */
                height: 30px;
            }
            QPushButton:hover {
                background-color: #A0B4CE;
            }
            QPushButton:pressed {
                background-color: #90A4BE;
            }
        """

        # Track the selected button
        self.selected_button = None

        # Mail Button (selected by default)
        self.mail_button = QPushButton("Mail")
        self.mail_button.setStyleSheet(self.selected_button_style)
        mail_icon_path = os.path.join(self.script_dir, "mail_icon.png")
        print(f"Mail icon exists: {os.path.exists(mail_icon_path)}")  # Debug statement
        self.mail_button.setIcon(QIcon(mail_icon_path))
        self.mail_button.setIconSize(QSize(32, 32))
        self.mail_button.clicked.connect(self.show_email_view)
        self.selected_button = self.mail_button
        self.left_layout.addWidget(self.mail_button)

        # Delineator
        delineator1 = QFrame()
        delineator1.setFrameShape(QFrame.HLine)
        delineator1.setFrameShadow(QFrame.Sunken)
        delineator1.setStyleSheet("background-color: #D0D0D0;")
        self.left_layout.addWidget(delineator1)

        # Calendar Button
        self.calendar_button = QPushButton("Calendar")
        self.calendar_button.setStyleSheet(self.button_style)
        calendar_icon_path = os.path.join(self.script_dir, "calendar_icon.png")
        print(f"Calendar icon exists: {os.path.exists(calendar_icon_path)}")  # Debug statement
        self.calendar_button.setIcon(QIcon(calendar_icon_path))
        self.calendar_button.setIconSize(QSize(32, 32))
        self.calendar_button.clicked.connect(self.show_calendar_view)
        self.left_layout.addWidget(self.calendar_button)

        # Delineator
        delineator2 = QFrame()
        delineator2.setFrameShape(QFrame.HLine)
        delineator2.setFrameShadow(QFrame.Sunken)
        delineator2.setStyleSheet("background-color: #D0D0D0;")
        self.left_layout.addWidget(delineator2)

        # Contacts Button
        self.contacts_button = QPushButton("Contacts")
        self.contacts_button.setStyleSheet(self.button_style)
        contacts_icon_path = os.path.join(self.script_dir, "contacts_icon.png")
        print(f"Contacts icon exists: {os.path.exists(contacts_icon_path)}")  # Debug statement
        self.contacts_button.setIcon(QIcon(contacts_icon_path))
        self.contacts_button.setIconSize(QSize(32, 32))
        self.contacts_button.clicked.connect(lambda: self.set_selected_button(self.contacts_button))
        self.left_layout.addWidget(self.contacts_button)

        # Delineator
        delineator3 = QFrame()
        delineator3.setFrameShape(QFrame.HLine)
        delineator3.setFrameShadow(QFrame.Sunken)
        delineator3.setStyleSheet("background-color: #D0D0D0;")
        self.left_layout.addWidget(delineator3)

        # Tasks Button
        self.tasks_button = QPushButton("Tasks")
        self.tasks_button.setStyleSheet(self.button_style)
        tasks_icon_path = os.path.join(self.script_dir, "tasks_icon.png")
        print(f"Tasks icon exists: {os.path.exists(tasks_icon_path)}")  # Debug statement
        self.tasks_button.setIcon(QIcon(tasks_icon_path))
        self.tasks_button.setIconSize(QSize(32, 32))
        self.tasks_button.clicked.connect(lambda: self.set_selected_button(self.tasks_button))
        self.left_layout.addWidget(self.tasks_button)

        self.splitter.addWidget(self.left_pane)

        # Create email view (toolbar, middle, and right panes)
        self.email_view_widget = QWidget()
        self.email_view_layout = QVBoxLayout(self.email_view_widget)
        self.email_view_layout.setContentsMargins(0,0,0,0)
        self.email_view_layout.setSpacing(0)

        # Create email view (middle and right panes)
        self.email_splitter = QSplitter(Qt.Horizontal)
        self.email_splitter.setStyleSheet("QSplitter::handle { background-color: #808080; width: 1px; }")

        # Middle pane: Email List
        self.email_middle_pane = QWidget()
        self.email_middle_layout = QVBoxLayout(self.email_middle_pane)
        self.email_middle_pane.setStyleSheet("background-color: #DDE5ED; border: 1px solid #808080;")
        self.email_middle_layout.setContentsMargins(2, 2, 2, 2)
        self.email_middle_layout.setSpacing(0)

        self.email_list = QTreeWidget()
        self.email_list.setColumnCount(3)
        self.email_list.setHeaderLabels(["From", "Subject", "Received"])
        self.email_list.setStyleSheet("""
            QTreeWidget {
                font: 8pt Tahoma;
                background-color: white;
                border: none;
                alternate-background-color: #F5F5F5;
            }
            QTreeWidget::item {
                padding: 1px;
                border-bottom: 1px solid #D0D0D0;
            }
            QTreeWidget::item:hover {
                background-color: #F0F0F0;
            }
            QTreeWidget::item:selected {
                background-color: #316AC5;
                color: white;
            }
            QTreeWidget::branch {
                background-color: #E6E6E6;
                border-bottom: 1px solid #D0D0D0;
            }
        """)
        self.email_list.header().setStyleSheet("""
            QHeaderView::section {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #D0D0D0, stop:1 #C0C0C0);
                font: 8pt Tahoma;
                font-weight: bold;
                border: 1px solid #808080;
                padding: 2px;
            }
        """)
        self.email_list.setColumnWidth(0, 150)
        self.email_list.setColumnWidth(1, 300)
        self.email_list.setColumnWidth(2, 100)
        self.email_list.header().setStretchLastSection(True)
        self.email_list.setAlternatingRowColors(True)
        self.email_list.setSortingEnabled(True)
        self.email_list.sortByColumn(2, Qt.DescendingOrder)

        today_header = QTreeWidgetItem(self.email_list, ["Today"])
        today_header.setFont(0, self.email_list.font())
        today_header.setFlags(Qt.ItemIsEnabled)

        yesterday_header = QTreeWidgetItem(self.email_list, ["Yesterday"])
        yesterday_header.setFont(0, self.email_list.font())
        today_header.setFlags(Qt.ItemIsEnabled)

        self.email_list.expandAll()
        self.email_middle_layout.addWidget(self.email_list)

        self.email_splitter.addWidget(self.email_middle_pane)

        # Right pane: Email Preview
        self.email_right_pane = QWidget()
        self.email_right_layout = QVBoxLayout(self.email_right_pane)
        self.email_right_pane.setStyleSheet("background-color: white; border: 1px solid #808080;")
        self.email_right_layout.setContentsMargins(2, 2, 2, 2)
        self.email_right_layout.setSpacing(0)

        self.preview_header = QLabel("")
        self.preview_header.setStyleSheet("font: 10pt Tahoma; font-weight: bold; background-color: #E6E6E6; padding: 5px;")
        self.email_right_layout.addWidget(self.preview_header)

        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setStyleSheet("font: 9pt Tahoma; border: none; padding: 5px;")
        self.email_right_layout.addWidget(self.preview_text)

        self.email_splitter.addWidget(self.email_right_pane)
        self.email_view_layout.addWidget(self.email_splitter)

        # Create calendar view (initially hidden)
        self.calendar_pane = QWidget()
        self.calendar_layout = QVBoxLayout(self.calendar_pane)
        self.calendar_pane.setStyleSheet("background-color: #DDE5ED; border: 1px solid #808080;")
        self.calendar_layout.setContentsMargins(2, 2, 2, 2)

        self.calendar = QCalendarWidget()
        self.calendar.setFirstDayOfWeek(Qt.Sunday)
        self.calendar.setGridVisible(True)
        self.calendar.setHorizontalHeaderFormat(QCalendarWidget.ShortDayNames)
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.calendar.setStyleSheet("""
            QCalendarWidget {
                font: 8pt Tahoma;
                background-color: white;
            }
            QCalendarWidget QAbstractItemView {
                background-color: white;
                selection-background-color: #316AC5;
                font: 8pt Tahoma;
            }
            QCalendarWidget QAbstractItemView:enabled {
                color: black;
            }
            QCalendarWidget QAbstractItemView:disabled {
                color: #808080;
            }
            QCalendarWidget QWidget#qt_calendar_navigationbar {
                background-color: #C0C0C0;
            }
            QCalendarWidget QTableView {
                border: 1px solid #D0D0D0;
                gridline-color: #D0D0D0;
            }
            QCalendarWidget QTableView::item {
                border: 1px solid #D0D0D0;
                padding: 2px;
            }
            QCalendarWidget QTableView::item:selected {
                background-color: #316AC5;
                color: white;
            }
            QCalendarWidget QTableView::item:today {
                background-color: #D0E0FF;
            }
            QCalendarWidget QToolButton {
                font: 8pt Tahoma;
                background-color: #C0C0C0;
                border: none;
            }
            QCalendarWidget QToolButton:hover {
                background-color: #A0A0A0;
            }
            QCalendarWidget QToolButton:pressed {
                background-color: #808080;
            }
            QCalendarWidget QHeaderView::section {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #D0D0D0, stop:1 #C0C0C0);
                font: 8pt Tahoma;
                font-weight: bold;
                border: 1px solid #808080;
                padding: 2px;
                text-align: center;
            }
        """)
        self.calendar_layout.addWidget(self.calendar)

        # Initially add the email view widget to the splitter
        self.splitter.addWidget(self.email_view_widget)

        # Add splitter to main layout
        self.main_layout.addWidget(self.splitter)

        # Notification bar at the bottom
        email_address = "No account loaded"
        if self.settings["accounts"]:
            email_address = self.settings["accounts"][0]["email"].upper()
        self.notification_bar = QLabel(f"INBOX LOADED FOR {email_address}")
        self.notification_bar.setStyleSheet("""
            QLabel {
                background-color: #003087;
                color: white;
                font: 8pt Tahoma;
                font-weight: bold;
                padding: 1px;
            }
        """)
        self.notification_bar.setMaximumHeight(20)
        self.main_layout.addWidget(self.notification_bar)

        # Set splitter sizes: left pane 15%, middle 60%, right 25%
        total_width = 800
        left_width = int(total_width * 0.15)  # 120px
        middle_width = int(total_width * 0.60)  # 480px
        right_width = int(total_width * 0.25)  # 200px
        self.splitter.setSizes([left_width, middle_width, right_width])

        # Initially show email view
        self.current_view = "email"
        self.show_email_view()

        # Load emails from database
        self.load_emails_from_db()

    def init_database(self):
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()

        # Check for schema updates (missing account_email column)
        cursor.execute("PRAGMA table_info(emails)")
        columns = [col[1] for col in cursor.fetchall()]
        if columns and "account_email" not in columns:
            print("Detected outdated schema (missing account_email). Recreating emails table.")
            cursor.execute("DROP TABLE emails")
            columns = []

        if columns and "cc_addr" not in columns:
            print("Adding 'cc_addr' column to emails table.")
            cursor.execute("ALTER TABLE emails ADD COLUMN cc_addr TEXT")

        # Create emails table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                account_email TEXT NOT NULL,
                uid TEXT NOT NULL,
                folder TEXT NOT NULL,
                subject TEXT,
                from_addr TEXT,
                to_addr TEXT,
                cc_addr TEXT,
                date TEXT,
                body BLOB,
                flags TEXT,
                UNIQUE(account_email, uid, folder)
            )
        """)
        conn.commit()
        conn.close()

    def sync_emails(self, account):
        try:
            email_address = account["email"]
            auth_method = account.get("auth_method", "oauth") # Default to oauth for old accounts

            imap = imaplib.IMAP4_SSL(account["imap_server"], account["imap_port"])

            if auth_method == "oauth":
                # Retrieve refresh token from keyring
                refresh_token = keyring.get_password("OpenOutlook_RefreshToken", email_address)
                if not refresh_token:
                    print(f"Refresh token not found for {email_address}")
                    return

                config_data = GOOGLE_CLIENT_CONFIG["installed"]
                client_id = config_data["client_id"]
                client_secret = config_data["client_secret"]

                creds = Credentials(
                    None,
                    refresh_token=refresh_token,
                    token_uri="https://oauth2.googleapis.com/token",
                    client_id=client_id,
                    client_secret=client_secret,
                    scopes=["https://mail.google.com/"],
                )

                # Refresh the credentials
                creds.refresh(Request())

                # Connect to IMAP server
                auth_string = f"user={email_address}\1auth=Bearer {creds.token}\1\1"
                imap.authenticate("XOAUTH2", lambda x: auth_string.encode())
            
            elif auth_method == "password":
                app_password = keyring.get_password("OpenOutlook_AppPassword", email_address)
                if not app_password:
                    print(f"App password not found for {email_address}")
                    return
                imap.login(email_address, app_password)

            imap.select("INBOX")
            _, message_numbers = imap.search(None, "ALL")
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()

            def decode_header_str(header):
                if header is None:
                    return ""
                decoded_parts = decode_header(header)
                header_str = ""
                for part, encoding in decoded_parts:
                    if isinstance(part, bytes):
                        header_str += part.decode(encoding or 'utf-8', errors='ignore')
                    else:
                        header_str += part
                return header_str

            for num in message_numbers[0].split():
                _, msg_data = imap.fetch(num, "(RFC822)")
                if msg_data[0] is None:
                    continue
                
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                # Decode headers
                subject = decode_header_str(msg["Subject"])
                from_addr = decode_header_str(msg["From"])
                to_addr = decode_header_str(msg["To"])
                cc_addr = decode_header_str(msg["Cc"])
                date = decode_header_str(msg["Date"])

                # Extract body
                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            payload = part.get_payload(decode=True)
                            charset = part.get_content_charset() or 'utf-8'
                            body = payload.decode(charset, errors='ignore')
                            break # Found plain text part, stop
                else: # Not multipart
                    payload = msg.get_payload(decode=True)
                    charset = msg.get_content_charset() or 'utf-8'
                    body = payload.decode(charset, errors='ignore')

                if not body:
                    body = "Email content could not be displayed."

                # Encrypt the email body
                encrypted_body = self.cipher.encrypt(body.encode('utf-8'))
                
                # Store in database, ignoring if it already exists
                cursor.execute("""
                    INSERT OR IGNORE INTO emails (account_email, uid, folder, subject, from_addr, to_addr, cc_addr, date, body, flags)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (email_address, num.decode(), "INBOX", subject, from_addr, to_addr, cc_addr, date, encrypted_body, "UNREAD"))
            
            conn.commit()
            conn.close()
            imap.logout()
            # Refresh the email list
            self.load_emails_from_db()
        except Exception as e:
            print(f"Failed to sync emails: {e}") # This will now print more specific errors from parsing

    def load_emails_from_db(self):
        self.email_list.clear()
        if not self.settings["accounts"]:
            return

        # For PoC, load emails for the first account
        account_email = self.settings["accounts"][0]["email"]

        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute("SELECT from_addr, to_addr, cc_addr, subject, date, body, uid FROM emails WHERE folder='INBOX' AND account_email=?", (account_email,))
        emails = cursor.fetchall()
        today_header = QTreeWidgetItem(self.email_list, ["Today"])
        today_header.setFlags(Qt.ItemIsEnabled)
        for email_record in emails:
            from_addr, to_addr, cc_addr, subject, date_str, encrypted_body, uid = email_record
            # Decrypt the email body
            try:
                body = self.cipher.decrypt(encrypted_body).decode('utf-8')
            except Exception as e:
                print(f"Failed to decrypt email UID {uid}: {e}")
                body = "Failed to decrypt email"
            
            item = SortableTreeWidgetItem(today_header, [from_addr, subject, date_str])
            
            # Parse and store the datetime object for sorting
            try:
                dt_obj = parsedate_to_datetime(date_str)
                item.setData(0, Qt.UserRole + 1, dt_obj)
            except Exception as e:
                print(f"Could not parse date '{date_str}': {e}")
                item.setData(0, Qt.UserRole + 1, None)
            
            email_data = {
                'from': from_addr,
                'to': to_addr,
                'cc': cc_addr,
                'subject': subject,
                'body': body,
                'date': date_str,
                'uid': uid
            }
            item.setData(0, Qt.UserRole, email_data)
        self.email_list.expandAll()
        conn.close()
        # Connect the email list to update the preview pane
        self.email_list.itemClicked.connect(self.update_preview)

    def update_preview(self, item, column):
        if not item or not item.parent(): # Ignore header clicks
            return
        email_data = item.data(0, Qt.UserRole)
        if not email_data: return
        body = email_data.get('body', '')
        from_addr = item.text(0)
        subject = item.text(1)
        self.preview_header.setText(f"<b>{subject}</b><br>From: {from_addr}")
        self.preview_text.setText(body)

    def load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "r") as f:
                    self.settings = json.load(f)
                    if "accounts" not in self.settings:
                        self.settings["accounts"] = []
            except (json.JSONDecodeError, IOError) as e:
                print(f"Error loading settings: {e}. Using default settings.")
                self.settings = {"accounts": []}
        else:
            self.settings = {"accounts": []}

    def save_settings(self):
        try:
            with open(self.settings_file, "w") as f:
                settings_to_save = {"accounts": []}
                for acc in self.settings["accounts"]:
                    acc_copy = acc.copy()
                    acc_copy.pop("credentials", None)
                    settings_to_save["accounts"].append(acc_copy)
                json.dump(settings_to_save, f, indent=4)
        except IOError as e:
            print(f"Error saving settings: {e}")

    def open_account_settings(self):
        dialog = AccountSettingsDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            new_settings = dialog.get_settings()
            email_addr = new_settings["email"]

            self.settings["accounts"] = [acc for acc in self.settings.get("accounts", []) if acc["email"] != email_addr]

            account_info = {
                "email": email_addr,
                "auth_method": new_settings["auth_method"],
                "imap_server": new_settings["imap_server"],
                "imap_port": new_settings["imap_port"],
                "smtp_server": new_settings["smtp_server"],
                "smtp_port": new_settings["smtp_port"],
            }

            if new_settings["auth_method"] == "oauth":
                creds = new_settings["credentials"]
                if creds and creds.refresh_token:
                    keyring.set_password("OpenOutlook_RefreshToken", email_addr, creds.refresh_token)
            else:
                password = new_settings["password"]
                if password:
                    keyring.set_password("OpenOutlook_AppPassword", email_addr, password)

            self.settings["accounts"].append(account_info)
            self.save_settings()
            self.sync_emails(account_info)

    def open_compose_window(self, mode='new'):
        if not self.settings["accounts"]:
            QMessageBox.warning(self, "No Account", "Please configure an email account first.")
            return
        
        account = self.settings["accounts"][0]
        
        original_msg = None
        if mode != 'new':
            selected_item = self.email_list.currentItem()
            if not selected_item or not selected_item.parent():
                QMessageBox.warning(self, "No Email Selected", "Please select an email to " + mode + ".")
                return
            original_msg = selected_item.data(0, Qt.UserRole)

        compose_dialog = ComposeWindow(self, account=account, mode=mode, original_msg=original_msg)
        compose_dialog.exec_()

    def send_email(self, account, to_addrs, cc_addrs, subject, body):
        email_address = account["email"]
        
        msg = MIMEMultipart()
        msg['From'] = email_address
        msg['To'] = to_addrs
        if cc_addrs:
            msg['Cc'] = cc_addrs
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        recipients = to_addrs.split(',') + cc_addrs.split(',')
        recipients = [r.strip() for r in recipients if r.strip()]

        try:
            smtp = smtplib.SMTP_SSL(account["smtp_server"], account["smtp_port"])
            auth_method = account.get("auth_method", "oauth")

            if auth_method == "oauth":
                refresh_token = keyring.get_password("OpenOutlook_RefreshToken", email_address)
                if not refresh_token:
                    QMessageBox.critical(self, "Auth Error", "OAuth refresh token not found.")
                    return

                config_data = GOOGLE_CLIENT_CONFIG["installed"]
                creds = Credentials(None, refresh_token=refresh_token, token_uri="https://oauth2.googleapis.com/token", client_id=config_data["client_id"], client_secret=config_data["client_secret"], scopes=["https://mail.google.com/"])
                creds.refresh(Request())
                
                auth_string = f"user={email_address}\1auth=Bearer {creds.token}\1\1"
                smtp.auth('XOAUTH2', lambda: auth_string.encode('utf-8'))
            elif auth_method == "password":
                app_password = keyring.get_password("OpenOutlook_AppPassword", email_address)
                if not app_password:
                    QMessageBox.critical(self, "Auth Error", "App password not found.")
                    return
                smtp.login(email_address, app_password)

            smtp.sendmail(email_address, recipients, msg.as_string())
            smtp.quit()
            QMessageBox.information(self, "Success", "Email sent successfully!")
        except Exception as e:
            print(f"Failed to send email: {e}")
            QMessageBox.critical(self, "Send Error", f"Failed to send email: {e}")

    def show_about_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("About OpenOutlook")
        layout = QVBoxLayout()
        
        if os.path.exists(os.path.join(self.script_dir, "About.jpg")):
            label = QLabel()
            pixmap = QPixmap(os.path.join(self.script_dir, "About.jpg"))
            label.setPixmap(pixmap)
            layout.addWidget(label)
        else:
            layout.addWidget(QLabel("OpenOutlook v1.0\nA Python-based Outlook Clone."))
            
        dialog.setLayout(layout)
        dialog.exec_()

    def print_email(self):
        QMessageBox.information(self, "Print", "Printing functionality is not yet implemented.")

    def delete_email(self):
        item = self.email_list.currentItem()
        if not item or not item.parent():
            return

        email_data = item.data(0, Qt.UserRole)
        if not email_data: return
        uid = email_data.get('uid')
        if not uid: return

        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        # Assuming first account for now
        if self.settings["accounts"]:
            account_email = self.settings["accounts"][0]["email"]
            cursor.execute("DELETE FROM emails WHERE uid=? AND account_email=?", (uid, account_email))
            conn.commit()
        conn.close()

        item.parent().removeChild(item)
        self.preview_header.setText("")
        self.preview_text.clear()

    def send_receive_all(self):
        if self.settings.get("accounts"):
            for account in self.settings["accounts"]:
                self.sync_emails(account)
            QMessageBox.information(self, "Sync", "Send/Receive completed.")
        else:
            QMessageBox.warning(self, "Sync", "No accounts configured.")

    def open_find_dialog(self):
        QMessageBox.information(self, "Find", "Advanced Find dialog is coming soon.")

    def quick_search(self):
        text = self.search_input.text().lower()
        root = self.email_list.invisibleRootItem()
        for i in range(root.childCount()):
            group_item = root.child(i)
            for j in range(group_item.childCount()):
                item = group_item.child(j)
                subject = item.text(1).lower()
                sender = item.text(0).lower()
                item.setHidden(text not in subject and text not in sender)

    def load_or_create_encryption_key(self):
        key = keyring.get_password("OpenOutlook_EncryptionKey", "default_user")
        if key:
            self.encryption_key = key.encode()
        else:
            self.encryption_key = Fernet.generate_key()
            keyring.set_password("OpenOutlook_EncryptionKey", "default_user", self.encryption_key.decode())
        self.cipher = Fernet(self.encryption_key)

    def set_selected_button(self, button):
        if self.selected_button:
            self.selected_button.setStyleSheet(self.button_style)
        self.selected_button = button
        self.selected_button.setStyleSheet(self.selected_button_style)

    def show_email_view(self):
        if self.current_view == "calendar":
            self.splitter.replaceWidget(1, self.email_view_widget)
            self.calendar_pane.setParent(None)
        self.email_view_widget.show()
        self.current_view = "email"
        self.set_selected_button(self.mail_button)

    def show_calendar_view(self):
        if self.current_view == "email":
            self.splitter.replaceWidget(1, self.calendar_pane)
            self.email_view_widget.setParent(None)
        self.calendar_pane.show()
        self.current_view = "calendar"
        self.set_selected_button(self.calendar_button)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Tahoma"))
    window = OutlookLookalike()
    window.show()
    sys.exit(app.exec_())