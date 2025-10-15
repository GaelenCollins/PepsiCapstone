import sys
import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.patheffects
from pandas.core.frame import com
import seaborn as sns
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, 
                           QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, 
                           QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
                           QScrollArea, QFrame, QTextEdit, QMessageBox, QHeaderView,
                           QComboBox, QDateEdit, QCheckBox, QGroupBox, QSpinBox,
                           QProgressBar, QSplitter)
from PyQt5.QtCore import Qt, pyqtSignal, QDate, QThread, pyqtSlot, QTimer
from PyQt5.QtGui import QFont, QIcon, QColor, QPalette, QPixmap
import warnings
warnings.filterwarnings('ignore')

class ModernTheme:                                      #Dark theme for UI
    def __init__(self):
        # Always use dark theme
        self.current_theme = "dark"
        self.themes = {
            "dark": {
                "bg": "#1e1e1e",
                "card_bg": "#2d2d2d", 
                "header_bg": "#2d2d2d",
                "header_fg": "#ffffff",
                "fg": "#ffffff",
                "text_secondary": "#b0b0b0",
                "accent": "#007acc",
                "success": "#28a745",
                "danger": "#dc3545", 
                "warning": "#ffc107",
                "info": "#17a2b8",
                "purple": "#6f42c1",
                "hover": "#404040"
            }
        }
    
    def get_color(self, color_name):
        return self.themes[self.current_theme].get(color_name, "#000000")

class MetricCard(QFrame):
    def __init__(self, title, value, color, theme):
        super().__init__()
        self.theme = theme
        self.color = color
        self.setup_ui(title, value)
        
    def setup_ui(self, title, value):
        self.setFrameStyle(QFrame.Box)
        self.setLineWidth(0)
        self.setContentsMargins(0, 0, 0, 0)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)
        
        # Value label (bigger)
        value_label = QLabel(value)
        value_label.setAlignment(Qt.AlignCenter)
        value_label.setFont(QFont("Segoe UI", 32, QFont.Bold))
        value_label.setStyleSheet("color: white; background: transparent;")
        
        # Title label (bigger)
        title_label = QLabel(title)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        title_label.setStyleSheet("color: white; background: transparent;")
        
        layout.addWidget(value_label)
        layout.addWidget(title_label)
        self.setLayout(layout)
        
        # Set background color
        self.setStyleSheet(f"""
            MetricCard {{
                background-color: {self.color};
                border-radius: 8px;
                border: none;
            }}
        """)

class ModernCard(QFrame):
    def __init__(self, title=None, theme=None):
        super().__init__()
        self.theme = theme
        self.setup_ui(title)
        
    def setup_ui(self, title):
        self.setFrameStyle(QFrame.Box)
        self.setLineWidth(1)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        if title:
            title_label = QLabel(title)
            title_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
            title_label.setContentsMargins(20, 15, 20, 10)
            layout.addWidget(title_label)
        
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout()
        self.content_layout.setContentsMargins(20, 10, 20, 20)
        self.content_widget.setLayout(self.content_layout)
        layout.addWidget(self.content_widget)
        
        self.setLayout(layout)
        self.update_theme()
        
    def update_theme(self):
        if self.theme:
            self.setStyleSheet(f"""
                ModernCard {{
                    background-color: {self.theme.get_color('card_bg')};
                    border: 1px solid {self.theme.get_color('hover')};
                    border-radius: 8px;
                }}
                QLabel {{
                    color: {self.theme.get_color('fg')};
                    background: transparent;
                }}
            """)

class RejectedUnitsAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.theme = ModernTheme()
        self.current_data = None
        self.filtered_data = None
        self.analysis_results = None
        self.current_file = None
        
        self.setup_ui()
        self.apply_theme()
        
    def setup_ui(self):
        self.setWindowTitle("Gaelen Collins - E80 Rejected Units Analyzer")
        self.setGeometry(100, 100, 1600, 1000)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Header
        self.create_header(main_layout)
        
        # Global Filters Section (persistent across all tabs)
        self.create_global_filters(main_layout)
        
        # Tab widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabPosition(QTabWidget.North)
        main_layout.addWidget(self.tab_widget)
        
        # Create all the analysis tabs
        self.create_upload_tab()
        self.create_dashboard_tab()
        self.create_trends_tab()
        self.create_production_analysis_tab()
        self.create_production_lines_tab()
        self.create_advanced_tracking_tab()
        self.create_dimensional_rejects_tab()
        self.create_tag_tracking_rejects_tab()
        self.create_time_analysis_tab()
        self.create_sku_analysis_tab()
        self.create_rejection_rate_tab()
        
        central_widget.setLayout(main_layout)
        
    def create_global_filters(self, parent_layout): #Global filters that affect all tabs
        # Create a custom card with minimize button
        self.filters_card = QFrame()
        self.filters_card.setFrameStyle(QFrame.Box)
        self.filters_card.setLineWidth(1)
        self.filters_card.setMaximumHeight(300)
        
        # Main layout for the card
        card_layout = QVBoxLayout()
        card_layout.setContentsMargins(0, 0, 0, 0)
        card_layout.setSpacing(0)
        
        # Title bar with minimize button
        title_bar = QFrame()
        title_bar.setFixedHeight(40)
        title_bar_layout = QHBoxLayout()
        title_bar_layout.setContentsMargins(20, 10, 20, 10)
        
        title_label = QLabel("🔍 Global Filters")
        title_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        
        # Minimize/expand button
        self.filters_minimize_btn = QPushButton("−")
        self.filters_minimize_btn.setFixedSize(25, 25)
        self.filters_minimize_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme.get_color('accent')};
                color: white;
                border: none;
                border-radius: 12px;
                font-weight: bold;
                font-size: 14px;
            }}
            QPushButton:hover {{
                background-color: {self.theme.get_color('accent_hover')};
            }}
        """)
        self.filters_minimize_btn.clicked.connect(self.toggle_filters_minimize)
        
        title_bar_layout.addWidget(title_label)
        title_bar_layout.addStretch()
        title_bar_layout.addWidget(self.filters_minimize_btn)
        title_bar.setLayout(title_bar_layout)
        
        # Content area
        self.filters_content_widget = QWidget()
        self.filters_content_layout = QHBoxLayout()
        self.filters_content_layout.setContentsMargins(20, 10, 20, 20)
        self.filters_content_widget.setLayout(self.filters_content_layout)
        
        # Add to card layout
        card_layout.addWidget(title_bar)
        card_layout.addWidget(self.filters_content_widget)
        self.filters_card.setLayout(card_layout)
        
        # Apply theme to the card
        self.filters_card.setStyleSheet(f"""
            QFrame {{
                background-color: {self.theme.get_color('card_bg')};
                border: 1px solid {self.theme.get_color('accent')};
                border-radius: 5px;
                color: {self.theme.get_color('fg')};
            }}
        """)
        
        # Initialize as expanded
        self.filters_minimized = False
        
        filters_layout = self.filters_content_layout
        
        # Period filter
        period_group = QGroupBox("Period Selection")
        period_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 1px solid {self.theme.get_color('accent')};
                border-radius: 3px;
                margin-top: 1ex;
                color: {self.theme.get_color('fg')};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }}
        """)
        period_layout = QVBoxLayout()
        
        # Period filter buttons
        period_buttons_layout = QHBoxLayout()
        self.period_select_all_btn = QPushButton("Select All")
        self.period_select_all_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme.get_color('accent')};
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 3px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {self.theme.get_color('accent_hover')};
            }}
        """)
        self.period_select_all_btn.clicked.connect(self.select_all_periods)
        
        self.period_reset_btn = QPushButton("Reset")
        self.period_reset_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #e74c3c;
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 3px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #c0392b;
            }}
        """)
        self.period_reset_btn.clicked.connect(self.reset_periods)
        
        period_buttons_layout.addWidget(self.period_select_all_btn)
        period_buttons_layout.addWidget(self.period_reset_btn)
        period_buttons_layout.addStretch()
        period_layout.addLayout(period_buttons_layout)
        
        self.global_period_checkboxes = {}
        self.global_period_scroll = QScrollArea()
        self.global_period_scroll.setWidgetResizable(True)
        self.global_period_scroll.setMaximumHeight(120)
        global_period_widget = QWidget()
        self.global_period_widget_layout = QVBoxLayout()
        global_period_widget.setLayout(self.global_period_widget_layout)
        self.global_period_scroll.setWidget(global_period_widget)
        period_layout.addWidget(self.global_period_scroll)
        period_group.setLayout(period_layout)
        
        # Line filter
        line_group = QGroupBox("Production Lines")
        line_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 1px solid {self.theme.get_color('accent')};
                border-radius: 3px;
                margin-top: 1ex;
                color: {self.theme.get_color('fg')};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }}
        """)
        line_layout = QVBoxLayout()
        
        # Line filter buttons
        line_buttons_layout = QHBoxLayout()
        self.line_select_all_btn = QPushButton("Select All")
        self.line_select_all_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme.get_color('accent')};
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 3px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {self.theme.get_color('accent_hover')};
            }}
        """)
        self.line_select_all_btn.clicked.connect(self.select_all_lines)
        
        self.line_reset_btn = QPushButton("Reset")
        self.line_reset_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #e74c3c;
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 3px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #c0392b;
            }}
        """)
        self.line_reset_btn.clicked.connect(self.reset_lines)
        
        line_buttons_layout.addWidget(self.line_select_all_btn)
        line_buttons_layout.addWidget(self.line_reset_btn)
        line_buttons_layout.addStretch()
        line_layout.addLayout(line_buttons_layout)
        
        self.global_line_checkboxes = {}
        self.global_line_scroll = QScrollArea()
        self.global_line_scroll.setWidgetResizable(True)
        self.global_line_scroll.setMaximumHeight(120)
        global_line_widget = QWidget()
        self.global_line_widget_layout = QVBoxLayout()
        global_line_widget.setLayout(self.global_line_widget_layout)
        self.global_line_scroll.setWidget(global_line_widget)
        line_layout.addWidget(self.global_line_scroll)
        line_group.setLayout(line_layout)
        
        # Product filter
        sku_group = QGroupBox("Products")
        sku_group.setStyleSheet(f"""
            QGroupBox {{
                font-weight: bold;
                border: 1px solid {self.theme.get_color('accent')};
                border-radius: 3px;
                margin-top: 1ex;
                color: {self.theme.get_color('fg')};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }}
        """)
        sku_layout = QVBoxLayout()
        
        # SKU filter buttons
        sku_buttons_layout = QHBoxLayout()
        self.sku_select_all_btn = QPushButton("Select All")
        self.sku_select_all_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {self.theme.get_color('accent')};
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 3px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {self.theme.get_color('accent_hover')};
            }}
        """)
        self.sku_select_all_btn.clicked.connect(self.select_all_skus)
        
        self.sku_reset_btn = QPushButton("Reset")
        self.sku_reset_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: #e74c3c;
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 3px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: #c0392b;
            }}
        """)
        self.sku_reset_btn.clicked.connect(self.reset_skus)
        
        sku_buttons_layout.addWidget(self.sku_select_all_btn)
        sku_buttons_layout.addWidget(self.sku_reset_btn)
        sku_buttons_layout.addStretch()
        sku_layout.addLayout(sku_buttons_layout)
        
        self.global_sku_checkboxes = {}
        self.global_sku_scroll = QScrollArea()
        self.global_sku_scroll.setWidgetResizable(True)
        self.global_sku_scroll.setMaximumHeight(120)
        global_sku_widget = QWidget()
        self.global_sku_widget_layout = QVBoxLayout()
        global_sku_widget.setLayout(self.global_sku_widget_layout)
        self.global_sku_scroll.setWidget(global_sku_widget)
        sku_layout.addWidget(self.global_sku_scroll)
        sku_group.setLayout(sku_layout)
        
        filters_layout.addWidget(period_group)
        filters_layout.addWidget(line_group)
        filters_layout.addWidget(sku_group)
        
        parent_layout.addWidget(self.filters_card)
        
    def toggle_filters_minimize(self):
        """Toggle the minimize/expand state of the filters box"""
        if self.filters_minimized:
            # Expand
            self.filters_content_widget.setVisible(True)
            self.filters_card.setMaximumHeight(300)
            self.filters_minimize_btn.setText("−")
            self.filters_minimized = False
        else:
            # Minimize
            self.filters_content_widget.setVisible(False)
            self.filters_card.setMaximumHeight(40)  # Just enough for the title bar
            self.filters_minimize_btn.setText("+")
            self.filters_minimized = True
        
    def create_header(self, main_layout):
        header = QFrame()
        header.setFixedHeight(40)  # Slightly increased to accommodate larger UGA logo
        
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(40, 5, 40, 5)  # Reduced margins for logos
        
        # PepsiCo Logo
        pepsico_logo = QLabel()
        pepsico_pixmap = QPixmap("PepsiCo_logo.png")
        if not pepsico_pixmap.isNull():
            # Scale logo to fit header height (max 30px)
            scaled_pepsico = pepsico_pixmap.scaledToHeight(30, Qt.SmoothTransformation)
            pepsico_logo.setPixmap(scaled_pepsico)
        pepsico_logo.setAlignment(Qt.AlignCenter)
        
        # Title
        title = QLabel("UGA Capstone - E80 Rejected Units Analyzer")
        title.setFont(QFont("Segoe UI", 14, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        
        # UGA Logo
        uga_logo = QLabel()
        uga_pixmap = QPixmap("ugaengineering.png")
        if not uga_pixmap.isNull():
            # Scale logo to fit header height (max 38px - larger for UGA)
            scaled_uga = uga_pixmap.scaledToHeight(90, Qt.SmoothTransformation)
            uga_logo.setPixmap(scaled_uga)
        uga_logo.setAlignment(Qt.AlignCenter)

        header_layout.addWidget(pepsico_logo)
        header_layout.addStretch()
        header_layout.addWidget(title)
        header_layout.addStretch()
        header_layout.addWidget(uga_logo)
        
        header.setLayout(header_layout)
        main_layout.addWidget(header)
                
    def apply_theme(self):
        bg = self.theme.get_color('bg')
        fg = self.theme.get_color('fg')
        card_bg = self.theme.get_color('card_bg')
        accent = self.theme.get_color('accent')
        hover = self.theme.get_color('hover')
        success = self.theme.get_color('success')
        
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {bg};
                color: {fg};
            }}
            QTabWidget::pane {{
                border: none;
                background-color: {bg};
            }}
            QTabBar::tab {{
                background-color: {card_bg};
                color: {fg};
                padding: 8px 20px;
                margin: 2px;
                border-radius: 6px;
                font-weight: bold;
                min-width: 160px;
            }}
            QTabBar::tab:selected {{
                background-color: {accent};
                color: white;
            }}
            QTabBar::tab:hover {{
                background-color: {hover};
            }}
            QFrame {{
                background-color: {bg};
                color: {fg};
            }}
            QLabel {{
                color: {fg};
                background: transparent;
            }}
            QPushButton {{
                background-color: {card_bg};
                color: {fg};
                border: 1px solid {hover};
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {hover};
            }}
            QPushButton:pressed {{
                background-color: {accent};
                color: white;
            }}
            QPushButton#process_btn_ready {{
                background-color: {success};
                color: white;
                border: 1px solid {success};
            }}
            QPushButton#process_btn_ready:hover {{
                background-color: #27ae60;
            }}
            QTableWidget {{
                background-color: {card_bg};
                color: {fg};
                gridline-color: {hover};
                border: 1px solid {hover};
                border-radius: 4px;
            }}
            QTableWidget::item {{
                padding: 8px;
            }}
            QTableWidget::item:selected {{
                background-color: {accent};
                color: white;
            }}
            QHeaderView::section {{
                background-color: {accent};
                color: white;
                padding: 8px;
                border: none;
                font-weight: bold;
            }}
            QTextEdit {{
                background-color: {card_bg};
                color: {fg};
                border: 1px solid {hover};
                border-radius: 4px;
                padding: 10px;
            }}
            QScrollArea {{
                background-color: {bg};
                border: none;
            }}
            QComboBox {{
                background-color: {card_bg};
                color: {fg};
                border: 1px solid {hover};
                border-radius: 4px;
                padding: 5px;
            }}
            QComboBox::drop-down {{
                border: none;
            }}
            QComboBox::down-arrow {{
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid {fg};
                margin-right: 5px;
            }}
            QDateEdit {{
                background-color: {card_bg};
                color: {fg};
                border: 1px solid {hover};
                border-radius: 4px;
                padding: 5px;
            }}
            QCheckBox {{
                color: {fg};
                spacing: 8px;
            }}
            QCheckBox::indicator {{
                width: 18px;
                height: 18px;
                border: 2px solid {hover};
                border-radius: 3px;
                background-color: {card_bg};
            }}
            QCheckBox::indicator:checked {{
                background-color: {accent};
                border-color: {accent};
            }}
            QGroupBox {{
                font-weight: bold;
                border: 2px solid {hover};
                border-radius: 8px;
                margin-top: 1ex;
                padding-top: 10px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }}
        """)
        
    def create_upload_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)
        
        # Main file upload card
        main_card = ModernCard("📊 Load Rejected Units Data", self.theme)        

        # Icon
        icon_label = QLabel("📁")
        icon_label.setFont(QFont("Segoe UI", 64))
        icon_label.setAlignment(Qt.AlignCenter)
        main_card.content_layout.addWidget(icon_label)
        
        # Description
        desc = QLabel("Select your E80 Rejected Units Excel file (.xlsx)")
        desc.setFont(QFont("Segoe UI", 18))
        desc.setAlignment(Qt.AlignCenter)
        main_card.content_layout.addWidget(desc)
        
        # Upload button
        self.upload_btn = QPushButton("Choose File")
        self.upload_btn.clicked.connect(self.select_file)
        main_card.content_layout.addWidget(self.upload_btn)
        
        # Clear file button (initially hidden)
        self.clear_file_btn = QPushButton("Clear Selection")
        self.clear_file_btn.clicked.connect(self.clear_selected_file)
        self.clear_file_btn.setVisible(False)
        self.clear_file_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        main_card.content_layout.addWidget(self.clear_file_btn)
        
        # Status
        self.status_label = QLabel("Ready to upload rejected units data...")
        self.status_label.setFont(QFont("Segoe UI", 10))
        self.status_label.setAlignment(Qt.AlignCenter)
        main_card.content_layout.addWidget(self.status_label)
        
        layout.addWidget(main_card)
        
        # Process button
        self.process_btn = QPushButton("Analyze Data")
        self.process_btn.clicked.connect(self.process_data)
        self.process_btn.setEnabled(False)
    
        layout.addWidget(self.process_btn)
        
        layout.addStretch()
        
        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "📁 Upload Data")
        
    def create_dashboard_tab(self):
        self.dashboard_tab = QWidget()
        
        # Create scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"background-color: {self.theme.get_color('bg')}; border: none;")
        
        # Create main widget for scroll area
        main_widget = QWidget()
        self.dashboard_layout = QVBoxLayout()
        self.dashboard_layout.setContentsMargins(20, 20, 20, 20)
        
        # Metrics grid (bigger cards)
        self.metrics_widget = QWidget()
        self.metrics_layout = QGridLayout()
        self.metrics_widget.setLayout(self.metrics_layout)
        self.dashboard_layout.addWidget(self.metrics_widget)
        
        main_widget.setLayout(self.dashboard_layout)
        scroll.setWidget(main_widget)
        
        # Set scroll area as the tab content
        tab_layout = QVBoxLayout()
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(scroll)
        self.dashboard_tab.setLayout(tab_layout)
        self.tab_widget.addTab(self.dashboard_tab, "📊 Dashboard")
        
    def create_analysis_tab(self):
        tab = QWidget()
        
        # Create scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"background-color: {self.theme.get_color('bg')}; border: none;")
        
        # Create main widget for scroll area
        main_widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Filters section
        filters_card = ModernCard("🔍 Analysis Filters", self.theme)
        filters_layout = QHBoxLayout()
        
        # Period selection filter
        period_group = QGroupBox("Period Selection")
        period_layout = QVBoxLayout()
        
        # Period checkboxes
        self.period_checkboxes = {}
        self.period_scroll = QScrollArea()
        self.period_scroll.setWidgetResizable(True)
        self.period_scroll.setMaximumHeight(150)
        period_widget = QWidget()
        self.period_widget_layout = QVBoxLayout()
        period_widget.setLayout(self.period_widget_layout)
        self.period_scroll.setWidget(period_widget)
        period_layout.addWidget(self.period_scroll)
        
        # Add period definitions
        period_definitions = {
            'Period 1': 'Dec 29, 2024 - Jan 25, 2025',
            'Period 2': 'Jan 26 - Feb 22, 2025',
            'Period 3': 'Feb 23 - Mar 22, 2025',
            'Period 4': 'Mar 23 - Apr 19, 2025',
            'Period 5': 'Apr 20 - May 17, 2025',
            'Period 6': 'May 18 - Jun 14, 2025',
            'Period 7': 'Jun 15 - Jul 12, 2025',
            'Period 8': 'Jul 13 - Aug 9, 2025',
            'Period 9': 'Aug 10 - Sep 6, 2025',
            'Period 10': 'Sep 7 - Oct 4, 2025',
            'Period 11': 'Oct 5 - Nov 1, 2025',
            'Period 12': 'Nov 2 - Nov 29, 2025',
            'Period 13': 'Nov 30 - Dec 27, 2025'
        }
        
        for period, date_range in period_definitions.items():
            checkbox = QCheckBox(f"{period}: {date_range}")
            checkbox.setChecked(True)  # Default to all selected
            checkbox.stateChanged.connect(self.apply_filters)  # Connect to filter function
            self.period_checkboxes[period] = checkbox
            self.period_widget_layout.addWidget(checkbox)
        
        period_group.setLayout(period_layout)
        
        # Line filter (multi-select)
        line_group = QGroupBox("Production Lines")
        line_layout = QVBoxLayout()
        self.line_checkboxes = {}  # Dictionary to store checkboxes
        self.line_scroll = QScrollArea()
        self.line_scroll.setWidgetResizable(True)
        self.line_scroll.setMaximumHeight(150)
        line_widget = QWidget()
        self.line_widget_layout = QVBoxLayout()
        line_widget.setLayout(self.line_widget_layout)
        self.line_scroll.setWidget(line_widget)
        line_layout.addWidget(self.line_scroll)
        line_group.setLayout(line_layout)
        
        # Rejection reason filter
        reason_group = QGroupBox("Rejection Reason")
        reason_layout = QHBoxLayout()
        self.reason_filter = QComboBox()
        self.reason_filter.addItem("All Reasons")
        reason_layout.addWidget(QLabel("Reason:"))
        reason_layout.addWidget(self.reason_filter)
        reason_group.setLayout(reason_layout)
        
        filters_layout.addWidget(period_group)
        filters_layout.addWidget(line_group)
        filters_layout.addWidget(reason_group)
        
        # Apply filters button
        self.apply_filters_btn = QPushButton("Apply Filters")
        self.apply_filters_btn.clicked.connect(self.apply_filters)
        filters_layout.addWidget(self.apply_filters_btn)
        
        filters_card.content_layout.addLayout(filters_layout)
        layout.addWidget(filters_card)
        
        # Analysis results table
        self.analysis_table = QTableWidget()
        self.analysis_table.setColumnCount(7)
        self.analysis_table.setHorizontalHeaderLabels([
            "Reject DateTime", "Source", "SKU", "Reject Reason", "LPN", 
            "Log Text", "Quantity"
        ])
        
        header = self.analysis_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        self.analysis_table.verticalHeader().setVisible(False)
        
        layout.addWidget(self.analysis_table)
        main_widget.setLayout(layout)
        scroll.setWidget(main_widget)
        
        # Set scroll area as the tab content
        tab_layout = QVBoxLayout()
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(scroll)
        tab.setLayout(tab_layout)
        self.tab_widget.addTab(tab, "🔍 Analysis")
        
    def create_trends_tab(self):
        tab = QWidget()
        
        # Create scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"background-color: {self.theme.get_color('bg')}; border: none;")
        
        # Create main widget for scroll area
        main_widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        
        # Charts section
        charts_card = ModernCard("📈 Period Trends", self.theme)
        
        # Create matplotlib figure for trends (only 2 charts now)
        self.trends_figure = Figure(figsize=(12, 6), dpi=100)
        self.trends_canvas = FigureCanvas(self.trends_figure)
        charts_card.content_layout.addWidget(self.trends_canvas)
        
        layout.addWidget(charts_card)
        main_widget.setLayout(layout)
        scroll.setWidget(main_widget)
        
        # Set scroll area as the tab content
        tab_layout = QVBoxLayout()
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(scroll)
        tab.setLayout(tab_layout)
        self.tab_widget.addTab(tab, "📈 Trends")
        
    def create_production_analysis_tab(self):
        tab = QWidget()
        
        # Create scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"background-color: {self.theme.get_color('bg')}; border: none;")
        
        # Create main widget for scroll area
        main_widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Production analysis
        production_card = ModernCard("🏭 Production Line & Product Analysis", self.theme)
        
        # Create matplotlib figure for production analysis
        self.production_figure = Figure(figsize=(12, 8), dpi=100)
        self.production_canvas = FigureCanvas(self.production_figure)
        production_card.content_layout.addWidget(self.production_canvas)
        
        layout.addWidget(production_card)
        main_widget.setLayout(layout)
        scroll.setWidget(main_widget)
        
        # Set scroll area as the tab content
        tab_layout = QVBoxLayout()
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(scroll)
        tab.setLayout(tab_layout)
        self.tab_widget.addTab(tab, "🏭 Production Analysis")
        
    def create_production_lines_tab(self):
        tab = QWidget()
        
        # Create scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"background-color: {self.theme.get_color('bg')}; border: none;")
        
        # Create main widget for scroll area
        main_widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Production lines consolidation
        lines_card = ModernCard("🏭 Production Lines Consolidation", self.theme)
        
        # Create matplotlib figure for production lines
        self.production_lines_figure = Figure(figsize=(12, 8), dpi=100)
        self.production_lines_canvas = FigureCanvas(self.production_lines_figure)
        lines_card.content_layout.addWidget(self.production_lines_canvas)
        
        layout.addWidget(lines_card)
        main_widget.setLayout(layout)
        scroll.setWidget(main_widget)
        
        # Set scroll area as the tab content
        tab_layout = QVBoxLayout()
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(scroll)
        tab.setLayout(tab_layout)
        self.tab_widget.addTab(tab, "🏭 Production Lines")
        
    def create_advanced_tracking_tab(self):
        tab = QWidget()
        
        # Create scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"background-color: {self.theme.get_color('bg')}; border: none;")
        
        # Create main widget for scroll area
        main_widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Category analysis
        category_card = ModernCard("🔍 Rejection Category Analysis", self.theme)
        
        # Create matplotlib figure for category analysis
        self.category_figure = Figure(figsize=(12, 8), dpi=100)
        self.category_canvas = FigureCanvas(self.category_figure)
        category_card.content_layout.addWidget(self.category_canvas)
        
        layout.addWidget(category_card)
        main_widget.setLayout(layout)
        scroll.setWidget(main_widget)
        
        # Set scroll area as the tab content
        tab_layout = QVBoxLayout()
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(scroll)
        tab.setLayout(tab_layout)
        self.tab_widget.addTab(tab, "🔍 Advanced Tracking")
        
    def create_dimensional_rejects_tab(self):
        tab = QWidget()
        
        # Create scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"background-color: {self.theme.get_color('bg')}; border: none;")
        
        # Create main widget for scroll area
        main_widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Dimensional rejects analysis
        dimensional_card = ModernCard("📏 Dimensional Rejects Analysis", self.theme)
        
        # Create matplotlib figure for dimensional analysis
        self.dimensional_figure = Figure(figsize=(12, 8), dpi=100)
        self.dimensional_canvas = FigureCanvas(self.dimensional_figure)
        dimensional_card.content_layout.addWidget(self.dimensional_canvas)
        
        layout.addWidget(dimensional_card)
        main_widget.setLayout(layout)
        scroll.setWidget(main_widget)
        
        # Set scroll area as the tab content
        tab_layout = QVBoxLayout()
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(scroll)
        tab.setLayout(tab_layout)
        self.tab_widget.addTab(tab, "📏 Dimensional Rejects")
        
    def create_tag_tracking_rejects_tab(self):
        tab = QWidget()
        
        # Create scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"background-color: {self.theme.get_color('bg')}; border: none;")
        
        # Create main widget for scroll area
        main_widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Tag/Tracking rejects analysis
        tag_tracking_card = ModernCard("🏷️ Tag/Tracking Rejects Analysis", self.theme)
        
        # Create matplotlib figure for tag/tracking analysis
        self.tag_tracking_figure = Figure(figsize=(12, 8), dpi=100)
        self.tag_tracking_canvas = FigureCanvas(self.tag_tracking_figure)
        tag_tracking_card.content_layout.addWidget(self.tag_tracking_canvas)
        
        layout.addWidget(tag_tracking_card)
        main_widget.setLayout(layout)
        scroll.setWidget(main_widget)
        
        # Set scroll area as the tab content
        tab_layout = QVBoxLayout()
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(scroll)
        tab.setLayout(tab_layout)
        self.tab_widget.addTab(tab, "🏷️ Tag/Tracking Rejects")
        
    def create_time_analysis_tab(self):
        tab = QWidget()
        
        # Create scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(f"background-color: {self.theme.get_color('bg')}; border: none;")
        
        # Create main widget for scroll area
        main_widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Time analysis
        time_card = ModernCard("⏰ Time-of-Day Analysis", self.theme)
        
        # Create matplotlib figure for time analysis
        self.time_figure = Figure(figsize=(12, 8), dpi=100)
        self.time_canvas = FigureCanvas(self.time_figure)
        time_card.content_layout.addWidget(self.time_canvas)
        
        layout.addWidget(time_card)
        main_widget.setLayout(layout)
        scroll.setWidget(main_widget)
        
        # Set scroll area as the tab content
        tab_layout = QVBoxLayout()
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(scroll)
        tab.setLayout(tab_layout)
        self.tab_widget.addTab(tab, "⏰ Time Analysis")
        
    def create_sku_analysis_tab(self):
        """Create SKU analysis tab"""
        sku_tab = QWidget()
        sku_layout = QVBoxLayout()
        sku_tab.setLayout(sku_layout)
        
        # Add scroll area
        sku_scroll = QScrollArea()
        sku_scroll.setWidgetResizable(True)
        sku_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        sku_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        sku_content = QWidget()
        sku_content_layout = QVBoxLayout()
        sku_content.setLayout(sku_content_layout)
        
        # Product Analysis title
        title = QLabel("📦 Product Analysis")
        title.setStyleSheet(f"""
            font-size: 24px;
            font-weight: bold;
            color: {self.theme.get_color('fg')};
            margin: 20px 0;
        """)
        title.setAlignment(Qt.AlignCenter)
        sku_content_layout.addWidget(title)
        
        # Create matplotlib figure for SKU analysis
        self.sku_figure = Figure(figsize=(12, 8), facecolor=self.theme.get_color('card_bg'))
        self.sku_canvas = FigureCanvas(self.sku_figure)
        self.sku_canvas.setStyleSheet(f"background-color: {self.theme.get_color('card_bg')};")
        
        sku_content_layout.addWidget(self.sku_canvas)
        
        sku_scroll.setWidget(sku_content)
        sku_layout.addWidget(sku_scroll)
        
        self.tab_widget.addTab(sku_tab, "📦 Product Analysis")
    
    def create_rejection_rate_tab(self):
        """Create rejection rate analysis tab"""
        rejection_tab = QWidget()
        rejection_layout = QVBoxLayout()
        rejection_tab.setLayout(rejection_layout)
        
        # Create scroll area for the tab content
        rejection_scroll = QScrollArea()
        rejection_scroll.setWidgetResizable(True)
        rejection_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        rejection_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        rejection_content = QWidget()
        rejection_content_layout = QVBoxLayout()
        rejection_content.setLayout(rejection_content_layout)
        
        # Rejection Rate Analysis title
        title = QLabel("📊 Rejection Rate Analysis")
        title.setStyleSheet(f"""
            font-size: 24px;
            font-weight: bold;
            color: {self.theme.get_color('fg')};
            margin: 20px 0;
            padding: 10px;
        """)
        rejection_content_layout.addWidget(title)
        
        # Create matplotlib figure for rejection rate analysis
        self.rejection_figure = Figure(figsize=(12, 10), facecolor=self.theme.get_color('card_bg'))
        self.rejection_canvas = FigureCanvas(self.rejection_figure)
        self.rejection_canvas.setStyleSheet(f"background-color: {self.theme.get_color('card_bg')};")
        
        rejection_content_layout.addWidget(self.rejection_canvas)
        
        rejection_scroll.setWidget(rejection_content)
        rejection_layout.addWidget(rejection_scroll)
        
        self.tab_widget.addTab(rejection_tab, "📊 Rejection Rates")
        
    def select_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Select Rejected Units Excel File", "", 
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        
        if filename:
            self.current_file = filename
            self.status_label.setText(f"📁 File loaded: {os.path.basename(filename)}")
            self.process_btn.setEnabled(True)
            
            # Show clear button
            self.clear_file_btn.setVisible(True)
            
            # Make button green when file is selected
            self.process_btn.setObjectName("process_btn_ready")
            self.process_btn.setStyleSheet("")
            self.apply_theme()
    
    def clear_selected_file(self):
        self.current_file = None
        self.current_data = None
        self.filtered_data = None
        self.analysis_results = None
        self.status_label.setText("Ready to upload rejected units data...")
        self.process_btn.setEnabled(False)
        self.clear_file_btn.setVisible(False)
        self.process_btn.setObjectName("")
        self.apply_theme()
        
        # Clear all displays
        self.clear_displays()
        
    def clear_displays(self):
        # Clear metrics
        for i in reversed(range(self.metrics_layout.count())):
            self.metrics_layout.itemAt(i).widget().setParent(None)
            
        # Clear summary
        self.summary_text.setText("No data loaded. Please upload a rejected units file.")
        
        
        # Clear trends chart
        self.trends_figure.clear()
        self.trends_canvas.draw()
        
        # Clear production analysis chart
        self.production_figure.clear()
        self.production_canvas.draw()
        
        # Clear production lines chart
        self.production_lines_figure.clear()
        self.production_lines_canvas.draw()
        
        # Clear advanced tracking chart
        self.category_figure.clear()
        self.category_canvas.draw()
        
        # Clear dimensional rejects chart
        self.dimensional_figure.clear()
        self.dimensional_canvas.draw()
        
        # Clear tag/tracking rejects chart
        self.tag_tracking_figure.clear()
        self.tag_tracking_canvas.draw()
        
        # Clear time analysis chart
        self.time_figure.clear()
        self.time_canvas.draw()
        
        # Clear SKU analysis chart
        self.sku_figure.clear()
        self.sku_canvas.draw()
        
        # Clear rejection rate analysis chart
        self.rejection_figure.clear()
        self.rejection_canvas.draw()
        
        
    def process_data(self):
        if not self.current_file:
            QMessageBox.warning(self, "Error", "Please select a file first")
            return
            
        try:
            # Load the Excel file and detect the correct sheet
            excel_file = pd.ExcelFile(self.current_file)
            sheet_names = excel_file.sheet_names
            
            # Look for the sheet with the actual data (contains 'Reject datetime' column)
            data_sheet = None
            for sheet_name in sheet_names:
                try:
                    # Read just the header to check columns
                    temp_df = pd.read_excel(self.current_file, sheet_name=sheet_name, nrows=0)
                    if 'Reject datetime' in temp_df.columns:
                        data_sheet = sheet_name
                        break
                except:
                    continue
            
            if data_sheet is None:
                QMessageBox.warning(self, "Invalid Data Format", 
                                  "Could not find a sheet with 'Reject datetime' column.\n\n"
                                  "Please ensure this is a valid E80 rejected units export.")
                return
                
            # Load the data from the correct sheet
            df = pd.read_excel(self.current_file, sheet_name=data_sheet)
            
            # Basic validation - check for expected columns (E80 format)
            expected_columns = ['Reject datetime', 'Source', 'Reject reason']
            missing_columns = [col for col in expected_columns if col not in df.columns]
            
            if missing_columns:
                QMessageBox.warning(self, "Invalid Data Format", 
                                  f"Missing expected columns: {', '.join(missing_columns)}\n\n"
                                  f"Please ensure this is a valid E80 rejected units export.\n"
                                  f"Expected columns: Reject datetime, Source, Reject reason, Lpn, Sku, Log text")
                return
                
            # Clean and process the data
            df = self.clean_data(df)
            self.current_data = df
            self.filtered_data = df.copy()  # Initialize filtered data
            
            # Perform analysis
            self.analysis_results = self.perform_analysis(df)
            
            # Update filters first, then all displays
            self.update_filters()
            self.update_dashboard()
            self.update_trends()
            self.update_production_analysis()
            self.update_production_lines()
            self.update_advanced_tracking()
            self.update_dimensional_rejects()
            self.update_tag_tracking_rejects()
            self.update_time_analysis()
            self.update_sku_analysis()
            self.update_rejection_rate_analysis()
            
            # Filter checkboxes are now populated in update_filters method
            
            # Switch to Dashboard tab (2nd page)
            self.tab_widget.setCurrentIndex(1)
            
            self.status_label.setText(f"Data analyzed successfully! (Sheet: {data_sheet}) - Switched to Dashboard tab")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error processing file: {str(e)}")
            
    def clean_data(self, df):
        """Clean and standardize the data"""
        # Convert date column to datetime (E80 format uses 'Reject datetime')
        if 'Reject datetime' in df.columns:
            df['Reject datetime'] = pd.to_datetime(df['Reject datetime'], errors='coerce')
            # Create a standard 'Date' column for compatibility - handle NaT values
            df['Date'] = df['Reject datetime'].apply(lambda x: x.date() if pd.notna(x) else None)
            
        # Fill missing values
        df = df.fillna('Unknown')
        
        # Standardize text columns (E80 format)
        text_columns = ['Source', 'Reject reason', 'Lpn', 'Sku', 'Log text']
        for col in text_columns:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
        
        # Replace SKU numbers with product descriptions
        df = self.replace_skus_with_descriptions(df)
                
        # Create quantity column (each row represents 1 rejected unit)
        df['Quantity'] = 1
        
        return df
    
    def replace_skus_with_descriptions(self, df):
        """Replace SKU numbers with product descriptions from the master file"""
        if 'Sku' not in df.columns:
            return df
            
        try:
            # Load the SKU master file
            sku_master_path = 'E80 Item Master - Master Excel.xlsx'
            sku_master = pd.read_excel(sku_master_path)
            
            # Create a mapping dictionary from SKU number to description
            # Convert both to strings for consistent matching
            sku_mapping = dict(zip(sku_master['Name'].astype(str), sku_master['Description']))
            
            # Convert rejected SKUs to proper format (they come as floats ending with .0)
            # Handle the .0 format and 'Unknown' values properly
            df['Sku'] = df['Sku'].astype(str).str.replace('.0', '').str.replace('nan', 'Unknown')
            
            # Only convert numeric SKUs to int, leave 'Unknown' as is
            numeric_mask = df['Sku'].str.match(r'^\d+$', na=False)
            df.loc[numeric_mask, 'Sku'] = df.loc[numeric_mask, 'Sku'].astype(int).astype(str)
            
            # Replace SKU numbers with descriptions
            df['Sku'] = df['Sku'].map(sku_mapping).fillna(df['Sku'])
            
            # Count how many were successfully mapped
            mapped_count = len(df[df['Sku'].isin(sku_master['Description'])])
            total_skus = len(df[df['Sku'] != '0'])  # Exclude NaN values converted to '0'
            
            print(f"Successfully mapped {mapped_count} out of {total_skus} SKU records to descriptions")
            print(f"Available mappings: {len(sku_mapping)} SKU numbers")
            
        except Exception as e:
            print(f"Warning: Could not load SKU master file: {e}")
            print("Continuing with original SKU numbers...")
            
        return df
        
    def perform_analysis(self, df):
        """Perform comprehensive analysis of rejected units"""
        analysis = {}
        
        # Basic statistics
        analysis['total_rejections'] = len(df)
        analysis['total_quantity'] = df['Quantity'].sum() if 'Quantity' in df.columns else 0
        analysis['date_range'] = (df['Date'].min(), df['Date'].max()) if 'Date' in df.columns else (None, None)
        
        # Rejection reasons analysis (E80 format uses 'Reject reason')
        if 'Reject reason' in df.columns:
            analysis['rejection_reasons'] = df['Reject reason'].value_counts().to_dict()
            
        # Source/Line analysis (E80 format uses 'Source')
        if 'Source' in df.columns:
            analysis['line_breakdown'] = df['Source'].value_counts().to_dict()
            
        # Time-based analysis - convert to periods
        if 'Date' in df.columns and 'Reject datetime' in df.columns:
            # Use the datetime column for proper time analysis
            df['Month'] = df['Reject datetime'].dt.to_period('M')
            analysis['monthly_trends'] = df.groupby('Month')['Quantity'].sum().to_dict()
            
            # Add period analysis based on Pepsi's period calendar
            df['Period'] = self.get_period_from_date(df['Reject datetime'])
            analysis['period_trends'] = df.groupby('Period')['Quantity'].sum().to_dict()
            
        # Product analysis (E80 format uses 'Sku')
        if 'Sku' in df.columns:
            analysis['product_breakdown'] = df['Sku'].value_counts().to_dict()
            
        return analysis
        
    def get_period_from_date(self, date_series):
        """Convert dates to Pepsi period numbers based on the period calendar"""
        periods = []
        for date in date_series:
            if pd.isna(date):
                periods.append('Unknown')
                continue
                
            # Convert to date if it's datetime
            if hasattr(date, 'date'):
                date = date.date()
            
            # Period definitions based on the calendar
            if date >= pd.Timestamp('2024-12-29').date() and date <= pd.Timestamp('2025-01-25').date():
                periods.append('Period 1')
            elif date >= pd.Timestamp('2025-01-26').date() and date <= pd.Timestamp('2025-02-22').date():
                periods.append('Period 2')
            elif date >= pd.Timestamp('2025-02-23').date() and date <= pd.Timestamp('2025-03-22').date():
                periods.append('Period 3')
            elif date >= pd.Timestamp('2025-03-23').date() and date <= pd.Timestamp('2025-04-19').date():
                periods.append('Period 4')
            elif date >= pd.Timestamp('2025-04-20').date() and date <= pd.Timestamp('2025-05-17').date():
                periods.append('Period 5')
            elif date >= pd.Timestamp('2025-05-18').date() and date <= pd.Timestamp('2025-06-14').date():
                periods.append('Period 6')
            elif date >= pd.Timestamp('2025-06-15').date() and date <= pd.Timestamp('2025-07-12').date():
                periods.append('Period 7')
            elif date >= pd.Timestamp('2025-07-13').date() and date <= pd.Timestamp('2025-08-09').date():
                periods.append('Period 8')
            elif date >= pd.Timestamp('2025-08-10').date() and date <= pd.Timestamp('2025-09-06').date():
                periods.append('Period 9')
            elif date >= pd.Timestamp('2025-09-07').date() and date <= pd.Timestamp('2025-10-04').date():
                periods.append('Period 10')
            elif date >= pd.Timestamp('2025-10-05').date() and date <= pd.Timestamp('2025-11-01').date():
                periods.append('Period 11')
            elif date >= pd.Timestamp('2025-11-02').date() and date <= pd.Timestamp('2025-11-29').date():
                periods.append('Period 12')
            elif date >= pd.Timestamp('2025-11-30').date() and date <= pd.Timestamp('2025-12-27').date():
                periods.append('Period 13')
            else:
                periods.append('Unknown')
                
        return periods
        
    def update_dashboard(self):
        if self.filtered_data is None:
            return
            
        # Clear existing metrics
        for i in reversed(range(self.metrics_layout.count())):
            self.metrics_layout.itemAt(i).widget().setParent(None)
            
        # Create metric cards with filtered data
        total_rejections = len(self.filtered_data)
        
        # Get top rejection reason from filtered data
        top_reason = "Unknown"
        if 'Reject reason' in self.filtered_data.columns:
            rejection_counts = self.filtered_data['Reject reason'].value_counts()
            if len(rejection_counts) > 0:
                top_reason = rejection_counts.index[0]
        
        # Calculate date range from filtered data
        date_range = "No Data"
        if 'Reject datetime' in self.filtered_data.columns:
            # Remove NaT values and get valid dates
            valid_dates = self.filtered_data['Reject datetime'].dropna()
            if len(valid_dates) > 0:
                min_date = valid_dates.min().strftime('%m/%d/%Y')
                max_date = valid_dates.max().strftime('%m/%d/%Y')
                date_range = f"{min_date} - {max_date}"
            
        metrics = [
            ("Total Rejections", f"{total_rejections:,}", self.theme.get_color('danger')),
            ("Top Reason", top_reason[:20], self.theme.get_color('info')),
            ("Date Range", date_range, self.theme.get_color('success')),
            ("Est. Lost Time", "TBD - Ask Oscar", self.theme.get_color('warning')),
            ("Est. Cost Impact", "TBD - Ask Oscar", self.theme.get_color('purple'))
        ]
        
        for i, (title, value, color) in enumerate(metrics):
            card = MetricCard(title, value, color, self.theme)
            self.metrics_layout.addWidget(card, 0, i)
        
    def update_trends(self):
        if self.filtered_data is None:
            return
            
        # Clear existing plots
        self.trends_figure.clear()
        
        # Create subplots (only 2 charts now)
        ax1 = self.trends_figure.add_subplot(1, 2, 1)
        ax2 = self.trends_figure.add_subplot(1, 2, 2)
        
        # Set dark theme for plots
        self.trends_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Use the globally filtered data
        filtered_data = self.filtered_data.copy()
        
        # Plot 1: Rejection reasons pie chart (filtered) with accurate percentages
        if 'Reject reason' in filtered_data.columns:
            rejection_counts = filtered_data['Reject reason'].value_counts()
            total_rejections = rejection_counts.sum()
            
            # Get top 5 reasons
            top_reasons = rejection_counts.head(5)
            other_count = rejection_counts.iloc[5:].sum() if len(rejection_counts) > 5 else 0
            
            # Prepare data for pie chart
            pie_labels = list(top_reasons.index)
            pie_counts = list(top_reasons.values)
            
            # Add "Other" slice if there are more than 5 reasons
            if other_count > 0:
                pie_labels.append('One of the other 20 Reasons')
                pie_counts.append(other_count)
            
            colors = ['#e74c3c', '#f39c12', '#f1c40f', '#2ecc71', '#3498db', '#9b59b6']
            wedges, texts, autotexts = ax1.pie(pie_counts, labels=pie_labels, autopct='%1.1f%%', 
                                             colors=colors[:len(pie_labels)], 
                                             textprops={'color': 'white', 'fontweight': 'bold'})
            ax1.set_title('Top Rejection Reasons', color='white', fontweight='bold', fontsize=16)
            
        # Plot 2: Period trends (filtered)
        if 'Reject datetime' in filtered_data.columns:
            # Calculate period trends for filtered data
            filtered_data['Period'] = self.get_period_from_date(filtered_data['Reject datetime'])
            period_trends = filtered_data.groupby('Period')['Quantity'].sum()
            
            # Sort periods numerically to ensure correct chronological order
            # Extract number from "Period X" format for proper sorting
            periods = sorted(period_trends.keys(), key=lambda x: int(x.split()[-1]))
           
            quantities = [period_trends[period] for period in periods]
            
            # Create clean period labels (just the numbers)
            period_labels = [period.split()[-1] for period in periods]
            
            ax2.plot(periods, quantities, marker='o', color='#e74c3c', linewidth=3, markersize=8)
            ax2.set_title('Period Rejection Trends', color='white', fontweight='bold', fontsize=16)
            ax2.set_xlabel('Period', color='white', fontsize=12)
            ax2.set_ylabel('Rejected Quantity', color='white', fontsize=12)
            ax2.set_xticks(range(len(periods)))
            ax2.set_xticklabels(period_labels)
            ax2.tick_params(axis='x', rotation=0, colors='white')
            ax2.tick_params(axis='y', colors='white')
            ax2.grid(True, alpha=0.3, color='white')

            # Add value labels on line plot points
            for i, (period, quantity) in enumerate(zip(periods, quantities)):
                ax2.text(i, quantity + max(quantities) * 0.02, f'{int(quantity)}', 
                        ha='center', va='bottom', color='white', fontweight='bold', fontsize=8,
                        bbox=dict(boxstyle="round,pad=0.3", facecolor='black', alpha=0.8))
            
        # Style all axes
        for ax in [ax1, ax2]:
            ax.tick_params(colors='white')
            for spine in ax.spines.values():
                spine.set_color('white')
                
        self.trends_figure.tight_layout()
        self.trends_canvas.draw()
        
    def update_production_analysis(self):
        """Update production line and product analysis"""
        if self.filtered_data is None:
            return
            
        # Clear existing plots
        self.production_figure.clear()
        
        # Create subplots
        ax1 = self.production_figure.add_subplot(1, 2, 1)
        ax2 = self.production_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.production_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Use the globally filtered data
        filtered_data = self.filtered_data.copy()
        
        # Plot 1: Line breakdown (using filtered data)
        if 'Source' in filtered_data.columns:
            line_counts = filtered_data['Source'].value_counts()
            
            bars = ax1.bar(line_counts.index, line_counts.values, color='#3498db', edgecolor='white', linewidth=1)
            ax1.set_title('Rejections by Production Line', color='white', fontweight='bold', fontsize=16)
            ax1.set_ylabel('Number of Rejections', color='white', fontsize=12)
            ax1.tick_params(axis='x', colors='white', rotation=45)
            ax1.tick_params(axis='y', colors='white')
            ax1.grid(True, alpha=0.3, color='white', axis='y')
            
            # Add value labels on bars
            for bar in bars:
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(height)}', ha='center', va='bottom', color='white', fontweight='bold',
                        bbox=dict(boxstyle="round,pad=0.3", facecolor='black', alpha=0.8))
            
        # Plot 2: Product breakdown (using filtered data)
        if 'Sku' in filtered_data.columns:
            product_counts = filtered_data['Sku'].value_counts().head(8)
            
            bars = ax2.barh(product_counts.index, product_counts.values, color='#9b59b6', edgecolor='white', linewidth=1)
            ax2.set_title('Rejections by Product', color='white', fontweight='bold', fontsize=16)
            ax2.set_xlabel('Number of Rejections', color='white', fontsize=12)
            ax2.tick_params(axis='x', colors='white')
            ax2.tick_params(axis='y', colors='white')
            ax2.grid(True, alpha=0.3, color='white', axis='x')
            
            # Add value labels on horizontal bars
            for bar in bars:
                width = bar.get_width()
                ax2.text(width + 0.1, bar.get_y() + bar.get_height()/2.,
                        f'{int(width)}', ha='left', va='center', color='white', fontweight='bold',
                        bbox=dict(boxstyle="round,pad=0.3", facecolor='black', alpha=0.8))
            
        # Style all axes
        for ax in [ax1, ax2]:
            for spine in ax.spines.values():
                spine.set_color('white')
                
        self.production_figure.tight_layout()
        self.production_canvas.draw()
        
    def update_production_lines(self):
        """Update production lines consolidation analysis"""
        if self.filtered_data is None:
            return
            
        # Clear existing plots
        self.production_lines_figure.clear()
        
        # Create subplots
        ax1 = self.production_lines_figure.add_subplot(1, 2, 1)
        ax2 = self.production_lines_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.production_lines_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Use the globally filtered data
        filtered_data = self.filtered_data.copy()
        
        # Apply period filter
        if hasattr(self, 'production_lines_period_checkboxes') and self.production_lines_period_checkboxes:
            selected_periods = [period for period, checkbox in self.production_lines_period_checkboxes.items() if checkbox.isChecked()]
            if selected_periods:
                filtered_data['Period'] = self.get_period_from_date(filtered_data['Reject datetime'])
                filtered_data = filtered_data[filtered_data['Period'].isin(selected_periods)]
        
        # Apply line filter
        if hasattr(self, 'production_lines_line_checkboxes') and self.production_lines_line_checkboxes:
            selected_lines = [line for line, checkbox in self.production_lines_line_checkboxes.items() if checkbox.isChecked()]
            if selected_lines:
                filtered_data = filtered_data[filtered_data['Source'].isin(selected_lines)]
        
        # Define line consolidation mapping based on EOL numbers
        def consolidate_line(source):
            """Consolidate production lines based on EOL numbers"""
            if pd.isna(source):
                return 'Unknown'
            
            source_str = str(source).upper()
            
            # Check for IBC first (Inbound Conveyor by Car Wash)
            if 'IBC' in source_str:
                return 'IBC (Inbound Conveyor)'
            # Check for EOL patterns
            elif 'EOL01' in source_str or 'EOL_01' in source_str:
                return 'Aquafina/Propel 1'
            elif 'EOL02' in source_str or 'EOL_02' in source_str:
                return 'Aquafina/Propel 2'
            elif 'EOL03' in source_str or 'EOL_03' in source_str:
                return 'Can Line 3'
            elif 'EOL04' in source_str or 'EOL_04' in source_str:
                return 'Can Line 4'
            elif 'EOL05' in source_str or 'EOL_05' in source_str:
                return 'Bottle Line 5'
            elif 'EOL06' in source_str or 'EOL_06' in source_str:
                return 'Bottle Line 6'
            # Check for simple number patterns (but not if it's IBC)
            elif source_str == '1' or source_str.endswith('1'):
                return 'Aquafina/Propel 1'
            elif source_str == '2' or source_str.endswith('2'):
                return 'Aquafina/Propel 2'
            elif source_str == '3' or source_str.endswith('3'):
                return 'Can Line 3'
            elif source_str == '4' or source_str.endswith('4'):
                return 'Can Line 4'
            elif source_str == '5' or source_str.endswith('5'):
                return 'Bottle Line 5'
            elif source_str == '6' or source_str.endswith('6'):
                return 'Bottle Line 6'
            else:
                return source_str  # Keep original if no pattern matches
        
        # Create consolidated data
        if 'Source' in filtered_data.columns:
            # Map individual lines to consolidated lines
            filtered_data['Consolidated_Line'] = filtered_data['Source'].apply(consolidate_line)
            
            # Plot 1: Consolidated production lines
            consolidated_counts = filtered_data['Consolidated_Line'].value_counts()
            
            colors = ['#e74c3c', '#f39c12', '#2ecc71', '#9b59b6', '#1abc9c', '#34495e']
            bars = ax1.bar(consolidated_counts.index, consolidated_counts.values, 
                          color=colors[:len(consolidated_counts)], edgecolor='white', linewidth=1)
            ax1.set_title('Rejections by Consolidated Production Line', color='white', fontweight='bold', fontsize=14)
            ax1.set_ylabel('Number of Rejections', color='white', fontsize=12)
            ax1.tick_params(axis='x', colors='white', rotation=45)
            ax1.tick_params(axis='y', colors='white')
            ax1.grid(True, alpha=0.3, color='white', axis='y')
            
            # Add value labels on bars
            for bar in bars:
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(height)}', ha='center', va='bottom', color='white', fontweight='bold',
                        bbox=dict(boxstyle="round,pad=0.3", facecolor='black', alpha=0.8))
            
            # Plot 2: Detailed breakdown showing individual lines within consolidated groups
            detailed_data = {}
            detailed_labels = []
            detailed_counts = []
            detailed_colors = []
            
            color_palette = ['#e74c3c', '#f39c12', '#2ecc71', '#9b59b6', '#1abc9c', '#34495e', '#e67e22', '#95a5a6']
            
            for consolidated_line in consolidated_counts.index:
                # Get individual lines that map to this consolidated line
                individual_lines = filtered_data[filtered_data['Consolidated_Line'] == consolidated_line]['Source'].value_counts()
                
                for i, (individual_line, count) in enumerate(individual_lines.items()):
                    detailed_labels.append(f"{individual_line}\n({consolidated_line})")
                    detailed_counts.append(count)
                    detailed_colors.append(color_palette[i % len(color_palette)])
            
            bars = ax2.bar(range(len(detailed_labels)), detailed_counts, color=detailed_colors, edgecolor='white', linewidth=1)
            ax2.set_title('Individual Lines Within Consolidated Groups', color='white', fontweight='bold', fontsize=14)
            ax2.set_ylabel('Number of Rejections', color='white', fontsize=12)
            ax2.set_xticks(range(len(detailed_labels)))
            ax2.set_xticklabels(detailed_labels, rotation=45, ha='right')
            ax2.tick_params(axis='x', colors='white')
            ax2.tick_params(axis='y', colors='white')
            ax2.grid(True, alpha=0.3, color='white', axis='y')
            
            # Add value labels on bars
            for bar in bars:
                height = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(height)}', ha='center', va='bottom', color='white', fontweight='bold', fontsize=8,
                        bbox=dict(boxstyle="round,pad=0.2", facecolor='black', alpha=0.8))
        else:
            # No data available
            ax1.text(0.5, 0.5, 'No Production Line Data Available', ha='center', va='center', 
                    transform=ax1.transAxes, color='white', fontsize=16)
            ax2.text(0.5, 0.5, 'No Production Line Data Available', ha='center', va='center', 
                    transform=ax2.transAxes, color='white', fontsize=16)
        
        # Style all axes
        for ax in [ax1, ax2]:
            for spine in ax.spines.values():
                spine.set_color('white')
                
        self.production_lines_figure.tight_layout()
        self.production_lines_canvas.draw()
        
    def update_advanced_tracking(self):
        """Update advanced tracking with rejection reason categories"""
        if self.filtered_data is None:
            return
            
        # Clear existing plots
        self.category_figure.clear()
        
        # Create subplots
        ax1 = self.category_figure.add_subplot(1, 2, 1)
        ax2 = self.category_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.category_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Use the globally filtered data
        filtered_data = self.filtered_data.copy()
        
        # Categorize rejection reasons using filtered data
        dimensional_issues = []
        tag_tracking_issues = []
        uncategorized_issues = []
        
        if 'Reject reason' in filtered_data.columns:
            rejection_counts = filtered_data['Reject reason'].value_counts()
            for reason, count in rejection_counts.items():
                if pd.isna(reason):
                    continue
                reason_lower = str(reason).lower()
                if any(keyword in reason_lower for keyword in ['dimension', 'size', 'measurement', 'weight', 'position', 'height', 'width', 'length', 'tolerance', 'maximum']):
                    dimensional_issues.append((reason, count))
                elif any(keyword in reason_lower for keyword in ['tag', 'label', 'lpn', 'barcode', 'duplicate', 'unit data not found', 'tracking', 'expected', 'exist', 'system', 'error', 'timeout', 'failed', 'check error']):
                    tag_tracking_issues.append((reason, count))
                else:
                    uncategorized_issues.append((reason, count))
        
        # Show uncategorized issues for user to categorize
        if uncategorized_issues:
            print("Uncategorized rejection reasons that need categorization:")
            for reason, count in uncategorized_issues:
                print(f"  - '{reason}': {count} occurrences")
        
        # Plot 1: Category breakdown (two categories now)
        categories = ['Dimensional Issues', 'Tag/Tracking/System Issues']
        category_counts = [sum(count for _, count in dimensional_issues),
                          sum(count for _, count in tag_tracking_issues)]
        
        colors = ['#e74c3c', '#f39c12']
        wedges, texts, autotexts = ax1.pie(category_counts, labels=categories, autopct='%1.1f%%', 
                                          colors=colors, textprops={'color': 'white', 'fontweight': 'bold'})
        ax1.set_title('Rejection Categories', color='white', fontweight='bold', fontsize=14)
        
        # Plot 2: Top reasons by category
        all_issues = dimensional_issues + tag_tracking_issues
        all_issues.sort(key=lambda x: x[1], reverse=True)
        
        top_reasons = all_issues[:8]
        reasons = [reason for reason, _ in top_reasons]
        counts = [count for _, count in top_reasons]
        
        bars = ax2.barh(reasons, counts, color='#3498db', edgecolor='white', linewidth=1)
        ax2.set_title('Top Rejection Reasons', color='white', fontweight='bold', fontsize=14)
        ax2.set_xlabel('Number of Rejections', color='white', fontsize=12)
        ax2.tick_params(axis='x', colors='white')
        ax2.tick_params(axis='y', colors='white')
        ax2.grid(True, alpha=0.3, color='white', axis='x')
        
        # Add value labels on horizontal bars
        for bar in bars:
            width = bar.get_width()
            ax2.text(width + 0.1, bar.get_y() + bar.get_height()/2.,
                    f'{int(width)}', ha='left', va='center', color='white', fontweight='bold',
                    bbox=dict(boxstyle="round,pad=0.3", facecolor='black', alpha=0.8))
        
        # Style axes
        for ax in [ax1, ax2]:
            for spine in ax.spines.values():
                spine.set_color('white')
                
        self.category_figure.tight_layout()
        self.category_canvas.draw()
        
    def update_dimensional_rejects(self):
        """Update dimensional rejects analysis"""
        if self.filtered_data is None:
            return
            
        # Clear existing plots
        self.dimensional_figure.clear()
        
        # Create subplots
        ax1 = self.dimensional_figure.add_subplot(1, 2, 1)
        ax2 = self.dimensional_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.dimensional_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Use the globally filtered data
        filtered_data = self.filtered_data.copy()
        
        # Apply period filter
        if hasattr(self, 'dimensional_period_checkboxes') and self.dimensional_period_checkboxes:
            selected_periods = [period for period, checkbox in self.dimensional_period_checkboxes.items() if checkbox.isChecked()]
            if selected_periods:
                filtered_data['Period'] = self.get_period_from_date(filtered_data['Reject datetime'])
                filtered_data = filtered_data[filtered_data['Period'].isin(selected_periods)]
        
        # Apply line filter
        if hasattr(self, 'dimensional_line_checkboxes') and self.dimensional_line_checkboxes:
            selected_lines = [line for line, checkbox in self.dimensional_line_checkboxes.items() if checkbox.isChecked()]
            if selected_lines:
                filtered_data = filtered_data[filtered_data['Source'].isin(selected_lines)]
        
        # Define dimensional keywords
        dimensional_keywords = ['dimension', 'size', 'measurement', 'weight', 'position', 'height', 'width', 'length', 'tolerance', 'maximum']
        
        # Filter for dimensional rejects
        dimensional_data = filtered_data[
            filtered_data['Reject reason'].str.contains('|'.join(dimensional_keywords), case=False, na=False)
        ]
        
        if len(dimensional_data) > 0:
            # Plot 1: Dimensional rejection reasons breakdown
            reason_counts = dimensional_data['Reject reason'].value_counts().head(10)
            
            bars = ax1.barh(range(len(reason_counts)), reason_counts.values, color='#e74c3c', edgecolor='white', linewidth=1)
            ax1.set_yticks(range(len(reason_counts)))
            ax1.set_yticklabels([reason[:30] + '...' if len(reason) > 30 else reason for reason in reason_counts.index], color='white')
            ax1.set_xlabel('Number of Rejections', color='white', fontsize=12)
            ax1.set_title('Dimensional Rejection Reasons', color='white', fontweight='bold', fontsize=16)
            ax1.tick_params(axis='x', colors='white')
            ax1.grid(True, alpha=0.3, color='white', axis='x')
            
            # Add value labels on horizontal bars
            for i, bar in enumerate(bars):
                width = bar.get_width()
                ax1.text(width + 0.1, bar.get_y() + bar.get_height()/2.,
                        f'{int(width)}', ha='left', va='center', color='white', fontweight='bold',
                        bbox=dict(boxstyle="round,pad=0.3", facecolor='black', alpha=0.8))
            
            # Plot 2: Dimensional rejects by production line
            line_counts = dimensional_data['Source'].value_counts()
            
            bars = ax2.bar(line_counts.index, line_counts.values, color='#f39c12', edgecolor='white', linewidth=1)
            ax2.set_xlabel('Production Line', color='white', fontsize=12)
            ax2.set_ylabel('Number of Rejections', color='white', fontsize=12)
            ax2.set_title('Dimensional Rejects by Production Line', color='white', fontweight='bold', fontsize=16)
            ax2.tick_params(axis='x', colors='white', rotation=45)
            ax2.tick_params(axis='y', colors='white')
            ax2.grid(True, alpha=0.3, color='white', axis='y')
            
            # Add value labels on bars
            for bar in bars:
                height = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(height)}', ha='center', va='bottom', color='white', fontweight='bold',
                        bbox=dict(boxstyle="round,pad=0.3", facecolor='black', alpha=0.8))
        else:
            # No dimensional data
            ax1.text(0.5, 0.5, 'No Dimensional Rejects Found', ha='center', va='center', 
                    transform=ax1.transAxes, color='white', fontsize=16)
            ax2.text(0.5, 0.5, 'No Dimensional Rejects Found', ha='center', va='center', 
                    transform=ax2.transAxes, color='white', fontsize=16)
        
        # Style all axes
        for ax in [ax1, ax2]:
            for spine in ax.spines.values():
                spine.set_color('white')
                
        self.dimensional_figure.tight_layout()
        self.dimensional_canvas.draw()
        
    def update_tag_tracking_rejects(self):
        """Update tag/tracking rejects analysis"""
        if self.filtered_data is None:
            return
            
        # Clear existing plots
        self.tag_tracking_figure.clear()
        
        # Create subplots
        ax1 = self.tag_tracking_figure.add_subplot(1, 2, 1)
        ax2 = self.tag_tracking_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.tag_tracking_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Use the globally filtered data
        filtered_data = self.filtered_data.copy()
        
        # Apply period filter
        if hasattr(self, 'tag_tracking_period_checkboxes') and self.tag_tracking_period_checkboxes:
            selected_periods = [period for period, checkbox in self.tag_tracking_period_checkboxes.items() if checkbox.isChecked()]
            if selected_periods:
                filtered_data['Period'] = self.get_period_from_date(filtered_data['Reject datetime'])
                filtered_data = filtered_data[filtered_data['Period'].isin(selected_periods)]
        
        # Apply line filter
        if hasattr(self, 'tag_tracking_line_checkboxes') and self.tag_tracking_line_checkboxes:
            selected_lines = [line for line, checkbox in self.tag_tracking_line_checkboxes.items() if checkbox.isChecked()]
            if selected_lines:
                filtered_data = filtered_data[filtered_data['Source'].isin(selected_lines)]
        
        # Define tag/tracking keywords
        tag_tracking_keywords = ['tag', 'label', 'lpn', 'barcode', 'duplicate', 'unit data not found', 'tracking', 'expected', 'exist']
        
        # Filter for tag/tracking rejects
        tag_tracking_data = filtered_data[
            filtered_data['Reject reason'].str.contains('|'.join(tag_tracking_keywords), case=False, na=False)
        ]
        
        if len(tag_tracking_data) > 0:
            # Plot 1: Tag/Tracking rejection reasons breakdown
            reason_counts = tag_tracking_data['Reject reason'].value_counts().head(10)
            
            bars = ax1.barh(range(len(reason_counts)), reason_counts.values, color='#3498db', edgecolor='white', linewidth=1)
            ax1.set_yticks(range(len(reason_counts)))
            ax1.set_yticklabels([reason[:30] + '...' if len(reason) > 30 else reason for reason in reason_counts.index], color='white')
            ax1.set_xlabel('Number of Rejections', color='white', fontsize=12)
            ax1.set_title('Tag/Tracking Rejection Reasons', color='white', fontweight='bold', fontsize=16)
            ax1.tick_params(axis='x', colors='white')
            ax1.grid(True, alpha=0.3, color='white', axis='x')
            
            # Add value labels on horizontal bars
            for i, bar in enumerate(bars):
                width = bar.get_width()
                ax1.text(width + 0.1, bar.get_y() + bar.get_height()/2.,
                        f'{int(width)}', ha='left', va='center', color='white', fontweight='bold',
                        bbox=dict(boxstyle="round,pad=0.3", facecolor='black', alpha=0.8))
            
            # Plot 2: Tag/Tracking rejects by production line
            line_counts = tag_tracking_data['Source'].value_counts()
            
            bars = ax2.bar(line_counts.index, line_counts.values, color='#9b59b6', edgecolor='white', linewidth=1)
            ax2.set_xlabel('Production Line', color='white', fontsize=12)
            ax2.set_ylabel('Number of Rejections', color='white', fontsize=12)
            ax2.set_title('Tag/Tracking Rejects by Production Line', color='white', fontweight='bold', fontsize=16)
            ax2.tick_params(axis='x', colors='white', rotation=45)
            ax2.tick_params(axis='y', colors='white')
            ax2.grid(True, alpha=0.3, color='white', axis='y')
            
            # Add value labels on bars
            for bar in bars:
                height = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(height)}', ha='center', va='bottom', color='white', fontweight='bold',
                        bbox=dict(boxstyle="round,pad=0.3", facecolor='black', alpha=0.8))
        else:
            # No tag/tracking data
            ax1.text(0.5, 0.5, 'No Tag/Tracking Rejects Found', ha='center', va='center', 
                    transform=ax1.transAxes, color='white', fontsize=16)
            ax2.text(0.5, 0.5, 'No Tag/Tracking Rejects Found', ha='center', va='center', 
                    transform=ax2.transAxes, color='white', fontsize=16)
        
        # Style all axes
        for ax in [ax1, ax2]:
            for spine in ax.spines.values():
                spine.set_color('white')
                
        self.tag_tracking_figure.tight_layout()
        self.tag_tracking_canvas.draw()
        
    def update_time_analysis(self):
        """Update time-of-day analysis"""
        if self.filtered_data is None:
            return
            
        # Clear existing plots
        self.time_figure.clear()
        
        # Create subplots
        ax1 = self.time_figure.add_subplot(1, 2, 1)
        ax2 = self.time_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.time_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Use the globally filtered data
        filtered_data = self.filtered_data.copy()
        
        # Apply period filter
        if hasattr(self, 'time_period_checkboxes') and self.time_period_checkboxes:
            selected_periods = [period for period, checkbox in self.time_period_checkboxes.items() if checkbox.isChecked()]
            if selected_periods:
                filtered_data['Period'] = self.get_period_from_date(filtered_data['Reject datetime'])
                filtered_data = filtered_data[filtered_data['Period'].isin(selected_periods)]
        
        # Apply line filter
        if hasattr(self, 'time_line_checkboxes') and self.time_line_checkboxes:
            selected_lines = [line for line, checkbox in self.time_line_checkboxes.items() if checkbox.isChecked()]
            if selected_lines:
                filtered_data = filtered_data[filtered_data['Source'].isin(selected_lines)]
        
        # Extract hour from datetime
        if 'Reject datetime' in filtered_data.columns:
            filtered_data['Hour'] = filtered_data['Reject datetime'].dt.hour
            filtered_data['DayOfWeek'] = filtered_data['Reject datetime'].dt.day_name()
            
            # Plot 1: Rejections by hour
            hourly_counts = filtered_data['Hour'].value_counts().sort_index()
            bars = ax1.bar(hourly_counts.index, hourly_counts.values, color='#e74c3c', 
                          edgecolor='white', linewidth=1)
            ax1.set_title('Rejections by Hour of Day', color='white', fontweight='bold', fontsize=12)
            ax1.set_xlabel('Hour', color='white')
            ax1.set_ylabel('Number of Rejections', color='white')
            ax1.tick_params(axis='x', colors='white')
            ax1.tick_params(axis='y', colors='white')
            ax1.grid(True, alpha=0.3, color='white', axis='y')

            worst_hour = hourly_counts.idxmax()
            best_hour = hourly_counts.idxmin()
            
            comparison_labels = [f'Worst Hour\n({worst_hour}:00)'+'\n', f'Best Hour\n({best_hour}:00)'+'\n']
            colors = ['#e74c3c', '#2ecc71']
            
            max_height = bars[0].get_height()
            min_height = bars[0].get_height()
            for bar in bars:
                if bar.get_height() > max_height:
                    max_height = bar.get_height()
                if bar.get_height() < min_height:
                    min_height = bar.get_height()

            # Add value labels on top of bars
            for bar in bars:
                height = bar.get_height()
                if height == max_height:
                    color = colors[0]
                    ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        comparison_labels[0] + f'{int(height)}', ha='center', va='bottom', color='white', fontweight='bold', fontsize=8,
                        bbox=dict(boxstyle="round,pad=0.2", facecolor=color, alpha=0.8))
                elif height == min_height:
                    color = colors[1]
                    ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        comparison_labels[1] + f'{int(height)}', ha='center', va='bottom', color='white', fontweight='bold', fontsize=8,
                        bbox=dict(boxstyle="round,pad=0.2", facecolor='green', alpha=0.8))
                else:
                    ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                            f'{int(height)}', ha='center', va='bottom', color='white', fontweight='bold', fontsize=8,
                            bbox=dict(boxstyle="round,pad=0.2", facecolor='black', alpha=0.8))
                
            
            # Plot 2: Rejections by day of week
            dow_counts = filtered_data['DayOfWeek'].value_counts()
            dow_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
            dow_counts = dow_counts.reindex([d for d in dow_order if d in dow_counts.index])
            
            bars = ax2.bar(dow_counts.index, dow_counts.values, color='#3498db', 
                          edgecolor='white', linewidth=1)
            ax2.set_title('Rejections by Day of Week', color='white', fontweight='bold', fontsize=12)
            ax2.set_xlabel('Day of Week', color='white')
            ax2.set_ylabel('Number of Rejections', color='white')
            ax2.tick_params(axis='x', colors='white', rotation=45)
            ax2.tick_params(axis='y', colors='white')
            ax2.grid(True, alpha=0.3, color='white', axis='y')
            
            # Add value labels on bars
            for bar in bars:
                height = bar.get_height()
                ax2.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(height)}', ha='center', va='bottom', color='white', fontweight='bold', fontsize=8,
                        bbox=dict(boxstyle="round,pad=0.2", facecolor='black', alpha=0.8))

        # Style all axes
        for ax in [ax1, ax2]:
            for spine in ax.spines.values():
                spine.set_color('white')
                
        self.time_figure.tight_layout()
        self.time_canvas.draw()
        


    def update_page_from_prefix(self, prefix):
        """Update the appropriate page based on the prefix"""
        if prefix == "production":
            self.update_production_analysis()
        elif prefix == "production_lines":
            self.update_production_lines()
        elif prefix == "advanced_tracking":
            self.update_advanced_tracking()
        elif prefix == "dimensional":
            self.update_dimensional_rejects()
        elif prefix == "tag_tracking":
            self.update_tag_tracking_rejects()
        elif prefix == "time":
            self.update_time_analysis()
        elif prefix == "sku":
            self.update_sku_analysis()
        elif prefix == "rejection":
            self.update_rejection_rate_analysis()

    def update_sku_analysis(self):
        """Update Product analysis"""
        if self.filtered_data is None:
            return
            
        # Clear existing plots
        self.sku_figure.clear()
        
        # Create subplots
        ax1 = self.sku_figure.add_subplot(1, 1, 1)
        # ax2 = self.sku_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.sku_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Use the globally filtered data
        filtered_data = self.filtered_data.copy()
        
        # Plot 1: Top Products by rejection count (with full product descriptions)
        num_products = 20
        total_rejections = filtered_data['Sku'].value_counts().sum()
        top_rejections = filtered_data['Sku'].value_counts().head(num_products).sum()
        percentage_top = (top_rejections / total_rejections) * 100
        if 'Sku' in filtered_data.columns:
            sku_counts = filtered_data['Sku'].value_counts().head(num_products)
            
            bars = ax1.bar(range(len(sku_counts)), sku_counts.values, color='#e74c3c', edgecolor='white', linewidth=1)
            ax1.set_title('Top Products by Rejection Count (Top ' + str(num_products) + ' account for ' + f'{percentage_top:.3f}' + '% of all rejections)', color='white', fontsize=12, fontweight='bold')
            ax1.set_xlabel('Product', color='white')
            ax1.set_ylabel('Number of Rejections', color='white')
            ax1.set_xticks(range(len(sku_counts)))
            ax1.set_xticklabels(sku_counts.index, rotation=-90, color='white')
            ax1.tick_params(colors='white')
            ax1.grid(True, alpha=0.3, color='white')
            
            # Add value labels on bars
            for i, (bar, value) in enumerate(zip(bars, sku_counts.values)):
                ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5, 
                        str(value), ha='center', va='bottom', color='white', fontweight='bold',
                        bbox=dict(boxstyle="round,pad=0.3", facecolor='black', alpha=0.8))
            
            # Set dark theme for subplot
            ax1.set_facecolor(self.theme.get_color('card_bg'))
            for spine in ax1.spines.values():
                spine.set_color('white')
        
        # # Plot 2: Pie chart of all Products
        # if 'Sku' in filtered_data.columns:
        #     sku_counts = filtered_data['Sku'].value_counts()
            
        #     # Get top 8 SKUs and group the rest as "Other"
        #     top_skus = sku_counts.head(8)
        #     other_count = sku_counts.iloc[8:].sum() if len(sku_counts) > 8 else 0
            
        #     if other_count > 0:
        #         pie_data = list(top_skus.values) + [other_count]
        #         pie_labels = list(top_skus.index) + ['Other']
        #     else:
        #         pie_data = list(top_skus.values)
        #         pie_labels = list(top_skus.index)
            
        #     colors = ['#e74c3c', '#3498db', '#f39c12', '#2ecc71', '#9b59b6', '#1abc9c', '#34495e', '#e67e22', '#95a5a6']
            
        #     wedges, texts, autotexts = ax2.pie(pie_data, labels=pie_labels, autopct='%1.1f%%', 
        #                                       colors=colors[:len(pie_data)], startangle=90)
        #     ax2.set_title('Product Distribution', color='white', fontsize=12, fontweight='bold')
            
        #     # Set text colors to white
        #     for text in texts:
        #         text.set_color('white')
        #     for autotext in autotexts:
        #         autotext.set_color('white')
        #         autotext.set_fontweight('bold')
            
        #     # Set dark theme for subplot
        #     ax2.set_facecolor(self.theme.get_color('card_bg'))
        
        # # Plot 3: Product rejection rate by production line (with full product descriptions)
        # if 'Sku' in filtered_data.columns and 'Source' in filtered_data.columns:
        #     # Get top 5 SKUs
        #     top_skus = filtered_data['Sku'].value_counts().head(5).index
            
        #     # Create heatmap-style data
        #     line_sku_data = filtered_data.groupby(['Source', 'Sku']).size().unstack(fill_value=0)
        #     line_sku_data = line_sku_data[top_skus]  # Only show top SKUs
            
        #     # Create bar chart showing SKU rejections by line
        #     x = np.arange(len(line_sku_data.index))
        #     width = 0.15
            
        #     for i, sku in enumerate(top_skus):
        #         ax3.bar(x + i * width, line_sku_data[sku], width, 
        #                label=sku, edgecolor='white', linewidth=1)
            
        #     ax3.set_title('Product Rejections by Production Line', color='white', fontsize=12, fontweight='bold')
        #     ax3.set_xlabel('Production Line', color='white')
        #     ax3.set_ylabel('Number of Rejections', color='white')
        #     ax3.set_xticks(x + width * 2)
        #     ax3.set_xticklabels(line_sku_data.index, rotation=45, color='white')
        #     ax3.tick_params(colors='white')
        #     ax3.legend(fontsize=8)
        #     ax3.grid(True, alpha=0.3, color='white')
            
        #     # Set dark theme for subplot
        #     ax3.set_facecolor(self.theme.get_color('card_bg'))
        #     for spine in ax3.spines.values():
        #         spine.set_color('white')
        
        # # Plot 4: Product rejection trends by period (with full product descriptions)
        # if 'Sku' in filtered_data.columns and 'Period' in filtered_data.columns:
        #     # Get top 3 SKUs for trend analysis
        #     top_skus = filtered_data['Sku'].value_counts().head(3).index
            
        #     # Create period series data
        #     for i, sku in enumerate(top_skus):
        #         sku_data = filtered_data[filtered_data['Sku'] == sku]
        #         period_counts = sku_data.groupby('Period').size()
                
        #         # Sort periods numerically
        #         period_counts = period_counts.sort_index()
                
        #         ax4.plot(period_counts.index, period_counts.values, 
        #                 marker='o', linewidth=2, label=sku, markersize=4)
            
        #     ax4.set_title('Product Rejection Trends by Period', color='white', fontsize=12, fontweight='bold')
        #     ax4.set_xlabel('Period', color='white')
        #     ax4.set_ylabel('Rejections per Period', color='white')
        #     ax4.tick_params(colors='white')
        #     ax4.legend(fontsize=8)
        #     ax4.grid(True, alpha=0.3, color='white')
            
        #     # Set dark theme for subplot
        #     ax4.set_facecolor(self.theme.get_color('card_bg'))
        #     for spine in ax4.spines.values():
        #         spine.set_color('white')
        
        self.sku_figure.tight_layout()
        self.sku_canvas.draw()
    
    def update_rejection_rate_analysis(self):
        """Update rejection rate analysis"""
        if self.filtered_data is None:
            return
            
        # Clear existing plots
        self.rejection_figure.clear()
        
        # Create 2 subplots - overall rate and per production line
        ax1 = self.rejection_figure.add_subplot(1, 2, 1)
        ax2 = self.rejection_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.rejection_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Use the globally filtered data
        filtered_data = self.filtered_data.copy()
        
        # Production data - EXACT numbers Oscar provided
        production_totals = {
            'Aquafina/Propel 1': 36720,  # Aquafina/Propel 1
            'Aquafina/Propel 2': 34407,  # Aquafina/Propel 2
            'Can Line 3': 54269,         # Can line 3
            'Can Line 4': 70981,         # Can line 4
            'Bottle Line 5': 62106,      # Bottle line 5
            'Bottle Line 6': 53596       # Bottle line 6
        }
        
        # Use the EXACT same consolidation method as the production lines tab
        def consolidate_line(source):
            """Consolidate production lines based on EOL numbers - EXACT COPY from production lines tab"""
            if pd.isna(source):
                return 'Unknown'
            
            source_str = str(source).upper()
            
            # Check for IBC first (Inbound Conveyor by Car Wash)
            if 'IBC' in source_str:
                return 'IBC (Inbound Conveyor)'
            # Check for EOL patterns
            elif 'EOL01' in source_str or 'EOL_01' in source_str:
                return 'Aquafina/Propel 1'
            elif 'EOL02' in source_str or 'EOL_02' in source_str:
                return 'Aquafina/Propel 2'
            elif 'EOL03' in source_str or 'EOL_03' in source_str:
                return 'Can Line 3'
            elif 'EOL04' in source_str or 'EOL_04' in source_str:
                return 'Can Line 4'
            elif 'EOL05' in source_str or 'EOL_05' in source_str:
                return 'Bottle Line 5'
            elif 'EOL06' in source_str or 'EOL_06' in source_str:
                return 'Bottle Line 6'
            # Check for simple number patterns (but not if it's IBC)
            elif source_str == '1' or source_str.endswith('1'):
                return 'Aquafina/Propel 1'
            elif source_str == '2' or source_str.endswith('2'):
                return 'Aquafina/Propel 2'
            elif source_str == '3' or source_str.endswith('3'):
                return 'Can Line 3'
            elif source_str == '4' or source_str.endswith('4'):
                return 'Can Line 4'
            elif source_str == '5' or source_str.endswith('5'):
                return 'Bottle Line 5'
            elif source_str == '6' or source_str.endswith('6'):
                return 'Bottle Line 6'
            else:
                return source_str  # Keep original if no pattern matches
        
        # Calculate rejection rates by line
        if 'Source' in filtered_data.columns:
            # Use the EXACT same consolidation method as production lines tab
            filtered_data['Consolidated_Line'] = filtered_data['Source'].apply(consolidate_line)
            consolidated_rejections = filtered_data['Consolidated_Line'].value_counts().to_dict()
            
            # Calculate rejection rates for all production lines (including those with 0 rejections)
            rejection_rates = {}
            rejection_details = {}  # Store raw numbers for display
            
            # Initialize all production lines
            for line in production_totals:
                rejection_count = consolidated_rejections.get(line, 0)  # 0 if no rejections
                production_total = production_totals[line]
                rejection_rate = (rejection_count / production_total) * 100
                rejection_rates[line] = rejection_rate
                rejection_details[line] = {
                    'rejections': rejection_count,
                    'production': production_total,
                    'rate': rejection_rate
                }
            
            # Plot 1: Overall rejection rate (single bar chart)
            total_rejections = len(filtered_data)
            total_production = sum(production_totals.values())
            overall_rate = (total_rejections / total_production) * 100
            
            categories = ['Overall Rejection Rate']
            values = [overall_rate]
            colors = ['#e74c3c']
            
            bars = ax1.bar(categories, values, color=colors, edgecolor='white', linewidth=1)
            ax1.set_title(f'Overall Rejection Rate: {overall_rate:.2f}%', color='white', fontsize=16, fontweight='bold')
            ax1.set_ylabel('Rejection Rate (%)', color='white')
            ax1.tick_params(colors='white')
            ax1.grid(True, alpha=0.3, color='white')
            
            # Add value label on bar with raw numbers
            for bar, value in zip(bars, values):
                label_text = f'{value:.2f}%\n({total_rejections:,}/{total_production:,})'
                ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
                        label_text, ha='center', va='bottom', color='white', fontweight='bold',
                        bbox=dict(boxstyle='round,pad=0.3', facecolor='black', alpha=0.7), fontsize=12)
            
            # Set dark theme for subplot
            ax1.set_facecolor(self.theme.get_color('card_bg'))
            for spine in ax1.spines.values():
                spine.set_color('white')
            
            # Plot 2: Rejection rates by production line
            if rejection_rates:
                lines = list(rejection_rates.keys())
                rates = list(rejection_rates.values())
                
                bars = ax2.bar(range(len(lines)), rates, color='#e74c3c', edgecolor='white', linewidth=1)
                ax2.set_title('Rejection Rate by Production Line', color='white', fontsize=16, fontweight='bold')
                ax2.set_xlabel('Production Line', color='white')
                ax2.set_ylabel('Rejection Rate (%)', color='white')
                ax2.set_xticks(range(len(lines)))
                ax2.set_xticklabels(lines, rotation=45, color='white')
                ax2.tick_params(colors='white')
                ax2.grid(True, alpha=0.3, color='white')
                
                # Add value labels on bars with raw numbers
                for i, (bar, value) in enumerate(zip(bars, rates)):
                    line = lines[i]
                    details = rejection_details[line]
                    label_text = f'{value:.2f}%\n({details["rejections"]:,}/{details["production"]:,})'
                    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max(rates)*0.01,
                            label_text, ha='center', va='bottom', color='white', fontweight='bold',
                            bbox=dict(boxstyle='round,pad=0.3', facecolor='black', alpha=0.7), fontsize=10)
                
                # Set dark theme for subplot
                ax2.set_facecolor(self.theme.get_color('card_bg'))
                for spine in ax2.spines.values():
                    spine.set_color('white')
            else:
                # If no rejection rates calculated, show a message
                ax2.text(0.5, 0.5, 'No production line data available\nfor rejection rate calculation', 
                        ha='center', va='center', transform=ax2.transAxes, 
                        color='white', fontsize=14, fontweight='bold')
                ax2.set_title('Rejection Rate by Production Line', color='white', fontsize=16, fontweight='bold')
                ax2.set_facecolor(self.theme.get_color('card_bg'))
        
        self.rejection_figure.tight_layout()
        self.rejection_canvas.draw()
        
    def update_filters(self):
        if self.current_data is None:
            return
            
        # Calculate periods for the data
        self.current_data['Period'] = self.get_period_from_date(self.current_data['Reject datetime'])
        available_periods = set(self.current_data['Period'].unique())
        available_lines = sorted(self.current_data['Source'].unique())
        available_skus = sorted(self.current_data['Sku'].unique())
        
        # Period definitions
        period_definitions = {
            'Period 1': 'Dec 29, 2024 - Jan 25, 2025',
            'Period 2': 'Jan 26, 2025 - Feb 22, 2025',
            'Period 3': 'Feb 23, 2025 - Mar 22, 2025',
            'Period 4': 'Mar 23, 2025 - Apr 19, 2025',
            'Period 5': 'Apr 20, 2025 - May 17, 2025',
            'Period 6': 'May 18, 2025 - Jun 14, 2025',
            'Period 7': 'Jun 15, 2025 - Jul 12, 2025',
            'Period 8': 'Jul 13, 2025 - Aug 9, 2025',
            'Period 9': 'Aug 10, 2025 - Sep 6, 2025',
            'Period 10': 'Sep 7, 2025 - Oct 4, 2025',
            'Period 11': 'Oct 5, 2025 - Nov 1, 2025',
            'Period 12': 'Nov 2, 2025 - Nov 29, 2025',
            'Period 13': 'Nov 30, 2025 - Dec 27, 2025'
        }
        
        # Update global period checkboxes
        for checkbox in self.global_period_checkboxes.values():
            checkbox.setParent(None)
        self.global_period_checkboxes.clear()
        
        # Create periods in correct order (1-13)
        for i in range(1, 14):
            period = f"Period {i}"
            if period in period_definitions:
                date_range = period_definitions[period]
                
                # Create container for checkbox and solo button
                period_container = QWidget()
                period_container_layout = QHBoxLayout()
                period_container_layout.setContentsMargins(0, 0, 0, 0)
                
                checkbox = QCheckBox(f"{period}: {date_range}")
                if period in available_periods:
                    checkbox.setChecked(True)
                    checkbox.setEnabled(True)
                    checkbox.setStyleSheet("color: white;")
                else:
                    checkbox.setChecked(False)
                    checkbox.setEnabled(False)
                    checkbox.setStyleSheet("color: gray;")
                checkbox.stateChanged.connect(self.update_all_tabs)
                
                # Add solo button
                solo_btn = QPushButton("Solo")
                solo_btn.setMaximumWidth(50)
                solo_btn.setStyleSheet(f"""
                    QPushButton {{
                        background-color: #f39c12;
                        color: white;
                        border: none;
                        padding: 2px 5px;
                        border-radius: 2px;
                        font-size: 10px;
                        font-weight: bold;
                    }}
                    QPushButton:hover {{
                        background-color: #e67e22;
                    }}
                """)
                solo_btn.clicked.connect(lambda checked, p=period: self.solo_period(p))
                
                period_container_layout.addWidget(checkbox)
                period_container_layout.addWidget(solo_btn)
                period_container_layout.addStretch()
                period_container.setLayout(period_container_layout)
                
                self.global_period_checkboxes[period] = checkbox
                self.global_period_widget_layout.addWidget(period_container)
        
        # Update global line checkboxes
        for checkbox in self.global_line_checkboxes.values():
            checkbox.setParent(None)
        self.global_line_checkboxes.clear()
        
        for line in available_lines:
            if pd.notna(line):
                # Create container for checkbox and solo button
                line_container = QWidget()
                line_container_layout = QHBoxLayout()
                line_container_layout.setContentsMargins(0, 0, 0, 0)
                
                checkbox = QCheckBox(str(line))
                # Set IBC01_SHAPE to be unchecked by default
                if str(line) == "IBC01_SHAPE":
                    checkbox.setChecked(False)
                else:
                    checkbox.setChecked(True)
                checkbox.stateChanged.connect(self.update_all_tabs)
                
                # Add solo button
                solo_btn = QPushButton("Solo")
                solo_btn.setMaximumWidth(50)
                solo_btn.setStyleSheet(f"""
                    QPushButton {{
                        background-color: #f39c12;
                        color: white;
                        border: none;
                        padding: 2px 5px;
                        border-radius: 2px;
                        font-size: 10px;
                        font-weight: bold;
                    }}
                    QPushButton:hover {{
                        background-color: #e67e22;
                    }}
                """)
                solo_btn.clicked.connect(lambda checked, l=line: self.solo_line(l))
                
                line_container_layout.addWidget(checkbox)
                line_container_layout.addWidget(solo_btn)
                line_container_layout.addStretch()
                line_container.setLayout(line_container_layout)
                
                self.global_line_checkboxes[line] = checkbox
                self.global_line_widget_layout.addWidget(line_container)
        
        # Update global SKU checkboxes
        for checkbox in self.global_sku_checkboxes.values():
            checkbox.setParent(None)
        self.global_sku_checkboxes.clear()
        
        for sku in available_skus:
            if pd.notna(sku):
                # Create container for checkbox and solo button
                sku_container = QWidget()
                sku_container_layout = QHBoxLayout()
                sku_container_layout.setContentsMargins(0, 0, 0, 0)
                
                checkbox = QCheckBox(str(sku))
                checkbox.setChecked(True)
                checkbox.stateChanged.connect(self.update_all_tabs)
                
                # Add solo button
                solo_btn = QPushButton("Solo")
                solo_btn.setMaximumWidth(50)
                solo_btn.setStyleSheet(f"""
                    QPushButton {{
                        background-color: #f39c12;
                        color: white;
                        border: none;
                        padding: 2px 5px;
                        border-radius: 2px;
                        font-size: 10px;
                        font-weight: bold;
                    }}
                    QPushButton:hover {{
                        background-color: #e67e22;
                    }}
                """)
                solo_btn.clicked.connect(lambda checked, s=sku: self.solo_sku(s))
                
                sku_container_layout.addWidget(checkbox)
                sku_container_layout.addWidget(solo_btn)
                sku_container_layout.addStretch()
                sku_container.setLayout(sku_container_layout)
                
                self.global_sku_checkboxes[sku] = checkbox
                self.global_sku_widget_layout.addWidget(sku_container)
        
        # Initialize filtered data with all data selected
        self.apply_filters()
        
    def select_all_periods(self):
        """Select all period checkboxes"""
        for checkbox in self.global_period_checkboxes.values():
            if checkbox.isEnabled():
                checkbox.setChecked(True)
        self.update_all_tabs()
        
    def reset_periods(self):
        """Reset all period checkboxes to unchecked"""
        for checkbox in self.global_period_checkboxes.values():
            checkbox.setChecked(False)
        self.update_all_tabs()
        
    def select_all_lines(self):
        """Select all production line checkboxes"""
        for checkbox in self.global_line_checkboxes.values():
            checkbox.setChecked(True)
        self.update_all_tabs()
        
    def reset_lines(self):
        """Reset all production line checkboxes to unchecked"""
        for checkbox in self.global_line_checkboxes.values():
            checkbox.setChecked(False)
        self.update_all_tabs()
        
    def select_all_skus(self):
        """Select all SKU checkboxes"""
        for checkbox in self.global_sku_checkboxes.values():
            checkbox.setChecked(True)
        self.update_all_tabs()
        
    def reset_skus(self):
        """Reset all SKU checkboxes to unchecked"""
        for checkbox in self.global_sku_checkboxes.values():
            checkbox.setChecked(False)
        self.update_all_tabs()
        
    def solo_period(self, period):
        """Solo a specific period (uncheck all others)"""
        for p, checkbox in self.global_period_checkboxes.items():
            checkbox.setChecked(p == period)
        self.update_all_tabs()
        
    def solo_line(self, line):
        """Solo a specific production line (uncheck all others)"""
        for l, checkbox in self.global_line_checkboxes.items():
            checkbox.setChecked(l == line)
        self.update_all_tabs()
        
    def solo_sku(self, sku):
        """Solo a specific SKU (uncheck all others)"""
        for s, checkbox in self.global_sku_checkboxes.items():
            checkbox.setChecked(s == sku)
        self.update_all_tabs()
        
    def update_all_tabs(self):
        """Update all tabs when global filters change"""
        # Update filtered data
        self.apply_filters()
        
        # Update all tabs
        self.update_dashboard()
        self.update_trends()
        self.update_production_analysis()
        self.update_production_lines()
        self.update_advanced_tracking()
        self.update_dimensional_rejects()
        self.update_tag_tracking_rejects()
        self.update_time_analysis()
        self.update_sku_analysis()
            
    def apply_filters(self):
        if self.current_data is None:
            return
            
        # Apply global filters
        filtered_data = self.current_data.copy()
        
        # Apply period filter
        if 'Reject datetime' in filtered_data.columns and hasattr(self, 'global_period_checkboxes'):
            # Get selected periods
            selected_periods = [period for period, checkbox in self.global_period_checkboxes.items() if checkbox.isChecked()]
            
            if selected_periods:
                # Calculate periods for all data
                filtered_data['Period'] = self.get_period_from_date(filtered_data['Reject datetime'])
                # Filter to only selected periods
                filtered_data = filtered_data[filtered_data['Period'].isin(selected_periods)]
        
        # Apply line filter (E80 format uses 'Source') - multi-select
        if hasattr(self, 'global_line_checkboxes'):
            selected_lines = [line for line, checkbox in self.global_line_checkboxes.items() if checkbox.isChecked()]
            if selected_lines:
                filtered_data = filtered_data[filtered_data['Source'].isin(selected_lines)]
        
        # Apply SKU filter - multi-select
        if hasattr(self, 'global_sku_checkboxes'):
            selected_skus = [sku for sku, checkbox in self.global_sku_checkboxes.items() if checkbox.isChecked()]
            if selected_skus:
                filtered_data = filtered_data[filtered_data['Sku'].isin(selected_skus)]
            
            
        # Store filtered data
        self.filtered_data = filtered_data
        print(f"Filtered data: {len(filtered_data)} rows (original: {len(self.current_data)} rows)")
            
        

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Set application properties
    app.setApplicationName("E80 Rejected Units Analyzer")
    app.setApplicationVersion("1.0")
    app.setOrganizationName("PepsiCo - Stone Mountain")
    
    window = RejectedUnitsAnalyzer()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
