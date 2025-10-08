import sys
import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.patheffects
import seaborn as sns
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget, 
                           QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, 
                           QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
                           QScrollArea, QFrame, QTextEdit, QMessageBox, QHeaderView,
                           QComboBox, QDateEdit, QCheckBox, QGroupBox, QSpinBox,
                           QProgressBar, QSplitter)
from PyQt5.QtCore import Qt, pyqtSignal, QDate, QThread, pyqtSlot
from PyQt5.QtGui import QFont, QIcon, QColor, QPalette
import warnings
warnings.filterwarnings('ignore')

class ModernTheme:
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
        self.analysis_results = None
        self.current_file = None
        
        self.setup_ui()
        self.apply_theme()
        
    def setup_ui(self):
        self.setWindowTitle("E80 Rejected Units Analyzer - Root Cause Analysis")
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
        
        # Tab widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabPosition(QTabWidget.North)
        main_layout.addWidget(self.tab_widget)
        
        # Create initial tabs
        self.create_upload_tab()
        self.create_dashboard_tab()
        self.create_analysis_tab()
        self.create_trends_tab()
        self.create_production_analysis_tab()
        self.create_advanced_tracking_tab()
        self.create_dimensional_rejects_tab()
        self.create_tag_tracking_rejects_tab()
        self.create_time_analysis_tab()
        self.create_reports_tab()
        
        central_widget.setLayout(main_layout)
        
    def create_header(self, main_layout):
        header = QFrame()
        header.setFixedHeight(80)
        
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(40, 20, 40, 20)
        
        # Title
        title = QLabel("E80 Rejected Units Analyzer")
        title.setFont(QFont("Segoe UI", 24, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        
        # Subtitle
        subtitle = QLabel("Root Cause Analysis & Production Optimization")
        subtitle.setFont(QFont("Segoe UI", 12))
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet(f"color: {self.theme.get_color('text_secondary')};")
        
        # Remove theme button since we only use dark mode
        
        header_layout.addStretch()
        header_layout.addWidget(title)
        header_layout.addStretch()
        
        header.setLayout(header_layout)
        main_layout.addWidget(header)
        
    # Removed theme toggle functions since we only use dark mode
        
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
        desc.setFont(QFont("Segoe UI", 12))
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
        self.dashboard_layout = QVBoxLayout()
        self.dashboard_layout.setContentsMargins(20, 20, 20, 20)
        
        # Metrics grid (bigger cards)
        self.metrics_widget = QWidget()
        self.metrics_layout = QGridLayout()
        self.metrics_widget.setLayout(self.metrics_layout)
        self.dashboard_layout.addWidget(self.metrics_widget)
        
        self.dashboard_tab.setLayout(self.dashboard_layout)
        self.tab_widget.addTab(self.dashboard_tab, "📊 Dashboard")
        
    def create_analysis_tab(self):
        tab = QWidget()
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
        self.analysis_table.setColumnCount(8)
        self.analysis_table.setHorizontalHeaderLabels([
            "Reject DateTime", "Source", "SKU", "Reject Reason", "LPN", 
            "Log Text", "Root Cause", "Quantity"
        ])
        
        header = self.analysis_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        self.analysis_table.verticalHeader().setVisible(False)
        
        layout.addWidget(self.analysis_table)
        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "🔍 Analysis")
        
    def create_trends_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Filters section
        filters_card = ModernCard("🔍 Trend Filters", self.theme)
        filters_layout = QHBoxLayout()
        
        # Period filter for trends
        period_group = QGroupBox("Period Selection")
        period_layout = QVBoxLayout()
        self.trend_period_checkboxes = {}
        self.trend_period_scroll = QScrollArea()
        self.trend_period_scroll.setWidgetResizable(True)
        self.trend_period_scroll.setMaximumHeight(100)
        trend_period_widget = QWidget()
        self.trend_period_widget_layout = QVBoxLayout()
        trend_period_widget.setLayout(self.trend_period_widget_layout)
        self.trend_period_scroll.setWidget(trend_period_widget)
        period_layout.addWidget(self.trend_period_scroll)
        period_group.setLayout(period_layout)
        
        # Line filter (multi-select for trends)
        line_group = QGroupBox("Production Lines")
        line_layout = QVBoxLayout()
        self.trend_line_checkboxes = {}
        self.trend_line_scroll = QScrollArea()
        self.trend_line_scroll.setWidgetResizable(True)
        self.trend_line_scroll.setMaximumHeight(100)
        trend_line_widget = QWidget()
        self.trend_line_widget_layout = QVBoxLayout()
        trend_line_widget.setLayout(self.trend_line_widget_layout)
        self.trend_line_scroll.setWidget(trend_line_widget)
        line_layout.addWidget(self.trend_line_scroll)
        line_group.setLayout(line_layout)
        
        filters_layout.addWidget(period_group)
        filters_layout.addWidget(line_group)
        filters_card.content_layout.addLayout(filters_layout)
        layout.addWidget(filters_card)
        
        # Charts section
        charts_card = ModernCard("📈 Period Trends", self.theme)
        
        # Create matplotlib figure for trends (only 2 charts now)
        self.trends_figure = Figure(figsize=(12, 6), dpi=100)
        self.trends_canvas = FigureCanvas(self.trends_figure)
        charts_card.content_layout.addWidget(self.trends_canvas)
        
        layout.addWidget(charts_card)
        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "📈 Trends")
        
    def create_production_analysis_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Production analysis
        production_card = ModernCard("🏭 Production Line & Product Analysis", self.theme)
        
        # Create matplotlib figure for production analysis
        self.production_figure = Figure(figsize=(12, 8), dpi=100)
        self.production_canvas = FigureCanvas(self.production_figure)
        production_card.content_layout.addWidget(self.production_canvas)
        
        layout.addWidget(production_card)
        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "🏭 Production Analysis")
        
    def create_advanced_tracking_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Category analysis
        category_card = ModernCard("🔍 Rejection Category Analysis", self.theme)
        
        # Create matplotlib figure for category analysis
        self.category_figure = Figure(figsize=(12, 8), dpi=100)
        self.category_canvas = FigureCanvas(self.category_figure)
        category_card.content_layout.addWidget(self.category_canvas)
        
        layout.addWidget(category_card)
        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "🔍 Advanced Tracking")
        
    def create_dimensional_rejects_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Dimensional rejects analysis
        dimensional_card = ModernCard("📏 Dimensional Rejects Analysis", self.theme)
        
        # Create matplotlib figure for dimensional analysis
        self.dimensional_figure = Figure(figsize=(12, 8), dpi=100)
        self.dimensional_canvas = FigureCanvas(self.dimensional_figure)
        dimensional_card.content_layout.addWidget(self.dimensional_canvas)
        
        layout.addWidget(dimensional_card)
        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "📏 Dimensional Rejects")
        
    def create_tag_tracking_rejects_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Tag/Tracking rejects analysis
        tag_tracking_card = ModernCard("🏷️ Tag/Tracking Rejects Analysis", self.theme)
        
        # Create matplotlib figure for tag/tracking analysis
        self.tag_tracking_figure = Figure(figsize=(12, 8), dpi=100)
        self.tag_tracking_canvas = FigureCanvas(self.tag_tracking_figure)
        tag_tracking_card.content_layout.addWidget(self.tag_tracking_canvas)
        
        layout.addWidget(tag_tracking_card)
        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "🏷️ Tag/Tracking Rejects")
        
    def create_time_analysis_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Time analysis
        time_card = ModernCard("⏰ Time-of-Day Analysis", self.theme)
        
        # Create matplotlib figure for time analysis
        self.time_figure = Figure(figsize=(12, 8), dpi=100)
        self.time_canvas = FigureCanvas(self.time_figure)
        time_card.content_layout.addWidget(self.time_canvas)
        
        layout.addWidget(time_card)
        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "⏰ Time Analysis")
        
    def create_reports_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Export options
        export_card = ModernCard("📋 Generate Reports", self.theme)
        
        # Report type selection
        report_layout = QHBoxLayout()
        
        self.summary_report_btn = QPushButton("Summary Report")
        self.summary_report_btn.clicked.connect(self.generate_summary_report)
        
        self.detailed_report_btn = QPushButton("Detailed Analysis")
        self.detailed_report_btn.clicked.connect(self.generate_detailed_report)
        
        self.trend_report_btn = QPushButton("Trend Report")
        self.trend_report_btn.clicked.connect(self.generate_trend_report)
        
        report_layout.addWidget(self.summary_report_btn)
        report_layout.addWidget(self.detailed_report_btn)
        report_layout.addWidget(self.trend_report_btn)
        
        export_card.content_layout.addLayout(report_layout)
        layout.addWidget(export_card)
        
        # Report preview
        preview_card = ModernCard("Report Preview", self.theme)
        self.report_preview = QTextEdit()
        self.report_preview.setFont(QFont("Consolas", 10))
        self.report_preview.setReadOnly(True)
        preview_card.content_layout.addWidget(self.report_preview)
        layout.addWidget(preview_card)
        
        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "📋 Reports")
        
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
        
        # Clear analysis table
        self.analysis_table.setRowCount(0)
        
        # Clear trends chart
        self.trends_figure.clear()
        self.trends_canvas.draw()
        
        # Clear production analysis chart
        self.production_figure.clear()
        self.production_canvas.draw()
        
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
        
        # Clear report preview
        self.report_preview.setText("No data available for report generation.")
        
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
            
            # Perform analysis
            self.analysis_results = self.perform_analysis(df)
            
            # Update all displays
            self.update_dashboard()
            self.update_analysis()
            self.update_trends()
            self.update_production_analysis()
            self.update_advanced_tracking()
            self.update_dimensional_rejects()
            self.update_tag_tracking_rejects()
            self.update_time_analysis()
            self.update_filters()
            
            # Switch to Analysis tab (2nd page)
            self.tab_widget.setCurrentIndex(1)
            
            self.status_label.setText(f"Data analyzed successfully! (Sheet: {data_sheet}) - Switched to Analysis tab")
            
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
                
        # Create quantity column (each row represents 1 rejected unit)
        df['Quantity'] = 1
        
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
        if self.analysis_results is None:
            return
            
        # Clear existing metrics
        for i in reversed(range(self.metrics_layout.count())):
            self.metrics_layout.itemAt(i).widget().setParent(None)
            
        # Create metric cards with placeholders
        total_rejections = self.analysis_results.get('total_rejections', 0)
        total_quantity = self.analysis_results.get('total_quantity', 0)
        
        # Get top rejection reason
        top_reason = "Unknown"
        if 'rejection_reasons' in self.analysis_results:
            top_reason = list(self.analysis_results['rejection_reasons'].keys())[0]
            
        metrics = [
            ("Total Rejections", f"{total_rejections:,}", self.theme.get_color('danger')),
            ("Top Reason", top_reason[:20], self.theme.get_color('info')),
            ("Est. Lost Time", "TBD - Ask Oscar", self.theme.get_color('warning')),
            ("Est. Cost Impact", "TBD - Ask Oscar", self.theme.get_color('purple'))
        ]
        
        for i, (title, value, color) in enumerate(metrics):
            card = MetricCard(title, value, color, self.theme)
            self.metrics_layout.addWidget(card, 0, i)
        
    def update_analysis(self):
        if self.current_data is None:
            return
            
        # Populate analysis table
        self.analysis_table.setRowCount(len(self.current_data))
        
        for row, (_, data) in enumerate(self.current_data.iterrows()):
            self.analysis_table.setItem(row, 0, QTableWidgetItem(str(data.get('Reject datetime', ''))))
            self.analysis_table.setItem(row, 1, QTableWidgetItem(str(data.get('Source', ''))))
            self.analysis_table.setItem(row, 2, QTableWidgetItem(str(data.get('Sku', ''))))
            self.analysis_table.setItem(row, 3, QTableWidgetItem(str(data.get('Reject reason', ''))))
            self.analysis_table.setItem(row, 4, QTableWidgetItem(str(data.get('Lpn', ''))))
            self.analysis_table.setItem(row, 5, QTableWidgetItem(str(data.get('Log text', ''))))
            
            # Root cause analysis (placeholder - would be enhanced with ML)
            root_cause = self.determine_root_cause(data)
            self.analysis_table.setItem(row, 6, QTableWidgetItem(root_cause))
            
            self.analysis_table.setItem(row, 7, QTableWidgetItem(str(data.get('Quantity', ''))))
            
    def determine_root_cause(self, data):
        """Determine likely root cause based on data patterns"""
        reason = str(data.get('Reject reason', '')).lower()
        log_text = str(data.get('Log text', '')).lower()
        
        # E80-specific root cause analysis
        if 'duplicated lpn' in reason or 'duplicate' in reason:
            return "Labeling System - Duplicate LPN"
        elif 'unit data not found' in reason or 'su_data' in log_text:
            return "System Integration - Missing Data"
        elif 'dimension' in reason or 'size' in reason or 'measurement' in reason:
            return "Equipment Calibration"
        elif 'tag' in reason or 'label' in reason or 'barcode' in reason:
            return "Labeling System"
        elif 'system' in reason or 'error' in reason or 'timeout' in reason:
            return "System Integration"
        elif 'quality' in reason or 'defect' in reason or 'damage' in reason:
            return "Quality Control"
        elif 'weight' in reason or 'scale' in reason:
            return "Equipment Calibration - Weight"
        elif 'position' in reason or 'location' in reason:
            return "Equipment Calibration - Position"
        else:
            return "Unknown/Other"
            
    def update_trends(self):
        if self.analysis_results is None:
            return
            
        # Clear existing plots
        self.trends_figure.clear()
        
        # Create subplots (only 2 charts now)
        ax1 = self.trends_figure.add_subplot(1, 2, 1)
        ax2 = self.trends_figure.add_subplot(1, 2, 2)
        
        # Set dark theme for plots
        self.trends_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Filter data based on selected periods and production lines
        filtered_data = self.current_data.copy()
        
        # Apply period filter
        if hasattr(self, 'trend_period_checkboxes') and self.trend_period_checkboxes:
            selected_periods = [period for period, checkbox in self.trend_period_checkboxes.items() if checkbox.isChecked()]
            if selected_periods:
                filtered_data['Period'] = self.get_period_from_date(filtered_data['Reject datetime'])
                filtered_data = filtered_data[filtered_data['Period'].isin(selected_periods)]
        
        # Apply line filter
        if hasattr(self, 'trend_line_checkboxes') and self.trend_line_checkboxes:
            selected_lines = [line for line, checkbox in self.trend_line_checkboxes.items() if checkbox.isChecked()]
            if selected_lines:
                filtered_data = filtered_data[filtered_data['Source'].isin(selected_lines)]
        
        # Plot 1: Rejection reasons pie chart (filtered)
        if 'Reject reason' in filtered_data.columns:
            rejection_counts = filtered_data['Reject reason'].value_counts()
            reasons = list(rejection_counts.keys())[:5]
            counts = list(rejection_counts.values)[:5]
            
            colors = ['#e74c3c', '#f39c12', '#f1c40f', '#2ecc71', '#3498db']
            wedges, texts, autotexts = ax1.pie(counts, labels=reasons, autopct='%1.1f%%', colors=colors[:len(reasons)], 
                                             textprops={'color': 'white', 'fontweight': 'bold'})
            ax1.set_title('Top Rejection Reasons', color='white', fontweight='bold', fontsize=16)
            
        # Plot 2: Period trends (filtered)
        if 'Reject datetime' in filtered_data.columns:
            # Calculate period trends for filtered data
            filtered_data['Period'] = self.get_period_from_date(filtered_data['Reject datetime'])
            period_trends = filtered_data.groupby('Period')['Quantity'].sum()
            
            periods = list(period_trends.keys())
            quantities = list(period_trends.values)
            
            ax2.plot(periods, quantities, marker='o', color='#e74c3c', linewidth=3, markersize=8)
            ax2.set_title('Period Rejection Trends', color='white', fontweight='bold', fontsize=16)
            ax2.set_ylabel('Rejected Quantity', color='white', fontsize=12)
            ax2.tick_params(axis='x', rotation=45, colors='white')
            ax2.tick_params(axis='y', colors='white')
            ax2.grid(True, alpha=0.3, color='white')
            
        # Style all axes
        for ax in [ax1, ax2]:
            ax.tick_params(colors='white')
            for spine in ax.spines.values():
                spine.set_color('white')
                
        self.trends_figure.tight_layout()
        self.trends_canvas.draw()
        
    def update_production_analysis(self):
        """Update production line and product analysis"""
        if self.analysis_results is None:
            return
            
        # Clear existing plots
        self.production_figure.clear()
        
        # Create subplots
        ax1 = self.production_figure.add_subplot(1, 2, 1)
        ax2 = self.production_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.production_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Plot 1: Line breakdown
        if 'line_breakdown' in self.analysis_results:
            lines = list(self.analysis_results['line_breakdown'].keys())
            line_counts = list(self.analysis_results['line_breakdown'].values())
            
            bars = ax1.bar(lines, line_counts, color='#3498db', edgecolor='white', linewidth=1)
            ax1.set_title('Rejections by Production Line', color='white', fontweight='bold', fontsize=16)
            ax1.set_ylabel('Number of Rejections', color='white', fontsize=12)
            ax1.tick_params(axis='x', colors='white', rotation=45)
            ax1.tick_params(axis='y', colors='white')
            ax1.grid(True, alpha=0.3, color='white', axis='y')
            
        # Plot 2: Product breakdown
        if 'product_breakdown' in self.analysis_results:
            products = list(self.analysis_results['product_breakdown'].keys())[:8]
            product_counts = list(self.analysis_results['product_breakdown'].values())[:8]
            
            bars = ax2.barh(products, product_counts, color='#9b59b6', edgecolor='white', linewidth=1)
            ax2.set_title('Rejections by Product', color='white', fontweight='bold', fontsize=16)
            ax2.set_xlabel('Number of Rejections', color='white', fontsize=12)
            ax2.tick_params(axis='x', colors='white')
            ax2.tick_params(axis='y', colors='white')
            ax2.grid(True, alpha=0.3, color='white', axis='x')
            
        # Style all axes
        for ax in [ax1, ax2]:
            for spine in ax.spines.values():
                spine.set_color('white')
                
        self.production_figure.tight_layout()
        self.production_canvas.draw()
        
    def update_advanced_tracking(self):
        """Update advanced tracking with rejection reason categories"""
        if self.analysis_results is None:
            return
            
        # Clear existing plots
        self.category_figure.clear()
        
        # Create subplots
        ax1 = self.category_figure.add_subplot(1, 2, 1)
        ax2 = self.category_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.category_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Categorize rejection reasons
        dimensional_issues = []
        tag_tracking_issues = []
        system_issues = []
        uncategorized_issues = []
        
        if 'rejection_reasons' in self.analysis_results:
            for reason, count in self.analysis_results['rejection_reasons'].items():
                reason_lower = reason.lower()
                if any(keyword in reason_lower for keyword in ['dimension', 'size', 'measurement', 'weight', 'position', 'height', 'width', 'length', 'tolerance', 'maximum']):
                    dimensional_issues.append((reason, count))
                elif any(keyword in reason_lower for keyword in ['tag', 'label', 'lpn', 'barcode', 'duplicate', 'unit data not found', 'tracking', 'expected', 'exist']):
                    tag_tracking_issues.append((reason, count))
                elif any(keyword in reason_lower for keyword in ['system', 'error', 'timeout', 'failed', 'check error']):
                    system_issues.append((reason, count))
                else:
                    uncategorized_issues.append((reason, count))
        
        # Show uncategorized issues for user to categorize
        if uncategorized_issues:
            print("Uncategorized rejection reasons that need categorization:")
            for reason, count in uncategorized_issues:
                print(f"  - '{reason}': {count} occurrences")
        
        # Plot 1: Category breakdown (three categories now)
        categories = ['Dimensional Issues', 'Tag/Tracking Issues', 'System Issues']
        category_counts = [sum(count for _, count in dimensional_issues),
                          sum(count for _, count in tag_tracking_issues),
                          sum(count for _, count in system_issues)]
        
        colors = ['#e74c3c', '#f39c12', '#9b59b6']
        wedges, texts, autotexts = ax1.pie(category_counts, labels=categories, autopct='%1.1f%%', 
                                          colors=colors, textprops={'color': 'white', 'fontweight': 'bold'})
        ax1.set_title('Rejection Categories', color='white', fontweight='bold', fontsize=14)
        
        # Plot 2: Top reasons by category
        all_issues = dimensional_issues + tag_tracking_issues + system_issues
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
        
        # Style axes
        for ax in [ax1, ax2]:
            for spine in ax.spines.values():
                spine.set_color('white')
                
        self.category_figure.tight_layout()
        self.category_canvas.draw()
        
    def update_dimensional_rejects(self):
        """Update dimensional rejects analysis"""
        if self.current_data is None:
            return
            
        # Clear existing plots
        self.dimensional_figure.clear()
        
        # Create subplots
        ax1 = self.dimensional_figure.add_subplot(1, 2, 1)
        ax2 = self.dimensional_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.dimensional_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Define dimensional keywords
        dimensional_keywords = ['dimension', 'size', 'measurement', 'weight', 'position', 'height', 'width', 'length', 'tolerance', 'maximum']
        
        # Filter for dimensional rejects
        dimensional_data = self.current_data[
            self.current_data['Reject reason'].str.contains('|'.join(dimensional_keywords), case=False, na=False)
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
                        f'{int(height)}', ha='center', va='bottom', color='white', fontweight='bold')
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
        if self.current_data is None:
            return
            
        # Clear existing plots
        self.tag_tracking_figure.clear()
        
        # Create subplots
        ax1 = self.tag_tracking_figure.add_subplot(1, 2, 1)
        ax2 = self.tag_tracking_figure.add_subplot(1, 2, 2)
        
        # Set dark theme
        self.tag_tracking_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Define tag/tracking keywords
        tag_tracking_keywords = ['tag', 'label', 'lpn', 'barcode', 'duplicate', 'unit data not found', 'tracking', 'expected', 'exist']
        
        # Filter for tag/tracking rejects
        tag_tracking_data = self.current_data[
            self.current_data['Reject reason'].str.contains('|'.join(tag_tracking_keywords), case=False, na=False)
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
                        f'{int(height)}', ha='center', va='bottom', color='white', fontweight='bold')
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
        if self.current_data is None:
            return
            
        # Clear existing plots
        self.time_figure.clear()
        
        # Create subplots
        ax1 = self.time_figure.add_subplot(2, 2, 1)
        ax2 = self.time_figure.add_subplot(2, 2, 2)
        ax3 = self.time_figure.add_subplot(2, 2, 3)
        ax4 = self.time_figure.add_subplot(2, 2, 4)
        
        # Set dark theme
        self.time_figure.patch.set_facecolor(self.theme.get_color('card_bg'))
        
        # Extract hour from datetime
        if 'Reject datetime' in self.current_data.columns:
            self.current_data['Hour'] = self.current_data['Reject datetime'].dt.hour
            self.current_data['DayOfWeek'] = self.current_data['Reject datetime'].dt.day_name()
            
            # Plot 1: Rejections by hour
            hourly_counts = self.current_data['Hour'].value_counts().sort_index()
            bars = ax1.bar(hourly_counts.index, hourly_counts.values, color='#e74c3c', 
                          edgecolor='white', linewidth=1)
            ax1.set_title('Rejections by Hour of Day', color='white', fontweight='bold', fontsize=12)
            ax1.set_xlabel('Hour', color='white')
            ax1.set_ylabel('Number of Rejections', color='white')
            ax1.tick_params(axis='x', colors='white')
            ax1.tick_params(axis='y', colors='white')
            ax1.grid(True, alpha=0.3, color='white', axis='y')
            
            # Plot 2: Rejections by day of week
            dow_counts = self.current_data['DayOfWeek'].value_counts()
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
            
            # Plot 3: Best vs Worst hours
            worst_hour = hourly_counts.idxmax()
            best_hour = hourly_counts.idxmin()
            
            comparison_data = [hourly_counts[worst_hour], hourly_counts[best_hour]]
            comparison_labels = [f'Worst Hour\n({worst_hour}:00)', f'Best Hour\n({best_hour}:00)']
            colors = ['#e74c3c', '#2ecc71']
            
            bars = ax3.bar(comparison_labels, comparison_data, color=colors, edgecolor='white', linewidth=1)
            ax3.set_title('Best vs Worst Hours', color='white', fontweight='bold', fontsize=12)
            ax3.set_ylabel('Number of Rejections', color='white')
            ax3.tick_params(axis='x', colors='white')
            ax3.tick_params(axis='y', colors='white')
            ax3.grid(True, alpha=0.3, color='white', axis='y')
            
            # Plot 4: Hourly trend line
            ax4.plot(hourly_counts.index, hourly_counts.values, marker='o', color='#9b59b6', 
                    linewidth=3, markersize=6)
            ax4.set_title('Hourly Rejection Trend', color='white', fontweight='bold', fontsize=12)
            ax4.set_xlabel('Hour', color='white')
            ax4.set_ylabel('Number of Rejections', color='white')
            ax4.tick_params(axis='x', colors='white')
            ax4.tick_params(axis='y', colors='white')
            ax4.grid(True, alpha=0.3, color='white')
            
        # Style all axes
        for ax in [ax1, ax2, ax3, ax4]:
            for spine in ax.spines.values():
                spine.set_color('white')
                
        self.time_figure.tight_layout()
        self.time_canvas.draw()
        
    def update_filters(self):
        if self.current_data is None:
            return
            
        # Update period checkboxes - show which periods have data
        if 'Reject datetime' in self.current_data.columns:
            # Calculate periods for the data
            self.current_data['Period'] = self.get_period_from_date(self.current_data['Reject datetime'])
            available_periods = set(self.current_data['Period'].unique())
            
            # Update checkboxes to show which periods have data
            for period, checkbox in self.period_checkboxes.items():
                if period in available_periods:
                    checkbox.setEnabled(True)
                    checkbox.setStyleSheet("color: white;")  # Normal color for available periods
                else:
                    checkbox.setEnabled(False)
                    checkbox.setStyleSheet("color: gray;")  # Grayed out for periods with no data
                    checkbox.setChecked(False)  # Uncheck periods with no data
            
        # Update line checkboxes (E80 format uses 'Source')
        # Clear existing checkboxes
        for checkbox in self.line_checkboxes.values():
            checkbox.setParent(None)
        self.line_checkboxes.clear()
        
        if 'Source' in self.current_data.columns:
            lines = sorted(self.current_data['Source'].unique())
            for line in lines:
                if line != 'Unknown':
                    checkbox = QCheckBox(line)
                    checkbox.setChecked(True)  # Default to all selected
                    self.line_checkboxes[line] = checkbox
                    self.line_widget_layout.addWidget(checkbox)
            
        # Update reason filter (E80 format uses 'Reject reason')
        self.reason_filter.clear()
        self.reason_filter.addItem("All Reasons")
        if 'Reject reason' in self.current_data.columns:
            reasons = sorted(self.current_data['Reject reason'].unique())
            self.reason_filter.addItems([reason for reason in reasons if reason != 'Unknown'])
            
        # Update trend period checkboxes
        if 'Reject datetime' in self.current_data.columns:
            # Calculate periods for the data
            self.current_data['Period'] = self.get_period_from_date(self.current_data['Reject datetime'])
            available_periods = set(self.current_data['Period'].unique())
            
            # Clear existing checkboxes
            for checkbox in self.trend_period_checkboxes.values():
                checkbox.setParent(None)
            self.trend_period_checkboxes.clear()
            
            # Add period definitions for trends
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
                if period in available_periods:
                    checkbox.setChecked(True)  # Default to all selected
                    checkbox.setEnabled(True)
                    checkbox.setStyleSheet("color: white;")
                else:
                    checkbox.setChecked(False)
                    checkbox.setEnabled(False)
                    checkbox.setStyleSheet("color: gray;")
                checkbox.stateChanged.connect(self.update_trends)  # Connect to update function
                self.trend_period_checkboxes[period] = checkbox
                self.trend_period_widget_layout.addWidget(checkbox)
        
        # Update trend line checkboxes
        for checkbox in self.trend_line_checkboxes.values():
            checkbox.setParent(None)
        self.trend_line_checkboxes.clear()
        
        if 'Source' in self.current_data.columns:
            lines = sorted(self.current_data['Source'].unique())
            for line in lines:
                if line != 'Unknown':
                    checkbox = QCheckBox(line)
                    checkbox.setChecked(True)  # Default to all selected
                    checkbox.stateChanged.connect(self.update_trends)  # Connect to update function
                    self.trend_line_checkboxes[line] = checkbox
                    self.trend_line_widget_layout.addWidget(checkbox)
            
    def apply_filters(self):
        if self.current_data is None:
            return
            
        # Apply period filter
        filtered_data = self.current_data.copy()
        
        if 'Reject datetime' in filtered_data.columns and hasattr(self, 'period_checkboxes'):
            # Get selected periods
            selected_periods = [period for period, checkbox in self.period_checkboxes.items() if checkbox.isChecked()]
            
            if selected_periods:
                # Calculate periods for all data
                filtered_data['Period'] = self.get_period_from_date(filtered_data['Reject datetime'])
                # Filter to only selected periods
                filtered_data = filtered_data[filtered_data['Period'].isin(selected_periods)]
        
        # Apply line filter (E80 format uses 'Source') - multi-select
        selected_lines = [line for line, checkbox in self.line_checkboxes.items() if checkbox.isChecked()]
        if selected_lines:
            filtered_data = filtered_data[filtered_data['Source'].isin(selected_lines)]
            
        # Apply reason filter (E80 format uses 'Reject reason')
        selected_reason = self.reason_filter.currentText()
        if selected_reason != "All Reasons":
            filtered_data = filtered_data[filtered_data['Reject reason'] == selected_reason]
            
        # Update analysis table with filtered data
        self.analysis_table.setRowCount(len(filtered_data))
        
        for row, (_, data) in enumerate(filtered_data.iterrows()):
            self.analysis_table.setItem(row, 0, QTableWidgetItem(str(data.get('Reject datetime', ''))))
            self.analysis_table.setItem(row, 1, QTableWidgetItem(str(data.get('Source', ''))))
            self.analysis_table.setItem(row, 2, QTableWidgetItem(str(data.get('Sku', ''))))
            self.analysis_table.setItem(row, 3, QTableWidgetItem(str(data.get('Reject reason', ''))))
            self.analysis_table.setItem(row, 4, QTableWidgetItem(str(data.get('Lpn', ''))))
            self.analysis_table.setItem(row, 5, QTableWidgetItem(str(data.get('Log text', ''))))
            
            root_cause = self.determine_root_cause(data)
            self.analysis_table.setItem(row, 6, QTableWidgetItem(root_cause))
            
            self.analysis_table.setItem(row, 7, QTableWidgetItem(str(data.get('Quantity', ''))))
            
    def generate_summary_report(self):
        if self.analysis_results is None:
            QMessageBox.warning(self, "No Data", "Please load and analyze data first.")
            return
            
        report = f"""
E80 REJECTED UNITS ANALYSIS REPORT
==================================
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

EXECUTIVE SUMMARY
================
Total Rejected Units: {self.analysis_results.get('total_rejections', 0):,}
Total Rejected Quantity: {self.analysis_results.get('total_quantity', 0):,}
Analysis Period: {self.analysis_results.get('date_range', (None, None))[0]} to {self.analysis_results.get('date_range', (None, None))[1]}

KEY FINDINGS
============
"""
        
        if 'rejection_reasons' in self.analysis_results:
            report += "\nTop Rejection Reasons:\n"
            for i, (reason, count) in enumerate(list(self.analysis_results['rejection_reasons'].items())[:5]):
                percentage = (count / self.analysis_results.get('total_rejections', 1)) * 100
                report += f"{i+1}. {reason}: {count:,} ({percentage:.1f}%)\n"
                
        if 'line_breakdown' in self.analysis_results:
            report += "\nProduction Line Impact:\n"
            for line, count in self.analysis_results['line_breakdown'].items():
                percentage = (count / self.analysis_results.get('total_rejections', 1)) * 100
                report += f"• {line}: {count:,} rejections ({percentage:.1f}%)\n"
                
        report += f"""

RECOMMENDATIONS
===============
1. Focus improvement efforts on the top rejection reasons
2. Investigate production lines with highest rejection rates
3. Implement preventive maintenance schedules
4. Enhance operator training programs
5. Consider system upgrades for identified root causes

ESTIMATED COST IMPACT
=====================
Based on industry standards, each rejected pallet costs approximately $50-100 in rework and disposal.
With {self.analysis_results.get('total_quantity', 0):,} rejected units, estimated annual impact: 
${self.analysis_results.get('total_quantity', 0) * 75:,.0f}
"""
        
        self.report_preview.setText(report)
        
    def generate_detailed_report(self):
        if self.current_data is None:
            QMessageBox.warning(self, "No Data", "Please load and analyze data first.")
            return
            
        # Generate detailed analysis report
        report = f"""
DETAILED ROOT CAUSE ANALYSIS REPORT
===================================
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

DATA OVERVIEW
=============
Total Records: {len(self.current_data):,}
Date Range: {self.current_data['Date'].min()} to {self.current_data['Date'].max()}
Unique Products: {self.current_data['Product'].nunique() if 'Product' in self.current_data.columns else 'N/A'}
Unique Lines: {self.current_data['Line'].nunique() if 'Line' in self.current_data.columns else 'N/A'}

DETAILED BREAKDOWN
==================
"""
        
        # Add detailed statistics for each category
        if 'Reject reason' in self.current_data.columns:
            report += "\nRejection Reasons Analysis:\n"
            reason_stats = self.current_data['Reject reason'].value_counts()
            for reason, count in reason_stats.head(10).items():
                percentage = (count / len(self.current_data)) * 100
                report += f"• {reason}: {count:,} ({percentage:.1f}%)\n"
                
        if 'Source' in self.current_data.columns:
            report += "\nProduction Line Analysis:\n"
            line_stats = self.current_data['Source'].value_counts()
            for line, count in line_stats.items():
                percentage = (count / len(self.current_data)) * 100
                report += f"• {line}: {count:,} ({percentage:.1f}%)\n"
                
        # Add time-based analysis
        if 'Reject datetime' in self.current_data.columns:
            self.current_data['DayOfWeek'] = self.current_data['Reject datetime'].dt.day_name()
            self.current_data['Hour'] = self.current_data['Reject datetime'].dt.hour
            
            report += "\nTemporal Analysis:\n"
            dow_stats = self.current_data['DayOfWeek'].value_counts()
            report += "Rejections by Day of Week:\n"
            for day, count in dow_stats.items():
                percentage = (count / len(self.current_data)) * 100
                report += f"• {day}: {count:,} ({percentage:.1f}%)\n"
                
        self.report_preview.setText(report)
        
    def generate_trend_report(self):
        if self.analysis_results is None:
            QMessageBox.warning(self, "No Data", "Please load and analyze data first.")
            return
            
        report = f"""
TREND ANALYSIS REPORT
====================
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

MONTHLY TRENDS
==============
"""
        
        if 'monthly_trends' in self.analysis_results:
            for month, quantity in self.analysis_results['monthly_trends'].items():
                report += f"• {month}: {quantity:,} rejected units\n"
                
        # Calculate trend direction
        if len(self.analysis_results.get('monthly_trends', {})) >= 2:
            quantities = list(self.analysis_results['monthly_trends'].values())
            if len(quantities) >= 2:
                recent_trend = quantities[-1] - quantities[-2]
                if recent_trend > 0:
                    trend_direction = "INCREASING"
                    trend_color = "🔴"
                elif recent_trend < 0:
                    trend_direction = "DECREASING"
                    trend_color = "🟢"
                else:
                    trend_direction = "STABLE"
                    trend_color = "🟡"
                    
                report += f"\nTREND DIRECTION: {trend_color} {trend_direction}\n"
                report += f"Change from previous month: {recent_trend:+,} units\n"
                
        report += f"""

FORECASTING
===========
Based on current trends, projected annual rejection rate: 
{self.analysis_results.get('total_quantity', 0) * 12 / max(1, len(self.analysis_results.get('monthly_trends', {}))):,.0f} units

RECOMMENDATIONS
===============
1. Monitor trend direction closely
2. Implement corrective actions for increasing trends
3. Replicate successful practices from improving periods
4. Set up automated alerts for trend changes
"""
        
        self.report_preview.setText(report)

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
