#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
群体出行人员筛查系统
供公安民警使用的桌面应用
"""

import sys
import os
from datetime import datetime, date
import pandas as pd
import warnings


def setup_qt_environment():
    """设置Qt环境变量，确保应用能正常启动"""

    if getattr(sys, "frozen", False):
        current_dir = os.path.dirname(sys.executable)
        is_frozen = True
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        is_frozen = False

    possible_plugin_paths = []

    if is_frozen:
        possible_plugin_paths.extend(
            [
                os.path.join(current_dir, "PySide2", "plugins"),
                os.path.join(current_dir, "lib", "PySide2", "plugins"),
                os.path.join(current_dir, "Lib", "site-packages", "PySide2", "plugins"),
            ]
        )
    else:
        possible_plugin_paths.extend(
            [
                os.path.join(
                    current_dir,
                    "venv_win7",
                    "Lib",
                    "site-packages",
                    "PySide2",
                    "plugins",
                ),
                os.path.join(sys.prefix, "Lib", "site-packages", "PySide2", "plugins"),
                os.path.join(
                    os.path.dirname(sys.executable),
                    "Lib",
                    "site-packages",
                    "PySide2",
                    "plugins",
                ),
            ]
        )

    qt_plugin_path = None
    for path in possible_plugin_paths:
        if os.path.exists(path):
            qt_plugin_path = path
            break

    if qt_plugin_path:
        os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = qt_plugin_path
        print(f"设置Qt插件路径: {qt_plugin_path}")
    else:
        print("警告: 未找到Qt插件路径，可能会出现显示问题")
        print(f"搜索过的路径: {possible_plugin_paths}")

    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"

    if not os.environ.get("QT_QPA_PLATFORM"):
        os.environ["QT_QPA_PLATFORM"] = "windows"


setup_qt_environment()

from PySide2.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QComboBox,
    QDateEdit,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QMessageBox,
    QGroupBox,
    QProgressBar,
    QAbstractItemView,
    QCheckBox,
    QTextEdit,
    QLineEdit,
    QSplitter,
    QDialog,
    QDialogButtonBox,
    QScrollArea,
    QFrame,
    QGridLayout,
    QScrollBar,
)
from PySide2.QtCore import Qt, QDate, QThread, Signal, QTimer
from PySide2.QtGui import QFont, QPalette, QColor, QBrush, QPixmap, QPainter
from PySide2.QtWidgets import QDesktopWidget

warnings.filterwarnings("ignore")


class PersonDetailDialog(QDialog):

    def __init__(self, person_id, all_data, parent=None):
        super().__init__(parent)
        self.person_id = person_id
        self.all_data = all_data
        self.person_records = None

        self.init_ui()
        self.load_person_data()

    def init_ui(self):
        self.setWindowTitle(f"人员详情 - {self.person_id}")
        self.setModal(True)
        self.resize(1000, 850)

        self.setStyleSheet(
            """
            QDialog {
                background-color: #f8f9fa;
            }
            QGroupBox {
                font-size: 14px;
                font-weight: bold;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                margin-top: 15px;
                padding-top: 15px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 8px 0 8px;
                color: #495057;
            }
            QLabel {
                font-size: 13px;
                color: #495057;
            }
            .info-label {
                font-weight: bold;
                color: #212529;
            }
            .value-label {
                color: #6c757d;
                padding: 5px;
                background-color: #f8f9fa;
                border-radius: 4px;
            }
            QTableWidget {
                border: 1px solid #dee2e6;
                border-radius: 6px;
                background-color: white;
                gridline-color: #e9ecef;
            }
            QTableWidget::item {
                padding: 8px;
                border-bottom: 1px solid #e9ecef;
            }
            QTableWidget::item:selected {
                background-color: #e3f2fd;
                color: #1976d2;
            }
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #004085;
            }
        """
        )

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

        title_layout = QHBoxLayout()
        title_icon = QLabel("👤")
        title_icon.setStyleSheet("font-size: 24px;")
        title_label = QLabel("人员详细信息")
        title_label.setStyleSheet(
            """
            font-size: 20px;
            font-weight: bold;
            color: #212529;
            margin-left: 10px;
        """
        )
        title_layout.addWidget(title_icon)
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        main_layout.addLayout(title_layout)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)

        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(15)

        self.basic_info_group = QGroupBox("📋 基本信息")
        self.setup_basic_info_area()
        scroll_layout.addWidget(self.basic_info_group)

        self.stats_group = QGroupBox("📊 出行统计")
        self.setup_stats_area()
        scroll_layout.addWidget(self.stats_group)

        self.records_group = QGroupBox("✈️ 出行记录详情")
        self.setup_records_area()
        scroll_layout.addWidget(self.records_group)

        self.timeline_group = QGroupBox("📅 时间线")
        self.setup_timeline_area()
        scroll_layout.addWidget(self.timeline_group)

        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)

        button_layout = QHBoxLayout()

        self.export_btn = QPushButton("📄 导出个人报告")
        self.export_btn.clicked.connect(self.export_person_report)

        self.related_btn = QPushButton("🔍 关联分析")
        self.related_btn.clicked.connect(self.show_related_analysis)

        close_btn = QPushButton("❌ 关闭")
        close_btn.clicked.connect(self.accept)

        button_layout.addWidget(self.export_btn)
        button_layout.addWidget(self.related_btn)
        button_layout.addStretch()
        button_layout.addWidget(close_btn)

        main_layout.addLayout(button_layout)

    def setup_basic_info_area(self):
        layout = QGridLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        self.name_label = QLabel()
        self.id_label = QLabel()
        self.total_records_label = QLabel()
        self.date_range_label = QLabel()

        for label in [
            self.name_label,
            self.id_label,
            self.total_records_label,
            self.date_range_label,
        ]:
            label.setProperty("class", "value-label")

        layout.addWidget(QLabel("姓名："), 0, 0)
        layout.addWidget(self.name_label, 0, 1)
        layout.addWidget(QLabel("证件号："), 0, 2)
        layout.addWidget(self.id_label, 0, 3)

        layout.addWidget(QLabel("记录总数："), 1, 0)
        layout.addWidget(self.total_records_label, 1, 1)
        layout.addWidget(QLabel("出行时间范围："), 1, 2)
        layout.addWidget(self.date_range_label, 1, 3)

        layout.setColumnStretch(1, 1)
        layout.setColumnStretch(3, 1)

        self.basic_info_group.setLayout(layout)

    def setup_stats_area(self):
        layout = QGridLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        self.cities_label = QLabel()
        self.flights_label = QLabel()
        self.status_label = QLabel()
        self.sources_label = QLabel()

        for label in [
            self.cities_label,
            self.flights_label,
            self.status_label,
            self.sources_label,
        ]:
            label.setProperty("class", "value-label")

        layout.addWidget(QLabel("涉及城市："), 0, 0)
        layout.addWidget(self.cities_label, 0, 1)
        layout.addWidget(QLabel("航班车次："), 0, 2)
        layout.addWidget(self.flights_label, 0, 3)

        layout.addWidget(QLabel("状态分布："), 1, 0)
        layout.addWidget(self.status_label, 1, 1)
        layout.addWidget(QLabel("数据来源："), 1, 2)
        layout.addWidget(self.sources_label, 1, 3)

        layout.setColumnStretch(1, 1)
        layout.setColumnStretch(3, 1)

        self.stats_group.setLayout(layout)

    def setup_records_area(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        self.records_table = QTableWidget()
        self.records_table.setAlternatingRowColors(True)
        self.records_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.records_table.horizontalHeader().setStretchLastSection(True)
        self.records_table.setMinimumHeight(250)
        self.records_table.verticalHeader().setDefaultSectionSize(35)
        self.records_table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        layout.addWidget(self.records_table)
        self.records_group.setLayout(layout)

    def setup_timeline_area(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        self.timeline_text = QTextEdit()
        self.timeline_text.setReadOnly(True)
        self.timeline_text.setMinimumHeight(300)
        self.timeline_text.setMaximumHeight(500)
        self.timeline_text.setStyleSheet(
            """
            QTextEdit {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 6px;
                padding: 10px;
                font-family: 'Courier New', monospace;
                font-size: 14px;
                line-height: 1.5;
            }
        """
        )

        layout.addWidget(self.timeline_text)
        self.timeline_group.setLayout(layout)

    def load_person_data(self):
        if self.all_data is None or self.all_data.empty:
            QMessageBox.warning(self, "警告", "没有可用的数据")
            return

        self.person_records = self.all_data[
            self.all_data["证件号"].astype(str).str.strip()
            == str(self.person_id).strip()
        ].copy()

        if self.person_records.empty:
            QMessageBox.warning(self, "警告", f"未找到证件号为 {self.person_id} 的记录")
            return

        if "出发日期" in self.person_records.columns:
            self.person_records = self.person_records.sort_values("出发日期")

        self.update_basic_info()
        self.update_stats_info()
        self.update_records_table()
        self.update_timeline()

    def update_basic_info(self):
        """更新基本信息"""
        if self.person_records.empty:
            return

        # 获取姓名（取第一条记录的姓名）
        name = (
            str(self.person_records.iloc[0]["姓名"])
            if "姓名" in self.person_records.columns
            else "未知"
        )
        self.name_label.setText(name)

        # 证件号
        self.id_label.setText(str(self.person_id))

        # 记录总数
        self.total_records_label.setText(f"{len(self.person_records)} 条")

        # 出行时间范围
        if "出发日期" in self.person_records.columns:
            dates = pd.to_datetime(
                self.person_records["出发日期"], errors="coerce"
            ).dropna()
            if not dates.empty:
                min_date = dates.min().strftime("%Y-%m-%d")
                max_date = dates.max().strftime("%Y-%m-%d")
                if min_date == max_date:
                    date_range = min_date
                else:
                    date_range = f"{min_date} 至 {max_date}"
                self.date_range_label.setText(date_range)
            else:
                self.date_range_label.setText("日期信息不完整")
        else:
            self.date_range_label.setText("无日期信息")

    def update_stats_info(self):
        """更新统计信息"""
        if self.person_records.empty:
            return

        # 涉及城市统计
        cities = set()
        for col in ["发站", "到站"]:
            if col in self.person_records.columns:
                city_list = self.person_records[col].dropna().astype(str)
                cities.update([city for city in city_list if city.strip()])

        cities_text = f"{len(cities)} 个：" + "、".join(list(cities)[:5])
        if len(cities) > 5:
            cities_text += f"... (共{len(cities)}个)"
        self.cities_label.setText(cities_text if cities else "无")

        # 航班车次统计
        if "航班车次" in self.person_records.columns:
            flights = self.person_records["航班车次"].dropna().astype(str)
            unique_flights = flights.unique()
            flights_text = f"{len(unique_flights)} 个：" + "、".join(unique_flights[:3])
            if len(unique_flights) > 3:
                flights_text += f"... (共{len(unique_flights)}个)"
            self.flights_label.setText(
                flights_text if len(unique_flights) > 0 else "无"
            )
        else:
            self.flights_label.setText("无")

        # 状态分布统计
        if "状态类型" in self.person_records.columns:
            status_counts = self.person_records["状态类型"].value_counts()
            status_parts = [
                f"{status}({count})" for status, count in status_counts.items()
            ]
            self.status_label.setText("、".join(status_parts))
        else:
            self.status_label.setText("无状态信息")

        # 数据来源统计
        if "数据源" in self.person_records.columns:
            source_counts = self.person_records["数据源"].value_counts()
            source_parts = [
                f"{source}({count})" for source, count in source_counts.items()
            ]
            self.sources_label.setText("、".join(source_parts))
        else:
            self.sources_label.setText("无来源信息")

    def update_records_table(self):
        """更新记录表格"""
        if self.person_records.empty:
            return

        # 定义要显示的列
        columns = [
            "出发日期",
            "出发时间",
            "航班车次",
            "发站",
            "到站",
            "变更操作",
            "状态类型",
            "数据源",
        ]

        # 过滤出实际存在的列
        available_columns = [
            col for col in columns if col in self.person_records.columns
        ]

        # 设置表格
        self.records_table.setRowCount(len(self.person_records))
        self.records_table.setColumnCount(len(available_columns))
        self.records_table.setHorizontalHeaderLabels(available_columns)

        # 填充数据
        for row_idx, (_, record) in enumerate(self.person_records.iterrows()):
            for col_idx, col in enumerate(available_columns):
                value = record[col]

                if pd.isna(value):
                    value = ""
                elif col == "出发日期":
                    try:
                        value = pd.to_datetime(value).strftime("%Y-%m-%d")
                    except:
                        value = str(value)
                else:
                    value = str(value)

                item = QTableWidgetItem(value)

                # 根据状态设置颜色
                if col == "状态类型":
                    if value == "已确认":
                        item.setForeground(QBrush(QColor("#28a745")))
                        item.setBackground(QBrush(QColor("#d4edda")))
                    elif value == "待确认":
                        item.setForeground(QBrush(QColor("#fd7e14")))
                        item.setBackground(QBrush(QColor("#fff3cd")))
                elif col == "变更操作":
                    # 根据变更操作类型设置颜色
                    if value in ["登机", "值机"]:
                        item.setForeground(QBrush(QColor("#28a745")))
                    elif value in ["出票", "改期"]:
                        item.setForeground(QBrush(QColor("#17a2b8")))

                self.records_table.setItem(row_idx, col_idx, item)

        # 调整列宽
        self.records_table.resizeColumnsToContents()

    def update_timeline(self):
        """更新时间线显示"""
        if self.person_records.empty:
            self.timeline_text.setText("无时间线数据")
            return

        timeline_html = (
            "<div style='font-family: Consolas, monospace; font-size: 14px;'>"
        )
        timeline_html += (
            "<h4 style='color: #495057; margin-bottom: 15px;'>📅 出行时间线</h4>"
        )

        # 按日期分组显示
        if "出发日期" in self.person_records.columns:
            # 转换日期并排序
            records_with_date = self.person_records.copy()
            records_with_date["出发日期_dt"] = pd.to_datetime(
                records_with_date["出发日期"], errors="coerce"
            )
            records_sorted = records_with_date.dropna(
                subset=["出发日期_dt"]
            ).sort_values("出发日期_dt")

            current_date = None
            for _, record in records_sorted.iterrows():
                record_date = record["出发日期_dt"].strftime("%Y-%m-%d")

                # 如果是新的日期，添加日期标题
                if current_date != record_date:
                    if current_date is not None:
                        timeline_html += "<br>"
                    timeline_html += f"<div style='background-color: #e9ecef; padding: 8px; margin: 5px 0; border-radius: 4px; font-weight: bold; color: #495057;'>"
                    timeline_html += f"📅 {record_date}"
                    timeline_html += "</div>"
                    current_date = record_date

                # 添加记录项
                time_str = str(record.get("出发时间", "")).strip() or "时间未知"
                flight = str(record.get("航班车次", "")).strip() or "航班未知"
                route = f"{record.get('发站', '')} → {record.get('到站', '')}"
                operation = str(record.get("变更操作", "")).strip() or "操作未知"
                status = str(record.get("状态类型", "")).strip() or "状态未知"

                # 根据状态设置图标和颜色
                if status == "已确认":
                    status_icon = "✅"
                    status_color = "#28a745"
                elif status == "待确认":
                    status_icon = "⏳"
                    status_color = "#fd7e14"
                else:
                    status_icon = "❓"
                    status_color = "#6c757d"

                timeline_html += f"<div style='margin-left: 20px; padding: 6px; border-left: 3px solid {status_color}; margin-bottom: 8px;'>"
                timeline_html += f"<span style='color: {status_color}; font-weight: bold;'>{status_icon} {time_str}</span> - "
                timeline_html += f"<span style='color: #007bff;'>{flight}</span> "
                timeline_html += f"<span style='color: #495057;'>{route}</span><br>"
                timeline_html += f"<small style='color: #6c757d; margin-left: 15px;'>操作: {operation} | 状态: {status}</small>"
                timeline_html += "</div>"
        else:
            timeline_html += (
                "<p style='color: #6c757d;'>无法生成时间线：缺少日期信息</p>"
            )

        timeline_html += "</div>"
        self.timeline_text.setHtml(timeline_html)

    def export_person_report(self):
        """导出个人详情报告"""
        if self.person_records is None or self.person_records.empty:
            QMessageBox.warning(self, "警告", "没有可导出的数据")
            return

        # 获取姓名用于文件名
        name = (
            str(self.person_records.iloc[0]["姓名"])
            if "姓名" in self.person_records.columns
            else "未知姓名"
        )
        default_filename = f"个人出行报告_{name}_{self.person_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存个人报告", default_filename, "Excel文件 (*.xlsx)"
        )

        if file_path:
            try:
                with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                    # 基本信息表
                    basic_info = pd.DataFrame(
                        {
                            "项目": [
                                "姓名",
                                "证件号",
                                "记录总数",
                                "最早出行",
                                "最近出行",
                            ],
                            "信息": [
                                name,
                                self.person_id,
                                len(self.person_records),
                                (
                                    self.person_records["出发日期"].min()
                                    if "出发日期" in self.person_records.columns
                                    else "无"
                                ),
                                (
                                    self.person_records["出发日期"].max()
                                    if "出发日期" in self.person_records.columns
                                    else "无"
                                ),
                            ],
                        }
                    )
                    basic_info.to_excel(writer, sheet_name="基本信息", index=False)

                    # 详细记录表
                    self.person_records.to_excel(
                        writer, sheet_name="详细记录", index=False
                    )

                    # 统计汇总表
                    summary_data = []

                    # 城市统计
                    if "到站" in self.person_records.columns:
                        city_counts = self.person_records["到站"].value_counts()
                        for city, count in city_counts.items():
                            summary_data.append(["目的地", city, count])

                    # 状态统计
                    if "状态类型" in self.person_records.columns:
                        status_counts = self.person_records["状态类型"].value_counts()
                        for status, count in status_counts.items():
                            summary_data.append(["状态", status, count])

                    if summary_data:
                        summary_df = pd.DataFrame(
                            summary_data, columns=["类别", "项目", "数量"]
                        )
                        summary_df.to_excel(writer, sheet_name="统计汇总", index=False)

                QMessageBox.information(
                    self, "成功", f"个人报告已导出到：\n{file_path}"
                )

            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")

    def show_related_analysis(self):
        """显示关联分析"""
        if self.person_records is None or self.person_records.empty:
            QMessageBox.warning(self, "警告", "没有可分析的数据")
            return

        # 创建关联分析对话框
        analysis_dialog = QDialog(self)
        analysis_dialog.setWindowTitle(f"关联分析 - {self.person_id}")
        analysis_dialog.setModal(True)
        analysis_dialog.resize(600, 500)

        layout = QVBoxLayout(analysis_dialog)

        # 分析结果文本区域
        analysis_text = QTextEdit()
        analysis_text.setReadOnly(True)
        analysis_text.setStyleSheet(
            """
            QTextEdit {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 6px;
                padding: 15px;
                font-size: 13px;
                line-height: 1.5;
            }
        """
        )

        # 生成关联分析内容
        analysis_content = self.generate_related_analysis()
        analysis_text.setHtml(analysis_content)

        layout.addWidget(QLabel("🔍 关联分析结果"))
        layout.addWidget(analysis_text)

        # 关闭按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(analysis_dialog.accept)
        layout.addWidget(button_box)

        analysis_dialog.exec_()

    def generate_related_analysis(self):
        """生成关联分析内容"""
        html = "<div style='font-family: Arial, sans-serif;'>"
        html += "<h3 style='color: #495057; margin-bottom: 20px;'>🔍 关联分析报告</h3>"

        # 分析同航班人员
        if "航班车次" in self.person_records.columns and self.all_data is not None:
            html += (
                "<h4 style='color: #007bff; margin-top: 20px;'>✈️ 同航班人员分析</h4>"
            )

            person_flights = self.person_records["航班车次"].unique()
            related_people = set()

            for flight in person_flights:
                if pd.notna(flight) and str(flight).strip():
                    same_flight_records = self.all_data[
                        (self.all_data["航班车次"] == flight)
                        & (self.all_data["证件号"] != self.person_id)
                    ]

                    if not same_flight_records.empty:
                        html += f"<p style='margin-left: 15px;'><strong>航班 {flight}:</strong> 发现 {len(same_flight_records)} 名同行人员</p>"

                        # 列出同行人员（最多显示5个）
                        for _, record in same_flight_records.head(5).iterrows():
                            name = record.get("姓名", "未知")
                            id_num = (
                                str(record.get("证件号", ""))[:6] + "****"
                            )  # 隐藏部分证件号
                            html += f"<p style='margin-left: 30px; color: #6c757d;'>• {name} ({id_num})</p>"

                        if len(same_flight_records) > 5:
                            html += f"<p style='margin-left: 30px; color: #6c757d;'>... 还有 {len(same_flight_records) - 5} 人</p>"

                        related_people.update(same_flight_records["证件号"].astype(str))

            if not related_people:
                html += (
                    "<p style='margin-left: 15px; color: #6c757d;'>未发现同航班人员</p>"
                )

        # 分析出行模式
        html += "<h4 style='color: #28a745; margin-top: 20px;'>📊 出行模式分析</h4>"

        if "出发日期" in self.person_records.columns:
            dates = pd.to_datetime(
                self.person_records["出发日期"], errors="coerce"
            ).dropna()
            if len(dates) > 1:
                # 计算出行间隔
                date_diffs = dates.diff().dropna()
                avg_interval = date_diffs.mean().days
                html += f"<p style='margin-left: 15px;'>平均出行间隔: {avg_interval:.1f} 天</p>"

                # 分析出行频率
                if avg_interval < 7:
                    frequency_desc = "高频出行（平均间隔少于一周）"
                elif avg_interval < 30:
                    frequency_desc = "中频出行（平均间隔少于一月）"
                else:
                    frequency_desc = "低频出行（平均间隔超过一月）"

                html += f"<p style='margin-left: 15px;'>出行频率: {frequency_desc}</p>"
            else:
                html += "<p style='margin-left: 15px; color: #6c757d;'>出行记录不足，无法分析出行模式</p>"

        # 分析目的地偏好
        if "到站" in self.person_records.columns:
            destination_counts = self.person_records["到站"].value_counts()
            if not destination_counts.empty:
                html += (
                    "<h4 style='color: #dc3545; margin-top: 20px;'>🎯 目的地偏好</h4>"
                )

                for dest, count in destination_counts.head(5).items():
                    percentage = (count / len(self.person_records)) * 100
                    html += f"<p style='margin-left: 15px;'>{dest}: {count} 次 ({percentage:.1f}%)</p>"

        # 风险提示
        html += "<h4 style='color: #fd7e14; margin-top: 20px;'>⚠️ 风险提示</h4>"

        risk_factors = []

        # 高频出行检查
        if len(self.person_records) > 10:
            risk_factors.append("高频出行：记录数量较多，需要重点关注")

        # 多目的地检查
        if "到站" in self.person_records.columns:
            unique_destinations = self.person_records["到站"].nunique()
            if unique_destinations > 5:
                risk_factors.append(
                    f"多目的地出行：涉及 {unique_destinations} 个不同目的地"
                )

        # 状态异常检查
        if "状态类型" in self.person_records.columns:
            unconfirmed_count = len(
                self.person_records[self.person_records["状态类型"] == "待确认"]
            )
            if unconfirmed_count > 0:
                risk_factors.append(f"存在 {unconfirmed_count} 条待确认记录")

        if risk_factors:
            for factor in risk_factors:
                html += f"<p style='margin-left: 15px; color: #fd7e14;'>• {factor}</p>"
        else:
            html += (
                "<p style='margin-left: 15px; color: #28a745;'>未发现明显风险因素</p>"
            )

        html += "</div>"
        return html


class DataPreviewLoader(QThread):
    progress = Signal(int)
    message = Signal(str)
    finished = Signal(pd.DataFrame)
    error = Signal(str)

    def __init__(self):
        super().__init__()
        self.file1_path = ""
        self.file2_path = ""
        self.existing_data = None

        # 复用DataProcessor的状态优先级
        self.status_priority = {
            "登机": 1,
            "值机": 2,
            "进检": 3,
            "出票": 4,
            "座变": 5,
            "改期": 6,
            "段消": 7,
            "换开": 8,
            "证变": 9,
            "值拉": 10,
            "票务记录": 11,
        }

    def set_params(self, file1, file2, existing_data=None):
        """设置参数"""
        self.file1_path = file1
        self.file2_path = file2
        self.existing_data = existing_data

    def run(self):
        try:
            self.progress.emit(10)
            self.message.emit("正在预加载数据...")

            all_data = pd.DataFrame()

            if self.file1_path:
                self.message.emit("正在预加载票务全库数据...")
                df1 = self.read_ticket_data(self.file1_path)
                if not df1.empty:
                    all_data = pd.concat([all_data, df1], ignore_index=True)
                self.progress.emit(40)

            if self.file2_path:
                self.message.emit("正在预加载群体票务数据...")
                df2 = self.read_mixed_transport_data(self.file2_path)
                if not df2.empty:
                    all_data = pd.concat([all_data, df2], ignore_index=True)
                self.progress.emit(60)

            if all_data.empty:
                self.error.emit("没有读取到有效数据")
                return

            self.progress.emit(80)
            self.message.emit("正在整理数据...")

            # 去重处理
            all_data = self.final_dedup(all_data)

            # 如果有现有数据，进行合并
            if self.existing_data is not None and not self.existing_data.empty:
                self.message.emit("正在合并历史数据...")
                all_data = pd.concat([self.existing_data, all_data], ignore_index=True)
                all_data = self.final_dedup(all_data)

            self.progress.emit(100)
            self.message.emit("数据预加载完成！")

            self.finished.emit(all_data)

        except Exception as e:
            self.error.emit(f"预加载出错：{str(e)}")

    def read_ticket_data(self, file_path):
        """读取票务全库数据 - 简化版"""
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()  # 去除空格
            df["数据源"] = "票务全库"  # 添加数据源

            # 基本数据清理
            if "姓名" in df.columns:
                df = df[df["姓名"].notna() & (df["姓名"].astype(str).str.strip() != "")]
            if "证件号" in df.columns:
                df = df[
                    df["证件号"].notna() & (df["证件号"].astype(str).str.strip() != "")
                ]

            # 处理日期
            if "出发日期" in df.columns:
                df["出发日期"] = pd.to_datetime(df["出发日期"], errors="coerce")

            return df
        except Exception as e:
            raise Exception(f"读取票务数据失败：{str(e)}")

    def read_flight_data(self, file_path):
        """读取航班更新数据 - 简化版"""
        try:
            # 尝试读取航班工作表
            xl_file = pd.ExcelFile(file_path)
            if "航班" in xl_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name="航班")
            else:
                df = pd.read_excel(file_path)

            df.columns = df.columns.str.strip()
            df["数据源"] = "航班更新"

            # 添加状态处理
            if "变更操作" in df.columns:
                confirmed_operations = ["登机", "值机", "进检"]
                planned_operations = ["出票", "座变", "改期"]
                df["状态类型"] = df["变更操作"].apply(
                    lambda x: (
                        "已确认"
                        if x in confirmed_operations
                        else "待确认" if x in planned_operations else "其他"
                    )
                )

            # 映射列名
            column_mapping = {
                "航班号": "航班车次",
                "出发机场名称": "发站",
                "到达机场名称": "到站",
                "起飞时间": "出发时间",
            }
            for old_col, new_col in column_mapping.items():
                if old_col in df.columns and new_col not in df.columns:
                    df[new_col] = df[old_col]

            # 处理日期时间
            if "起飞时间" in df.columns:
                df["起飞时间_dt"] = pd.to_datetime(df["起飞时间"], errors="coerce")
                df["出发日期"] = df["起飞时间_dt"].dt.normalize()
                df["出发时间"] = df["起飞时间_dt"].dt.strftime("%H:%M")
            elif "出发日期" not in df.columns:
                df["出发日期"] = pd.NaT

            return df
        except Exception as e:
            raise Exception(f"读取航班数据失败：{str(e)}")

    def merge_data(self, df1, df2):
        """合并两份数据 - 简化版"""
        if df1 is None or df1.empty:
            if df2 is None or df2.empty:
                return pd.DataFrame()
            return df2

        if df2 is None or df2.empty:
            return df1

        # 票务全库添加默认字段
        if "变更操作" not in df1.columns:
            df1["变更操作"] = "票务记录"
        if "状态类型" not in df1.columns:
            df1["状态类型"] = "待确认"

        # 合并数据
        all_data = pd.concat([df1, df2], ignore_index=True)

        # 确保人员类型字段存在且有默认值
        if "人员类型" not in all_data.columns:
            all_data["人员类型"] = "未知"
        else:
            # 填充空值
            all_data["人员类型"] = all_data["人员类型"].fillna("未知")
            all_data["人员类型"] = all_data["人员类型"].replace("", "未知")

        # 基本去重
        dedup_columns = ["姓名", "证件号", "航班车次", "出发日期"]
        dedup_cols_exist = [col for col in dedup_columns if col in all_data.columns]

        if dedup_cols_exist:
            all_data = all_data.drop_duplicates(subset=dedup_cols_exist, keep="first")

        return all_data

    def read_mixed_transport_data(self, file_path):
        """简化版混合交通数据读取"""
        if not file_path:
            return pd.DataFrame()

        try:
            xl_file = pd.ExcelFile(file_path)
            all_data = []

            # 读取铁路数据
            if "铁路" in xl_file.sheet_names:
                railway_df = pd.read_excel(file_path, sheet_name="铁路")
                railway_df.columns = railway_df.columns.str.strip()

                # 基本字段映射
                if "证件编号" in railway_df.columns:
                    railway_df["证件号"] = railway_df["证件编号"]
                if "车次" in railway_df.columns:
                    railway_df["航班车次"] = railway_df["车次"]
                if "乘车日期" in railway_df.columns:
                    railway_df["出发日期"] = pd.to_datetime(
                        railway_df["乘车日期"], errors="coerce"
                    )
                if "乘车时间" in railway_df.columns:
                    railway_df["出发时间"] = railway_df["乘车时间"]

                railway_df["数据源"] = "铁路票务"
                railway_df["变更操作"] = "票务记录"
                railway_df["状态类型"] = "待确认"
                all_data.append(railway_df)

            # 读取航班数据
            if "航班" in xl_file.sheet_names:
                flight_df = pd.read_excel(file_path, sheet_name="航班")
                flight_df.columns = flight_df.columns.str.strip()

                # 基本字段映射
                if "航班号" in flight_df.columns:
                    flight_df["航班车次"] = flight_df["航班号"]
                if "出发机场名称" in flight_df.columns:
                    flight_df["发站"] = flight_df["出发机场名称"]
                if "到达机场名称" in flight_df.columns:
                    flight_df["到站"] = flight_df["到达机场名称"]
                if "起飞时间" in flight_df.columns:
                    flight_df["起飞时间_dt"] = pd.to_datetime(
                        flight_df["起飞时间"], errors="coerce"
                    )
                    flight_df["出发日期"] = flight_df["起飞时间_dt"].dt.normalize()
                    flight_df["出发时间"] = flight_df["起飞时间_dt"].dt.strftime(
                        "%H:%M"
                    )

                # 处理状态类型
                if "变更操作" in flight_df.columns:
                    confirmed_operations = ["登机", "值机", "进检"]
                    planned_operations = ["出票", "座变", "改期"]
                    flight_df["状态类型"] = flight_df["变更操作"].apply(
                        lambda x: (
                            "已确认"
                            if x in confirmed_operations
                            else "待确认" if x in planned_operations else "其他"
                        )
                    )

                flight_df["数据源"] = "航班更新"
                all_data.append(flight_df)

            # 合并所有数据
            if all_data:
                combined_df = pd.concat(all_data, ignore_index=True, sort=False)
                return combined_df
            else:
                return pd.DataFrame()

        except Exception as e:
            raise Exception(f"读取混合交通数据失败：{str(e)}")

    def final_dedup(self, df):
        """最终去重"""
        dedup_columns = ["姓名", "证件号", "航班车次", "出发日期"]
        dedup_cols_exist = [col for col in dedup_columns if col in df.columns]

        if dedup_cols_exist and "变更操作" in df.columns:
            df["优先级"] = df["变更操作"].map(self.status_priority).fillna(99)
            df = df.sort_values(by=dedup_cols_exist + ["优先级"])
            df = df.drop_duplicates(subset=dedup_cols_exist, keep="first")
            df = df.drop(columns=["优先级"])
        elif dedup_cols_exist:
            df = df.drop_duplicates(subset=dedup_cols_exist, keep="first")

        return df


class DataProcessor(QThread):
    progress = Signal(int)
    message = Signal(str)
    finished = Signal(pd.DataFrame)
    error = Signal(str)

    def __init__(self):
        super().__init__()
        self.file1_path = ""
        self.file2_path = ""
        self.start_date = None  # 改为开始日期
        self.end_date = None  # 新增结束日期
        self.target_city = ""
        self.min_people_count = 3  # 默认3人及以上
        self.existing_data = None  # 已存在的数据
        self.append_mode = False  # 是否为追加模式
        self.selected_person_types = []  # 新增：选中的人员类型列表

        # 定义状态优先级（数字越小优先级越高）
        self.status_priority = {
            "登机": 1,
            "值机": 2,
            "进检": 3,
            "出票": 4,
            "座变": 5,
            "改期": 6,
            "段消": 7,
            "换开": 8,
            "证变": 9,
            "值拉": 10,
            "票务记录": 11,  # 票务全库的默认状态
        }

    def set_params(
        self,
        file1,
        file2,
        start_date,
        end_date,
        target_city,
        min_people_count,
        existing_data=None,
        append_mode=False,
        selected_person_types=None,
    ):
        self.file1_path = file1
        self.file2_path = file2
        self.start_date = start_date  # 开始日期
        self.end_date = end_date  # 结束日期
        self.target_city = target_city
        self.min_people_count = min_people_count
        self.existing_data = existing_data
        self.append_mode = append_mode
        self.selected_person_types = selected_person_types or []  # 设置人员类型筛选

    def run(self):
        """简化的数据处理流程 - 针对两种固定表格格式优化"""
        try:
            self.progress.emit(10)
            self.message.emit("正在读取数据...")

            all_data = pd.DataFrame()

            # 处理票务全库数据（file1）
            if self.file1_path:
                self.message.emit("正在处理票务全库数据...")
                df1 = self.read_ticket_data(self.file1_path)
                if not df1.empty:
                    all_data = pd.concat([all_data, df1], ignore_index=True)
                    print(f"票务全库数据：{len(df1)} 条记录")
                self.progress.emit(30)

            # 处理群体票务数据（file2）- 包含铁路+航班
            if self.file2_path:
                self.message.emit("正在处理群体票务数据（铁路+航班）...")
                df2 = self.read_mixed_transport_data(self.file2_path)
                if not df2.empty:
                    all_data = pd.concat([all_data, df2], ignore_index=True)
                    print(f"群体票务数据：{len(df2)} 条记录")
                self.progress.emit(60)

            # 验证数据
            if all_data.empty:
                self.error.emit("没有读取到有效数据")
                return

            # 检查必要字段
            required_fields = ["姓名", "证件号", "航班车次", "出发日期", "到站"]
            missing_fields = [
                field for field in required_fields if field not in all_data.columns
            ]
            if missing_fields:
                self.error.emit(f"数据缺少必要字段：{', '.join(missing_fields)}")
                return

            self.progress.emit(70)
            self.message.emit("正在去重和整理数据...")

            # 统一去重处理
            all_data = self.final_dedup(all_data)
            print(f"去重后总数据：{len(all_data)} 条记录")

            # 如果有历史数据，进行合并
            if self.existing_data is not None and not self.existing_data.empty:
                self.message.emit("正在合并历史数据...")
                all_data = pd.concat([self.existing_data, all_data], ignore_index=True)
                all_data = self.final_dedup(all_data)
                print(f"合并历史数据后：{len(all_data)} 条记录")

            # 保存合并后的全量数据
            self.all_data = all_data

            self.progress.emit(80)
            self.message.emit("正在筛选目标人员...")

            # 筛选数据
            filtered_data = self.filter_data(
                all_data, self.start_date, self.end_date, self.target_city
            )

            self.progress.emit(90)
            self.message.emit("正在识别群体出行...")

            # 识别群体出行
            result = self.identify_groups(filtered_data)

            self.progress.emit(100)
            self.message.emit("筛查完成！")
            self.finished.emit(result)

        except Exception as e:
            self.error.emit(f"处理出错：{str(e)}")

    def read_ticket_data(self, file_path):
        """读取票务全库数据"""
        if not file_path:
            return pd.DataFrame()

        try:
            df = pd.read_excel(file_path)
            # 标准化列名
            df.columns = df.columns.str.strip()

            # 添加数据源标记
            df["数据源"] = "票务全库"

            # 确保必要的列存在
            required_cols = [
                "姓名",
                "证件号",
                "航班车次",
                "发站",
                "到站",
                "出发日期",
                "出发时间",
            ]
            # 保留的可选列（如果存在的话）
            optional_cols = ["人员类型", "方向", "交通工具", "入库时间"]

            for col in required_cols:
                if col not in df.columns:
                    if "航班号" in df.columns and col == "航班车次":
                        df["航班车次"] = df["航班号"]
                    elif "车次" in df.columns and col == "航班车次":
                        df["航班车次"] = df["车次"]
                    else:
                        df[col] = None

            # 处理可选列：如果存在就保留，不存在就不添加
            for col in optional_cols:
                if col in df.columns:
                    # 清理空值，但保留列
                    df[col] = df[col].fillna("").astype(str).str.strip()
                    print(f"保留可选字段：{col}，共有 {df[col].nunique()} 种不同值")

            # 如果存在人员类型字段，打印统计信息
            if "人员类型" in df.columns:
                person_type_counts = df["人员类型"].value_counts()
                print("票务数据中的人员类型分布：")
                for ptype, count in person_type_counts.items():
                    if ptype and ptype.strip():  # 只显示非空值
                        print(f"  {ptype}: {count} 条")

            # 数据验证
            # 1. 检查必填字段
            critical_cols = ["姓名", "证件号", "航班车次"]
            for col in critical_cols:
                if col in df.columns:
                    # 删除该列为空的记录
                    before_count = len(df)
                    df = df[df[col].notna() & (df[col].astype(str).str.strip() != "")]
                    after_count = len(df)
                    if before_count > after_count:
                        print(f"删除{col}为空的记录：{before_count - after_count}条")

            # 2. 处理姓名和航班车次的空值
            if "姓名" in df.columns:
                df["姓名"] = df["姓名"].fillna("").astype(str).str.strip()
            if "航班车次" in df.columns:
                df["航班车次"] = df["航班车次"].fillna("").astype(str).str.strip()
            if "证件号" in df.columns:
                df["证件号"] = df["证件号"].fillna("").astype(str).str.strip()

            # 处理日期格式 - 统一使用pandas datetime
            if "出发日期" in df.columns:
                df["出发日期"] = pd.to_datetime(df["出发日期"], errors="coerce")

            # 去重处理
            original_count = len(df)
            # 基于关键字段去重：姓名、证件号、航班车次、出发日期
            dedup_columns = ["姓名", "证件号", "航班车次", "出发日期"]
            # 确保去重列都存在
            dedup_cols_exist = [col for col in dedup_columns if col in df.columns]

            if dedup_cols_exist:
                df = df.drop_duplicates(subset=dedup_cols_exist, keep="first")
                after_count = len(df)
                if original_count > after_count:
                    print(
                        f"票务全库数据去重：{original_count} 条 -> {after_count} 条（删除 {original_count - after_count} 条重复）"
                    )
                else:
                    print(f"票务全库数据无重复记录：{original_count} 条")

            return df
        except Exception as e:
            raise Exception(f"读取票务全库数据失败：{str(e)}")

    def read_flight_data(self, file_path):
        """读取航班更新数据"""
        if not file_path:
            return pd.DataFrame()

        try:
            # 尝试读取航班工作表
            xl_file = pd.ExcelFile(file_path)
            if "航班" in xl_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name="航班")
            else:
                # 如果没有航班工作表，读取第一个工作表
                df = pd.read_excel(file_path)

            # 标准化列名
            df.columns = df.columns.str.strip()

            # 保留变更操作信息，不再过滤
            if "变更操作" in df.columns:
                # 统计各种变更操作的数量
                operation_counts = df["变更操作"].value_counts()
                print("航班数据变更操作统计：")
                for op, count in operation_counts.items():
                    print(f"  {op}: {count} 条")

                # 添加状态类型标记
                confirmed_operations = ["登机", "值机", "进检"]
                planned_operations = ["出票", "座变", "改期"]

                df["状态类型"] = df["变更操作"].apply(
                    lambda x: (
                        "已确认"
                        if x in confirmed_operations
                        else "待确认" if x in planned_operations else "其他"
                    )
                )

                # 统计状态类型
                status_counts = df["状态类型"].value_counts()
                print("\n状态类型统计：")
                for status, count in status_counts.items():
                    print(f"  {status}: {count} 条")
            else:
                df["变更操作"] = "未知"
                df["状态类型"] = "其他"

            # 添加数据源标记
            df["数据源"] = "航班更新"

            # 数据验证（与票务数据处理一致）
            # 1. 检查必填字段
            critical_cols = ["姓名", "证件号"]
            for col in critical_cols:
                if col in df.columns:
                    # 删除该列为空的记录
                    before_count = len(df)
                    df = df[df[col].notna() & (df[col].astype(str).str.strip() != "")]
                    after_count = len(df)
                    if before_count > after_count:
                        print(
                            f"航班数据：删除{col}为空的记录：{before_count - after_count}条"
                        )

            # 映射列名
            column_mapping = {
                "航班号": "航班车次",
                "出发机场名称": "发站",
                "到达机场名称": "到站",
                "起飞时间": "出发时间",
            }

            for old_col, new_col in column_mapping.items():
                if old_col in df.columns and new_col not in df.columns:
                    df[new_col] = df[old_col]

            # 保留的可选列（如果存在的话）
            optional_cols = ["人员类型", "方向", "交通工具"]
            for col in optional_cols:
                if col in df.columns:
                    # 清理空值，但保留列
                    df[col] = df[col].fillna("").astype(str).str.strip()
                    print(
                        f"航班数据保留可选字段：{col}，共有 {df[col].nunique()} 种不同值"
                    )

            # 如果存在人员类型字段，打印统计信息
            if "人员类型" in df.columns:
                person_type_counts = df["人员类型"].value_counts()
                print("航班数据中的人员类型分布：")
                for ptype, count in person_type_counts.items():
                    if ptype and ptype.strip():  # 只显示非空值
                        print(f"  {ptype}: {count} 条")

            # 2. 处理姓名、证件号和航班车次的空值
            if "姓名" in df.columns:
                df["姓名"] = df["姓名"].fillna("").astype(str).str.strip()
            if "证件号" in df.columns:
                df["证件号"] = df["证件号"].fillna("").astype(str).str.strip()
            if "航班车次" in df.columns:
                df["航班车次"] = df["航班车次"].fillna("").astype(str).str.strip()

            # 处理日期时间
            if "起飞时间" in df.columns:
                # 将起飞时间转换为datetime
                df["起飞时间_dt"] = pd.to_datetime(df["起飞时间"], errors="coerce")
                # 提取日期部分 - 保持为datetime类型，不转换为date
                df["出发日期"] = df["起飞时间_dt"].dt.normalize()
                # 提取时间部分
                df["出发时间"] = df["起飞时间_dt"].dt.strftime("%H:%M")
            elif "出发日期" not in df.columns:
                # 如果没有出发日期列，尝试从其他列获取
                df["出发日期"] = pd.NaT

            # 添加优先级列
            if "变更操作" in df.columns:
                df["优先级"] = (
                    df["变更操作"].map(self.status_priority).fillna(99)
                )  # 未知状态设为99
            else:
                df["优先级"] = 99

            # 按优先级去重处理
            original_count = len(df)

            # 基于关键字段去重：姓名、证件号、航班车次、出发日期
            dedup_columns = ["姓名", "证件号", "航班车次", "出发日期"]
            # 确保去重列都存在
            dedup_cols_exist = [col for col in dedup_columns if col in df.columns]

            if dedup_cols_exist and "优先级" in df.columns:
                # 出发日期已经在前面处理过，这里不需要再次转换
                # 先按关键字段和优先级排序，确保优先级高的排在前面
                df = df.sort_values(by=dedup_cols_exist + ["优先级"])
                # 保留每组中优先级最高（数字最小）的记录
                df = df.drop_duplicates(subset=dedup_cols_exist, keep="first")
                # 删除临时的优先级列
                df = df.drop(columns=["优先级"])

                after_count = len(df)
                if original_count > after_count:
                    print(
                        f"航班数据按优先级去重：{original_count} 条 -> {after_count} 条（删除 {original_count - after_count} 条重复）"
                    )
                    # 统计保留的状态分布
                    if "变更操作" in df.columns:
                        print("保留记录的状态分布：")
                        status_counts = df["变更操作"].value_counts()
                        for status, count in status_counts.items():
                            print(f"  {status}: {count} 条")
                else:
                    print(f"航班数据无重复记录：{original_count} 条")
            elif dedup_cols_exist:
                # 如果没有优先级信息，保持原有的去重逻辑
                df = df.drop_duplicates(subset=dedup_cols_exist, keep="first")
                after_count = len(df)
                if original_count > after_count:
                    print(
                        f"航班数据去重：{original_count} 条 -> {after_count} 条（删除 {original_count - after_count} 条重复）"
                    )
                else:
                    print(f"航班数据无重复记录：{original_count} 条")

            print(f"\n航班数据处理完成，共 {len(df)} 条记录（全部保留）")

            return df
        except Exception as e:
            raise Exception(f"读取航班更新数据失败：{str(e)}")

    def merge_data(self, df1, df2):
        """合并两份数据并智能去重"""
        print("\n开始合并数据...")

        # 处理单文件情况
        if df1 is None or df1.empty:
            if df2 is None or df2.empty:
                return pd.DataFrame()
            # 只有航班数据
            print("只有航班更新数据")
            return self.process_single_dataframe(df2, "航班更新")

        if df2 is None or df2.empty:
            # 只有票务数据
            print("只有票务全库数据")
            return self.process_single_dataframe(df1, "票务全库")

        # 以下是两个文件都存在的情况
        # 票务全库添加默认字段
        if "变更操作" not in df1.columns:
            df1["变更操作"] = "票务记录"
        if "状态类型" not in df1.columns:
            df1["状态类型"] = "待确认"

        # 选择需要的列
        cols = [
            "姓名",
            "证件号",
            "航班车次",
            "发站",
            "到站",
            "出发日期",
            "出发时间",
            "数据源",
            "变更操作",
            "状态类型",
            "人员类型",  # 新增：保留人员类型字段
        ]

        # 获取可用的列
        df1_cols = [col for col in cols if col in df1.columns]
        df2_cols = [col for col in cols if col in df2.columns]

        df1_subset = df1[df1_cols].copy()
        df2_subset = df2[df2_cols].copy()

        # 安全地转换字符串，处理空值
        def safe_str(val):
            """安全地将值转换为字符串，空值转为空字符串"""
            if pd.isna(val):
                return ""
            return str(val).strip()

        # 创建用于识别交集的键（包含日期维度）
        df1_subset["合并键"] = (
            df1_subset["姓名"].apply(safe_str)
            + "_"
            + df1_subset["证件号"].apply(safe_str)
            + "_"
            + df1_subset["航班车次"].apply(safe_str)
            + "_"
            + df1_subset["出发日期"].dt.strftime("%Y-%m-%d").fillna("")
        )
        df2_subset["合并键"] = (
            df2_subset["姓名"].apply(safe_str)
            + "_"
            + df2_subset["证件号"].apply(safe_str)
            + "_"
            + df2_subset["航班车次"].apply(safe_str)
            + "_"
            + df2_subset["出发日期"].dt.strftime("%Y-%m-%d").fillna("")
        )

        # 识别交集
        intersection_keys = set(df1_subset["合并键"]) & set(df2_subset["合并键"])

        print(f"票务全库记录数: {len(df1_subset)}")
        print(f"航班更新记录数: {len(df2_subset)}")
        print(f"交集记录数: {len(intersection_keys)}")

        # 处理交集：优先使用已确认状态的记录
        result_list = []

        # 处理交集记录
        for key in intersection_keys:
            df1_record = df1_subset[df1_subset["合并键"] == key].iloc[0]
            df2_record = df2_subset[df2_subset["合并键"] == key].iloc[0]

            # 如果航班表显示"已确认"状态，使用航班表数据
            if df2_record.get("状态类型") == "已确认":
                # 使用航班数据，但如果航班数据的人员类型为空，优先使用票务数据的人员类型
                if (
                    pd.isna(df2_record.get("人员类型"))
                    or df2_record.get("人员类型") == ""
                ):
                    if (
                        not pd.isna(df1_record.get("人员类型"))
                        and df1_record.get("人员类型") != ""
                    ):
                        df2_record["人员类型"] = df1_record.get("人员类型")
                result_list.append(df2_record)
            else:
                # 否则保留票务全库的记录，但更新变更操作信息
                df1_record["变更操作"] = df2_record.get(
                    "变更操作", df1_record.get("变更操作", "票务记录")
                )
                df1_record["状态类型"] = df2_record.get(
                    "状态类型", df1_record.get("状态类型", "待确认")
                )
                # 人员类型优先使用票务数据，如果为空则使用航班数据
                if (
                    pd.isna(df1_record.get("人员类型"))
                    or df1_record.get("人员类型") == ""
                ):
                    if (
                        not pd.isna(df2_record.get("人员类型"))
                        and df2_record.get("人员类型") != ""
                    ):
                        df1_record["人员类型"] = df2_record.get("人员类型")
                result_list.append(df1_record)

        # 添加两表独有的记录
        df1_unique = df1_subset[~df1_subset["合并键"].isin(intersection_keys)]
        df2_unique = df2_subset[~df2_subset["合并键"].isin(intersection_keys)]

        print(f"票务全库独有记录: {len(df1_unique)}")
        print(f"航班更新独有记录: {len(df2_unique)}")

        # 合并所有数据
        all_data = pd.concat(
            [pd.DataFrame(result_list), df1_unique, df2_unique], ignore_index=True
        )

        # 删除合并键列
        all_data = all_data.drop(columns=["合并键"])

        # 确保人员类型字段存在且有默认值
        if "人员类型" not in all_data.columns:
            all_data["人员类型"] = "未知"
        else:
            # 填充空值
            all_data["人员类型"] = all_data["人员类型"].fillna("未知")
            all_data["人员类型"] = all_data["人员类型"].replace("", "未知")

        print(f"合并后总记录数: {len(all_data)}")

        # 最终去重检查（按优先级）
        # 基于关键字段进行最终去重
        dedup_columns = ["姓名", "证件号", "航班车次", "出发日期"]
        dedup_cols_exist = [col for col in dedup_columns if col in all_data.columns]

        if dedup_cols_exist and "变更操作" in all_data.columns:
            before_final_dedup = len(all_data)

            # 添加优先级列
            all_data["优先级"] = (
                all_data["变更操作"].map(self.status_priority).fillna(99)
            )

            # 按关键字段和优先级排序
            all_data = all_data.sort_values(by=dedup_cols_exist + ["优先级"])

            # 保留每组中优先级最高的记录
            all_data = all_data.drop_duplicates(subset=dedup_cols_exist, keep="first")

            # 删除临时的优先级列
            all_data = all_data.drop(columns=["优先级"])

            after_final_dedup = len(all_data)
            if before_final_dedup > after_final_dedup:
                print(
                    f"最终按优先级去重：{before_final_dedup} 条 -> {after_final_dedup} 条（删除 {before_final_dedup - after_final_dedup} 条重复）"
                )
        elif dedup_cols_exist:
            before_final_dedup = len(all_data)
            all_data = all_data.drop_duplicates(subset=dedup_cols_exist, keep="first")
            after_final_dedup = len(all_data)
            if before_final_dedup > after_final_dedup:
                print(
                    f"最终去重：{before_final_dedup} 条 -> {after_final_dedup} 条（删除 {before_final_dedup - after_final_dedup} 条重复）"
                )

        # 统计合并后的状态分布
        if "状态类型" in all_data.columns:
            status_dist = all_data["状态类型"].value_counts()
            print("\n合并后状态分布：")
            for status, count in status_dist.items():
                print(f"  {status}: {count} 条")

        return all_data

    def process_single_dataframe(self, df, source_type):
        """处理单个数据框的情况"""
        if source_type == "票务全库":
            # 票务全库添加默认字段
            if "变更操作" not in df.columns:
                df["变更操作"] = "票务记录"
            if "状态类型" not in df.columns:
                df["状态类型"] = "待确认"

        # 选择需要的列
        cols = [
            "姓名",
            "证件号",
            "航班车次",
            "发站",
            "到站",
            "出发日期",
            "出发时间",
            "数据源",
            "变更操作",
            "状态类型",
            "人员类型",  # 新增：保留人员类型字段
        ]

        # 获取可用的列
        df_cols = [col for col in cols if col in df.columns]
        result = df[df_cols].copy()

        # 确保人员类型字段存在且有默认值
        if "人员类型" not in result.columns:
            result["人员类型"] = "未知"
        else:
            # 填充空值
            result["人员类型"] = result["人员类型"].fillna("未知")
            result["人员类型"] = result["人员类型"].replace("", "未知")

        print(f"{source_type}记录数: {len(result)}")

        # 统计状态分布
        if "状态类型" in result.columns:
            status_dist = result["状态类型"].value_counts()
            print(f"\n{source_type}状态分布：")
            for status, count in status_dist.items():
                print(f"  {status}: {count} 条")

        return result

    def final_dedup(self, df):
        """最终去重，用于追加数据后的去重"""
        # 基于关键字段进行最终去重
        dedup_columns = ["姓名", "证件号", "航班车次", "出发日期"]
        dedup_cols_exist = [col for col in dedup_columns if col in df.columns]

        if dedup_cols_exist and "变更操作" in df.columns:
            before_dedup = len(df)

            # 添加优先级列
            df["优先级"] = df["变更操作"].map(self.status_priority).fillna(99)

            # 按关键字段和优先级排序
            df = df.sort_values(by=dedup_cols_exist + ["优先级"])

            # 保留每组中优先级最高的记录
            df = df.drop_duplicates(subset=dedup_cols_exist, keep="first")

            # 删除临时的优先级列
            df = df.drop(columns=["优先级"])

            after_dedup = len(df)
            if before_dedup > after_dedup:
                print(
                    f"追加数据去重：{before_dedup} 条 -> {after_dedup} 条（删除 {before_dedup - after_dedup} 条重复）"
                )

        return df

    def filter_data(self, df, start_date, end_date, target_city):
        """根据日期和城市筛选数据"""
        # 判断是单日期还是日期区间
        if start_date == end_date:
            print(f"开始筛选数据，目标日期: {start_date}, 目标城市: {target_city}")
        else:
            print(
                f"开始筛选数据，日期区间: {start_date} 到 {end_date}, 目标城市: {target_city}"
            )

        print(f"筛选前数据量: {len(df)}")

        # 创建日期筛选条件
        if start_date == end_date:
            # 单日期筛选（保持向后兼容）
            date_filter = df["出发日期"].dt.date == start_date
        else:
            # 日期区间筛选
            date_filter = (df["出发日期"].dt.date >= start_date) & (
                df["出发日期"].dt.date <= end_date
            )

        df_filtered = df[date_filter]

        print(f"日期筛选后数据量: {len(df_filtered)}")

        # 筛选目标城市
        if target_city == "北京":
            city_keywords = ["北京", "首都", "大兴"]
        else:  # 福州
            city_keywords = ["福州", "长乐"]

        # 创建城市筛选条件
        city_filter = (
            df_filtered["到站"]
            .astype(str)
            .str.contains("|".join(city_keywords), case=False, na=False)
        )
        result = df_filtered[city_filter]

        print(f"城市筛选后数据量: {len(result)}")

        # 人员类型筛选
        if self.selected_person_types and "人员类型" in result.columns:
            # 如果选择了人员类型，进行筛选
            before_person_filter = len(result)
            person_type_filter = result["人员类型"].isin(self.selected_person_types)
            result = result[person_type_filter]
            print(
                f"人员类型筛选后数据量: {len(result)}（筛选条件：{', '.join(self.selected_person_types)}）"
            )

            # 显示筛选后的人员类型分布
            if len(result) > 0:
                person_type_dist = result["人员类型"].value_counts()
                print("筛选后人员类型分布：")
                for ptype, count in person_type_dist.items():
                    print(f"  {ptype}: {count} 条")
        elif self.selected_person_types:
            print("警告：选择了人员类型筛选，但数据中没有'人员类型'字段")

        # 添加目标城市列，用于群体分组
        result = result.copy()  # 避免修改原始数据
        result["目标城市"] = target_city

        return result

    def identify_groups(self, df):
        """简化的人员筛选（不再分群体，只要总人数达到阈值就全部筛出）"""
        if df.empty:
            return pd.DataFrame()

        # 新逻辑：不再分组，只统计总人数是否达到阈值
        total_people = len(df)

        # 如果总人数达到最少人数要求，返回所有数据
        if total_people >= self.min_people_count:
            result = df.copy()
            # 移除分组相关字段，不再需要
            # result["分组键"] = "全部人员"  # 可选：保留一个统一的标识
            # result["群体人数"] = total_people  # 移除群体人数概念
        else:
            # 如果总人数不足，返回空结果
            result = pd.DataFrame()

        # 确保姓名列是字符串类型
        if not result.empty and "姓名" in result.columns:
            result["姓名"] = result["姓名"].astype(str)

        # 排序 - 按出发日期和姓名排序
        if not result.empty:
            result = result.sort_values(["出发日期", "姓名"])

        return result

    def read_mixed_transport_data(self, file_path):
        """读取包含多种交通方式的数据文件（铁路+航班）"""
        if not file_path:
            return pd.DataFrame()

        try:
            # 检查文件中的工作表
            xl_file = pd.ExcelFile(file_path)
            print(f"发现工作表: {xl_file.sheet_names}")

            all_transport_data = []

            # 处理铁路数据
            if "铁路" in xl_file.sheet_names:
                print("正在处理铁路数据...")
                railway_df = pd.read_excel(file_path, sheet_name="铁路")
                railway_df.columns = railway_df.columns.str.strip()

                # 铁路数据字段映射
                railway_mapping = {
                    "证件编号": "证件号",
                    "车次": "航班车次",
                    "乘车日期": "出发日期",
                    "乘车时间": "出发时间",
                }

                for old_col, new_col in railway_mapping.items():
                    if (
                        old_col in railway_df.columns
                        and new_col not in railway_df.columns
                    ):
                        railway_df[new_col] = railway_df[old_col]

                # 添加数据源和交通方式标记
                railway_df["数据源"] = "铁路票务"
                railway_df["交通方式"] = "铁路"
                railway_df["变更操作"] = "票务记录"
                railway_df["状态类型"] = "待确认"

                # 数据清理
                if "姓名" in railway_df.columns:
                    railway_df["姓名"] = (
                        railway_df["姓名"].fillna("").astype(str).str.strip()
                    )
                if "证件号" in railway_df.columns:
                    railway_df["证件号"] = (
                        railway_df["证件号"].fillna("").astype(str).str.strip()
                    )
                if "航班车次" in railway_df.columns:
                    railway_df["航班车次"] = (
                        railway_df["航班车次"].fillna("").astype(str).str.strip()
                    )

                # 处理日期
                if "出发日期" in railway_df.columns:
                    railway_df["出发日期"] = pd.to_datetime(
                        railway_df["出发日期"], errors="coerce"
                    )

                # 去重
                dedup_columns = ["姓名", "证件号", "航班车次", "出发日期"]
                dedup_cols_exist = [
                    col for col in dedup_columns if col in railway_df.columns
                ]
                if dedup_cols_exist:
                    original_count = len(railway_df)
                    railway_df = railway_df.drop_duplicates(
                        subset=dedup_cols_exist, keep="first"
                    )
                    print(f"铁路数据去重：{original_count} -> {len(railway_df)} 条")

                # 统计人员类型
                if "人员类型" in railway_df.columns:
                    type_counts = railway_df["人员类型"].value_counts()
                    print("铁路数据人员类型分布：")
                    for ptype, count in type_counts.items():
                        if ptype and str(ptype).strip():
                            print(f"  {ptype}: {count} 条")

                all_transport_data.append(railway_df)
                print(f"铁路数据处理完成：{len(railway_df)} 条记录")

            # 处理航班数据
            if "航班" in xl_file.sheet_names:
                print("正在处理航班数据...")
                flight_df = pd.read_excel(file_path, sheet_name="航班")
                flight_df.columns = flight_df.columns.str.strip()

                # 航班数据字段映射（与原有逻辑保持一致）
                flight_mapping = {
                    "航班号": "航班车次",
                    "出发机场名称": "发站",
                    "到达机场名称": "到站",
                    "起飞时间": "出发时间",
                }

                for old_col, new_col in flight_mapping.items():
                    if (
                        old_col in flight_df.columns
                        and new_col not in flight_df.columns
                    ):
                        flight_df[new_col] = flight_df[old_col]

                # 添加数据源和交通方式标记
                flight_df["数据源"] = "航班更新"
                flight_df["交通方式"] = "航班"

                # 处理变更操作和状态类型（与原有逻辑保持一致）
                if "变更操作" in flight_df.columns:
                    confirmed_operations = ["登机", "值机", "进检"]
                    planned_operations = ["出票", "座变", "改期"]

                    flight_df["状态类型"] = flight_df["变更操作"].apply(
                        lambda x: (
                            "已确认"
                            if x in confirmed_operations
                            else "待确认" if x in planned_operations else "其他"
                        )
                    )

                    # 统计变更操作
                    operation_counts = flight_df["变更操作"].value_counts()
                    print("航班数据变更操作统计：")
                    for op, count in operation_counts.items():
                        print(f"  {op}: {count} 条")
                else:
                    flight_df["变更操作"] = "未知"
                    flight_df["状态类型"] = "其他"

                # 数据清理
                critical_cols = ["姓名", "证件号"]
                for col in critical_cols:
                    if col in flight_df.columns:
                        before_count = len(flight_df)
                        flight_df = flight_df[
                            flight_df[col].notna()
                            & (flight_df[col].astype(str).str.strip() != "")
                        ]
                        after_count = len(flight_df)
                        if before_count > after_count:
                            print(
                                f"航班数据：删除{col}为空的记录：{before_count - after_count}条"
                            )

                # 处理字段
                if "姓名" in flight_df.columns:
                    flight_df["姓名"] = (
                        flight_df["姓名"].fillna("").astype(str).str.strip()
                    )
                if "证件号" in flight_df.columns:
                    flight_df["证件号"] = (
                        flight_df["证件号"].fillna("").astype(str).str.strip()
                    )
                if "航班车次" in flight_df.columns:
                    flight_df["航班车次"] = (
                        flight_df["航班车次"].fillna("").astype(str).str.strip()
                    )

                # 处理日期时间
                if "起飞时间" in flight_df.columns:
                    flight_df["起飞时间_dt"] = pd.to_datetime(
                        flight_df["起飞时间"], errors="coerce"
                    )
                    flight_df["出发日期"] = flight_df["起飞时间_dt"].dt.normalize()
                    flight_df["出发时间"] = flight_df["起飞时间_dt"].dt.strftime(
                        "%H:%M"
                    )
                elif "出发日期" not in flight_df.columns:
                    flight_df["出发日期"] = pd.NaT

                # 按优先级去重
                if "变更操作" in flight_df.columns:
                    flight_df["优先级"] = (
                        flight_df["变更操作"].map(self.status_priority).fillna(99)
                    )
                else:
                    flight_df["优先级"] = 99

                dedup_columns = ["姓名", "证件号", "航班车次", "出发日期"]
                dedup_cols_exist = [
                    col for col in dedup_columns if col in flight_df.columns
                ]

                if dedup_cols_exist and "优先级" in flight_df.columns:
                    original_count = len(flight_df)
                    flight_df = flight_df.sort_values(by=dedup_cols_exist + ["优先级"])
                    flight_df = flight_df.drop_duplicates(
                        subset=dedup_cols_exist, keep="first"
                    )
                    flight_df = flight_df.drop(columns=["优先级"])
                    print(
                        f"航班数据按优先级去重：{original_count} -> {len(flight_df)} 条"
                    )

                # 统计人员类型
                if "人员类型" in flight_df.columns:
                    type_counts = flight_df["人员类型"].value_counts()
                    print("航班数据人员类型分布：")
                    for ptype, count in type_counts.items():
                        if ptype and str(ptype).strip():
                            print(f"  {ptype}: {count} 条")

                all_transport_data.append(flight_df)
                print(f"航班数据处理完成：{len(flight_df)} 条记录")

            # 合并所有交通数据
            if not all_transport_data:
                return pd.DataFrame()

            combined_df = pd.concat(all_transport_data, ignore_index=True, sort=False)

            # 确保人员类型字段存在
            if "人员类型" not in combined_df.columns:
                combined_df["人员类型"] = "未知"
            else:
                combined_df["人员类型"] = combined_df["人员类型"].fillna("未知")
                combined_df["人员类型"] = combined_df["人员类型"].replace("", "未知")

            # 最终统计
            print(f"\n=== 混合交通数据处理完成 ===")
            print(f"总记录数: {len(combined_df)}")
            if "交通方式" in combined_df.columns:
                transport_counts = combined_df["交通方式"].value_counts()
                print("交通方式分布：")
                for transport, count in transport_counts.items():
                    print(f"  {transport}: {count} 条")

            if "状态类型" in combined_df.columns:
                status_counts = combined_df["状态类型"].value_counts()
                print("状态类型分布：")
                for status, count in status_counts.items():
                    print(f"  {status}: {count} 条")

            return combined_df

        except Exception as e:
            raise Exception(f"读取混合交通数据失败：{str(e)}")


class GroupTravelChecker(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file1_path = ""
        self.file2_path = ""
        self.result_data = None
        self.processor = DataProcessor()
        self.preview_loader = DataPreviewLoader()

        self.merged_data = None
        self.append_mode = False
        self.original_data_count = 0
        self.has_file1_selected = False
        self.has_file2_selected = False

        self.all_merged_data = None
        self.filtered_data = None
        self.search_results = None
        self.search_history = []
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(self.perform_search)

        self.available_person_types = []
        self.person_type_checkboxes = {}

        self.init_ui()
        self.setup_connections()

    def init_ui(self):
        self.setWindowTitle("群体出行人员筛查系统 v1.1")
        self.setGeometry(100, 100, 1200, 700)
        self.setMinimumSize(800, 600)

        screen = QDesktopWidget().screenGeometry()
        if screen.width() < 1200 or screen.height() < 700:
            self.resize(
                min(screen.width() - 100, 1200), min(screen.height() - 100, 700)
            )
        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #f5f5f5;
                font-size: 16px;
            }
            QGroupBox {
                font-size: 18px;
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 12px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 10px 0 10px;
                font-size: 18px;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-size: 16px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
            QLabel {
                font-size: 16px;
            }
            QComboBox, QDateEdit {
                padding: 6px;
                border: 1px solid #cccccc;
                border-radius: 5px;
                font-size: 16px;
                min-height: 30px;
            }
            QLineEdit {
                font-size: 16px;
                min-height: 30px;
            }
            QTableWidget {
                font-size: 14px;
            }
            QTableWidget::item {
                padding: 8px;
            }
            QHeaderView::section {
                font-size: 16px;
                font-weight: bold;
                padding: 6px;
                min-height: 35px;
            }
            QTextEdit {
                font-size: 14px;
            }
            QCheckBox {
                font-size: 16px;
                spacing: 6px;
            }
            /* 滚动条样式 */
            QScrollBar:vertical {
                background: #f0f0f0;
                width: 15px;
                margin: 0;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                min-height: 30px;
                border-radius: 7px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a0a0a0;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0;
            }
        """
        )

        # 创建中心部件和滚动区域
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 创建主布局（用于容纳滚动区域）
        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(0, 0, 0, 0)

        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameStyle(QFrame.NoFrame)
        central_layout.addWidget(scroll_area)

        # 创建滚动内容容器
        scroll_content = QWidget()
        scroll_area.setWidget(scroll_content)

        # 主布局 - 现在添加到滚动内容容器中
        main_layout = QVBoxLayout(scroll_content)
        main_layout.setSpacing(15)  # 减小间距
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 标题
        title_label = QLabel("群体出行人员筛查系统")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(24)  # 减小标题字体
        title_font.setBold(True)
        title_label.setFont(title_font)
        main_layout.addWidget(title_label)

        # 文件选择区域
        file_group = QGroupBox("数据导入")
        file_layout = QVBoxLayout()

        # 添加数据状态显示
        self.data_status_label = QLabel("当前无数据")
        self.data_status_label.setStyleSheet(
            """
            QLabel {
                background-color: #f0f0f0;
                border: 1px solid #cccccc;
                border-radius: 5px;
                padding: 12px;
                font-size: 16px;
                color: #666666;
            }
        """
        )
        file_layout.addWidget(self.data_status_label)

        # 添加自动合并提示
        merge_tip_label = QLabel(
            "💡 提示：新选择的文件会自动与现有数据合并，无需额外操作"
        )
        merge_tip_label.setStyleSheet(
            """
            QLabel {
                color: #1976D2;
                font-size: 14px;
                padding: 8px;
                background-color: #E3F2FD;
                border-radius: 5px;
                margin: 5px 0;
            }
        """
        )
        file_layout.addWidget(merge_tip_label)
        file_layout.addSpacing(8)

        # 文件1选择
        file1_layout = QHBoxLayout()
        file1_label = QLabel("票务全库数据：")
        file1_label.setFixedWidth(150)
        file1_label.setToolTip("选择单工作表的票务全库Excel文件")
        self.file1_edit = QLabel("未选择文件")
        self.file1_edit.setStyleSheet(
            """
            QLabel {
                background-color: white;
                border: 1px solid #cccccc;
                border-radius: 5px;
                padding: 6px;
                font-size: 14px;
            }
        """
        )
        self.file1_btn = QPushButton("选择票务全库")
        self.file1_btn.setFixedWidth(120)
        self.file1_btn.setToolTip("选择包含完整票务信息的单工作表Excel文件")
        file1_layout.addWidget(file1_label)
        file1_layout.addWidget(self.file1_edit, 1)  # 让文件路径占据剩余空间
        file1_layout.addWidget(self.file1_btn)

        # 文件2选择
        file2_layout = QHBoxLayout()
        file2_label = QLabel("群体票务数据：")
        file2_label.setFixedWidth(150)
        file2_label.setToolTip("选择包含'铁路'和'航班'两个工作表的Excel文件")
        self.file2_edit = QLabel("未选择文件")
        self.file2_edit.setStyleSheet(
            """
            QLabel {
                background-color: white;
                border: 1px solid #cccccc;
                border-radius: 5px;
                padding: 6px;
                font-size: 14px;
            }
        """
        )
        self.file2_btn = QPushButton("选择群体票务")
        self.file2_btn.setFixedWidth(120)
        self.file2_btn.setToolTip("选择包含铁路+航班两个工作表的Excel文件")
        file2_layout.addWidget(file2_label)
        file2_layout.addWidget(self.file2_edit, 1)  # 让文件路径占据剩余空间
        file2_layout.addWidget(self.file2_btn)

        file_layout.addLayout(file1_layout)
        file_layout.addLayout(file2_layout)

        # 添加文件历史记录区域
        file_layout.addSpacing(8)
        history_label = QLabel("导入历史：")
        file_layout.addWidget(history_label)

        self.file_history_text = QTextEdit()
        self.file_history_text.setReadOnly(True)
        self.file_history_text.setMaximumHeight(80)
        self.file_history_text.setPlaceholderText("文件导入历史将显示在这里...")
        self.file_history_text.setStyleSheet(
            """
            QTextEdit {
                background-color: #fafafa;
                border: 1px solid #e0e0e0;
                border-radius: 5px;
                padding: 8px;
                font-size: 13px;
            }
        """
        )
        file_layout.addWidget(self.file_history_text)

        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        search_group = QGroupBox("数据搜索")
        search_layout = QVBoxLayout()

        # 添加提示标签
        search_tip_label = QLabel("💡 提示：可在导入数据后任意阶段进行搜索，无需先筛查")
        search_tip_label.setStyleSheet(
            """
            QLabel {
                color: #1976D2;
                font-size: 14px;
                padding: 8px;
                background-color: #E3F2FD;
                border-radius: 5px;
                margin-bottom: 8px;
            }
        """
        )
        search_layout.addWidget(search_tip_label)

        # 搜索控件行
        search_row = QHBoxLayout()

        # 搜索范围选择 - 新增功能
        search_scope_label = QLabel("搜索范围：")
        search_scope_label.setFixedWidth(100)
        self.search_scope_combo = QComboBox()
        self.search_scope_combo.addItems(["全部导入数据", "筛查结果"])
        self.search_scope_combo.setFixedWidth(150)
        self.search_scope_combo.setEnabled(False)  # 初始禁用

        # 搜索类型选择
        search_type_label = QLabel("搜索类型：")
        search_type_label.setFixedWidth(100)
        self.search_type_combo = QComboBox()
        self.search_type_combo.addItems(["姓名", "证件号", "航班车次", "全字段"])
        self.search_type_combo.setFixedWidth(150)

        # 搜索输入框
        search_input_label = QLabel("搜索内容：")
        search_input_label.setFixedWidth(100)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入搜索内容...")
        self.search_input.setStyleSheet(
            """
            QLineEdit {
                padding: 10px;
                border: 2px solid #cccccc;
                border-radius: 5px;
                font-size: 16px;
                background-color: white;
                min-height: 30px;
            }
            QLineEdit:focus {
                border-color: #2196F3;
            }
            QLineEdit:disabled {
                background-color: #f5f5f5;
                color: #999999;
            }
        """
        )

        # 搜索按钮
        self.search_data_btn = QPushButton("搜索")
        self.search_data_btn.setFixedWidth(100)
        self.search_data_btn.setStyleSheet(
            """
            QPushButton {
                background-color: #9C27B0;
                font-size: 16px;
                padding: 8px;
            }
            QPushButton:hover {
                background-color: #7B1FA2;
            }
        """
        )

        # 清除搜索按钮
        self.clear_search_btn = QPushButton("清除")
        self.clear_search_btn.setFixedWidth(100)
        self.clear_search_btn.setStyleSheet(
            """
            QPushButton {
                background-color: #607D8B;
                font-size: 16px;
                padding: 8px;
            }
            QPushButton:hover {
                background-color: #455A64;
            }
        """
        )

        # 搜索历史按钮
        self.search_history_btn = QPushButton("历史")
        self.search_history_btn.setFixedWidth(100)
        self.search_history_btn.setStyleSheet(
            """
            QPushButton {
                background-color: #795548;
                font-size: 16px;
                padding: 8px;
            }
            QPushButton:hover {
                background-color: #5D4037;
            }
        """
        )

        # 布局搜索控件
        search_row.addWidget(search_scope_label)
        search_row.addWidget(self.search_scope_combo)
        search_row.addSpacing(15)
        search_row.addWidget(search_type_label)
        search_row.addWidget(self.search_type_combo)
        search_row.addSpacing(15)
        search_row.addWidget(search_input_label)
        search_row.addWidget(self.search_input, 1)  # 让搜索框占据剩余空间
        search_row.addSpacing(15)
        search_row.addWidget(self.search_data_btn)
        search_row.addWidget(self.clear_search_btn)
        search_row.addWidget(self.search_history_btn)

        # 搜索状态和统计信息
        search_status_row = QHBoxLayout()
        self.search_status_label = QLabel("准备搜索")
        self.search_status_label.setStyleSheet(
            """
            QLabel {
                color: #666666;
                font-size: 18px;
                padding: 10px;
            }
        """
        )

        # 当前搜索范围指示标签
        self.search_scope_indicator = QLabel("")
        self.search_scope_indicator.setStyleSheet(
            """
            QLabel {
                color: #FF5722;
                font-size: 18px;
                font-weight: bold;
                padding: 10px;
            }
        """
        )

        search_status_row.addWidget(self.search_status_label)
        search_status_row.addWidget(self.search_scope_indicator)
        search_status_row.addStretch()

        search_layout.addLayout(search_row)
        search_layout.addLayout(search_status_row)
        search_group.setLayout(search_layout)

        # 初始状态：禁用搜索功能
        self.search_input.setEnabled(False)
        self.search_type_combo.setEnabled(False)
        self.search_data_btn.setEnabled(False)
        self.clear_search_btn.setEnabled(False)
        self.search_history_btn.setEnabled(False)

        main_layout.addWidget(search_group)

        # 筛选条件区域
        filter_group = QGroupBox("筛选条件")
        filter_layout = QVBoxLayout()

        # 第一行：时间选择模式和日期
        first_row = QHBoxLayout()

        # 时间选择模式
        time_mode_label = QLabel("时间模式：")
        time_mode_label.setFixedWidth(100)
        self.time_mode_combo = QComboBox()
        self.time_mode_combo.addItems(["单日期", "时间区间"])
        self.time_mode_combo.setFixedWidth(150)

        # 日期选择区域
        date_container = QWidget()
        self.date_layout = QHBoxLayout(date_container)
        self.date_layout.setContentsMargins(0, 0, 0, 0)

        # 单日期选择
        self.single_date_widget = QWidget()
        single_date_layout = QHBoxLayout(self.single_date_widget)
        single_date_layout.setContentsMargins(0, 0, 0, 0)

        single_date_label = QLabel("出发日期：")
        single_date_label.setFixedWidth(100)
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        self.date_edit.setFixedWidth(150)

        single_date_layout.addWidget(single_date_label)
        single_date_layout.addWidget(self.date_edit)
        single_date_layout.addStretch()

        # 日期区间选择
        self.date_range_widget = QWidget()
        date_range_layout = QHBoxLayout(self.date_range_widget)
        date_range_layout.setContentsMargins(0, 0, 0, 0)

        start_date_label = QLabel("开始日期：")
        start_date_label.setFixedWidth(100)
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(QDate.currentDate())
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_date_edit.setFixedWidth(150)

        end_date_label = QLabel("结束日期：")
        end_date_label.setFixedWidth(100)
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDate(QDate.currentDate().addDays(7))  # 默认7天后
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.end_date_edit.setFixedWidth(150)

        date_range_layout.addWidget(start_date_label)
        date_range_layout.addWidget(self.start_date_edit)
        date_range_layout.addSpacing(30)
        date_range_layout.addWidget(end_date_label)
        date_range_layout.addWidget(self.end_date_edit)
        date_range_layout.addStretch()

        # 初始状态设置
        self.date_range_widget.setVisible(False)  # 默认隐藏区间选择

        # 添加到日期布局
        self.date_layout.addWidget(self.single_date_widget)
        self.date_layout.addWidget(self.date_range_widget)
        self.date_layout.addStretch()

        first_row.addWidget(time_mode_label)
        first_row.addWidget(self.time_mode_combo)
        first_row.addSpacing(20)
        first_row.addWidget(date_container)
        first_row.addStretch()

        # 第二行：城市和最少人数
        second_row = QHBoxLayout()

        # 城市选择
        city_label = QLabel("目标城市：")
        city_label.setFixedWidth(100)
        self.city_combo = QComboBox()
        self.city_combo.addItems(["北京", "福州"])
        self.city_combo.setFixedWidth(150)

        # 最少人数选择
        people_label = QLabel("最少人数：")
        people_label.setFixedWidth(100)
        self.people_combo = QComboBox()
        self.people_combo.addItems(["1人及以上", "2人及以上", "3人及以上"])
        self.people_combo.setCurrentIndex(2)  # 默认选择"3人及以上"
        self.people_combo.setFixedWidth(150)

        second_row.addWidget(city_label)
        second_row.addWidget(self.city_combo)
        second_row.addSpacing(40)
        second_row.addWidget(people_label)
        second_row.addWidget(self.people_combo)
        second_row.addStretch()

        person_type_row = QHBoxLayout()

        # 人员类型筛选标签
        person_type_label = QLabel("人员类型：")
        person_type_label.setFixedWidth(100)

        # 全选/取消全选复选框
        self.select_all_person_types_cb = QCheckBox("全选")
        self.select_all_person_types_cb.setFixedWidth(80)
        self.select_all_person_types_cb.setChecked(True)  # 默认全选

        # 人员类型复选框容器
        self.person_type_container = QWidget()
        self.person_type_layout = QHBoxLayout(self.person_type_container)
        self.person_type_layout.setContentsMargins(0, 0, 0, 0)
        self.person_type_layout.setSpacing(12)

        # 人员类型复选框字典，用于管理
        self.person_type_checkboxes = {}

        # 人员类型发现状态标签
        self.person_type_status_label = QLabel("请先导入数据以发现人员类型")
        self.person_type_status_label.setStyleSheet(
            """
            QLabel {
                color: #888888;
                font-style: italic;
                font-size: 14px;
            }
        """
        )

        person_type_row.addWidget(person_type_label)
        person_type_row.addWidget(self.select_all_person_types_cb)
        person_type_row.addWidget(self.person_type_container)
        person_type_row.addWidget(self.person_type_status_label)
        person_type_row.addStretch()

        # 第三行：筛选选项和筛查按钮
        third_row = QHBoxLayout()

        # 筛查按钮
        self.search_btn = QPushButton("开始筛查")
        self.search_btn.setFixedWidth(150)
        self.search_btn.setFixedHeight(40)
        self.search_btn.setStyleSheet(
            """
            QPushButton {
                background-color: #2196F3;
                font-size: 18px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """
        )
        self.search_btn.setEnabled(False)  # 初始禁用，选择文件后启用

        # 清空数据按钮
        self.clear_btn = QPushButton("清空数据")
        self.clear_btn.setFixedWidth(150)
        self.clear_btn.setFixedHeight(40)
        self.clear_btn.setStyleSheet(
            """
            QPushButton {
                background-color: #F44336;
                font-size: 18px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #D32F2F;
            }
        """
        )
        self.clear_btn.setEnabled(False)  # 初始禁用，有数据后启用

        third_row.addWidget(self.search_btn)
        third_row.addSpacing(15)
        third_row.addWidget(self.clear_btn)
        third_row.addStretch()

        filter_layout.addLayout(first_row)
        filter_layout.addLayout(second_row)
        filter_layout.addLayout(person_type_row)
        filter_layout.addLayout(third_row)
        filter_group.setLayout(filter_layout)
        main_layout.addWidget(filter_group)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_label = QLabel("")
        self.progress_label.setVisible(False)
        main_layout.addWidget(self.progress_label)
        main_layout.addWidget(self.progress_bar)

        # 结果展示区域
        result_group = QGroupBox("筛查结果")
        result_layout = QVBoxLayout()

        # 统计信息
        self.stats_label = QLabel("请先导入数据并开始筛查")
        self.stats_label.setStyleSheet("font-size: 16px; color: #666666;")
        result_layout.addWidget(self.stats_label)

        # 结果表格
        self.result_table = QTableWidget()
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.result_table.horizontalHeader().setStretchLastSection(True)
        # 设置最小高度，确保能显示多行数据
        self.result_table.setMinimumHeight(400)  # 设置最小高度为400像素
        # 设置行高，让表格显示更紧凑
        self.result_table.verticalHeader().setDefaultSectionSize(35)
        # 设置表格的垂直滚动条策略
        self.result_table.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        result_layout.addWidget(self.result_table)

        # 导出按钮
        export_layout = QHBoxLayout()
        export_layout.addStretch()
        self.export_btn = QPushButton("导出结果")
        self.export_btn.setEnabled(False)
        self.export_btn.setFixedWidth(150)
        self.export_btn.setFixedHeight(40)
        export_layout.addWidget(self.export_btn)
        result_layout.addLayout(export_layout)

        result_group.setLayout(result_layout)
        main_layout.addWidget(result_group)

    def setup_connections(self):
        """设置信号连接"""
        self.file1_btn.clicked.connect(self.select_file1)
        self.file2_btn.clicked.connect(self.select_file2)
        self.search_btn.clicked.connect(self.start_search)
        self.clear_btn.clicked.connect(self.clear_data)

        # 时间模式切换
        self.time_mode_combo.currentTextChanged.connect(self.on_time_mode_changed)

        # 日期变化事件
        self.start_date_edit.dateChanged.connect(self.validate_date_range)
        self.end_date_edit.dateChanged.connect(self.validate_date_range)

        # 数据处理信号
        self.processor.progress.connect(self.update_progress)
        self.processor.message.connect(self.update_message)
        self.processor.finished.connect(self.show_results)
        self.processor.error.connect(self.show_error)

        # 新增：数据预加载信号
        self.preview_loader.progress.connect(self.update_progress)
        self.preview_loader.message.connect(self.update_message)
        self.preview_loader.finished.connect(self.on_data_preview_loaded)
        self.preview_loader.error.connect(self.show_preview_error)

        self.search_input.textChanged.connect(self.on_search_text_changed)
        self.search_data_btn.clicked.connect(self.search_data)
        self.clear_search_btn.clicked.connect(self.clear_search)
        self.search_history_btn.clicked.connect(self.show_search_history)
        self.search_scope_combo.currentTextChanged.connect(self.on_search_scope_changed)

        self.select_all_person_types_cb.stateChanged.connect(
            self.on_select_all_person_types_changed
        )

        self.result_table.cellClicked.connect(self.on_table_cell_clicked)

        # 导出按钮连接
        self.export_btn.clicked.connect(self.export_results)

    def discover_person_types(self, data_df):
        """发现数据中的人员类型"""
        if data_df is None or data_df.empty or "人员类型" not in data_df.columns:
            self.available_person_types = []
            self.update_person_type_ui(preserve_selection=False)
            return

        # 获取所有非空的人员类型
        person_types = data_df["人员类型"].dropna().astype(str).str.strip()
        person_types = person_types[person_types != ""].unique().tolist()
        person_types.sort()  # 排序

        # 检查是否已经有人员类型数据（避免重复创建UI）
        had_person_types = bool(self.available_person_types)

        self.available_person_types = person_types
        print(f"发现 {len(person_types)} 种人员类型：{person_types}")

        # 如果之前已经有人员类型，保持用户的选择状态
        self.update_person_type_ui(preserve_selection=had_person_types)

    def update_person_type_ui(self, preserve_selection=False):
        """更新人员类型UI

        Args:
            preserve_selection: 是否保持用户的选择状态，默认为False（初次创建时全选）
        """
        # 保存当前的选择状态（如果需要保持）
        current_selections = {}
        if preserve_selection and self.person_type_checkboxes:
            for person_type, checkbox in self.person_type_checkboxes.items():
                current_selections[person_type] = checkbox.isChecked()

        # 清理现有的复选框
        for checkbox in self.person_type_checkboxes.values():
            checkbox.setParent(None)
            checkbox.deleteLater()
        self.person_type_checkboxes.clear()

        if not self.available_person_types:
            self.person_type_status_label.setText(
                "数据中未发现人员类型字段或无有效数据"
            )
            self.person_type_status_label.setVisible(True)
            self.select_all_person_types_cb.setVisible(False)
            return

        # 更新状态标签
        self.person_type_status_label.setText(
            f"发现 {len(self.available_person_types)} 种人员类型"
        )
        self.person_type_status_label.setVisible(True)
        self.select_all_person_types_cb.setVisible(True)

        # 创建新的复选框
        for person_type in self.available_person_types:
            checkbox = QCheckBox(person_type)

            # 决定初始选择状态
            if preserve_selection and person_type in current_selections:
                # 保持之前的选择状态
                checkbox.setChecked(current_selections[person_type])
            else:
                # 新的人员类型默认选中，初次创建时全选
                checkbox.setChecked(True)

            checkbox.stateChanged.connect(self.on_person_type_selection_changed)

            self.person_type_layout.addWidget(checkbox)
            self.person_type_checkboxes[person_type] = checkbox

        # 更新全选框的状态（根据实际的选择情况）
        if self.person_type_checkboxes:
            all_checked = all(
                checkbox.isChecked()
                for checkbox in self.person_type_checkboxes.values()
            )
            self.select_all_person_types_cb.blockSignals(True)
            self.select_all_person_types_cb.setChecked(all_checked)
            self.select_all_person_types_cb.blockSignals(False)

    def on_select_all_person_types_changed(self, state):
        """全选/取消全选人员类型

        新的逻辑：
        - 只有两种状态：Qt.Checked（全选）和 Qt.Unchecked（全不选）
        - 勾选时：所有人员类型都被勾选
        - 取消勾选时：所有人员类型都被取消勾选
        - 任何时候点击全选复选框，都立即同步所有类型的勾选状态
        """
        checked = state == Qt.Checked
        for checkbox in self.person_type_checkboxes.values():
            checkbox.setChecked(checked)

    def on_person_type_selection_changed(self):
        """人员类型选择变化处理

        逻辑：
        - 当所有类型都被选中时，全选框自动勾选
        - 当有任何类型被取消时，全选框自动取消勾选
        - 保持全选框与单个类型选择的同步
        """
        if not self.person_type_checkboxes:
            return

        # 检查是否所有类型都被选中
        all_checked = all(
            checkbox.isChecked() for checkbox in self.person_type_checkboxes.values()
        )

        # 暂时断开全选框的信号，避免循环触发
        self.select_all_person_types_cb.blockSignals(True)

        # 根据所有类型的选择状态更新全选框
        self.select_all_person_types_cb.setChecked(all_checked)

        # 重新连接信号
        self.select_all_person_types_cb.blockSignals(False)

    def get_selected_person_types(self):
        """获取选中的人员类型列表"""
        selected_types = []
        for person_type, checkbox in self.person_type_checkboxes.items():
            if checkbox.isChecked():
                selected_types.append(person_type)
        return selected_types

    def select_file1(self):
        """选择文件1"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择票务全库数据", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.file1_path = file_path
            self.file1_edit.setText(os.path.basename(file_path))
            self.has_file1_selected = True
            self.update_button_states()

            # 新增：触发数据预加载
            self.trigger_data_preview()

    def select_file2(self):
        """选择文件2"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择铁路航班分开数据", "", "Excel文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.file2_path = file_path
            self.file2_edit.setText(os.path.basename(file_path))
            self.has_file2_selected = True
            self.update_button_states()

            # 新增：触发数据预加载
            self.trigger_data_preview()

    def on_time_mode_changed(self, mode):
        """时间模式切换处理"""
        if mode == "单日期":
            self.single_date_widget.setVisible(True)
            self.date_range_widget.setVisible(False)
        else:  # 时间区间
            self.single_date_widget.setVisible(False)
            self.date_range_widget.setVisible(True)
            # 确保日期范围合理
            self.validate_date_range()

    def validate_date_range(self):
        """验证日期范围的合理性"""
        if self.time_mode_combo.currentText() == "时间区间":
            start_date = self.start_date_edit.date()
            end_date = self.end_date_edit.date()

            # 如果开始日期晚于结束日期，自动调整
            if start_date > end_date:
                # 将结束日期设置为开始日期后7天
                self.end_date_edit.setDate(start_date.addDays(7))

                # 显示提示
                QMessageBox.information(
                    self, "日期调整", "开始日期不能晚于结束日期，已自动调整结束日期。"
                )

    def get_selected_dates(self):
        """获取选择的日期（单日期或日期区间）"""
        if self.time_mode_combo.currentText() == "单日期":
            selected_date = self.date_edit.date().toPython()
            return selected_date, selected_date
        else:  # 时间区间
            start_date = self.start_date_edit.date().toPython()
            end_date = self.end_date_edit.date().toPython()
            return start_date, end_date

    def start_search(self):
        """开始筛查"""
        # 检查文件是否已选择 - 至少要有一个文件
        if not self.file1_path and not self.file2_path:
            QMessageBox.warning(self, "警告", "请至少选择一份数据文件！")
            return

        # 正常筛查模式（移除追加模式，因为文件选择已经有自动合并功能）
        self.append_mode = False

        # 添加文件历史
        file1_name = os.path.basename(self.file1_path) if self.file1_path else "无"
        file2_name = os.path.basename(self.file2_path) if self.file2_path else "无"

        # 如果已有数据，则标记为追加数据
        if self.merged_data is not None and not self.merged_data.empty:
            operation = "数据筛查（基于已有数据）"
        else:
            operation = "初始数据导入"

        self.add_file_history(
            file1_name,
            file2_name,
            operation,
        )

        # 禁用控件
        self.search_btn.setEnabled(False)
        self.clear_btn.setEnabled(False)
        self.export_btn.setEnabled(False)

        # 显示进度条
        self.progress_bar.setVisible(True)
        self.progress_label.setVisible(True)
        self.progress_bar.setValue(0)

        # 获取最少人数
        people_text = self.people_combo.currentText()
        min_people_count = int(people_text[0])  # 提取数字部分

        # 获取选择的日期（单日期或日期区间）
        start_date, end_date = self.get_selected_dates()

        # 获取选中的人员类型
        selected_person_types = self.get_selected_person_types()

        # 设置参数并启动处理线程
        # 注意：现在不再需要特别的append_mode，因为文件选择时已经自动合并了数据
        self.processor.set_params(
            self.file1_path,
            self.file2_path,
            start_date,  # 传递开始日期
            end_date,  # 传递结束日期
            self.city_combo.currentText(),
            min_people_count,
            None,  # 不再通过这里传递existing_data，而是通过文件选择的预加载机制
            False,  # 简化：不再需要append_mode
            selected_person_types,  # 传递人员类型筛选参数
        )
        self.processor.start()

    def clear_data(self):
        """清空数据"""
        reply = QMessageBox.question(
            self,
            "确认清空",
            "确定要清空所有数据吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )

        if reply == QMessageBox.Yes:
            # 停止预加载线程（如果正在运行）
            if self.preview_loader.isRunning():
                self.preview_loader.terminate()
                self.preview_loader.wait()

            # 清空所有数据
            self.merged_data = None
            self.result_data = None
            self.original_data_count = 0
            self.append_mode = False

            self.available_person_types = []
            self.update_person_type_ui(preserve_selection=False)

            self.search_input.clear()
            self.search_results = None
            self.search_history = []
            self.search_status_label.setText("请先导入数据以启用搜索功能")
            self.search_scope_indicator.setText("")
            self.search_scope_combo.setCurrentText("全部导入数据")

            # 清空表格
            self.result_table.setRowCount(0)

            # 更新统计信息
            self.stats_label.setText("请先导入数据并开始筛查")

            # 更新数据状态
            self.update_data_status()

            # 清空文件路径显示
            self.file1_edit.setText("未选择文件")
            self.file2_edit.setText("未选择文件")
            self.file1_path = ""
            self.file2_path = ""
            self.has_file1_selected = False
            self.has_file2_selected = False

            # 确保按钮文字正确
            self.search_btn.setText("开始筛查")
            self.search_btn.setStyleSheet(
                """
                QPushButton {
                    background-color: #2196F3;
                    font-size: 16px;
                    padding: 10px;
                }
                QPushButton:hover {
                    background-color: #1976D2;
                }
            """
            )

            # 更新按钮状态
            self.update_button_states()

            QMessageBox.information(self, "提示", "数据已清空")

    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)

    def update_message(self, message):
        """更新进度消息"""
        self.progress_label.setText(message)

    def show_results(self, result_df):
        """显示筛查结果"""
        # 保存原始数据
        self.result_data = result_df.copy()

        # 隐藏进度条
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)

        # 启用控件
        self.search_btn.setEnabled(True)
        self.clear_btn.setEnabled(True)

        # 第四阶段：获取合并后的全量数据，用于人员类型发现
        if hasattr(self.processor, "all_data"):
            self.merged_data = self.processor.all_data
            # 发现人员类型（基于全量数据，而不是筛选后的结果）
            self.discover_person_types(self.merged_data)

            # 更新搜索功能状态 - 现在有了全量数据，可以启用搜索
            self.enable_search_features(True)

        if result_df.empty:
            self.stats_label.setText("未发现符合条件的出行人员")
            self.result_table.setRowCount(0)
            self.update_data_status()
            self.update_button_states()
            return

        # 打印调试信息
        print(f"筛查结果数据形状: {result_df.shape}")
        print(f"筛查结果列名: {result_df.columns.tolist()}")
        print(f"前5行数据:\n{result_df.head()}")

        # 更新统计信息
        total_people = len(result_df)
        # 移除群体统计，因为新逻辑不再分组
        # total_groups = result_df["分组键"].nunique()

        # 获取当前的人数阈值
        min_people = int(self.people_combo.currentText()[0])

        # 根据人数阈值调整描述
        if min_people == 1:
            threshold_desc = "1人及以上"
        else:
            threshold_desc = f"{min_people}人及以上"

        # 统计状态分布
        status_info = ""
        if "状态类型" in result_df.columns:
            status_counts = result_df["状态类型"].value_counts()
            status_parts = []
            for status, count in status_counts.items():
                status_parts.append(f"{status} {count} 人")
            status_info = f"（{', '.join(status_parts)}）"

        # 显示数据量信息
        data_info = ""
        if self.append_mode and self.original_data_count > 0:
            # 获取merged_data（需要从processor获取）
            if hasattr(self.processor, "all_data"):
                self.merged_data = self.processor.all_data
                new_data_count = len(self.merged_data) - self.original_data_count
                data_info = f"\n原有数据: {self.original_data_count} 条，新增数据: {new_data_count} 条，合并后总数据: {len(self.merged_data)} 条"

                # 如果是追加模式，显示数据变化提示 - 使用更友善的弹窗
                append_msg_box = QMessageBox(self)
                append_msg_box.setWindowTitle("✅ 数据追加成功")
                append_msg_box.setIcon(QMessageBox.Information)
                append_msg_box.setText("🎉 数据追加操作已完成！")
                append_msg_box.setInformativeText(
                    f"📊 数据统计信息：\n"
                    f"📁 原有数据：{self.original_data_count} 条\n"
                    f"➕ 新增数据：{new_data_count} 条\n"
                    f"📈 当前总数据：{len(self.merged_data)} 条\n\n"
                    f"✨ 所有数据已成功合并！"
                )
                append_msg_box.setStandardButtons(QMessageBox.Ok)
                append_msg_box.button(QMessageBox.Ok).setText("好的")

                # 设置友善的样式
                append_msg_box.setStyleSheet(
                    """
                    QMessageBox {
                        background-color: #f8fff8;
                        border: 2px solid #4caf50;
                        border-radius: 8px;
                    }
                    QMessageBox QLabel {
                        color: #2e7d32;
                        font-size: 14px;
                        padding: 10px;
                    }
                    QMessageBox QPushButton {
                        background-color: #4caf50;
                        color: white;
                        border: none;
                        padding: 8px 20px;
                        border-radius: 4px;
                        font-weight: bold;
                        min-width: 80px;
                    }
                    QMessageBox QPushButton:hover {
                        background-color: #45a049;
                    }
                """
                )

                append_msg_box.exec_()
        else:
            # 首次筛查，保存merged_data
            if hasattr(self.processor, "all_data"):
                self.merged_data = self.processor.all_data

        # 第五阶段：添加详情查看提示
        detail_tip = "\n💡 提示：点击证件号可查看该人员的详细出行记录"

        self.stats_label.setText(
            f"共发现 {total_people} 人符合筛选条件（{threshold_desc}阈值）{status_info}{data_info}{detail_tip}"
        )

        # 设置表格
        columns = [
            "姓名",
            "证件号",
            "航班车次",
            "发站",
            "到站",
            "出发日期",
            "出发时间",
            "人员类型",  # 第四阶段：添加人员类型列
            "变更操作",
            "状态类型",
            "数据源",
        ]

        # 过滤出实际存在的列
        available_columns = [col for col in columns if col in result_df.columns]

        self.result_table.setRowCount(len(result_df))
        self.result_table.setColumnCount(len(available_columns))
        self.result_table.setHorizontalHeaderLabels(available_columns)

        # 启用表格排序功能
        self.result_table.setSortingEnabled(False)  # 先禁用，填充数据后再启用

        # 重置索引以确保正确的行号
        result_df = result_df.reset_index(drop=True)

        # 填充数据
        for row_idx in range(len(result_df)):
            for col_idx, col in enumerate(available_columns):
                value = result_df.iloc[row_idx][col]
                if pd.isna(value):
                    value = ""
                elif col == "出发日期":
                    value = str(value)[:10]
                else:
                    value = str(value)

                item = QTableWidgetItem(value)

                # 为排序设置正确的数据类型
                if col == "出发日期":
                    # 为日期列设置正确的排序数据
                    try:
                        date_value = pd.to_datetime(result_df.iloc[row_idx][col])
                        # 使用时间戳作为排序依据
                        item.setData(Qt.UserRole, date_value.timestamp())
                    except:
                        item.setData(Qt.UserRole, 0)  # 无效日期设为最小值
                elif col == "姓名":
                    # 为姓名设置排序数据（使用原始值）
                    item.setData(Qt.UserRole, str(value))

                # 根据状态类型设置文字颜色
                if col == "状态类型":
                    if value == "已确认":
                        item.setForeground(QColor("#2E7D32"))  # 深绿色
                    elif value == "待确认":
                        item.setForeground(QColor("#F57C00"))  # 橙色
                    else:
                        item.setForeground(QColor("#757575"))  # 灰色

                # 第五阶段：为证件号列设置特殊样式，表示可点击
                elif col == "证件号":
                    item.setForeground(QColor("#1976D2"))  # 蓝色表示可点击
                    item.setBackground(QColor("#E3F2FD"))  # 浅蓝色背景
                    # 设置工具提示
                    item.setToolTip("点击查看该人员的详细出行记录")
                    # 设置字体为下划线，类似链接样式
                    font = item.font()
                    font.setUnderline(True)
                    item.setFont(font)

                # 使用行索引设置背景色，让不同行有不同颜色
                if col != "证件号":  # 证件号列已经设置了特殊背景色
                    colors = ["#FFE6E6", "#E6F3FF", "#E6FFE6", "#FFFFE6", "#F3E6FF"]
                    color = colors[row_idx % len(colors)]

                    # 如果是已确认状态，使用稍深的背景色
                    if result_df.iloc[row_idx].get("状态类型") == "已确认":
                        base_color = QColor(color)
                        base_color = base_color.darker(110)
                        item.setBackground(base_color)
                    else:
                        item.setBackground(QColor(color))

                self.result_table.setItem(row_idx, col_idx, item)

        # 数据填充完成后启用排序功能
        self.result_table.setSortingEnabled(True)

        # 设置表头样式，提示可排序
        header = self.result_table.horizontalHeader()
        header.setStyleSheet(
            """
            QHeaderView::section {
                background-color: #f0f0f0;
                border: 1px solid #cccccc;
                padding: 8px;
                font-weight: bold;
                font-size: 13px;
            }
            QHeaderView::section:hover {
                background-color: #e0e0e0;
                cursor: pointer;
            }
        """
        )

        # 添加排序提示到统计标签
        sort_tip = "\n🔄 提示：点击列标题可按该列排序"
        current_text = self.stats_label.text()
        if "🔄 提示：" not in current_text:
            self.stats_label.setText(current_text + sort_tip)

        # 调整列宽
        self.result_table.resizeColumnsToContents()

        # 启用导出按钮
        self.export_btn.setEnabled(True)

        # 更新数据状态显示
        self.update_data_status()

        # 更新按钮状态
        self.update_button_states()

        # 第三阶段：有数据后启用搜索功能
        self.enable_search_features(True)

        # 如果现在有筛查结果，默认选择搜索范围为全部导入数据
        if self.result_data is not None and not self.result_data.empty:
            self.search_scope_combo.setCurrentText("全部导入数据")

    def show_error(self, error_msg):
        """显示错误信息"""
        # 隐藏进度条
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)

        # 启用控件
        self.search_btn.setEnabled(True)
        if self.merged_data is not None:
            self.clear_btn.setEnabled(True)

        # 更新按钮状态
        self.update_button_states()

        # 显示错误信息
        QMessageBox.critical(self, "错误", error_msg)

    def export_results(self):
        """导出结果"""
        if self.result_data is None or self.result_data.empty:
            return

        # 选择保存位置
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存筛查结果",
            f"群体出行筛查结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel文件 (*.xlsx)",
        )

        if file_path:
            try:
                # 创建Excel写入器
                with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                    # 写入详细数据
                    self.result_data.to_excel(
                        writer, sheet_name="筛查结果", index=False
                    )

                    # 创建汇总表
                    summary = (
                        self.result_data.groupby(["到站"])
                        .agg(
                            {
                                "姓名": "count",
                                "航班车次": lambda x: ", ".join(x.unique()),
                                "出发日期": "first",
                                "状态类型": lambda x: (
                                    ", ".join(x.value_counts().index.tolist())
                                    if "状态类型" in self.result_data.columns
                                    else ""
                                ),
                            }
                        )
                        .rename(columns={"姓名": "人数"})
                    )

                    summary.to_excel(writer, sheet_name="群体汇总")

                    # 创建状态统计表
                    if "状态类型" in self.result_data.columns:
                        # 按状态类型统计
                        status_summary = self.result_data["状态类型"].value_counts()
                        status_df = pd.DataFrame(
                            {
                                "状态类型": status_summary.index,
                                "数量": status_summary.values,
                                "占比": (
                                    status_summary.values / len(self.result_data) * 100
                                ).round(2),
                            }
                        )
                        status_df.to_excel(writer, sheet_name="状态统计", index=False)

                        # 变更操作统计
                        if "变更操作" in self.result_data.columns:
                            operation_stats = self.result_data[
                                "变更操作"
                            ].value_counts()
                            operation_df = pd.DataFrame(
                                {
                                    "变更操作": operation_stats.index,
                                    "数量": operation_stats.values,
                                    "占比": (
                                        operation_stats.values
                                        / len(self.result_data)
                                        * 100
                                    ).round(2),
                                }
                            )
                            operation_df.to_excel(
                                writer, sheet_name="变更操作统计", index=False
                            )

                QMessageBox.information(self, "成功", f"结果已导出到：\n{file_path}")

            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")

    # 第三阶段：搜索功能相关方法
    def on_search_text_changed(self, text):
        """搜索文本变化时的处理（防抖搜索）"""
        if text.strip():
            self.search_timer.start(300)  # 300ms延迟
        else:
            self.clear_search()

    def perform_search(self):
        """执行搜索操作"""
        search_text = self.search_input.text().strip()
        if not search_text:
            return

        self.search_data()

    def search_data(self):
        """搜索数据"""
        search_text = self.search_input.text().strip()
        search_type = self.search_type_combo.currentText()

        if not search_text:
            QMessageBox.warning(self, "提示", "请输入搜索内容！")
            return

        # 根据搜索范围选择数据源
        search_scope = self.search_scope_combo.currentText()
        search_source = None

        if search_scope == "全部导入数据":
            # 优先搜索全部导入数据
            if self.merged_data is not None and not self.merged_data.empty:
                search_source = self.merged_data
                search_scope_desc = "全部导入数据"
            else:
                QMessageBox.warning(self, "提示", "请先导入数据！")
                return
        else:  # 筛查结果
            # 搜索筛查结果
            if self.result_data is not None and not self.result_data.empty:
                search_source = self.result_data
                search_scope_desc = "筛查结果"
            else:
                QMessageBox.warning(self, "提示", "请先进行筛查！")
                return

        try:
            # 执行搜索
            search_results = self.execute_search(
                search_source, search_text, search_type
            )

            if search_results.empty:
                self.search_status_label.setText(
                    f"在{search_scope_desc}中未找到匹配 '{search_text}' 的记录"
                )
                self.search_scope_indicator.setText(
                    f"当前搜索范围：{search_scope_desc}"
                )
                QMessageBox.information(
                    self, "搜索结果", f"在{search_scope_desc}中未找到匹配的记录"
                )
                return

            # 显示搜索结果
            self.show_search_results(
                search_results, search_text, search_type, search_scope_desc
            )

            # 添加到搜索历史
            self.add_search_history(search_text, search_type)

            # 更新状态标签
            self.search_status_label.setText(
                f"在{search_scope_desc}中找到 {len(search_results)} 条匹配记录"
            )
            self.search_scope_indicator.setText(f"当前搜索范围：{search_scope_desc}")

        except Exception as e:
            QMessageBox.critical(self, "搜索错误", f"搜索过程中发生错误：{str(e)}")

    def execute_search(self, data_df, search_text, search_type):
        """执行搜索逻辑"""
        search_text_lower = search_text.lower()

        if search_type == "姓名":
            mask = (
                data_df["姓名"]
                .astype(str)
                .str.lower()
                .str.contains(search_text_lower, case=False, na=False)
            )
        elif search_type == "证件号":
            mask = (
                data_df["证件号"]
                .astype(str)
                .str.contains(search_text, case=False, na=False)
            )
        elif search_type == "航班车次":
            mask = (
                data_df["航班车次"]
                .astype(str)
                .str.contains(search_text, case=False, na=False)
            )
        elif search_type == "全字段":
            # 在所有主要字段中搜索
            main_fields = ["姓名", "证件号", "航班车次", "发站", "到站"]
            available_fields = [
                field for field in main_fields if field in data_df.columns
            ]

            mask = pd.Series([False] * len(data_df), index=data_df.index)
            for field in available_fields:
                field_mask = (
                    data_df[field]
                    .astype(str)
                    .str.lower()
                    .str.contains(search_text_lower, case=False, na=False)
                )
                mask = mask | field_mask

        return data_df[mask].reset_index(drop=True)

    def show_search_results(
        self, search_results, search_text, search_type, search_scope
    ):
        """显示搜索结果"""
        # 保存搜索结果
        self.search_results = search_results

        # 更新表格显示
        columns = [
            "姓名",
            "证件号",
            "航班车次",
            "发站",
            "到站",
            "出发日期",
            "出发时间",
            "变更操作",
            "状态类型",
            "数据源",
        ]

        # 过滤出实际存在的列
        available_columns = [col for col in columns if col in search_results.columns]

        self.result_table.setRowCount(len(search_results))
        self.result_table.setColumnCount(len(available_columns))
        self.result_table.setHorizontalHeaderLabels(available_columns)

        # 启用表格排序功能
        self.result_table.setSortingEnabled(False)  # 先禁用，填充数据后再启用

        # 填充搜索结果数据（带高亮）
        for row_idx in range(len(search_results)):
            for col_idx, col in enumerate(available_columns):
                value = search_results.iloc[row_idx][col]
                if pd.isna(value):
                    value = ""
                elif col == "出发日期":
                    value = str(value)[:10]
                else:
                    value = str(value)

                item = QTableWidgetItem(value)

                # 为排序设置正确的数据类型
                if col == "出发日期":
                    try:
                        date_value = pd.to_datetime(search_results.iloc[row_idx][col])
                        item.setData(Qt.UserRole, date_value.timestamp())
                    except:
                        item.setData(Qt.UserRole, 0)
                elif col == "姓名":
                    item.setData(Qt.UserRole, str(value))

                # 高亮显示匹配的内容
                if self.should_highlight_cell(value, search_text, search_type, col):
                    # 设置高亮背景色
                    item.setBackground(QBrush(QColor("#FFEB3B")))  # 黄色高亮
                    item.setForeground(QBrush(QColor("#E65100")))  # 深橙色文字
                else:
                    # 设置普通背景色
                    colors = ["#E8F5E9", "#E3F2FD", "#FFF3E0", "#F3E5F5", "#E0F2F1"]
                    color = colors[row_idx % len(colors)]
                    item.setBackground(QBrush(QColor(color)))

                # 证件号列特殊样式（搜索结果中也可点击查看详情）
                if col == "证件号":
                    item.setForeground(QBrush(QColor("#1976D2")))
                    item.setToolTip("点击查看该人员的详细出行记录")
                    font = item.font()
                    font.setUnderline(True)
                    item.setFont(font)

                self.result_table.setItem(row_idx, col_idx, item)

        # 数据填充完成后启用排序功能
        self.result_table.setSortingEnabled(True)

        # 设置表头样式
        header = self.result_table.horizontalHeader()
        header.setStyleSheet(
            """
            QHeaderView::section {
                background-color: #f0f0f0;
                border: 1px solid #cccccc;
                padding: 8px;
                font-weight: bold;
                font-size: 13px;
            }
            QHeaderView::section:hover {
                background-color: #e0e0e0;
                cursor: pointer;
            }
        """
        )

        # 调整列宽
        self.result_table.resizeColumnsToContents()

        # 更新统计信息
        search_tip = f"搜索结果：在{search_scope}中找到 {len(search_results)} 条匹配 '{search_text}' 的记录"
        sort_tip = "\n🔄 提示：点击列标题可按该列排序"
        self.stats_label.setText(search_tip + sort_tip)

    def should_highlight_cell(self, cell_value, search_text, search_type, column_name):
        """判断是否应该高亮该单元格"""
        if not cell_value or not search_text:
            return False

        cell_value_lower = str(cell_value).lower()
        search_text_lower = search_text.lower()

        if search_type == "全字段":
            return search_text_lower in cell_value_lower
        elif search_type == "姓名" and column_name == "姓名":
            return search_text_lower in cell_value_lower
        elif search_type == "证件号" and column_name == "证件号":
            return search_text in str(cell_value)
        elif search_type == "航班车次" and column_name == "航班车次":
            return search_text in str(cell_value)

        return False

    def clear_search(self):
        """清除搜索"""
        self.search_input.clear()
        self.search_results = None

        # 恢复显示原始数据
        if self.result_data is not None and not self.result_data.empty:
            self.show_results(self.result_data)

        self.search_status_label.setText("搜索已清除")
        self.search_scope_indicator.setText("")  # 清除搜索范围指示

    def add_search_history(self, search_text, search_type):
        """添加搜索历史"""
        history_item = {
            "text": search_text,
            "type": search_type,
            "time": datetime.now().strftime("%H:%M:%S"),
        }

        # 避免重复添加相同的搜索
        if history_item not in self.search_history:
            self.search_history.insert(0, history_item)  # 添加到开头

            # 最多保留10条历史记录
            if len(self.search_history) > 10:
                self.search_history = self.search_history[:10]

    def show_search_history(self):
        """显示搜索历史"""
        if not self.search_history:
            QMessageBox.information(self, "搜索历史", "暂无搜索历史记录")
            return

        # 创建历史记录对话框
        from PySide2.QtWidgets import QDialog, QListWidget, QDialogButtonBox

        dialog = QDialog(self)
        dialog.setWindowTitle("搜索历史")
        dialog.setModal(True)
        dialog.resize(400, 300)

        layout = QVBoxLayout(dialog)

        # 历史记录列表
        history_list = QListWidget()
        for item in self.search_history:
            history_text = f"[{item['time']}] {item['type']}: {item['text']}"
            history_list.addItem(history_text)

        layout.addWidget(QLabel("点击选择历史搜索："))
        layout.addWidget(history_list)

        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        layout.addWidget(button_box)

        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)

        # 双击或确定时执行搜索
        def on_history_selected():
            current_row = history_list.currentRow()
            if current_row >= 0:
                selected_item = self.search_history[current_row]
                self.search_input.setText(selected_item["text"])
                self.search_type_combo.setCurrentText(selected_item["type"])
                dialog.accept()
                self.search_data()

        history_list.itemDoubleClicked.connect(on_history_selected)

        if dialog.exec_() == QDialog.Accepted:
            on_history_selected()

    def enable_search_features(self, enable=True):
        """启用或禁用搜索功能"""
        self.search_input.setEnabled(enable)
        self.search_type_combo.setEnabled(enable)
        self.search_data_btn.setEnabled(enable)
        self.clear_search_btn.setEnabled(enable)
        self.search_history_btn.setEnabled(enable)

        # 根据数据状态设置搜索范围下拉框
        if enable:
            has_merged_data = (
                self.merged_data is not None and not self.merged_data.empty
            )
            has_result_data = (
                self.result_data is not None and not self.result_data.empty
            )

            if has_merged_data and has_result_data:
                # 两种数据都有，启用搜索范围选择
                self.search_scope_combo.setEnabled(True)
                self.search_status_label.setText("搜索功能已启用 - 可选择搜索范围")
            elif has_merged_data:
                # 只有导入数据，固定为全部数据搜索
                self.search_scope_combo.setEnabled(False)
                self.search_scope_combo.setCurrentText("全部导入数据")
                self.search_status_label.setText("搜索功能已启用 - 搜索全部导入数据")
            elif has_result_data:
                # 只有筛查结果（理论上不应该出现这种情况）
                self.search_scope_combo.setEnabled(False)
                self.search_scope_combo.setCurrentText("筛查结果")
                self.search_status_label.setText("搜索功能已启用 - 搜索筛查结果")
        else:
            self.search_scope_combo.setEnabled(False)
            self.search_status_label.setText("请先导入数据以启用搜索功能")

    def update_button_states(self):
        """根据当前状态更新按钮可用性"""
        # 至少需要一个文件就可以筛查
        can_search = self.has_file1_selected or self.has_file2_selected
        self.search_btn.setEnabled(can_search)

        # 只有有数据后才能清空
        has_data = self.merged_data is not None
        self.clear_btn.setEnabled(has_data and not self.processor.isRunning())

        # 搜索功能状态管理 - 只要有导入的数据就启用搜索功能
        has_searchable_data = (
            self.merged_data is not None and not self.merged_data.empty
        )
        self.enable_search_features(has_searchable_data)

    def update_data_status(self):
        """更新数据状态显示"""
        if self.merged_data is None:
            self.data_status_label.setText("当前无数据")
            self.data_status_label.setStyleSheet(
                """
                QLabel {
                    background-color: #f0f0f0;
                    border: 1px solid #cccccc;
                    border-radius: 5px;
                    padding: 15px;
                    font-size: 20px;
                    color: #666666;
                }
            """
            )
        else:
            total_records = len(self.merged_data)
            status_text = f"当前数据总量：{total_records} 条记录"

            # 添加状态分布信息
            if "状态类型" in self.merged_data.columns:
                status_counts = self.merged_data["状态类型"].value_counts()
                status_parts = []
                for status, count in status_counts.items():
                    status_parts.append(f"{status} {count} 条")
                status_text += f"\n状态分布：{', '.join(status_parts)}"

            self.data_status_label.setText(status_text)
            self.data_status_label.setStyleSheet(
                """
                QLabel {
                    background-color: #e8f5e9;
                    border: 1px solid #4caf50;
                    border-radius: 5px;
                    padding: 15px;
                    font-size: 20px;
                    color: #2e7d32;
                }
            """
            )

    def add_file_history(
        self, file1_name, file2_name, operation="导入", record_count=0
    ):
        """添加文件导入历史记录"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        history_text = f"[{timestamp}] {operation}：\n"
        history_text += f"  票务数据：{file1_name}\n"
        history_text += f"  航班数据：{file2_name}\n"
        if record_count > 0:
            history_text += f"  处理记录数：{record_count} 条\n"
        history_text += "-" * 50 + "\n"

        current_text = self.file_history_text.toPlainText()
        self.file_history_text.setPlainText(history_text + current_text)

    def on_table_cell_clicked(self, row, column):
        """表格单元格点击事件"""
        if self.result_table.rowCount() == 0 or self.result_table.columnCount() == 0:
            return

        # 获取列名
        header_item = self.result_table.horizontalHeaderItem(column)
        if header_item is None:
            return

        column_name = header_item.text()

        # 只在点击证件号列时触发详情查看
        if column_name == "证件号":
            # 获取点击的证件号
            item = self.result_table.item(row, column)
            if item is None:
                return

            person_id = item.text().strip()
            if not person_id:
                QMessageBox.warning(self, "提示", "证件号信息为空")
                return

            # 显示人员详情
            self.show_person_detail(person_id)

    def show_person_detail(self, person_id):
        """显示人员详情对话框"""
        # 使用合并后的全量数据进行查询，确保能看到所有相关记录
        data_source = None

        if self.merged_data is not None and not self.merged_data.empty:
            data_source = self.merged_data
        elif self.result_data is not None and not self.result_data.empty:
            data_source = self.result_data
        else:
            QMessageBox.warning(self, "警告", "没有可用的数据进行详情查询")
            return

        try:
            # 创建并显示详情对话框
            detail_dialog = PersonDetailDialog(person_id, data_source, self)
            detail_dialog.exec_()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"显示人员详情时发生错误：{str(e)}")
            print(f"详情对话框错误：{e}")  # 调试信息

    def on_search_scope_changed(self, text):
        """搜索范围变化时的处理"""
        # 更新搜索范围指示器
        if text:
            self.search_scope_indicator.setText(f"当前搜索范围：{text}")

        # 如果有搜索内容，可以自动重新搜索
        if self.search_input.text().strip():
            # 使用防抖定时器，避免频繁搜索
            self.search_timer.start(300)

    def trigger_data_preview(self):
        """触发数据预加载"""
        # 检查是否有文件被选择
        if not self.file1_path and not self.file2_path:
            return

        # 如果预加载器正在运行，先停止
        if self.preview_loader.isRunning():
            self.preview_loader.terminate()
            self.preview_loader.wait()

        # 显示进度提示
        self.progress_bar.setVisible(True)
        self.progress_label.setVisible(True)
        self.progress_bar.setValue(0)
        self.progress_label.setText("正在预加载数据，请稍候...")

        # 临时禁用按钮
        self.search_btn.setEnabled(False)
        self.clear_btn.setEnabled(False)

        # 设置参数并启动预加载
        self.preview_loader.set_params(
            self.file1_path,
            self.file2_path,
            self.merged_data,  # 传入已有数据（如果有的话）
        )
        self.preview_loader.start()

    def on_data_preview_loaded(self, preview_data):
        """数据预加载完成的处理"""
        # 隐藏进度条
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)

        # 保存预加载的数据
        self.merged_data = preview_data

        # 发现人员类型
        if self.merged_data is not None and not self.merged_data.empty:
            self.discover_person_types(self.merged_data)

            # 启用搜索功能
            self.enable_search_features(True)

            # 更新数据状态显示
            self.update_data_status()

            # 显示成功消息 - 使用自定义弹窗，更友善温和
            total_records = len(self.merged_data)

            # 创建自定义消息框
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("✅ 数据加载成功")
            msg_box.setIcon(QMessageBox.Information)  # 使用信息图标
            msg_box.setText("🎉 数据预加载已完成！")
            msg_box.setInformativeText(
                f"✨ 成功加载了 {total_records} 条记录\n\n"
                f"📋 您现在可以进行以下操作：\n"
                f"🔍 使用搜索功能查找特定人员\n"
                f"🏷️ 选择人员类型进行筛选\n"
                f"⚙️ 设置筛查条件并开始筛查"
            )
            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.button(QMessageBox.Ok).setText("好的")

            # 设置更友善的样式
            msg_box.setStyleSheet(
                """
                QMessageBox {
                    background-color: #f8fff8;
                    border: 2px solid #4caf50;
                    border-radius: 8px;
                }
                QMessageBox QLabel {
                    color: #2e7d32;
                    font-size: 14px;
                    padding: 10px;
                }
                QMessageBox QPushButton {
                    background-color: #4caf50;
                    color: white;
                    border: none;
                    padding: 8px 20px;
                    border-radius: 4px;
                    font-weight: bold;
                    min-width: 80px;
                }
                QMessageBox QPushButton:hover {
                    background-color: #45a049;
                }
            """
            )

            msg_box.exec_()

            # 更新文件历史
            file1_name = os.path.basename(self.file1_path) if self.file1_path else "无"
            file2_name = os.path.basename(self.file2_path) if self.file2_path else "无"
            self.add_file_history(file1_name, file2_name, "预加载数据", total_records)

        # 恢复按钮状态
        self.update_button_states()

    def show_preview_error(self, error_msg):
        """显示预加载错误信息"""
        # 隐藏进度条
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)

        # 恢复按钮状态
        self.update_button_states()

        # 显示错误消息
        QMessageBox.warning(
            self,
            "数据预加载失败",
            f"数据预加载过程中发生错误：\n{error_msg}\n\n"
            f"您仍然可以使用筛查功能，但搜索功能需要数据预加载成功后才能使用。",
        )


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    window = GroupTravelChecker()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
