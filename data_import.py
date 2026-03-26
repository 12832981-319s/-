from pathlib import Path

import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QLabel, QPushButton, QTextEdit, QVBoxLayout, QWidget

from app_paths import sample_data_dir
from data_cleaning import clean_value, finalize_dataframe, normalize_text


class DataImportWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.dataframes = {}
        self.dataset_meta = {}
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.title_label = QLabel("数据导入与清洗")
        self.title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #2c3e50;")
        layout.addWidget(self.title_label)

        self.btn_import = QPushButton("导入 Excel 文件")
        self.btn_import.clicked.connect(self.import_files)
        self.btn_import.setStyleSheet(
            "padding: 10px; font-size: 14px; background-color: #3498db; color: white; border-radius: 5px;"
        )
        layout.addWidget(self.btn_import)

        self.btn_load_samples = QPushButton("加载 data 目录示例数据")
        self.btn_load_samples.clicked.connect(self.load_sample_files)
        self.btn_load_samples.setStyleSheet(
            "padding: 10px; font-size: 14px; background-color: #16a085; color: white; border-radius: 5px;"
        )
        layout.addWidget(self.btn_load_samples)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setPlaceholderText("导入后会在这里显示清洗结果和数据摘要。")
        layout.addWidget(self.log_text)

        self.setLayout(layout)

    def import_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择文件", "", "Excel Files (*.xls *.xlsx)")
        if files:
            self.load_files(files)

    def load_sample_files(self):
        sample_dir = sample_data_dir()
        files = sorted(str(path) for path in sample_dir.glob("*.xls*"))
        if not files:
            self.log_text.setText("未找到示例数据目录。")
            return
        self.load_files(files)

    def load_files(self, files):
        self.dataframes = {}
        self.dataset_meta = {}
        self.log_text.clear()
        self.log_text.append(f"开始处理 {len(files)} 个文件。\n")

        for file_path in files:
            filename = Path(file_path).name
            try:
                dataframe = self.clean_energy_data(file_path, filename)
                self.dataframes[filename] = dataframe
                self.dataset_meta[filename] = {"rows": len(dataframe), "columns": list(dataframe.columns)}
                self.log_text.append(self.build_summary(filename, dataframe))
            except Exception as exc:
                self.log_text.append(f"[失败] {filename}: {exc}\n")

        if self.dataframes:
            self.log_text.append("全部文件处理完成，可切换到其他标签页继续分析。")

    def build_summary(self, filename, dataframe: pd.DataFrame) -> str:
        preview_cols = "、".join(dataframe.columns[:4])
        return (
            f"[成功] {filename}\n"
            f"行数: {len(dataframe)}，列数: {len(dataframe.columns)}\n"
            f"字段: {preview_cols}\n"
        )

    def clean_energy_data(self, file_path, filename):
        if "综合能源平衡表" in filename or "09-03" in filename or "9-3" in filename:
            return self._load_balance_sheet(file_path)
        if "分行业能源消费量" in filename or "09-07" in filename or "9-7" in filename:
            return self._load_sector_consumption(file_path)
        if "能源消费弹性系数" in filename or "09-09" in filename or "9-9" in filename:
            return self._load_elasticity(file_path)
        if "能源消费总量及构成" in filename or "09-02" in filename or "9-2" in filename:
            return self._load_total_consumption(file_path)
        if "重点耗能企业单位产品能源消费" in filename or "9-12" in filename:
            return self._load_key_enterprises(file_path)
        return self._load_generic(file_path)

    def _load_balance_sheet(self, file_path):
        raw = pd.read_excel(file_path, header=None)
        df = raw.iloc[6:, :7].copy()
        df.columns = ["项目", "英文项目", "2020", "2021", "2022", "2023", "2024"]
        df = finalize_dataframe(df, text_columns=2)
        df = df[~df["项目"].str.contains("注|资料来源", na=False)]
        return df

    def _load_sector_consumption(self, file_path):
        raw = pd.read_excel(file_path, header=None)
        df = raw.iloc[8:, :13].copy()
        df.columns = [
            "行业",
            "英文行业",
            "能源消费总量",
            "煤炭",
            "焦炭",
            "原油",
            "汽油",
            "煤油",
            "柴油",
            "燃料油",
            "液化石油气",
            "天然气",
            "电力",
        ]
        return finalize_dataframe(df, text_columns=2)

    def _load_elasticity(self, file_path):
        raw = pd.read_excel(file_path, header=None)
        df = raw.iloc[10:, :5].copy()
        df.columns = ["年份", "能源消费增长率", "电力消费增长率", "能源消费弹性系数", "电力消费弹性系数"]
        return finalize_dataframe(df, text_columns=1)

    def _load_total_consumption(self, file_path):
        raw = pd.read_excel(file_path, header=None)
        df = raw.iloc[9:, :6].copy()
        df.columns = ["年份", "能源消费总量", "煤炭占比", "石油占比", "天然气占比", "一次电力及其他能源占比"]
        return finalize_dataframe(df, text_columns=1)

    def _load_key_enterprises(self, file_path):
        raw = pd.read_excel(file_path, header=None)
        df = raw.iloc[5:, :5].copy()
        df.columns = ["项目", "英文项目", "2020", "2023", "2024"]
        return finalize_dataframe(df, text_columns=2)

    def _load_generic(self, file_path):
        raw = pd.read_excel(file_path, header=None)
        non_empty_rows = [idx for idx in range(len(raw)) if raw.iloc[idx].notna().sum() >= 2]
        header_row = non_empty_rows[0]
        data_start = next((idx for idx in non_empty_rows if idx > header_row), header_row + 1)
        columns = [normalize_text(cell) or f"列{index + 1}" for index, cell in enumerate(raw.iloc[header_row])]
        df = raw.iloc[data_start:].copy()
        df.columns = columns[: len(df.columns)]

        for index, column in enumerate(df.columns):
            if index == 0:
                df[column] = df[column].map(normalize_text)
            else:
                df[column] = df[column].map(clean_value)

        df = df.dropna(how="all")
        df = df[df[df.columns[0]] != ""]
        return df.reset_index(drop=True)

    def get_data(self):
        return self.dataframes
