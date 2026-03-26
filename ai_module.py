import json
import os
from pathlib import Path

from openai import OpenAI
from PyQt5.QtCore import QObject, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QComboBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app_paths import config_file_path


BASE_URLS = {
    "中国大陆（北京）": "https://dashscope.aliyuncs.com/compatible-mode/v1",
    "新加坡": "https://dashscope-intl.aliyuncs.com/compatible-mode/v1",
    "美国（弗吉尼亚）": "https://dashscope-us.aliyuncs.com/compatible-mode/v1",
    "自定义": "",
}


class LLMWorker(QObject):
    finished = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, api_key, base_url, model, messages):
        super().__init__()
        self.api_key = api_key
        self.base_url = base_url
        self.model = model
        self.messages = messages

    def run(self):
        try:
            client = OpenAI(api_key=self.api_key, base_url=self.base_url)
            response = client.chat.completions.create(
                model=self.model,
                messages=self.messages,
            )
            content = response.choices[0].message.content or ""
            self.finished.emit(content.strip())
        except Exception as exc:
            self.failed.emit(str(exc))


class AIAnalysisWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.dataframes = {}
        self.thread = None
        self.worker = None
        self.config_path = config_file_path("llm_config.json")
        self.init_ui()
        self.load_config()

    def init_ui(self):
        layout = QVBoxLayout()

        title = QLabel("AI 决策建议")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #8e44ad;")
        layout.addWidget(title)

        desc = QLabel("支持调用通义千问生成节能降碳分析报告，并可保存本地配置。")
        layout.addWidget(desc)

        env_hint = QLabel("建议将 API Key 放到环境变量 DASHSCOPE_API_KEY，界面输入只作为当前会话覆盖值。")
        env_hint.setStyleSheet("color: #666666;")
        layout.addWidget(env_hint)

        form_layout = QFormLayout()

        self.api_input = QLineEdit()
        self.api_input.setPlaceholderText("请输入百炼 API Key")
        self.api_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("API Key", self.api_input)

        self.model_input = QLineEdit()
        self.model_input.setPlaceholderText("如 qwen-plus")
        self.model_input.setText("qwen-plus")
        form_layout.addRow("模型", self.model_input)

        self.region_combo = QComboBox()
        self.region_combo.addItems(list(BASE_URLS.keys()))
        self.region_combo.currentTextChanged.connect(self.on_region_changed)
        form_layout.addRow("地域", self.region_combo)

        self.base_url_input = QLineEdit()
        self.base_url_input.setPlaceholderText("自动填充，或手动输入兼容接口地址")
        form_layout.addRow("Base URL", self.base_url_input)

        layout.addLayout(form_layout)

        button_row = QHBoxLayout()

        self.btn_save = QPushButton("保存配置")
        self.btn_save.clicked.connect(self.save_config)
        button_row.addWidget(self.btn_save)

        self.btn_test = QPushButton("测试连接")
        self.btn_test.clicked.connect(self.test_connection)
        button_row.addWidget(self.btn_test)

        self.btn_generate = QPushButton("生成 AI 报告")
        self.btn_generate.clicked.connect(self.generate_ai_report)
        button_row.addWidget(self.btn_generate)

        layout.addLayout(button_row)

        self.ai_output = QTextEdit()
        self.ai_output.setReadOnly(True)
        self.ai_output.setPlaceholderText("导入数据后可测试连接或生成报告。")
        self.ai_output.setStyleSheet("background-color: #f6f8fb;")
        layout.addWidget(self.ai_output)

        self.setLayout(layout)
        self.on_region_changed(self.region_combo.currentText())

    def on_region_changed(self, region_name):
        default_url = BASE_URLS.get(region_name, "")
        self.base_url_input.setReadOnly(False)
        if region_name != "自定义":
            self.base_url_input.setText(default_url)

    def set_data(self, dataframes):
        self.dataframes = dataframes or {}

    def load_config(self):
        env_api_key = os.getenv("DASHSCOPE_API_KEY", "").strip()
        if env_api_key:
            self.api_input.setText(env_api_key)

        if not self.config_path.exists():
            return

        try:
            data = json.loads(self.config_path.read_text(encoding="utf-8"))
        except Exception:
            return

        self.model_input.setText(data.get("model", "qwen-plus"))
        region = data.get("region", "中国大陆（北京）")
        if region in BASE_URLS:
            self.region_combo.setCurrentText(region)
        self.base_url_input.setText(data.get("base_url", BASE_URLS.get(region, "")))

    def save_config(self):
        config = {
            "model": self.model_input.text().strip() or "qwen-plus",
            "region": self.region_combo.currentText(),
            "base_url": self.base_url_input.text().strip(),
        }
        self.config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
        self.ai_output.setPlainText(
            f"配置已保存到: {self.config_path}\n"
            "出于安全考虑，API Key 不会写入配置文件。"
        )

    def test_connection(self):
        if not self.validate_inputs(require_data=False):
            return

        messages = [
            {"role": "system", "content": "你是一个简洁可靠的助手。请用简洁中文回复。"},
            {"role": "user", "content": "请回复“通义千问连接测试成功”。"},
        ]
        self.run_llm_task(messages, "正在测试连接...")

    def generate_ai_report(self):
        if not self.validate_inputs(require_data=True):
            return

        prompt = self.build_analysis_prompt()
        messages = [
            {
                "role": "system",
                "content": (
                    "你是资深能源分析顾问和专业汇报写作者。"
                    "请基于给定数据摘要输出一份适合管理层阅读的中文 HTML 报告片段。"
                    "不要输出 markdown 代码块，不要输出 html、body 标签。"
                    "请使用 h1、h2、p、ul、li、table、tr、td、th、strong 等基础标签。"
                    "必须包含以下部分：现状概览、关键问题、行业机会、分阶段行动建议、结语。"
                    "整体风格要求简洁、专业、层次清晰，重点数字使用 strong 标签。"
                    "所有判断必须基于给定数据，不要虚构不存在的指标。"
                ),
            },
            {"role": "user", "content": prompt},
        ]
        self.run_llm_task(messages, "正在调用通义千问生成报告...")

    def validate_inputs(self, require_data):
        api_key = self.api_input.text().strip()
        model = self.model_input.text().strip()
        base_url = self.base_url_input.text().strip()

        if not api_key:
            QMessageBox.warning(self, "缺少配置", "请先输入 API Key。")
            return False
        if not model:
            QMessageBox.warning(self, "缺少配置", "请先输入模型名称。")
            return False
        if not base_url:
            QMessageBox.warning(self, "缺少配置", "请先输入 Base URL。")
            return False
        if require_data and not self.dataframes:
            QMessageBox.warning(self, "缺少数据", "请先导入数据。")
            return False
        return True

    def build_analysis_prompt(self):
        summary_blocks = []
        for name, dataframe in self.dataframes.items():
            summary_blocks.append(self.summarize_dataframe(name, dataframe))

        return (
            "请根据以下能源数据摘要，生成一份管理层简明版节能降碳分析报告。\n\n"
            "展示要求：\n"
            "1. 页面顶端需要主标题和副标题。\n"
            "2. 每个一级部分单独成节，内容不要过长。\n"
            "3. 行动建议用表格展示，列包含阶段、关键动作、责任主体、量化目标。\n"
            "4. 结尾给出一句明确的管理建议。\n\n"
            "数据摘要如下：\n"
            f"{chr(10).join(summary_blocks)}\n\n"
            "请注意：\n"
            "1. 结论必须引用摘要中的具体指标。\n"
            "2. 优先指出高耗能行业、结构变化和供需平衡风险。\n"
            "3. 建议要可执行，尽量分阶段。\n"
            "4. 如果发现数据可能存在口径问题，可以明确提示“建议核查”。"
        )

    def summarize_dataframe(self, name, dataframe):
        lines = [f"[{name}] 行数={len(dataframe)} 列={list(dataframe.columns)}"]

        if "能源消费总量及构成" in name and not dataframe.empty:
            latest = dataframe.iloc[-1]
            lines.append(
                f"最新年份={latest['年份']}，总量={latest['能源消费总量']:.2f}，"
                f"煤炭占比={latest['煤炭占比']:.1f}%，石油占比={latest['石油占比']:.1f}%，"
                f"天然气占比={latest['天然气占比']:.1f}%，"
                f"一次电力及其他能源占比={latest['一次电力及其他能源占比']:.1f}%"
            )

        elif "能源消费弹性系数" in name and not dataframe.empty:
            latest = dataframe.iloc[-1]
            lines.append(
                f"最新年份={latest['年份']}，能源弹性系数={latest['能源消费弹性系数']:.2f}，"
                f"电力弹性系数={latest['电力消费弹性系数']:.2f}"
            )

        elif "分行业能源消费量" in name and not dataframe.empty:
            top = dataframe[~dataframe["行业"].str.contains("消费总计", na=False)].sort_values(
                "能源消费总量", ascending=False
            ).head(5)
            for _, row in top.iterrows():
                lines.append(
                    f"行业={row['行业']}，综合能耗={row['能源消费总量']:.2f}，"
                    f"煤炭={row['煤炭']:.2f}，天然气={row['天然气']:.2f}，电力={row['电力']:.2f}"
                )

        elif "综合能源平衡表" in name and not dataframe.empty:
            latest_year = "2024" if "2024" in dataframe.columns else dataframe.columns[-1]
            focus = dataframe[
                dataframe["项目"].isin(["可供消费的能源总量", "一次能源生产量", "调入量", "调出量", "能源消费总量"])
            ]
            for _, row in focus.iterrows():
                lines.append(f"{row['项目']}={float(row[latest_year]):.2f}")

        elif "重点耗能企业单位产品能源消费" in name and not dataframe.empty:
            focus = dataframe.head(5)
            for _, row in focus.iterrows():
                lines.append(f"项目={row['项目']}，2020={row['2020']:.2f}，2023={row['2023']:.2f}，2024={row['2024']:.2f}")

        return "\n".join(lines)

    def run_llm_task(self, messages, waiting_text):
        self.set_buttons_enabled(False)
        self.ai_output.setPlainText(waiting_text)

        self.thread = QThread()
        self.worker = LLMWorker(
            api_key=self.api_input.text().strip(),
            base_url=self.base_url_input.text().strip(),
            model=self.model_input.text().strip(),
            messages=messages,
        )
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_llm_success)
        self.worker.failed.connect(self.on_llm_error)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.cleanup_thread)
        self.thread.start()

    def on_llm_success(self, content):
        cleaned = (content or "").strip()
        if self.looks_like_html(cleaned):
            self.ai_output.setHtml(self.wrap_html_report(cleaned))
        else:
            self.ai_output.setPlainText(cleaned or "调用成功，但模型未返回可显示内容。")
        self.set_buttons_enabled(True)

    def on_llm_error(self, error_text):
        self.ai_output.setPlainText(f"调用失败：\n{error_text}")
        self.set_buttons_enabled(True)

    def cleanup_thread(self):
        if self.worker is not None:
            self.worker.deleteLater()
            self.worker = None
        if self.thread is not None:
            self.thread.deleteLater()
            self.thread = None

    def set_buttons_enabled(self, enabled):
        self.btn_save.setEnabled(enabled)
        self.btn_test.setEnabled(enabled)
        self.btn_generate.setEnabled(enabled)

    def looks_like_html(self, text):
        lowered = text.lower()
        html_markers = ["<h1", "<h2", "<p", "<table", "<div", "<ul", "<ol", "<strong"]
        return any(marker in lowered for marker in html_markers)

    def wrap_html_report(self, content):
        return f"""
        <style>
            body {{
                font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
                color: #243447;
                background: #f4f7fb;
            }}
            .report {{
                background: #ffffff;
                border: 1px solid #d9e2ec;
                border-radius: 14px;
                padding: 24px 28px;
                line-height: 1.7;
            }}
            h1 {{
                font-size: 24px;
                margin: 0 0 8px 0;
                color: #16324f;
            }}
            h2 {{
                margin-top: 22px;
                margin-bottom: 10px;
                font-size: 18px;
                color: #1f5c7a;
                border-left: 4px solid #1f7a8c;
                padding-left: 10px;
            }}
            p {{
                margin: 8px 0;
            }}
            ul {{
                margin: 8px 0 8px 18px;
            }}
            li {{
                margin: 4px 0;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin: 12px 0 16px 0;
                background: #fbfdff;
            }}
            th {{
                background: #eaf3fb;
                color: #16324f;
                font-weight: 700;
            }}
            th, td {{
                border: 1px solid #d9e2ec;
                padding: 8px 10px;
                text-align: left;
                vertical-align: top;
            }}
            strong {{
                color: #b54708;
            }}
        </style>
        <div class="report">{content}</div>
        """
