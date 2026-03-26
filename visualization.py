import matplotlib

matplotlib.use("Qt5Agg")
matplotlib.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
matplotlib.rcParams["axes.unicode_minus"] = False

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)


DATASET_LABELS = {
    "total": "总量与结构",
    "elasticity": "弹性系数",
    "sector": "行业能耗",
    "balance": "供需平衡",
    "enterprise": "重点企业指标",
    "generic": "通用数据",
}


class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=11, height=7, dpi=100):
        self.figure = Figure(figsize=(width, height), dpi=dpi, facecolor="#ffffff")
        super().__init__(self.figure)
        self.setParent(parent)


class VisualizationWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.dataframes = {}
        self.current_dataset_type = "generic"
        self.chart_registry = {
            "total": [
                ("总量趋势", self.plot_total_trend),
                ("结构面积图", self.plot_energy_structure),
                ("结构玫瑰图", self.plot_energy_polar),
            ],
            "elasticity": [
                ("双折线", self.plot_elasticity_lines),
                ("增长对比柱图", self.plot_growth_bars),
                ("弹性散点图", self.plot_elasticity_scatter),
            ],
            "sector": [
                ("行业排名", self.plot_sector_ranking),
                ("能源构成热力图", self.plot_sector_heatmap),
                ("TOP5 占比环图", self.plot_sector_donut),
            ],
            "balance": [
                ("供需对比", self.plot_balance_supply_demand),
                ("历年平衡趋势", self.plot_balance_trend),
                ("供需瀑布图", self.plot_balance_waterfall),
            ],
            "enterprise": [
                ("指标改善斜率图", self.plot_enterprise_slope),
                ("2024 指标排名", self.plot_enterprise_bar),
                ("历年热力图", self.plot_enterprise_heatmap),
            ],
            "generic": [
                ("前几列趋势", self.plot_generic_line),
                ("末列排序", self.plot_generic_bar),
            ],
        }
        self.init_ui()

    def init_ui(self):
        root = QVBoxLayout()
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(14)

        title = QLabel("可视化分析驾驶舱")
        title.setStyleSheet("font-size: 24px; font-weight: 700; color: #16324f;")
        root.addWidget(title)

        subtitle = QLabel("根据已导入的数据自动推荐合适的图表类型，并展示关键摘要。")
        subtitle.setStyleSheet("font-size: 13px; color: #52606d;")
        root.addWidget(subtitle)

        content = QHBoxLayout()
        content.setSpacing(16)
        root.addLayout(content, 1)

        left_panel = self.create_card()
        left_panel.setMaximumWidth(360)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setSpacing(12)

        control_title = QLabel("图表控制")
        control_title.setStyleSheet("font-size: 16px; font-weight: 700; color: #102a43;")
        left_layout.addWidget(control_title)

        self.file_combo = QComboBox()
        self.file_combo.currentTextChanged.connect(self.on_file_changed)
        self.file_combo.setStyleSheet(self.combo_style())
        left_layout.addWidget(self.build_labeled_block("数据文件", self.file_combo))

        self.chart_combo = QComboBox()
        self.chart_combo.setStyleSheet(self.combo_style())
        left_layout.addWidget(self.build_labeled_block("图表类型", self.chart_combo))

        self.btn_plot = QPushButton("生成图表")
        self.btn_plot.clicked.connect(self.plot_chart)
        self.btn_plot.setStyleSheet(self.primary_button_style())
        left_layout.addWidget(self.btn_plot)

        self.dataset_badge = QLabel("未加载数据")
        self.dataset_badge.setAlignment(Qt.AlignCenter)
        self.dataset_badge.setStyleSheet(
            "background: #eaf2ff; color: #1d4ed8; border-radius: 14px; padding: 6px 10px; font-weight: 600;"
        )
        left_layout.addWidget(self.dataset_badge)

        metrics_card = self.create_inner_card()
        metrics_layout = QGridLayout(metrics_card)
        metrics_layout.setHorizontalSpacing(8)
        metrics_layout.setVerticalSpacing(8)
        self.metric_cards = []
        for idx in range(3):
            box = QLabel("--")
            box.setAlignment(Qt.AlignCenter)
            box.setWordWrap(True)
            box.setStyleSheet(
                "background: #f8fbff; border: 1px solid #d9e2ec; border-radius: 10px; padding: 12px; "
                "font-size: 13px; color: #243b53;"
            )
            metrics_layout.addWidget(box, idx, 0)
            self.metric_cards.append(box)
        left_layout.addWidget(metrics_card)

        summary_title = QLabel("数据摘要")
        summary_title.setStyleSheet("font-size: 15px; font-weight: 700; color: #102a43;")
        left_layout.addWidget(summary_title)

        self.summary_browser = QTextBrowser()
        self.summary_browser.setStyleSheet(
            "background: #fbfdff; border: 1px solid #d9e2ec; border-radius: 10px; padding: 8px; color: #243b53;"
        )
        left_layout.addWidget(self.summary_browser, 1)

        content.addWidget(left_panel)

        right_panel = self.create_card()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setSpacing(12)

        self.chart_title = QLabel("图表预览")
        self.chart_title.setStyleSheet("font-size: 18px; font-weight: 700; color: #102a43;")
        right_layout.addWidget(self.chart_title)

        self.chart_hint = QLabel("导入数据后，系统会根据文件自动推荐更合适的图表。")
        self.chart_hint.setStyleSheet("font-size: 13px; color: #52606d;")
        right_layout.addWidget(self.chart_hint)

        self.canvas = MplCanvas(self)
        right_layout.addWidget(self.canvas, 1)
        content.addWidget(right_panel, 1)

        self.setLayout(root)
        self.setStyleSheet("QWidget { background: #f4f7fb; }")
        self.reset_empty_state()

    def create_card(self):
        card = QFrame()
        card.setStyleSheet(
            "QFrame { background: #ffffff; border: 1px solid #d9e2ec; border-radius: 16px; }"
        )
        return card

    def create_inner_card(self):
        card = QFrame()
        card.setStyleSheet(
            "QFrame { background: #ffffff; border: 1px dashed #cbd5e1; border-radius: 12px; }"
        )
        return card

    def combo_style(self):
        return (
            "QComboBox { background: #ffffff; border: 1px solid #cbd5e1; border-radius: 10px; padding: 8px 10px; }"
            "QComboBox::drop-down { border: none; }"
        )

    def primary_button_style(self):
        return (
            "QPushButton { background: #0f766e; color: white; border: none; border-radius: 10px; "
            "padding: 10px 14px; font-size: 14px; font-weight: 600; }"
            "QPushButton:hover { background: #115e59; }"
        )

    def build_labeled_block(self, label_text, widget):
        wrapper = QWidget()
        layout = QVBoxLayout(wrapper)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        label = QLabel(label_text)
        label.setStyleSheet("font-size: 13px; color: #486581; font-weight: 600;")
        layout.addWidget(label)
        layout.addWidget(widget)
        return wrapper

    def set_data(self, dataframes):
        self.dataframes = dataframes or {}
        self.file_combo.blockSignals(True)
        self.file_combo.clear()
        if not self.dataframes:
            self.file_combo.addItem("请先导入数据")
            self.file_combo.blockSignals(False)
            self.reset_empty_state()
            return

        for name in self.dataframes.keys():
            self.file_combo.addItem(name)
        self.file_combo.blockSignals(False)
        self.on_file_changed(self.file_combo.currentText())

    def reset_empty_state(self):
        self.dataset_badge.setText("未加载数据")
        self.summary_browser.setHtml("<p style='color:#7b8794;'>请先导入 Excel 数据文件。</p>")
        self.chart_combo.clear()
        self.chart_combo.addItem("暂无可用图表")
        for card in self.metric_cards:
            card.setText("--")
        self.chart_title.setText("图表预览")
        self.chart_hint.setText("导入数据后，系统会根据文件自动推荐更合适的图表。")
        figure = self.canvas.figure
        figure.clear()
        ax = figure.add_subplot(111)
        ax.axis("off")
        ax.text(0.5, 0.5, "请先导入数据", ha="center", va="center", fontsize=18, color="#7b8794")
        self.canvas.draw()

    def on_file_changed(self, filename):
        if filename not in self.dataframes:
            return

        dataframe = self.dataframes[filename]
        self.current_dataset_type = self.infer_dataset_type(filename, dataframe)
        self.dataset_badge.setText(DATASET_LABELS.get(self.current_dataset_type, "通用数据"))
        self.chart_combo.clear()
        for label, _ in self.chart_registry[self.current_dataset_type]:
            self.chart_combo.addItem(label)

        self.update_summary_panel(filename, dataframe)
        self.plot_chart()

    def infer_dataset_type(self, filename, dataframe):
        if "能源消费总量及构成" in filename:
            return "total"
        if "能源消费弹性系数" in filename:
            return "elasticity"
        if "分行业能源消费量" in filename:
            return "sector"
        if "综合能源平衡表" in filename:
            return "balance"
        if "重点耗能企业单位产品能源消费" in filename:
            return "enterprise"

        columns = set(dataframe.columns)
        if {"年份", "能源消费总量", "煤炭占比", "天然气占比"}.issubset(columns):
            return "total"
        if {"年份", "能源消费增长率", "能源消费弹性系数", "电力消费弹性系数"}.issubset(columns):
            return "elasticity"
        if {"行业", "能源消费总量", "煤炭", "电力"}.issubset(columns):
            return "sector"
        if {"项目", "2020", "2023", "2024"}.issubset(columns):
            return "enterprise"
        if {"项目", "2024"}.issubset(columns):
            return "balance"
        return "generic"

    def update_summary_panel(self, filename, dataframe):
        summary_html, metrics = self.build_summary_html(filename, dataframe, self.current_dataset_type)
        self.summary_browser.setHtml(summary_html)
        for card, text in zip(self.metric_cards, metrics):
            card.setText(text)

    def build_summary_html(self, filename, dataframe, dataset_type):
        base_info = [
            f"<p><b>文件</b>：{filename}</p>",
            f"<p><b>行数</b>：{len(dataframe)}，<b>列数</b>：{len(dataframe.columns)}</p>",
        ]
        metrics = ["--", "--", "--"]

        if dataset_type == "total":
            latest = dataframe.iloc[-1]
            metrics = [
                f"总量\n{latest['能源消费总量']:.2f}",
                f"煤炭占比\n{latest['煤炭占比']:.1f}%",
                f"清洁占比\n{latest['天然气占比'] + latest['一次电力及其他能源占比']:.1f}%",
            ]
            base_info.append(
                "<p>这组数据适合看总量变化和能源结构迁移，系统优先推荐趋势线、面积图和玫瑰图。</p>"
            )

        elif dataset_type == "elasticity":
            latest = dataframe.iloc[-1]
            metrics = [
                f"年份\n{latest['年份']}",
                f"能源弹性\n{latest['能源消费弹性系数']:.2f}",
                f"电力弹性\n{latest['电力消费弹性系数']:.2f}",
            ]
            base_info.append("<p>这组数据适合看增长与弹性的联动变化，适合折线、柱图和散点图。</p>")

        elif dataset_type == "sector":
            filtered = dataframe[~dataframe["行业"].str.contains("消费总计", na=False)]
            top = filtered.sort_values("能源消费总量", ascending=False).head(1).iloc[0]
            metrics = [
                f"行业数\n{len(filtered)}",
                f"最高能耗\n{top['能源消费总量']:.2f}",
                f"TOP 行业\n{top['行业']}",
            ]
            base_info.append("<p>这组数据适合看行业分层和能源构成差异，推荐排名、热力图和环图。</p>")

        elif dataset_type == "balance":
            latest_year = "2024"
            available = dataframe.loc[dataframe["项目"] == "可供消费的能源总量", latest_year]
            consumption = dataframe.loc[dataframe["项目"] == "能源消费总量", latest_year]
            available_val = float(available.iloc[0]) if not available.empty else 0.0
            consumption_val = float(consumption.iloc[0]) if not consumption.empty else 0.0
            metrics = [
                f"供给\n{available_val:.2f}",
                f"消费\n{consumption_val:.2f}",
                f"差额\n{available_val - consumption_val:.2f}",
            ]
            base_info.append("<p>这组数据适合看供给、调入调出和消费之间的平衡关系，推荐对比图和瀑布图。</p>")

        elif dataset_type == "enterprise":
            latest_sorted = dataframe.sort_values("2024", ascending=False).head(1).iloc[0]
            metrics = [
                f"指标数\n{len(dataframe)}",
                f"最高 2024\n{latest_sorted['2024']:.2f}",
                f"重点指标\n{latest_sorted['项目'][:10]}",
            ]
            base_info.append("<p>这组数据适合看单位产品能耗改善趋势，推荐斜率图、排序图和热力图。</p>")

        else:
            metrics = [
                f"首列\n{dataframe.columns[0]}",
                f"末列\n{dataframe.columns[-1]}",
                f"数值列\n{max(len(dataframe.columns) - 1, 0)}",
            ]
            base_info.append("<p>这是通用数据集，系统会选择兼容性最高的趋势图和排序图。</p>")

        return "".join(base_info), metrics

    def plot_chart(self):
        filename = self.file_combo.currentText()
        if filename not in self.dataframes:
            return

        dataframe = self.dataframes[filename]
        chart_label = self.chart_combo.currentText()
        chart_map = dict(self.chart_registry[self.current_dataset_type])
        plotter = chart_map.get(chart_label)
        if plotter is None:
            return

        self.chart_title.setText(chart_label)
        self.chart_hint.setText(f"当前数据类型：{DATASET_LABELS.get(self.current_dataset_type, '通用数据')}")
        figure = self.canvas.figure
        figure.clear()

        try:
            plotter(figure, dataframe, filename)
        except Exception as exc:
            ax = figure.add_subplot(111)
            ax.axis("off")
            ax.text(0.5, 0.55, "绘图失败", ha="center", va="center", fontsize=20, color="#c2410c")
            ax.text(0.5, 0.45, str(exc), ha="center", va="center", fontsize=11, color="#7b8794")

        figure.tight_layout()
        self.canvas.draw()

    def style_axes(self, ax):
        ax.set_facecolor("#fbfdff")
        for spine in ax.spines.values():
            spine.set_color("#d9e2ec")
        ax.tick_params(colors="#486581")
        ax.title.set_color("#102a43")
        ax.grid(True, axis="y", linestyle="--", alpha=0.25, color="#9fb3c8")

    def plot_total_trend(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        years = dataframe["年份"]
        total = dataframe["能源消费总量"]
        ax.plot(years, total, color="#0f766e", linewidth=3, marker="o", markersize=6)
        ax.fill_between(years, total, color="#99f6e4", alpha=0.35)
        ax.set_title("能源消费总量趋势")
        ax.set_xlabel("年份")
        ax.set_ylabel("万吨标准煤")
        self.style_axes(ax)

    def plot_energy_structure(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        years = dataframe["年份"]
        columns = ["煤炭占比", "石油占比", "天然气占比", "一次电力及其他能源占比"]
        colors = ["#334e68", "#d97706", "#0f766e", "#2563eb"]
        values = [dataframe[col] for col in columns]
        ax.stackplot(years, values, labels=columns, colors=colors, alpha=0.92)
        ax.set_title("能源结构面积图")
        ax.set_xlabel("年份")
        ax.set_ylabel("占比 (%)")
        self.style_axes(ax)
        ax.legend(loc="upper right", frameon=False)

    def plot_energy_polar(self, figure, dataframe, filename):
        ax = figure.add_subplot(111, projection="polar")
        latest = dataframe.iloc[-1]
        labels = ["煤炭占比", "石油占比", "天然气占比", "一次电力及其他能源占比"]
        values = [float(latest[label]) for label in labels]
        angles = [n / float(len(labels)) * 2 * 3.1415926 for n in range(len(labels))]
        angles += angles[:1]
        values += values[:1]
        ax.plot(angles, values, color="#7c3aed", linewidth=2.5)
        ax.fill(angles, values, color="#c4b5fd", alpha=0.45)
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(["煤炭", "石油", "天然气", "一次电力"], color="#334e68")
        ax.set_title(f"{int(latest['年份'])} 年能源结构玫瑰图", pad=20, color="#102a43")
        ax.grid(color="#d9e2ec", alpha=0.6)

    def plot_elasticity_lines(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        years = dataframe["年份"]
        ax.plot(years, dataframe["能源消费弹性系数"], color="#0f766e", linewidth=2.5, marker="o", label="能源弹性")
        ax.plot(years, dataframe["电力消费弹性系数"], color="#2563eb", linewidth=2.5, marker="o", label="电力弹性")
        ax.axhline(1.0, color="#c2410c", linestyle="--", linewidth=1.5, label="弹性=1")
        ax.set_title("能源与电力弹性系数变化")
        ax.set_xlabel("年份")
        ax.set_ylabel("弹性系数")
        self.style_axes(ax)
        ax.legend(frameon=False)

    def plot_growth_bars(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        years = list(map(str, dataframe["年份"]))
        x = list(range(len(years)))
        width = 0.38
        ax.bar([i - width / 2 for i in x], dataframe["能源消费增长率"], width=width, color="#14b8a6", label="能源增长率")
        ax.bar([i + width / 2 for i in x], dataframe["电力消费增长率"], width=width, color="#60a5fa", label="电力增长率")
        ax.set_xticks(x)
        ax.set_xticklabels(years, rotation=45)
        ax.set_title("增长率对比")
        ax.set_ylabel("%")
        self.style_axes(ax)
        ax.legend(frameon=False)

    def plot_elasticity_scatter(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        scatter = ax.scatter(
            dataframe["能源消费弹性系数"],
            dataframe["电力消费弹性系数"],
            c=dataframe["年份"],
            cmap="viridis",
            s=70,
            alpha=0.85,
        )
        for _, row in dataframe.tail(8).iterrows():
            ax.text(row["能源消费弹性系数"], row["电力消费弹性系数"], str(int(row["年份"])), fontsize=8)
        ax.axvline(1.0, linestyle="--", color="#94a3b8")
        ax.axhline(1.0, linestyle="--", color="#94a3b8")
        ax.set_title("能源弹性与电力弹性散点图")
        ax.set_xlabel("能源消费弹性系数")
        ax.set_ylabel("电力消费弹性系数")
        self.style_axes(ax)
        figure.colorbar(scatter, ax=ax, label="年份")

    def plot_sector_ranking(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        filtered = dataframe[~dataframe["行业"].str.contains("消费总计", na=False)]
        top = filtered.sort_values("能源消费总量", ascending=False).head(10).sort_values("能源消费总量")
        colors = ["#0f766e" if value == top["能源消费总量"].max() else "#7dd3c7" for value in top["能源消费总量"]]
        ax.barh(top["行业"], top["能源消费总量"], color=colors)
        ax.set_title("行业能耗排名 TOP 10")
        ax.set_xlabel("万吨标准煤")
        self.style_axes(ax)
        ax.grid(True, axis="x", linestyle="--", alpha=0.25, color="#9fb3c8")

    def plot_sector_heatmap(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        filtered = dataframe[~dataframe["行业"].str.contains("消费总计", na=False)]
        top = filtered.sort_values("能源消费总量", ascending=False).head(8)
        cols = ["煤炭", "原油", "天然气", "电力"]
        matrix = top[cols].to_numpy()
        image = ax.imshow(matrix, cmap="YlGnBu", aspect="auto")
        ax.set_yticks(range(len(top)))
        ax.set_yticklabels(top["行业"])
        ax.set_xticks(range(len(cols)))
        ax.set_xticklabels(cols)
        ax.set_title("高耗能行业能源构成热力图")
        for i in range(matrix.shape[0]):
            for j in range(matrix.shape[1]):
                ax.text(j, i, f"{matrix[i, j]:.0f}", ha="center", va="center", fontsize=8, color="#102a43")
        figure.colorbar(image, ax=ax, fraction=0.03, pad=0.02)

    def plot_sector_donut(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        filtered = dataframe[~dataframe["行业"].str.contains("消费总计", na=False)]
        top = filtered.sort_values("能源消费总量", ascending=False).head(5)
        colors = ["#0f766e", "#14b8a6", "#2dd4bf", "#5eead4", "#99f6e4"]
        wedges, texts, autotexts = ax.pie(
            top["能源消费总量"],
            labels=top["行业"],
            autopct="%1.1f%%",
            startangle=90,
            colors=colors,
            pctdistance=0.8,
            wedgeprops=dict(width=0.38, edgecolor="white"),
        )
        ax.text(0, 0, "TOP5\n占比", ha="center", va="center", fontsize=16, color="#102a43", fontweight="bold")
        ax.set_title("高耗能行业 TOP5 占比")

    def plot_balance_supply_demand(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        latest_year = "2024"
        subset = dataframe[dataframe["项目"].isin(["可供消费的能源总量", "一次能源生产量", "调入量", "调出量", "能源消费总量"])]
        colors = ["#2563eb", "#0f766e", "#22c55e", "#f97316", "#7c3aed"]
        ax.bar(subset["项目"], subset[latest_year], color=colors)
        ax.set_title("2024 年供需平衡关键项对比")
        ax.set_ylabel("万吨标准煤")
        ax.tick_params(axis="x", rotation=20)
        self.style_axes(ax)

    def plot_balance_trend(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        year_columns = [column for column in dataframe.columns if str(column).isdigit()]
        targets = {
            "可供消费的能源总量": "#2563eb",
            "一次能源生产量": "#0f766e",
            "能源消费总量": "#7c3aed",
        }
        for item, color in targets.items():
            row = dataframe[dataframe["项目"] == item]
            if not row.empty:
                values = row.iloc[0][year_columns].tolist()
                ax.plot(year_columns, values, marker="o", linewidth=2.5, color=color, label=item)
        ax.set_title("历年供给与消费趋势")
        ax.set_ylabel("万吨标准煤")
        self.style_axes(ax)
        ax.legend(frameon=False)

    def plot_balance_waterfall(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        latest_year = "2024"
        items = ["一次能源生产量", "调入量", "调出量", "能源消费总量"]
        values = []
        labels = []
        for item in items:
            row = dataframe[dataframe["项目"] == item]
            if not row.empty:
                value = float(row.iloc[0][latest_year])
                if item == "调出量":
                    value = -value
                if item == "能源消费总量":
                    value = -value
                values.append(value)
                labels.append(item)

        cumulative = [0]
        for value in values[:-1]:
            cumulative.append(cumulative[-1] + value)

        colors = ["#0f766e" if value >= 0 else "#dc2626" for value in values]
        ax.bar(labels, values, bottom=cumulative, color=colors)
        ax.axhline(0, color="#94a3b8", linewidth=1)
        ax.set_title("2024 年供需瀑布图")
        ax.set_ylabel("万吨标准煤")
        ax.tick_params(axis="x", rotation=20)
        self.style_axes(ax)

    def plot_enterprise_slope(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        top = dataframe.sort_values("2024", ascending=False).head(8)
        for idx, (_, row) in enumerate(top.iterrows()):
            y_values = [row["2020"], row["2023"], row["2024"]]
            color = "#0f766e" if row["2024"] <= row["2020"] else "#dc2626"
            ax.plot([0, 1, 2], y_values, marker="o", linewidth=2, color=color, alpha=0.85)
            ax.text(2.05, y_values[-1], row["项目"][:12], fontsize=8, va="center")
        ax.set_xticks([0, 1, 2])
        ax.set_xticklabels(["2020", "2023", "2024"])
        ax.set_title("重点指标改善斜率图")
        ax.set_ylabel("单位产品能耗")
        self.style_axes(ax)

    def plot_enterprise_bar(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        top = dataframe.sort_values("2024", ascending=False).head(10).sort_values("2024")
        ax.barh(top["项目"], top["2024"], color="#2563eb")
        ax.set_title("2024 年重点企业指标排名")
        ax.set_xlabel("2024 指标值")
        self.style_axes(ax)
        ax.grid(True, axis="x", linestyle="--", alpha=0.25, color="#9fb3c8")

    def plot_enterprise_heatmap(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        top = dataframe.head(10)
        cols = ["2020", "2023", "2024"]
        matrix = top[cols].to_numpy()
        image = ax.imshow(matrix, cmap="YlOrRd", aspect="auto")
        ax.set_yticks(range(len(top)))
        ax.set_yticklabels([item[:12] for item in top["项目"]])
        ax.set_xticks(range(len(cols)))
        ax.set_xticklabels(cols)
        ax.set_title("重点指标历年热力图")
        for i in range(matrix.shape[0]):
            for j in range(matrix.shape[1]):
                ax.text(j, i, f"{matrix[i, j]:.1f}", ha="center", va="center", fontsize=8, color="#102a43")
        figure.colorbar(image, ax=ax, fraction=0.03, pad=0.02)

    def plot_generic_line(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        x_col = dataframe.columns[0]
        numeric_cols = dataframe.columns[1: min(4, len(dataframe.columns))]
        for column in numeric_cols:
            ax.plot(dataframe[x_col], dataframe[column], marker="o", linewidth=2, label=column)
        ax.set_title("通用趋势图")
        ax.set_xlabel(x_col)
        self.style_axes(ax)
        if len(numeric_cols) > 0:
            ax.legend(frameon=False)

    def plot_generic_bar(self, figure, dataframe, filename):
        ax = figure.add_subplot(111)
        x_col = dataframe.columns[0]
        y_col = dataframe.columns[-1]
        subset = dataframe.head(12)
        ax.bar(subset[x_col].astype(str), subset[y_col], color="#0f766e")
        ax.set_title("通用排序图")
        ax.set_xlabel(x_col)
        ax.set_ylabel(y_col)
        ax.tick_params(axis="x", rotation=30)
        self.style_axes(ax)
