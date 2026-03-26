from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)


class AnalysisWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.dataframes = {}
        self.init_ui()

    def init_ui(self):
        root = QVBoxLayout()
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(14)

        title = QLabel("节能诊断驾驶舱")
        title.setStyleSheet("font-size: 24px; font-weight: 700; color: #16324f;")
        root.addWidget(title)

        subtitle = QLabel("从总量、结构、行业、供需和数据质量五个维度给出诊断结果。")
        subtitle.setStyleSheet("font-size: 13px; color: #52606d;")
        root.addWidget(subtitle)

        self.btn_analyze = QPushButton("刷新诊断")
        self.btn_analyze.clicked.connect(self.run_analysis)
        self.btn_analyze.setStyleSheet(
            "QPushButton { background: #c2410c; color: white; border: none; border-radius: 10px; "
            "padding: 10px 14px; font-size: 14px; font-weight: 600; }"
            "QPushButton:hover { background: #9a3412; }"
        )
        root.addWidget(self.btn_analyze, alignment=Qt.AlignLeft)

        self.metric_grid = QGridLayout()
        self.metric_grid.setHorizontalSpacing(12)
        self.metric_grid.setVerticalSpacing(12)
        self.metric_cards = []
        for index in range(4):
            card = self.create_metric_card()
            self.metric_grid.addWidget(card, 0, index)
            self.metric_cards.append(card)
        root.addLayout(self.metric_grid)

        middle_row = QHBoxLayout()
        middle_row.setSpacing(14)
        root.addLayout(middle_row, 1)

        left_card = self.create_panel_card()
        left_layout = QVBoxLayout(left_card)
        left_layout.setSpacing(10)
        left_layout.addWidget(self.build_section_title("风险诊断"))

        self.risk_browser = QTextBrowser()
        self.risk_browser.setStyleSheet(self.browser_style())
        left_layout.addWidget(self.risk_browser, 1)
        middle_row.addWidget(left_card, 1)

        right_card = self.create_panel_card()
        right_layout = QVBoxLayout(right_card)
        right_layout.setSpacing(10)
        right_layout.addWidget(self.build_section_title("重点行业 TOP 5"))

        self.industry_table = QTableWidget()
        self.industry_table.setColumnCount(4)
        self.industry_table.setHorizontalHeaderLabels(["行业", "综合能耗", "电力", "判断"])
        self.industry_table.verticalHeader().setVisible(False)
        self.industry_table.setAlternatingRowColors(True)
        self.industry_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.industry_table.setSelectionMode(QTableWidget.NoSelection)
        self.industry_table.horizontalHeader().setStretchLastSection(True)
        self.industry_table.setStyleSheet(self.table_style())
        right_layout.addWidget(self.industry_table, 1)
        middle_row.addWidget(right_card, 1)

        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(14)
        root.addLayout(bottom_row, 1)

        advice_card = self.create_panel_card()
        advice_layout = QVBoxLayout(advice_card)
        advice_layout.setSpacing(10)
        advice_layout.addWidget(self.build_section_title("行动建议"))

        self.advice_browser = QTextBrowser()
        self.advice_browser.setStyleSheet(self.browser_style())
        advice_layout.addWidget(self.advice_browser, 1)
        bottom_row.addWidget(advice_card, 1)

        quality_card = self.create_panel_card()
        quality_layout = QVBoxLayout(quality_card)
        quality_layout.setSpacing(10)
        quality_layout.addWidget(self.build_section_title("数据质量与供需观察"))

        self.quality_browser = QTextBrowser()
        self.quality_browser.setStyleSheet(self.browser_style())
        quality_layout.addWidget(self.quality_browser, 1)
        bottom_row.addWidget(quality_card, 1)

        self.setLayout(root)
        self.setStyleSheet("QWidget { background: #f4f7fb; }")
        self.render_empty_state()

    def create_metric_card(self):
        card = QFrame()
        card.setStyleSheet(
            "QFrame { background: #ffffff; border: 1px solid #d9e2ec; border-radius: 14px; }"
        )
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(6)

        title = QLabel("--")
        title.setStyleSheet("font-size: 13px; color: #486581; font-weight: 600;")
        value = QLabel("--")
        value.setStyleSheet("font-size: 24px; color: #102a43; font-weight: 700;")
        note = QLabel("--")
        note.setWordWrap(True)
        note.setStyleSheet("font-size: 12px; color: #7b8794;")

        layout.addWidget(title)
        layout.addWidget(value)
        layout.addWidget(note)
        card.title_label = title
        card.value_label = value
        card.note_label = note
        return card

    def create_panel_card(self):
        card = QFrame()
        card.setStyleSheet(
            "QFrame { background: #ffffff; border: 1px solid #d9e2ec; border-radius: 16px; }"
        )
        return card

    def build_section_title(self, text):
        label = QLabel(text)
        label.setStyleSheet("font-size: 17px; font-weight: 700; color: #102a43;")
        return label

    def browser_style(self):
        return (
            "QTextBrowser { background: #fbfdff; border: 1px solid #d9e2ec; border-radius: 10px; "
            "padding: 8px; color: #243b53; }"
        )

    def table_style(self):
        return (
            "QTableWidget { background: #fbfdff; border: 1px solid #d9e2ec; border-radius: 10px; gridline-color: #d9e2ec; }"
            "QHeaderView::section { background: #eaf3fb; color: #16324f; border: none; padding: 8px; font-weight: 700; }"
        )

    def set_data(self, dataframes):
        self.dataframes = dataframes or {}

    def render_empty_state(self):
        metric_defaults = [
            ("总量状态", "--", "导入数据后显示"),
            ("结构状态", "--", "导入数据后显示"),
            ("行业状态", "--", "导入数据后显示"),
            ("供需状态", "--", "导入数据后显示"),
        ]
        for card, (title, value, note) in zip(self.metric_cards, metric_defaults):
            card.title_label.setText(title)
            card.value_label.setText(value)
            card.note_label.setText(note)

        self.risk_browser.setHtml("<p style='color:#7b8794;'>请先导入数据后再进行诊断。</p>")
        self.advice_browser.setHtml("<p style='color:#7b8794;'>系统将在这里生成分阶段行动建议。</p>")
        self.quality_browser.setHtml("<p style='color:#7b8794;'>系统将在这里展示供需平衡和数据质量观察。</p>")
        self.industry_table.setRowCount(0)

    def run_analysis(self):
        if not self.dataframes:
            self.render_empty_state()
            return

        total_df = self._find_dataset("能源消费总量及构成")
        elasticity_df = self._find_dataset("能源消费弹性系数")
        sector_df = self._find_dataset("分行业能源消费量")
        balance_df = self._find_dataset("综合能源平衡表")
        enterprise_df = self._find_dataset("重点耗能企业单位产品能源消费")

        dashboard = self.build_dashboard(total_df, elasticity_df, sector_df, balance_df, enterprise_df)
        self.update_metric_cards(dashboard["metrics"])
        self.risk_browser.setHtml(dashboard["risk_html"])
        self.advice_browser.setHtml(dashboard["advice_html"])
        self.quality_browser.setHtml(dashboard["quality_html"])
        self.populate_industry_table(dashboard["industry_rows"])

    def _find_dataset(self, keyword):
        for name, dataframe in self.dataframes.items():
            if keyword in name:
                return dataframe
        return None

    def build_dashboard(self, total_df, elasticity_df, sector_df, balance_df, enterprise_df):
        metrics = self.build_metric_data(total_df, elasticity_df, sector_df, balance_df)
        risk_html = self.build_risk_html(total_df, elasticity_df, sector_df, balance_df)
        advice_html = self.build_advice_html(total_df, elasticity_df, sector_df, enterprise_df)
        quality_html = self.build_quality_html(sector_df, balance_df, enterprise_df)
        industry_rows = self.build_industry_rows(sector_df)
        return {
            "metrics": metrics,
            "risk_html": risk_html,
            "advice_html": advice_html,
            "quality_html": quality_html,
            "industry_rows": industry_rows,
        }

    def build_metric_data(self, total_df, elasticity_df, sector_df, balance_df):
        metrics = []

        if total_df is not None and not total_df.empty:
            latest = total_df.iloc[-1]
            total_value = float(latest["能源消费总量"])
            metrics.append(
                ("总量状态", f"{total_value:.2f}", f"{int(latest['年份'])} 年总能耗（万吨标准煤）")
            )
            coal_share = float(latest["煤炭占比"])
            structure_note = "煤炭依赖偏高" if coal_share >= 50 else "结构已有优化空间"
            metrics.append(("结构状态", f"{coal_share:.1f}%", f"煤炭占比，{structure_note}"))
        else:
            metrics.extend([("总量状态", "--", "缺少总量数据"), ("结构状态", "--", "缺少结构数据")])

        if sector_df is not None and not sector_df.empty:
            filtered = sector_df[~sector_df["行业"].map(self.is_summary_industry)]
            top = filtered.sort_values("能源消费总量", ascending=False).head(1).iloc[0]
            metrics.append(("行业状态", top["行业"][:10], f"最高能耗行业 {top['能源消费总量']:.2f}"))
        else:
            metrics.append(("行业状态", "--", "缺少行业数据"))

        if balance_df is not None and not balance_df.empty:
            latest_year = "2024" if "2024" in balance_df.columns else balance_df.columns[-1]
            available = self.extract_balance_value(balance_df, "可供消费的能源总量", latest_year)
            consumption = self.extract_balance_value(balance_df, "能源消费总量", latest_year)
            gap = available - consumption
            status = "供需偏紧" if gap < 100 else "供需可控"
            metrics.append(("供需状态", f"{gap:.2f}", f"{latest_year} 年差额，{status}"))
        else:
            metrics.append(("供需状态", "--", "缺少供需数据"))

        if elasticity_df is not None and not elasticity_df.empty:
            latest = elasticity_df.iloc[-1]
            elasticity = float(latest["能源消费弹性系数"])
            metrics[1] = (metrics[1][0], metrics[1][1], f"{metrics[1][2]}；弹性系数 {elasticity:.2f}")

        return metrics[:4]

    def update_metric_cards(self, metrics):
        for card, (title, value, note) in zip(self.metric_cards, metrics):
            card.title_label.setText(title)
            card.value_label.setText(value)
            card.note_label.setText(note)

    def build_risk_html(self, total_df, elasticity_df, sector_df, balance_df):
        items = []

        if elasticity_df is not None and not elasticity_df.empty:
            latest = elasticity_df.iloc[-1]
            elasticity = float(latest["能源消费弹性系数"])
            if elasticity > 1:
                level = "高"
                color = "#b91c1c"
                conclusion = "能耗增长依赖较强，需立即压降高耗能增量。"
            elif elasticity >= 0.5:
                level = "中"
                color = "#b45309"
                conclusion = "处于相对脱钩区间，但仍需继续优化。"
            else:
                level = "低"
                color = "#0f766e"
                conclusion = "脱钩状态较好。"
            items.append(
                f"<p><b style='color:{color};'>[{level}] 宏观效率</b><br>"
                f"最新能源消费弹性系数为 <b>{elasticity:.2f}</b>。{conclusion}</p>"
            )

        if total_df is not None and not total_df.empty:
            latest = total_df.iloc[-1]
            coal_share = float(latest["煤炭占比"])
            clean_share = float(latest["天然气占比"]) + float(latest["一次电力及其他能源占比"])
            level = "高" if coal_share >= 55 else "中" if coal_share >= 45 else "低"
            color = "#b91c1c" if level == "高" else "#b45309" if level == "中" else "#0f766e"
            items.append(
                f"<p><b style='color:{color};'>[{level}] 能源结构</b><br>"
                f"煤炭占比 <b>{coal_share:.1f}%</b>，较清洁能源占比 <b>{clean_share:.1f}%</b>。"
                "建议继续推动天然气和一次电力替代。</p>"
            )

        if balance_df is not None and not balance_df.empty:
            latest_year = "2024" if "2024" in balance_df.columns else balance_df.columns[-1]
            available = self.extract_balance_value(balance_df, "可供消费的能源总量", latest_year)
            consumption = self.extract_balance_value(balance_df, "能源消费总量", latest_year)
            gap = available - consumption
            level = "高" if gap < 20 else "中" if gap < 100 else "低"
            color = "#b91c1c" if level == "高" else "#b45309" if level == "中" else "#0f766e"
            items.append(
                f"<p><b style='color:{color};'>[{level}] 供需平衡</b><br>"
                f"{latest_year} 年供需差额为 <b>{gap:.2f}</b> 万吨标准煤。"
                "缓冲空间过小意味着外部冲击下的保障能力需要重点跟踪。</p>"
            )

        if sector_df is not None and not sector_df.empty:
            filtered = sector_df[~sector_df["行业"].map(self.is_summary_industry)]
            top = filtered.sort_values("能源消费总量", ascending=False).head(3)
            names = "、".join(top["行业"].tolist())
            items.append(
                f"<p><b style='color:#b45309;'>[中] 行业集中度</b><br>"
                f"高耗能行业主要集中在 <b>{names}</b>，建议优先纳入专项改造清单。</p>"
            )

        return "".join(items) if items else "<p style='color:#7b8794;'>暂无可用风险诊断结果。</p>"

    def build_advice_html(self, total_df, elasticity_df, sector_df, enterprise_df):
        short_term = []
        mid_term = []
        long_term = []

        short_term.append("对高耗能行业开展数据口径核查，统一煤炭、电力和综合能耗的统计边界。")

        if sector_df is not None and not sector_df.empty:
            filtered = sector_df[~sector_df["行业"].map(self.is_summary_industry)]
            top = filtered.sort_values("能源消费总量", ascending=False).head(2)
            for _, row in top.iterrows():
                short_term.append(f"优先对 {row['行业']} 开展能效诊断和电气化改造方案评估。")

        if total_df is not None and not total_df.empty:
            latest = total_df.iloc[-1]
            coal_share = float(latest["煤炭占比"])
            if coal_share >= 45:
                mid_term.append("围绕煤炭占比较高的结构特征，制定天然气替代和绿电采购计划。")
            else:
                mid_term.append("继续扩大较清洁能源使用比例，巩固结构优化成果。")

        if enterprise_df is not None and not enterprise_df.empty:
            mid_term.append("把重点耗能企业单位产品能耗指标纳入月度跟踪，形成能效对标台账。")

        if elasticity_df is not None and not elasticity_df.empty:
            latest = elasticity_df.iloc[-1]
            if float(latest["电力消费弹性系数"]) > 1:
                long_term.append("建设区域级负荷监测与需求响应机制，提前应对用电弹性持续偏高。")

        long_term.append("建立年度弹性系数、煤炭占比和供需差额联动考核机制，推动结构优化长期化。")

        return (
            "<h3 style='color:#102a43;'>立即处理</h3><ul>"
            + "".join(f"<li>{item}</li>" for item in short_term)
            + "</ul><h3 style='color:#102a43;'>中期优化</h3><ul>"
            + "".join(f"<li>{item}</li>" for item in mid_term)
            + "</ul><h3 style='color:#102a43;'>长期机制</h3><ul>"
            + "".join(f"<li>{item}</li>" for item in long_term)
            + "</ul>"
        )

    def build_quality_html(self, sector_df, balance_df, enterprise_df):
        blocks = []

        if balance_df is not None and not balance_df.empty:
            latest_year = "2024" if "2024" in balance_df.columns else balance_df.columns[-1]
            available = self.extract_balance_value(balance_df, "可供消费的能源总量", latest_year)
            consumption = self.extract_balance_value(balance_df, "能源消费总量", latest_year)
            blocks.append(
                f"<p><b>供需差额</b>：{latest_year} 年可供消费总量为 <b>{available:.2f}</b>，"
                f"能源消费总量为 <b>{consumption:.2f}</b>，差额 <b>{available - consumption:.2f}</b>。</p>"
            )

        if sector_df is not None and not sector_df.empty:
            filtered = sector_df[~sector_df["行业"].str.contains("消费总计", na=False)]
            sector_total = filtered["能源消费总量"].sum()
            coal_total = filtered["煤炭"].sum()
            blocks.append(
                f"<p><b>行业统计观察</b>：分行业综合能耗合计 <b>{sector_total:.2f}</b>，"
                f"煤炭分项合计 <b>{coal_total:.2f}</b>。如分项显著高于综合能耗，建议核查折算口径。</p>"
            )

        if enterprise_df is not None and not enterprise_df.empty:
            improved = (enterprise_df["2024"] < enterprise_df["2020"]).sum()
            blocks.append(
                f"<p><b>重点企业指标</b>：在已加载指标中，有 <b>{int(improved)}</b> 项 2024 年优于 2020 年，"
                "说明技术改造已有成效，但仍需继续跟踪边际改善速度。</p>"
            )

        return "".join(blocks) if blocks else "<p style='color:#7b8794;'>暂无质量或供需观察结果。</p>"

    def build_industry_rows(self, sector_df):
        rows = []
        if sector_df is None or sector_df.empty:
            return rows

        filtered = sector_df[~sector_df["行业"].map(self.is_summary_industry)]
        top = filtered.sort_values("能源消费总量", ascending=False).head(5)

        for _, row in top.iterrows():
            electric_ratio = float(row["电力"]) / float(row["能源消费总量"]) if float(row["能源消费总量"]) else 0
            if electric_ratio >= 0.25:
                judgement = "高电耗，适合优先推进绿电替代"
            elif float(row["煤炭"]) > float(row["电力"]):
                judgement = "煤炭依赖较强，适合燃料替代"
            else:
                judgement = "结构相对均衡，可持续做能效优化"

            rows.append(
                [
                    row["行业"],
                    f"{float(row['能源消费总量']):.2f}",
                    f"{float(row['电力']):.2f}",
                    judgement,
                ]
            )
        return rows

    def populate_industry_table(self, rows):
        self.industry_table.setRowCount(len(rows))
        for row_index, row_data in enumerate(rows):
            for col_index, value in enumerate(row_data):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignCenter if col_index < 3 else Qt.AlignLeft | Qt.AlignVCenter)
                self.industry_table.setItem(row_index, col_index, item)
        self.industry_table.resizeColumnsToContents()

    def extract_balance_value(self, balance_df, item_name, year_column):
        row = balance_df[balance_df["项目"] == item_name]
        if row.empty:
            return 0.0
        return float(row.iloc[0][year_column])

    def is_summary_industry(self, text):
        normalized = str(text).replace(" ", "")
        return "消费总计" in normalized or normalized in {"合计", "总计"}
