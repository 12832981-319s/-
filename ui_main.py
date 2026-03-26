from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow, QTabWidget

from app_paths import resource_path
from ai_module import AIAnalysisWidget
from analysis_engine import AnalysisWidget
from data_import import DataImportWidget
from visualization import VisualizationWidget


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("能源系统智能分析与节能决策平台")
        icon_path = resource_path("assets", "app_icon.ico")
        self.setWindowIcon(QIcon(str(icon_path)) if icon_path.exists() else QIcon())

        import PyQt5.QtWidgets as QtWidgets

        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        width, height = int(screen.width() * 0.8), int(screen.height() * 0.8)
        self.resize(width, height)
        self.move((screen.width() - width) // 2, (screen.height() - height) // 2)

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.data_import_tab = DataImportWidget()
        self.visualization_tab = VisualizationWidget()
        self.analysis_tab = AnalysisWidget()
        self.ai_tab = AIAnalysisWidget()

        self.tabs.addTab(self.data_import_tab, "数据导入")
        self.tabs.addTab(self.visualization_tab, "可视化分析")
        self.tabs.addTab(self.analysis_tab, "节能诊断")
        self.tabs.addTab(self.ai_tab, "AI 决策建议")
        self.tabs.currentChanged.connect(self.on_tab_changed)

        self.statusBar().showMessage("先导入 Excel 数据，再切换到其他分析页面。")

    def on_tab_changed(self, index):
        data = self.data_import_tab.get_data()
        if index == 1:
            self.visualization_tab.set_data(data)
        elif index == 2:
            self.analysis_tab.set_data(data)
        elif index == 3:
            self.ai_tab.set_data(data)
