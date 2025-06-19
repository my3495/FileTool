import sys
from PySide6.QtWidgets import QApplication
from src.gui.main_window import MainWindow
from PySide6.QtCore import QTranslator, QLocale, QLibraryInfo


def main():
    """主程序入口"""
    app = QApplication(sys.argv)

    # 加载中文翻译
    translator = QTranslator()
    locale = QLocale.system().name()  # 获取系统语言
    qt_trans_path = QLibraryInfo.path(QLibraryInfo.TranslationsPath)
    if translator.load(f"qt_{locale}", qt_trans_path):
        app.installTranslator(translator)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main() 