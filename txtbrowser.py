import sys

from PyQt5.QtWidgets import QApplication, QTextBrowser

if __name__ == '__main__':
    app = QApplication(sys.argv)
    text_area = QTextBrowser()
    text_area.setText(u'<p> Jhon Doe <a href='"'mailto:jhon@compay.com'"'>jhon@compay.com</a>  </p>')
    text_area.setOpenExternalLinks(True)
    text_area.show()
    sys.exit(app.exec_())
