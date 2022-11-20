from logging import exception
from scraper import scrape
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QWidget, QPushButton, QVBoxLayout, QLabel, QPlainTextEdit, QLineEdit
from PyQt5 import QtGui
from PyQt5.QtCore import *
import sys
import os
import shutil
from datetime import datetime
import glob
from selenium import webdriver
#from analyze import analyse
from generate_report import generate, generate_overall
import time
from formatter import format_color
from openpyxl import load_workbook


class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.setGeometry(200, 200, 400, 400)
        self.setWindowTitle("Meroshare Automation Bot")
        self.setWindowIcon(QtGui.QIcon('icon.png'))
        self.initUI()

    def initUI(self):
        self.button1 = QPushButton('Update Data', self)
        self.button2 = QPushButton('Browse file', self)

        self.location = QLineEdit(self)
        self.location.setFixedSize(200, 31)
        self.location.move(50, 50)
        self.location.setStyleSheet("border:none")

        self.button2.setStyleSheet(
            'background-color:#2997FF;color:white;font:15;border:none')
        self.button2.move(250, 50)
        self.button2.clicked.connect(self.clicked)
        self.button1.setStyleSheet(
            "color:black;font:30px;")
        self.button1.setFixedSize(290, 85)
        self.button1.move(50, 150)
        self.button1.setFont(QtGui.QFont("Roboto"))
    # self.button.show()
        self.button1.clicked.connect(self.start)
        self.label = QtWidgets.QLabel(self)
        self.label.setStyleSheet("color:green;")
        self.label.setText('..... Press the button to update.....\n')
        self.label.move(100, 200)
        self.label.setFixedSize(400, 200)

    def start(self):
        self.label.setText("Downloading data...")
        user_accounts = [

            {'name': "",
             'id': 0,
             'password': ""
             },
            {'name': "",
             'id': 0,
             'password': ""},

            {'name': " ",
             'id': 0,
             'password': ""},

            {'name': "",
             'id': 0,
             'password': ""},
        ]
        try:
            current_dateTime = datetime.now()
            file_path = f"./files/Meroshare-{str(current_dateTime)[:10]}.xlsx"

            path = "./files"
            excel_files = glob.glob(path + "/**/*.xlsx", recursive=True)
            for file in excel_files:
                os.remove(file)
            shutil.copy('./template/Book1.xlsx', './files/')

            # for item in reversed(self.filename[0]):
            #     if (item == '/'):
            #         break
            #     file_name = file_name+item
            # new_filename = file_name[:: -1]

            os.rename(f"./files/Book1.xlsx", file_path)

            for item in user_accounts:
                length = scrape(item["id"], item["password"],
                                item["name"]+"_CDSC", file_path, "Portfolio_" + item["name"])
                # self.label.insertPlainText(f"\n...Downloaded {item['name']}\n")

                # analyse(item["name"]+"_CDSC", "Portfolio_" +
                # item["name"], length)
            for item in user_accounts:
                generate(self.filename[0], item["name"] +
                         "_CDSC", "Portfolio_" + item["name"])

            generate_overall()
            # Update analysis

            for item in user_accounts:
                format_color("Portfolio_" + item["name"])

            self.label.setText(
                "\n\n*Operation successfully completed*\n\n*Press update to update again*\n")

        # Update analysis
        except:
            self.label.setStyleSheet("color:red;font:20px")
            self.label.setText("***An error occured***")
            print(exception)

    def clicked(self):

        self.filename = QFileDialog.getOpenFileName()

        self.location.setText(self.filename[0])


def window():
    app = QApplication(sys.argv)
    win = MyWindow()

    win.show()
    sys.exit(app.exec_())


window()
