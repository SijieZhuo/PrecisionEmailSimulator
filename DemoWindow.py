import os
import random

from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QApplication, QMessageBox, QTableWidgetItem
from PySide2 import QtGui, QtCore, QtWidgets
from PySide2.QtWebEngineWidgets import QWebEngineView, QWebEnginePage
import pandas as pd
import datetime
from bs4 import BeautifulSoup
import codecs
import time
import calendar
import csv
from pathlib import Path

from win10toast import ToastNotifier
import threading
import webbrowser

from win32gui import GetWindowText, GetForegroundWindow


# pretended user: Jacob Smith jsmi485@aucklanduni.ac.nz
# currently the "event time" is set to be at Wed 10/08/2022
# if the date need to be changed, need to change the pdfs as well (catering, transport)
event_time = datetime.datetime.strptime("10_08_2022", "%d_%m_%Y")
now = datetime.datetime.now()
Path("./data").mkdir(parents=True, exist_ok=True)
column_names = ['ID', 'name', 'from', 'to', 'title', 'content', 'attachment', 'star', 'time', 'readState']

# condition for determine whether high workload session should start first or not, 0 for low workload as the first session
studyCondition = random.randint(0, 1)


# studyCondition = 0
# studyCondition = 1


# for changing resolution
# if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
#     QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
#
# if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
#     QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)


class DemoWindow(QtWidgets.QWidget):

    def __new__(cls):
        if not hasattr(cls, 'instance'):
            cls.instance = super(DemoWindow, cls).__new__(cls)
        return cls.instance

    def __init__(self):
        super(DemoWindow, self).__init__()
        self.ui = QUiLoader().load('UI_files/main_page.ui')

        self.folderPath = './data/'
        self.username = ''

        # load phishing emails

        self.current_emaillist = pd.read_csv('./P_emails.csv')
        self.current_emaillist.columns = column_names
        self.current_emaillist['star'] = self.current_emaillist['star'].astype('bool')

        self.previousEmails = []
        # create empty dataframe for unread emails
        self.unread_emaillist = pd.DataFrame(columns=column_names)
        self.reported_emaillist = pd.DataFrame(columns=column_names)

        self.set_up_email_timestamp()
        modify_html()

        # adding emails
        self.set_up_emailList_table()

        # set top buttons
        self.ui.starBtn.clicked.connect(self.star_btn_clicked)
        # self.ui.deleteBtn.clicked.connect(self.delete_btn_click)
        # self.ui.unreadBtn.clicked.connect(self.reflection_session)
        self.ui.reportBtn.clicked.connect(self.report_btn_click)
        self.ui.nextBtn.clicked.connect(self.next_btn_click)

        # set reply buttons
        self.ui.replyBtn.clicked.connect(lambda: self.respond_btn_clicked('reply'))
        self.ui.replyToAllBtn.clicked.connect(lambda: self.respond_btn_clicked('reply_to_all'))
        self.ui.forwardBtn.clicked.connect(lambda: self.respond_btn_clicked('forward'))

        # set event table
        self.set_table1()
        self.set_table2()
        self.set_table3()
        # self.ui.tableSaveBtn.clicked.connect(self.save_table_data)

        self.hovered_url = ""



        # self.raise_()
        # self.activateWindow()
        # print(self.isActiveWindow())
        #
        # print(GetWindowText(GetForegroundWindow()))

    def next_btn_click(self):


        if self.currentSession == "extra":
            # end of the session, unread email session will be loaded automatically, press next button will go to reflection session
            self.reflection_session()
        else:
            self.countDownCounter = 1

    # =======================  email set up ===============================================

    def setup_emails(self, dataList, pList, initialListNum):
        currentList = dataList.iloc[:initialListNum]
        dataList = dataList.iloc[initialListNum:, :]
        # put 1 phishing email in the currentlist
        currentList, pList = self.insert_phishing_email_to_list(currentList)
        # put 1 phishing email in the datalist
        dataList, pList = self.insert_phishing_email_to_list(dataList)
        return currentList, dataList, pList

    def insert_phishing_email_to_list(self, emaillist):
        ranpos = random.randrange(1, emaillist.shape[0])
        emaillist = Insert_row_(ranpos, emaillist, self.emailList_P_Data.iloc[0])
        self.emailList_P_Data = self.emailList_P_Data.iloc[1:, :]
        return emaillist, self.emailList_P_Data

    def add_email(self, emaillist):
        print('----------------- email add notification --------------------------')
        # print(emaillist)
        if emaillist.shape[0] > 0:
            item = emaillist.iloc[0]
            self.current_emaillist = self.current_emaillist.append(item, ignore_index=True)
            self.load_email_widget(item, True)
            emaillist = emaillist.iloc[1:, :]
            ToastNotifier().show_toast('An email has arrived', item['title'], icon_path='resources/icon.ico',
                                       duration=5,
                                       threaded=True)
        return emaillist

    # load emails, input is the row of email
    def load_email_widget(self, email, insertAtFront=False):
        font = QtGui.QFont("Calibri", 10, QtGui.QFont.Bold)
        if insertAtFront:
            current_time = datetime.datetime.now().strftime("%H:%M")
            self.current_emaillist.loc[self.current_emaillist.ID == email['ID'], 'time'] = current_time
            email['time'] = current_time
            self.set_unread_email_count()
            rowPos = 0
        else:
            rowPos = self.ui.emailList.rowCount()

        # item = self.current_emaillist.loc[self.current_emaillist['ID'] == email['ID']].iloc[0]
        self.ui.emailList.insertRow(rowPos)
        self.set_cell(self.ui.emailList, rowPos, 0, email['name'], font)
        self.set_cell(self.ui.emailList, rowPos, 1, email['title'], font)
        self.set_cell(self.ui.emailList, rowPos, 2, email['time'], font)
        self.ui.emailList.setRowHeight(rowPos, 45)
        self.change_row_background(rowPos, QtGui.QColor(200, 200, 235))

    def set_up_email_timestamp(self):
        lst = [random.randint(1, 10) for x in range(self.current_emaillist.shape[0] - 4)]

        # list sorted down from 10 to 1
        lst.sort(reverse=True)
        currentDay = datetime.date.today()
        timeList = []

        # set up the list of times for the emails (from oldest to newest)
        for i in lst:
            timeList.append((currentDay - datetime.timedelta(days=i)).strftime("%d %b"))

        timeList.append((datetime.datetime.now() - datetime.timedelta(hours=4, minutes=29)).strftime("%H:%M"))
        timeList.append((datetime.datetime.now() - datetime.timedelta(hours=3, minutes=15)).strftime("%H:%M"))
        timeList.append((datetime.datetime.now() - datetime.timedelta(hours=2, minutes=22)).strftime("%H:%M"))
        timeList.append((datetime.datetime.now() - datetime.timedelta(hours=0, minutes=17)).strftime("%H:%M"))

        for index, row in self.current_emaillist.iterrows():
            if timeList != []:
                self.current_emaillist.at[index, 'time'] = timeList.pop()
            else:
                self.current_emaillist.at[index, 'time'] = "-1"
            # del timeList[-1]

    def change_row_background(self, row, colour):
        self.ui.emailList.item(row, 0).setBackground(colour)
        self.ui.emailList.item(row, 1).setBackground(colour)
        self.ui.emailList.item(row, 2).setBackground(colour)

    def get_current_email(self):
        currentRow = self.ui.emailList.currentRow()
        subjectLine = self.ui.emailList.item(currentRow, 1).text()

        return self.current_emaillist.loc[self.current_emaillist['title'] == subjectLine].iloc[0]


    # ===================== record user name ============================

    def load_user_name(self):
        if self.username != '':
            print(self.username)
            # self.username = self.usernameWindow.nameText.text()

    def setup(self, username):
        self.username = username
        print(self.username)

    # ========================= logging ==========================================


    # ============================= Event table ====================================================

    def set_up_emailList_table(self):
        self.ui.emailList.setRowCount(0)
        self.ui.emailList.setColumnCount(3)

        self.ui.emailList.setHorizontalHeaderLabels(['Sender', 'Title', 'Time'])
        header = self.ui.emailList.horizontalHeader()
        self.ui.emailList.setColumnWidth(0, 70)
        self.ui.emailList.setColumnWidth(2, 50)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        font = QtGui.QFont("Calibri", 12, QtGui.QFont.Bold)
        self.ui.emailList.horizontalHeaderItem(0).setFont(font)
        self.ui.emailList.horizontalHeaderItem(1).setFont(font)
        self.ui.emailList.horizontalHeaderItem(2).setFont(font)
        # self.current_email_count = initial_email_number + 1

        for i in range(0, self.current_emaillist.shape[0]):
            # print(emailListData.iloc[i])
            self.load_email_widget(self.current_emaillist.iloc[i])

        self.ui.emailList.clicked.connect(self.emailTableClicked)
        self.ui.emailList.selectRow(0)

        self.currentEmail = self.get_current_email()
        self.display_email()

    def emailTableClicked(self):

        self.display_email()

        # print(self.previousEmails)


    def set_cell(self, table, row, column, value, style=None, mergeRow=-1, mergeCol=-1):

        if mergeRow != -1:
            table.setSpan(row, column, mergeRow, mergeCol)
        newItem = QTableWidgetItem(value)
        if style is not None:
            newItem.setFont(style)
        table.setItem(row, column, newItem)

    def set_table1(self):

        self.ui.table1.setRowCount(11)
        self.ui.table1.setColumnCount(4)

        self.ui.table1.horizontalHeader().setVisible(False)
        self.ui.table1.verticalHeader().setVisible(False)

        header = self.ui.table1.horizontalHeader()
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        self.ui.table1.setColumnWidth(0, 150)
        self.ui.table1.setColumnWidth(1, 120)
        self.ui.table1.setColumnWidth(2, 250)
        self.ui.table1.setColumnWidth(3, 100)

        self.ui.table1.resizeRowsToContents()

        # Row 1
        font = QtGui.QFont("Calibri", 11, QtGui.QFont.Bold)
        self.set_cell(self.ui.table1, 0, 0, "Speaker Name", font)
        self.set_cell(self.ui.table1, 0, 1, "Organization", font)
        self.set_cell(self.ui.table1, 0, 2, "Biography info", font)
        self.set_cell(self.ui.table1, 0, 3, "Presentation title", font)

    def set_table2(self):

        self.ui.table2.setRowCount(7)
        self.ui.table2.setColumnCount(10)

        self.ui.table2.horizontalHeader().setVisible(False)
        self.ui.table2.verticalHeader().setVisible(False)

        self.ui.table2.setColumnWidth(0, 180)
        self.ui.table2.setColumnWidth(1, 150)
        self.ui.table2.setColumnWidth(2, 150)
        self.ui.table2.setColumnWidth(3, 100)
        self.ui.table2.setColumnWidth(4, 70)
        self.ui.table2.setColumnWidth(5, 120)
        self.ui.table2.setColumnWidth(6, 120)
        self.ui.table2.setColumnWidth(7, 150)
        self.ui.table2.setColumnWidth(8, 150)
        self.ui.table2.setColumnWidth(9, 100)

        self.ui.table2.resizeRowsToContents()

        # row 1
        font = QtGui.QFont("Calibri", 11, QtGui.QFont.Bold)

        self.set_cell(self.ui.table2, 0, 0, "Speakers", font, 2, 1)
        self.set_cell(self.ui.table2, 0, 1, "Flight information", font, 1, 4)
        self.set_cell(self.ui.table2, 0, 5, "Hotel information", font, 1, 5)
        # self.set_cell(self.ui.table2, 0, 6, "Presentation", font, 1, 4)

        # row 2
        font = QtGui.QFont("Calibri", 10, QtGui.QFont.Bold)

        self.set_cell(self.ui.table2, 1, 1, "Flight date/time - To AKL", font)
        self.set_cell(self.ui.table2, 1, 2, "Flight date/time - Back", font)
        self.set_cell(self.ui.table2, 1, 3, "No. of Baggage", font)
        self.set_cell(self.ui.table2, 1, 4, "Cost", font)
        self.set_cell(self.ui.table2, 1, 5, "From date", font)
        self.set_cell(self.ui.table2, 1, 6, "To date", font)
        self.set_cell(self.ui.table2, 1, 7, "Hotel name", font)
        self.set_cell(self.ui.table2, 1, 8, "Room type", font)
        self.set_cell(self.ui.table2, 1, 9, "Cost", font)

    def set_table3(self):
        self.ui.table3.setRowCount(9)
        self.ui.table3.setColumnCount(2)

        self.ui.table3.horizontalHeader().setVisible(False)
        self.ui.table3.verticalHeader().setVisible(False)

        header = self.ui.table3.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.ui.table3.setColumnWidth(1, 250)

        self.ui.table3.resizeRowsToContents()

        # Row 1
        font = QtGui.QFont("Calibri", 11, QtGui.QFont.Bold)
        self.set_cell(self.ui.table3, 0, 0, "Type", font)
        self.set_cell(self.ui.table3, 0, 1, "info/budget", font)

        font = QtGui.QFont("Calibri", 10, QtGui.QFont.Bold)
        self.set_cell(self.ui.table3, 1, 0, "Event Room location", font)
        self.set_cell(self.ui.table3, 2, 0, "Room booked date and time", font)
        self.set_cell(self.ui.table3, 3, 0, "Event catering fee", font)
        self.set_cell(self.ui.table3, 4, 0, "Catering location", font)
        self.set_cell(self.ui.table3, 5, 0, "hotel fee", font)
        self.set_cell(self.ui.table3, 6, 0, "transportation fee", font)
        self.set_cell(self.ui.table3, 7, 0, "other fees (Chairs)", font)
        self.set_cell(self.ui.table3, 8, 0, "Fee total", font)

    def save_table_data(self, name):
        tables = [self.ui.table1, self.ui.table2, self.ui.table3]
        with open(self.folderPath + name, 'w', newline='') as stream:
            writer = csv.writer(stream)
            for table in tables:
                for row in range(table.rowCount()):
                    rowdata = []
                    for column in range(table.columnCount()):
                        item = table.item(row, column)
                        if item is not None:
                            rowdata.append(
                                item.text())
                        else:
                            rowdata.append('')
                    writer.writerow(rowdata)

    # ============================== respond window =============================================

    def set_window(self, type):
        if type == "reply":
            self.respondWindow = QUiLoader().load('reply.ui')
            self.respondWindow.setWindowTitle("Reply")
        else:
            self.respondWindow = QUiLoader().load('forward.ui')
            self.respondWindow.setWindowTitle("Forward")

        self.respondWindow.setWindowFlags(QtCore.Qt.WindowCloseButtonHint)
        self.respondWindow.deleteBtn.clicked.connect(self.respondWindow.reject)
        self.respondWindow.sendBtn.clicked.connect(lambda: self.reply_send_btn_clicked(type))

        self.respondWindow.show()

    def respond_btn_clicked(self, types):
        print('reply')
        currentEmail = self.get_current_email()

        # set up the sender, subject line etc.
        if types == 'reply':

            self.set_window("reply")

            self.respondWindow.toText.setText(currentEmail['name'])
            self.respondWindow.ccText.setHidden(True)
            self.respondWindow.ccLine.setHidden(True)
            self.respondWindow.ccLabel.setHidden(True)
            self.respondWindow.subjectLine.setText('Re: ' + currentEmail['title'])


        elif types == 'reply_to_all':

            self.set_window("reply")

            self.respondWindow.toText.setText(currentEmail['name'])
            toAddresses = currentEmail['to'].split(', ')
            if 'me' in toAddresses:
                toAddresses.remove('me')
            self.respondWindow.ccText.setText(', '.join(str(s) for s in toAddresses))
            self.respondWindow.subjectLine.setText('Re: ' + currentEmail['title'])

        else:

            self.set_window("forward")
            self.respondWindow.subjectLine.setText('Forward: ' + currentEmail['title'])

    def reply_send_btn_clicked(self, type):
        if type == "reply":
            if self.respondWindow.content.toPlainText() == '':
                self.send_popup('Please write something in the text field', 3)
            else:
                #  log data
                self.respondWindow.reject()
        else:
            if self.respondWindow.toBox.currentText() == '':
                self.send_popup('Please select where you want to forward the email', 3)
            else:
                #  log data

                self.respondWindow.reject()

    # ============================= email top bar buttons ==========================================

    def star_btn_clicked(self):

        # print('star_btn_clicked ====================================================')
        #
        # self.section_break()

        current = self.get_current_email()
        index = self.current_emaillist.index[self.current_emaillist['ID'] == current['ID']].tolist()[0]

        # check star state and toggle it
        if self.get_current_email()['star']:
            self.current_emaillist.at[index, 'star'] = False

        else:
            self.current_emaillist.at[index, 'star'] = True

        self.set_email_row_font_colour(self.get_current_email())
        self.update_star(self.current_emaillist.at[index, 'star'])

    def update_star(self, value):
        if value:
            self.ui.starBtn.setIcon(QtGui.QPixmap("resources/star_activate.png"))
        else:
            self.ui.starBtn.setIcon(QtGui.QPixmap("resources/star.png"))

    def delete_btn_click(self):
        # self.send_popup('The email has been deleted', 3)
        self.remove_current_selected_email()

        # logging

    def report_btn_click(self):

        # self.send_popup('The email has been reported', 3)
        currentEmail = self.get_current_email()
        self.reported_emaillist = pd.concat([self.reported_emaillist, currentEmail.to_frame().T])

        self.remove_current_selected_email()
        messageNotification(self, "You have reported the selected email", False)
        # logging

    def remove_current_selected_email(self):
        # print(self.current_emaillist)
        currentEmail = self.get_current_email()
        self.current_emaillist.drop(
            self.current_emaillist.index[self.current_emaillist['ID'] == self.get_current_email()['ID']], inplace=True)
        print('--------------------------------')
        print(self.previousEmails)

        if self.previousEmails[-1] == currentEmail.title:
            self.previousEmails[:] = (x for x in self.previousEmails if x != self.previousEmails[-1])
        print()
        if len(self.previousEmails) != 0:
            previousItem = self.ui.emailList.findItems(self.previousEmails[-1], QtCore.Qt.MatchContains)
            print(previousItem)
            if len(previousItem) == 1:  # if the previous email exist in the email table
                print('previous exist')
                self.ui.emailList.removeRow(self.ui.emailList.currentRow())
                previousRow = self.ui.emailList.findItems(self.previousEmails[-1], QtCore.Qt.MatchContains)[0].row()
                self.ui.emailList.selectRow(previousRow)

                self.previousEmails[:] = (x for x in self.previousEmails if x != self.previousEmails[-1])
            else:
                print('previous not exist')
                self.ui.emailList.removeRow(self.ui.emailList.currentRow())
                self.ui.emailList.selectRow(self.ui.emailList.currentRow())
        else:
            print('no previous')
            self.ui.emailList.removeRow(self.ui.emailList.currentRow())
            self.ui.emailList.selectRow(self.ui.emailList.currentRow())

        print("vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv")

        self.display_email()

    def unread_btn_click(self):
        # self.send_popup('The email is marked unread', 3)
        item = self.get_current_email()
        self.current_emaillist.loc[self.current_emaillist['ID'] == item['ID'], 'readState'] = False
        self.set_unread_email_count()

        self.set_row_font(self.ui.emailList.currentRow(), QtGui.QFont.Bold)
        self.change_row_background(self.ui.emailList.currentRow(), QtGui.QColor(200, 200, 235))

        # logging

    # =============================== display the email =============================================

    def display_email(self):
        self.previousEmails = self.previousEmails + [self.currentEmail.title]

        item = self.get_current_email()
        self.currentEmail = item

        self.current_emaillist.loc[self.current_emaillist['ID'] == item['ID'], 'readState'] = True

        self.set_unread_email_count()

        self.set_email_row_font_colour(item)

        # set subject line
        self.ui.emailSubjectLine.setText(item['title'])
        # set sender email address
        self.ui.fromAddress.setText(item['name'] + '  <' + item['from'] + '>')
        # set star state
        self.update_star(item['star'])
        # set to address
        self.ui.toAddress.setText('to ' + item['to'])

        # set the content
        clear_layout(self.ui.contentL)
        webEngineView = HtmlView(self)
        path = os.path.join(os.path.abspath(os.getcwd()), 'resources', 'html', item['content'])
        webEngineView.load(
            QtCore.QUrl().fromLocalFile(path))
        webEngineView.resize(self.ui.contentW.width(), self.ui.contentW.height())
        self.ui.contentL.addWidget(webEngineView)
        webEngineView.page().linkHovered.connect(self.link_hovered)
        # webEngineView.setZoomFactor(1.5)
        webEngineView.show()

        # set the attachment
        clear_layout(self.ui.attachmentLayout)

        if item['attachment'] != 'None':
            attachmentString = item['attachment']
            attachments = attachmentString.split(',')
            for a in attachments:
                if a[:2] == 'P_':
                    btn = self.create_phish_attachment_btn(a[2:])
                else:
                    btn = self.create_attachment_btn(a)
                self.ui.attachmentLayout.addWidget(btn)

            spaceItem = QtWidgets.QSpacerItem(150, 10, QtWidgets.QSizePolicy.Expanding)
            self.ui.attachmentLayout.addSpacerItem(spaceItem)

        if item['to'] == 'me':
            self.ui.replyToAllBtn.setHidden(True)
        else:
            self.ui.replyToAllBtn.setHidden(False)

    def link_hovered(self, link):
        self.hovered_url = link

        thread = threading.Thread(target=self.stay_hovered, args=(link,))
        thread.start()

    def stay_hovered(self, link):
        time.sleep(1)
        if self.hovered_url == link and link != "":
            print("hovered: " + link)
        else:
            return
        counter = 1
        while self.hovered_url == link and link != "":
            counter = counter + 0.1
            time.sleep(0.1)
        print("unhovered")

    def set_row_font(self, row, font, size=10):
        self.ui.emailList.item(row, 0).setFont(QtGui.QFont('Calibri', size, font))
        self.ui.emailList.item(row, 1).setFont(QtGui.QFont('Calibri', size, font))
        self.ui.emailList.item(row, 2).setFont(QtGui.QFont('Calibri', size, font))

    def set_email_row_font_colour(self, row):
        if row['star']:
            self.set_row_font(self.ui.emailList.currentRow(), QtGui.QFont.Bold, 10)
            self.change_row_background(self.ui.emailList.currentRow(), QtGui.QColor(235, 200, 200))
        else:
            self.set_row_font(self.ui.emailList.currentRow(), QtGui.QFont.Normal)
            self.change_row_background(self.ui.emailList.currentRow(), QtGui.QColor(255, 255, 255))

    # ================================== attachment ======================================

    # ===== legit attachment =====
    def create_attachment_btn(self, name):
        btn = QtWidgets.QPushButton(name)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(os.path.abspath(os.getcwd()) + r'\resources\attachment.png'))
        btn.setIcon(icon)
        btn.setStyleSheet(
            "border: 1px solid rgb(150, 150, 150); border-radius:2px; background:#56d5f9; margin: 10px; font-size: 18px; padding: 5px;")
        btn.clicked.connect(lambda: self.open_attachment(name))
        return btn

    def open_attachment(self, name):
        attachmentRoot = os.path.abspath(os.getcwd()) + r'\resources\Attachments'
        webbrowser.open(attachmentRoot + '/' + name)

    # ===== phishing attachment =====
    def create_phish_attachment_btn(self, name):
        btn = QtWidgets.QPushButton(name)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(os.path.abspath(os.getcwd()) + r'\resources\attachment.png'))
        btn.setIcon(icon)
        btn.setStyleSheet(
            "border: 1px solid rgb(150, 150, 150); border-radius:2px; background:#56d5f9; margin: 10px; font-size: 18px; padding: 5px;")
        btn.clicked.connect(lambda: self.phish_attachment_clicked(name))
        return btn

    def phish_attachment_clicked(self, name):

        title = name.split('.', 1)[0]
        win1 = QtWidgets.QWidget()
        win1.adjustSize()
        # screen_resolution = app.desktop().screenGeometry()
        win1.setGeometry(100, 100, 800 // 2, 600 // 2, )
        win1.setWindowTitle(title)
        time.sleep(1)
        win1.show()
        time.sleep(1.5)
        file_not_opened_warning(name)


    # ============================ pop up =====================================================
    def send_popup(self, text: str, time: int):
        # load and show the pop up for delete message
        ui_popup = QUiLoader().load('popup.ui')
        # print(ui_popup.parent())

        ui_popup.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        ui_popup.text.setText(text)
        ui_popup.closeBtn.clicked.connect(lambda: close_popup(ui_popup))
        ui_popup.show()
        # if not manually closed, it would close in 3 sec

        time = threading.Timer(time, ui_popup.reject())
        time.start()

    def set_unread_email_count(self):
        # print((self.current_emaillist['readState'] == False).sum())
        self.ui.unreadEmailCount.setText(str((self.current_emaillist['readState'] == False).sum()))


def close_popup(popup):
    print('pop up closed')
    popup.reject()


# ================================ Utils =======================================

def check_charset(file_path):
    import chardet
    with open(file_path, "rb") as f:
        data = f.read(4)
        charset = chardet.detect(data)['encoding']
    return charset


def modify_html():
    # edit the time for the event room booking
    timeObj = event_time
    replaceTime = calendar.day_name[timeObj.weekday()] + ', ' + timeObj.strftime('%d/%m/%Y')
    print(replaceTime)
    replace_html_text('resources/html/room.html', "A",
                      "<p class=MsoNormal id=\"A\">Date(s): " + replaceTime + "<br/></p>")

    # edit the time for invites to the speakers

    replaceTime = event_time.strftime('%d %B %Y')
    print(replaceTime)
    replace_html_text('resources/html/speaker1.html', "A",
                      "<p class=MsoNormal id=\"A\" style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'> The conference will be held at the University of Auckland at " + replaceTime + ". The deadline for papers is January 25, 2022 (HST).<br></p>")
    replace_html_text('resources/html/speaker2.html', "A",
                      "<p class=MsoNormal id=\"A\" style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'> The conference will be held at the University of Auckland at " + replaceTime + ". The deadline for papers is January 25, 2022 (HST).<br></p>")
    replace_html_text('resources/html/speaker3.html', "A",
                      "<p class=MsoNormal id=\"A\" style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'> The conference will be held at the University of Auckland at " + replaceTime + ". The deadline for papers is January 25, 2022 (HST).<br></p>")
    replace_html_text('resources/html/speaker4_no.html', "A",
                      "<p class=MsoNormal id=\"A\" style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'> The conference will be held at the University of Auckland at " + replaceTime + ". The deadline for papers is January 25, 2022 (HST).<br></p>")
    replace_html_text('resources/html/speaker5_no.html', "A",
                      "<p class=MsoNormal id=\"A\" style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'> The conference will be held at the University of Auckland at " + replaceTime + ". The deadline for papers is January 25, 2022 (HST).<br></p>")
    replace_html_text('resources/html/speaker6.html', "A",
                      "<p class=MsoNormal id=\"A\" style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'> The conference will be held at the University of Auckland at " + replaceTime + ". The deadline for papers is January 25, 2022 (HST).<br></p>")

    # # edit the time for the catering
    timeObj = datetime.datetime.today() + datetime.timedelta(days=-4)
    replaceTime = timeObj.strftime("%A, %B %d, %Y")
    replace_html_text('resources/html/catering.html', "A",
                      '<p class=MsoNormal id=\"A\"> <b>Sent:</b> ' + replaceTime + ' 1:50 PM<br></p>')
    timeObj = datetime.datetime.today() + datetime.timedelta(days=-5)
    replaceTime = timeObj.strftime("%A, %B %d, %Y")
    replace_html_text('resources/html/catering.html', "B",
                      '<p class=MsoNormal id=\"B\"> <b>Sent:</b> ' + replaceTime + ' 9:07 a.m.<br> </p>')
    timeObj = datetime.datetime.today() + datetime.timedelta(days=-12)
    replaceTime = timeObj.strftime("%A, %B %d, %Y")
    replace_html_text('resources/html/catering.html', "C",
                      '<p class=MsoNormal id=\"C\"> <b>Sent:</b> ' + replaceTime + ' 9:53 a.m.<br> </p>')

    # accommendation for two speakers
    timeObj = event_time
    replaceTime1 = timeObj.strftime('%A, %d %B')
    replaceTime6 = timeObj.strftime('%d %B')
    timeObj = timeObj - datetime.timedelta(days=1)
    replaceTime2 = timeObj.strftime('%A, %d %B')
    replaceTime4 = timeObj.strftime('%d %B')
    replaceTime5 = timeObj.strftime('%d %B')
    timeObj = timeObj + datetime.timedelta(days=2)
    replaceTime3 = timeObj.strftime('%A, %d %B')
    replace_html_text('resources/html/travel1.html', "A",
                      "<div id=\"A\"> <p class=MsoNormal>I would like to arrange your travel and accommodation for the CS Collaboration Event on " + replaceTime1 + " 2022.</p> </br> <p class=MsoNormal>Please let me know what flights would be convenient:</p> </br> <p class=MsoNormal><b>" + replaceTime2 + " 2022</b></p> <p class=MsoNormal>Dunedin - Auckland ETD 11.00am | 12.30pm | 2.00pm | 3.00pm |4.00pm (please advise if another time preferred)</p> </br> <p class=MsoNormal>Accommodation at the Cordis Hotel for " + replaceTime2 + " and " + replaceTime1 + " 2022 </p> </br> <p class=MsoNormal><b>" + replaceTime3 + " 2022</b></p> </div>")
    replace_html_text('resources/html/travel1.html', "B",
                      "<div id=\"B\">  <p class=\"MsoNormal\">    We would like to have the 2pm flight from Dunedin - Auckland on " + replaceTime4 + " and return flight at 5pm (or around this time) Auckland – Dunedin.   </p>   <p class=\"MsoNormal\">    For the hotel nights, would be possible to offer also to Davis both the night of " + replaceTime5 + " and " + replaceTime6 + " ?    </p>   </div>")

    # accommendation for the thrid speaker
    replace_html_text('resources/html/travel2.html', "A",
                      "<div id=\"A\"> <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'>I would like to arrange your travel and accommodation for the CS Collaboration Event on " + replaceTime1 + " 2022.</p> <br/> <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'>Please let me know what flight time would be convenient:</p> <br/> <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'><b>" + replaceTime2 + " 2022</b></p><p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'>Christchurch - Auckland  ETD 11.00am | 12.30pm | 2.00pm | 3.00pm |4.00pm (please advise if another time preferred)</p> <br/> <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'>Accommodation at the Cordis Hotel for " + replaceTime2 + " and " + replaceTime1 + " 2021</p><br/> <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'><b> </b></p> <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto'><b>" + replaceTime3 + " 2022</b></p> </div>")

    # phish account email
    today = datetime.date.today()
    replaceTime = today.strftime("%m/%d/%Y")
    replace_html_text('resources/html/P_account.html', "A",
                      "<span  style='font-size:10.0pt;font-family:\"Segoe UI\",sans-serif;color:#4A4A4A;  letter-spacing:-.1pt;border:none windowtext 1.0pt;mso-border-alt:none windowtext 0cm;  padding:0cm' id = \"A\"> at " + replaceTime + " 07:04:39 am, with subject line: \"IMPORTANT: pending fees need to be processed!\".</span>")



def replace_html_text(filePath, id, text):
    with codecs.open(filePath, encoding=check_charset(filePath), errors='replace') as html_file:
        txt = html_file.read()
        soup = BeautifulSoup(txt, features='html.parser')

        element = soup.find(id=id)
        replaceText = BeautifulSoup(text, features='html.parser')
        element.replace_with(replaceText)
        # Store prettified version of modified html
        new_text = soup.prettify()

    # Write new contents
    with open(filePath, mode='w', encoding="utf-8") as new_html_file:
        new_html_file.write(new_text)
    remove_empty_line(filePath)


def remove_empty_line(file):
    with open(file, 'r+', encoding="utf-8") as f:
        lines = [i for i in f.readlines() if i and i != '\n']
        f.seek(0)
        f.writelines(lines)
        f.truncate()


def clear_layout(layout):
    if layout is not None:
        while layout.count():
            child = layout.takeAt(0)
            if child.widget() is not None:
                child.widget().deleteLater()
            elif child.layout() is not None:
                clear_layout(child.layout())


def file_not_opened_warning(filename):
    print('warning')
    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Warning)
    msgBox.setText(
        "Could not open file: " + filename + ". Something unexpected happened during the execution. \nError code: 506")
    msgBox.setWindowTitle("Windows")
    msgBox.setStandardButtons(QMessageBox.Ok)

    returnValue = msgBox.exec()
    if returnValue == QMessageBox.Ok:
        print('OK clicked')


def messageNotification(context, text, newSection=True):
    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Information)
    msgBox.setText(text)
    msgBox.setWindowTitle("Notification")
    msgBox.setStandardButtons(QMessageBox.Ok)

    returnValue = msgBox.exec()
    if (returnValue == QMessageBox.Ok) and newSection:
        print('OK clicked')

        context.starting_Section2()


# Function to insert row in the dataframe
def Insert_row_(row_number, df, row_value):
    # Slice the upper half of the dataframe
    df1 = df[0:row_number]
    # Store the result of lower half of the dataframe
    df2 = df[row_number:]
    # Insert the row in the upper half dataframe
    df1.loc[row_number] = row_value
    # Concat the two dataframes
    df_result = pd.concat([df1, df2])
    # Reassign the index labels
    df_result.index = [*range(df_result.shape[0])]
    # Return the updated dataframe
    return df_result


class WebEnginePage(QWebEnginePage):
    def acceptNavigationRequest(self, url, _type, isMainFrame):
        if _type == QWebEnginePage.NavigationTypeLinkClicked:
            QtGui.QDesktopServices.openUrl(url);
            return False
        return True


class HtmlView(QWebEngineView):
    def __init__(self, *args, **kwargs):
        QWebEngineView.__init__(self, *args, **kwargs)
        self.setPage(WebEnginePage(self))
if __name__ == '__main__':
    app = QApplication([])
    mainWindow = DemoWindow()
    mainWindow.ui.show()
    app.exec_()
