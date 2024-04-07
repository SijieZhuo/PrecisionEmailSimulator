import csv
import datetime
import os
import random
import time
import webbrowser
from pathlib import Path
import ast

import pandas as pd
from PySide2 import QtGui, QtCore, QtWidgets, QtWebChannel, QtWebEngineWidgets
from PySide2.QtCore import QRectF, QSize, Qt
from PySide2.QtGui import QIcon, QPixmap, QTextDocument
from PySide2.QtUiTools import QUiLoader
from PySide2.QtWebEngineWidgets import QWebEngineView, QWebEnginePage
from PySide2.QtWidgets import QMessageBox, QTableWidgetItem, QStyleOptionViewItem, QApplication, QStyle
from pygame import mixer
from win10toast import ToastNotifier
import paramiko
import io
import requests

now = datetime.datetime.now()
# Path("./data").mkdir(parents=True, exist_ok=True)
column_names = ['ID', 'name', 'from', 'to', 'title', 'content', 'attachment', 'star', 'time', 'readState', 'category']

log_file_name = ''

# for changing resolution
# if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
#     QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
#
# if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
#     QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)


'''
For the local version
If running the generated exe file, must add the UI_file folder in the room folder
'''

myhost = '130.216.217.42'


# myusername = 'wang'
# mypassword = '123456'


# kFile = io.StringIO(k)


class TaskWindow(QtWidgets.QWidget):

    def __init__(self, userName, config):
        super(TaskWindow, self).__init__()

        self.username = userName
        self.config = config
        print(config)
        self.saveLocation = self.config.get('saveLocation')

        self.ui = QUiLoader().load('resources/UI_files/main_page.ui')

        if self.username == '':
            self.username = 'no_user_name'
        self.emails = pd.read_csv(self.config.get('emailListLocation'))
        # self.legitEmailData.columns = column_names
        self.emails['star'] = self.emails['star'].astype('bool')
        #
        # self.phishEmailData = pd.read_csv(self.config.get('phishEmailLocation'))
        # # self.phishEmailData.columns = column_names
        # self.phishEmailData['star'] = self.phishEmailData['star'].astype('bool')
        #
        # notSuffledEmails = []
        # for session in self.config.get('sessions'):
        #     session = self.config.get('sessions').get(session)
        #
        #     if not session.get('phishEmails').get('shuffleEmails'):
        #         pemails = session.get('phishEmails').get('emailList').split(',')
        #         pemails = [int(x) for x in pemails]
        #         notSuffledEmails = notSuffledEmails + pemails
        # self.phishEmailPool = self.phishEmailData.loc[~self.phishEmailData['ID'].isin(notSuffledEmails)]
        # self.phishEmailPool = self.phishEmailPool.sample(frac=1).reset_index(drop=True)

        self.current_emaillist = []
        self.incomingEmails = []
        self.incomingInterval = 0
        self.previousEmails = []
        self.currentEmail = None
        self.hovered_url = 'none'
        self.audioNotificationTimes = []
        self.sessionTimer = QtCore.QTimer(self)
        self.incomingEmailTimer = QtCore.QTimer(self)
        self.reported_emaillist = []
        # ========== set up interface elements ===========
        mixer.init()
        self.beep = mixer.Sound("./resources/beep.wav")

        # set top buttons
        self.ui.starBtn.clicked.connect(self.star_btn_clicked)
        self.ui.deleteBtn.clicked.connect(self.delete_btn_click)
        self.ui.unreadBtn.clicked.connect(self.unread_btn_click)
        self.ui.reportBtn.clicked.connect(self.report_btn_click)
        self.ui.nextBtn.clicked.connect(self.next_btn_click)

        # set reply buttons
        self.ui.replyBtn.clicked.connect(lambda: self.respond_btn_clicked('reply'))
        self.ui.replyToAllBtn.clicked.connect(lambda: self.respond_btn_clicked('reply_to_all'))
        self.ui.forwardBtn.clicked.connect(lambda: self.respond_btn_clicked('forward'))

        self.ui.emailList.clicked.connect(self.emailTableClicked)

        # ==============================
        self.create_log_file()

        self.sessionList = list(self.config.get('sessions').keys())
        self.currentSession = self.sessionList[0]

        self.setupSession()

    # ================================================================================================

    # ======== set up session ===================
    def getCurrentSession(self):
        return self.config.get('sessions').get(self.currentSession)

    def setupSessionTimer(self, sessionConfig):
        self.sessionTimer = QtCore.QTimer(self)
        self.countDownCounter = int(sessionConfig.get('duration')) * 60
        self.running = True
        self.sessionTimer.timeout.connect(self.timerCountDown)
        self.sessionTimer.start(1000)

    def timerCountDown(self):
        if self.running:
            self.countDownCounter -= 1

            if self.countDownCounter == 0:
                self.running = False
                print("completed")
                self.log_email("finish " + self.getCurrentSession().get('name'))
                self.savePrimaryTaskDataLocal()
                print(self.currentSession)

                if self.getCurrentSession().get('endSessionPopup') != '':
                    messageNotification(self,
                                        self.getCurrentSession().get('endSessionPopup'))

                # check if it is the last session

            # add some notification for countdown
            elif (self.countDownCounter / 60) in self.audioNotificationTimes:
                self.beep.play()

            if self.countDownCounter % 60 < 10:
                timerStr = str(int(self.countDownCounter / 60)) + ':0' + str(self.countDownCounter % 60)
            else:
                timerStr = str(int(self.countDownCounter / 60)) + ':' + str(self.countDownCounter % 60)

            self.ui.timerLabel.setText(timerStr)

    def next_btn_click(self):
        print('next button clicked')
        print(self.currentSession)
        if self.incomingEmailTimer.isActive():  # turn off the incoming email timer
            self.incomingEmailTimer.stop()
        self.countDownCounter = 1
        # if self.getCurrentSession().get('primaryTaskHtml') != '':
        #     print('xxxxxxxxxxxxxxxxxxxxxx')
        #     self.savePrimaryTaskDataLocal()

    def setupIncomingEmailTimer(self, sessionConfig):
        print('setting up incoming timer')

        # set up timers for in coming emails
        self.incomingEmailTimer = QtCore.QTimer()

        self.incomingEmailTimer.timeout.connect(self.incoming_timer)
        print('cccc')
        print(sessionConfig.get('incomingInterval'))
        self.incomingEmailTimer.start(1000 * 60 * float(sessionConfig.get('incomingInterval')))

    def incoming_timer(self):
        if self.incomingEmails.shape[0] > 0:
            print('addEmail')
            self.incomingEmails = self.add_email(self.incomingEmails)
            self.set_unread_email_count()

            if self.incomingEmails.shape[0] == 0:
                print("timer stoped")
                self.incomingEmailTimer.stop()

    def setupSession(self):
        sessionConfig = self.getCurrentSession()
        self.setupSessionTimer(sessionConfig)

        self.ui.URLDisplay.setHidden(True)

        if sessionConfig.get('audioNotification') != '':
            self.audioNotificationTimes = [int(x) for x in sessionConfig.get('audioNotification').split(',')]

        if sessionConfig.get('timeCountDown'):
            self.ui.timerLabel.setHidden(False)
        else:
            self.ui.timerLabel.setHidden(True)

        if sessionConfig.get('incomingEmails'):
            self.setupIncomingEmailTimer(sessionConfig)

        for display, btn in zip(
                [sessionConfig.get('starBtn'), sessionConfig.get('reportBtn'), sessionConfig.get('deleteBtn'),
                 sessionConfig.get('unreadBtn')],
                [self.ui.starBtn, self.ui.reportBtn, self.ui.deleteBtn, self.ui.unreadBtn]):
            if not display:
                btn.hide()
            else:
                btn.show()

        # setup emails in the initial inbox
        self.setupEmails(sessionConfig)
        self.set_up_email_timestamp()

        self.set_up_emailList_table()
        self.log_email("start " + self.currentSession)
        if sessionConfig.get('primaryTaskHtml') != '':
            self.ui.primaryTaskW.show()
            self.setupPrimaryTask()
        else:
            self.ui.primaryTaskW.hide()

    def setupEmails(self, sessionConfig):
        self.current_emaillist = self.emails[
            (self.emails['ID'] >= int(sessionConfig.get('legitEmails').get('emailListRange').get('start'))) &
            (self.emails['ID'] <= int(sessionConfig.get('legitEmails').get('emailListRange').get('finish')))]
        if sessionConfig.get('legitEmails').get('shuffleEmails'):
            self.current_emaillist = self.current_emaillist.sample(frac=1).reset_index(drop=True)

        if sessionConfig.get('incomingEmails'):
            self.incomingEmails = self.emails[
                (self.emails['ID'] >= int(sessionConfig.get('legitEmails').get('incomingRange').get('start'))) &
                (self.emails['ID'] <= int(sessionConfig.get('legitEmails').get('incomingRange').get('finish')))]
            if sessionConfig.get('legitEmails').get('shuffleEmails'):
                self.incomingEmails = self.incomingEmails.sample(frac=1).reset_index(drop=True)

        if sessionConfig.get('hasPhishEmails'):
            self.addPhishingEmailsToList(sessionConfig)

        self.current_emaillist = self.current_emaillist.sort_index().reset_index(drop=True)
        print(self.incomingEmails)
        print('bbbbb')
        print(len(self.incomingEmails))
        if len(self.incomingEmails) != 0:
            self.incomingEmails = self.incomingEmails.sort_index().reset_index(drop=True)

    def addPhishingEmailsToList(self, sessionConfig):
        pEmailInboxID = [int(x) for x in sessionConfig.get('phishEmails').get('emailList').split(',')]
        if sessionConfig.get('phishEmails').get('incomingList') != '':
            pEmailIncomingID = [int(x) for x in sessionConfig.get('phishEmails').get('incomingList').split(',')]
        else:
            pEmailIncomingID = []
        pEmailInbox = self.emails[self.emails['ID'].isin(pEmailInboxID)]
        pEmailIncoming = self.emails[self.emails['ID'].isin(pEmailIncomingID)]

        if sessionConfig.get('phishEmails').get('shuffleEmails'):
            print('shuffle p emails')
            pEmailInbox = pEmailInbox.sample(frac=1).reset_index(drop=True)
            pEmailInbox = pEmailInbox.iloc[:int(sessionConfig.get('phishEmails').get('emailListNum'))]
            pEmailIncoming = pEmailIncoming.sample(frac=1).reset_index(drop=True)
            pEmailIncoming = pEmailIncoming.iloc[:int(sessionConfig.get('phishEmails').get('incomingNum'))]


        self.current_emaillist = self.insertPEmailToList(pEmailInbox, self.current_emaillist, 'emailListLocations', sessionConfig)
        if sessionConfig.get('incomingEmails'):
            self.incomingEmails = self.insertPEmailToList(pEmailIncoming, self.incomingEmails, 'incomingLocations', sessionConfig)

    def insertPEmailToList(self, plist, elist, location, sessionConfig):
        if sessionConfig.get('phishEmails').get('randomLoc'):
            for index, row in plist.iterrows():
                ranInt = random.randint(0, elist.shape[0])
                elist.loc[ranInt + 0.5] = row
                elist = elist.sort_index().reset_index(drop=True)
        else:
            if sessionConfig.get('phishEmails').get(location) != '':
                locList = [int(x) for x in sessionConfig.get('phishEmails').get(location).split(',')]
            else:
                locList = []
            for i in range(0, len(locList)):
                elist.loc[locList[i] - 1.5] = plist.iloc[0]
                plist = plist.iloc[1:]
                elist = elist.sort_index().reset_index(drop=True)

        return elist

    # =======================  email set up ===============================================

    def add_email(self, emaillist):
        print('----------------- email add notification --------------------------')
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
        if insertAtFront:
            current_time = datetime.datetime.now().strftime("%H:%M")
            self.current_emaillist.loc[self.current_emaillist.ID == email['ID'], 'time'] = current_time
            email['time'] = current_time
            self.log_incoming_email(email)
            self.set_unread_email_count()
            rowPos = 0
        else:
            rowPos = self.ui.emailList.rowCount()

        self.ui.emailList.insertRow(rowPos)
        cell1 = str(email['name']) + '<br>' + str(email['title'])
        self.set_cell(self.ui.emailList, rowPos, 0, cell1, QtGui.QFont("Calibri", 12, QtGui.QFont.Bold))
        self.set_cell(self.ui.emailList, rowPos, 1, str(email['time']), QtGui.QFont("Calibri", 10, QtGui.QFont.Bold))
        self.ui.emailList.item(rowPos, 1).setTextAlignment(Qt.AlignHCenter)
        self.ui.emailList.setRowHeight(rowPos, 65)
        self.change_row_background(rowPos, QtGui.QColor(245, 250, 255))

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
            # del timeList[-1] #

    def change_row_background(self, row, colour):
        self.ui.emailList.item(row, 0).setBackground(colour)
        self.ui.emailList.item(row, 1).setBackground(colour)

    def get_current_email(self):
        currentRow = self.ui.emailList.currentRow()
        email = self.ui.emailList.item(currentRow, 0).text()
        subjectLine = email.split('<br>')[1]

        return self.current_emaillist.loc[self.current_emaillist['title'] == subjectLine].iloc[0]

    # ========================== sections ================================================

    def get_next_section(self):
        currentSessionLoc = self.sessionList.index(self.currentSession)
        print(currentSessionLoc)
        print(len(self.sessionList))
        if len(self.sessionList) > currentSessionLoc + 1:
            nextSession = self.sessionList[currentSessionLoc + 1]
            print(nextSession)
            self.currentSession = nextSession
            self.setupSession()
        else:
            self.ui.close()

    def setupPrimaryTask(self):
        # read html file
        clear_layout(self.ui.primaryTaskL)
        self.primaryTask = QtWebEngineWidgets.QWebEngineView()

        pTaskData = PrimaryTaskData(self)
        pTaskData.valueChanged.connect(self.getTaskData)

        self.channel = QtWebChannel.QWebChannel()
        self.channel.registerObject("data", pTaskData)

        self.primaryTask.page().setWebChannel(self.channel)

        self.primaryTask.setUrl(QtCore.QUrl.fromLocalFile(self.getCurrentSession().get('primaryTaskHtml')))

        self.ui.primaryTaskL.addWidget(self.primaryTask)

    def savePrimaryTaskDataLocal(self):
        self.primaryTask.page().runJavaScript(
            """
            tableToCSV();
            
        """
        )

    @QtCore.Slot(str)
    def getTaskData(self, value):
        print('....................')
        print(value)
        fileName = self.currentSession + '_task.csv'
        path = os.path.join(self.folderPath, fileName)
        data = pd.DataFrame(columns=['category', 'c1', 'c2', 'c3', 'c4', 'c5', 'c6'])
        vlist = value.split('\n')
        for v in vlist:
            row = v.split('#$%')
            print(row)
            print(len(row))
            if len(row) == 2:
                row = row + ['', '', '', '', '']

            data.loc[len(data)] = row

        data.to_csv(path, index=False, header=False)

    # ========================= logging ==========================================

    def create_log_file(self):
        print('create log file')
        self.folderPath = os.path.join(self.saveLocation, self.username)
        Path(self.folderPath).mkdir(parents=True, exist_ok=True)
        # self.folderPath = './data/' + self.username + '/'
        self.fileName = os.path.join(self.folderPath, now.strftime("%d-%m-%Y_%H-%M-%S") + '_log.csv')

        # self.fileName = self.folderPath + now.strftime("%d-%m-%Y_%H-%M-%S") + '_log.csv'
        global log_file_name
        log_file_name = self.fileName
        print("logfile name")
        print(log_file_name)
        with open(self.fileName, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["time", "timestamp", "username", "ID", "email", "action", "detail", "studyCondition"])
        # else:

    def log_email(self, action, detail=""):
        email = self.get_current_email()

        with open(self.fileName, 'a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(
                [datetime.datetime.now(), time.time() * 1000, self.username, email['ID'], email['title'],
                 action, detail, self.getCurrentSession().get('name')])

    def log_incoming_email(self, email):
        with open(self.fileName, 'a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(
                [datetime.datetime.now(), time.time() * 1000, self.username, email['ID'], email['title'],
                 "incoming email", "", self.getCurrentSession().get('name')])

    # ============================= Event table ====================================================

    def set_up_emailList_table(self):
        self.ui.emailList.setRowCount(0)
        self.ui.emailList.setColumnCount(2)

        self.ui.emailList.setHorizontalHeaderLabels(['Email', 'Time'])
        header = self.ui.emailList.horizontalHeader()
        # self.ui.emailList.setColumnWidth(0, 70)
        self.ui.emailList.setColumnWidth(1, 50)

        self.ui.emailList.setItemDelegateForColumn(0, ListDelegate())

        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)

        for i in range(0, self.current_emaillist.shape[0]):
            self.load_email_widget(self.current_emaillist.iloc[i])

        self.ui.emailList.selectRow(0)

        self.currentEmail = self.get_current_email()
        self.display_email()

    def emailTableClicked(self):
        self.display_email()

        self.log_email("email opened")

    def set_cell(self, table, row, column, value, style=None, mergeRow=-1, mergeCol=-1):
        if mergeRow != -1:
            table.setSpan(row, column, mergeRow, mergeCol)
        newItem = QTableWidgetItem(value)
        if style is not None:
            newItem.setFont(style)
        table.setItem(row, column, newItem)

    def set_window(self, type):
        if type == "reply":
            self.respondWindow = QUiLoader().load('resources/UI_files/reply.ui')
            self.respondWindow.setWindowTitle("Reply")
        else:
            self.respondWindow = QUiLoader().load('resources/UI_files/forward.ui')
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
            self.respondWindow.content.setFocus()
            self.log_email("reply button clicked")

        elif types == 'reply_to_all':

            self.set_window("reply")

            self.respondWindow.toText.setText(currentEmail['name'])
            toAddresses = currentEmail['to'].split(', ')
            if 'me' in toAddresses:
                toAddresses.remove('me')
            self.respondWindow.ccText.setText(', '.join(str(s) for s in toAddresses))
            self.respondWindow.subjectLine.setText('Re: ' + currentEmail['title'])
            self.respondWindow.content.setFocus()
            self.log_email("reply to all button clicked")
        else:

            self.set_window("forward")
            self.respondWindow.subjectLine.setText('Forward: ' + currentEmail['title'])
            self.log_email("forward button clicked")

    def reply_send_btn_clicked(self, type):
        if type == "reply":
            if self.respondWindow.content.toPlainText() == '':
                messageNotification(self,
                                    "Please write something in the text field",
                                    False)
            else:
                #  log data
                self.log_email("reply", self.respondWindow.content.toPlainText())
                self.respondWindow.reject()
        else:
            if self.respondWindow.toBox.toPlainText() == '':
                messageNotification(self,
                                    "Please select where you want to forward the email",
                                    False)
            else:
                #  log data
                self.log_email(
                    "forward to " + self.respondWindow.toBox.toPlainText(), self.respondWindow.content.toPlainText())
                self.respondWindow.reject()

    # ============================= email top bar buttons ==========================================

    def star_btn_clicked(self):
        current = self.get_current_email()
        index = self.current_emaillist.index[self.current_emaillist['ID'] == current['ID']].tolist()[0]

        # check star state and toggle it
        if self.get_current_email()['star']:
            self.current_emaillist.at[index, 'star'] = False
            self.log_email("email unstared")

        else:
            self.current_emaillist.at[index, 'star'] = True
            self.log_email("email stared")

        self.set_email_row_font_colour(self.get_current_email())
        self.update_star(self.current_emaillist.at[index, 'star'])

    def update_star(self, value):
        if value:
            self.ui.starBtn.setIcon(QtGui.QPixmap("resources/star_activate.png"))
        else:
            self.ui.starBtn.setIcon(QtGui.QPixmap("resources/star.png"))

    def delete_btn_click(self):
        # self.send_popup('The email has been deleted', 3)
        self.log_email("email deleted")
        self.remove_current_selected_email()

        # logging

    def report_btn_click(self):
        # self.send_popup('The email has been reported', 3)
        self.log_email("email reported")
        currentEmail = self.get_current_email()
        # self.reported_emaillist = pd.concat([self.reported_emaillist, currentEmail.to_frame().T])

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
        self.change_row_background(self.ui.emailList.currentRow(), QtGui.QColor(245, 250, 255))

        # logging
        self.log_email("email marked as unread")

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
        path = self.config.get('emailResourceLocation') + '/html/' + item['content']
        print(path)
        # htmlFile = requests.get(path)
        # print(htmlFile.content)
        # print(htmlFile.content.decode("utf-16"))
        # webEngineView.setHtml(htmlFile.content.decode("utf-16"))
        webEngineView.load(QtCore.QUrl().fromLocalFile(path))
        # webEngineView.load(QtCore.QUrl('http://cs791-hishing-ticket-python-resource.s3-website-ap-southeast-2.amazonaws.com/ads.html'))
        # webEngineView.load(QtCore.QUrl(path))
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
        if (link == "") and (self.hovered_url == 'none'):
            pass
        else:
            self.hovered_url = link
            print(self.hovered_url)

            if self.hovered_url != "":
                self.log_email("url hovered", link)
                self.ui.URLDisplay.setHidden(False)
                # if link == "https://iam.auckland.ac.nz/profile/SAML2/Redirect/SSO?execution=e1s1":
                #     link = "https://docs.google.com/spreadsheets/d/1I32l2q-FAGXPPx32jT2HkrtD8yxNOU7KGrDNHb5-dxM/edit?usp=sharing"
                self.ui.URLDisplay.setText(link)
                print('hovered')
            else:
                self.log_email("url unhovered")
                self.hovered_url = 'none'
                self.ui.URLDisplay.setHidden(True)

                print('unhovered')

    def set_row_font(self, row, font, size=10):
        self.ui.emailList.item(row, 0).setFont(QtGui.QFont('Calibri', size + 2, font))
        self.ui.emailList.item(row, 1).setFont(QtGui.QFont('Calibri', size, font))

    def set_email_row_font_colour(self, row):
        if row['star']:
            self.set_row_font(self.ui.emailList.currentRow(), QtGui.QFont.Bold)
            self.change_row_background(self.ui.emailList.currentRow(), QtGui.QColor(235, 200, 200))
        else:
            self.set_row_font(self.ui.emailList.currentRow(), QtGui.QFont.Normal)
            self.change_row_background(self.ui.emailList.currentRow(), QtGui.QColor(245, 250, 255))

    def set_unread_email_count(self):
        self.ui.unreadEmailCount.setText(str((self.current_emaillist['readState'] == False).sum()))

    def setup_email_css(self, category):
        self.ui.emailSubjectLine.setStyleSheet(self.getCurrentSession().get('cssStyles').get(category).get('header'))
        self.ui.fromAddress.setStyleSheet(self.getCurrentSession().get('cssStyles').get(category).get('sender'))
        self.ui.contentW.setStyleSheet(self.getCurrentSession().get('cssStyles').get(category).get('body'))

        if self.getCurrentSession().get('cssStyles').get(category).get('headerIcon') == '':
            self.ui.subjectIcon.hide()
        else:
            self.ui.subjectIcon.show()
            self.ui.subjectIcon.setStyleSheet(self.getCurrentSession().get('cssStyles').get(category).get('headerIcon'))

        if self.getCurrentSession().get('cssStyles').get(category).get('senderIcon') == '':
            self.ui.userIcon.setStyleSheet(
                'border:None; border-image: url(resources/sender.png) 0 0 0 0 stretch stretch;')
        else:
            self.ui.userIcon.setStyleSheet(
                self.getCurrentSession().get('cssStyles').get(category).get('senderIcon'))

    def reset_css(self):
        self.ui.emailSubjectLine.setStyleSheet('')
        self.ui.subjectIcon.hide()
        self.ui.fromAddress.setStyleSheet('')
        self.ui.userIcon.setStyleSheet(
            'border:None; border-image: url(resources/sender.png) 0 0 0 0 stretch stretch;')
        self.ui.contentW.setStyleSheet('')

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
        attachmentRoot = self.config.get('emailResourceLocation') + '/Attachments'
        webbrowser.open(attachmentRoot + '/' + name)
        self.log_email("open attachment", "legit attachment: " + name)

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

        self.log_email("open attachment", "phishing attachment: " + name)


# ================================ Utils =======================================

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

        context.get_next_section()


# Function to insert row in the dataframe
def Insert_row_(row_number, df, row_value):
    # Slice the upper half of the dataframe
    df1 = df[0:row_number]
    # Store the result of lower half of the dataframe
    df2 = df[row_number:]
    # Insert the row in the upper half dataframe
    df1 = pd.concat([df1, row_value])
    # Concat the two dataframes
    df_result = pd.concat([df1, df2])
    # Reassign the index labels
    df_result.index = [*range(df_result.shape[0])]
    # Return the updated dataframe
    return df_result


class EmailContentPage(QWebEnginePage):
    def acceptNavigationRequest(self, url, _type, isMainFrame):
        if _type == QWebEnginePage.NavigationTypeLinkClicked:
            QtGui.QDesktopServices.openUrl(url);
            log = pd.read_csv(log_file_name)
            row = log.iloc[-1]
            with open(log_file_name, 'a', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(
                    [datetime.datetime.now(), time.time() * 1000, row[2], row[3], row[4],
                     'link clicked', url.toString(), row[7]])
            return False
        return True


class HtmlView(QWebEngineView):
    def __init__(self, *args, **kwargs):
        QWebEngineView.__init__(self, *args, **kwargs)
        self.setPage(EmailContentPage(self))


class PrimaryTaskData(QtCore.QObject):
    valueChanged = QtCore.Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._value = ""

    @QtCore.Property(str)
    def value(self):
        return self._value

    @value.setter
    def value(self, v):
        self._value = v
        self.valueChanged.emit(v)


class ListDelegate(QtWidgets.QStyledItemDelegate):
    def paint(self, painter, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)

        painter.save()
        doc = QTextDocument()
        doc.setHtml(opt.text)
        doc.setDefaultFont(opt.font)
        opt.text = "";
        style = opt.widget.style()
        style.drawControl(QStyle.CE_ItemViewItem, opt, painter)
        painter.translate(opt.rect.left(), opt.rect.top())
        clip = QRectF(0, 0, opt.rect.width(), opt.rect.height())
        doc.drawContents(painter, clip)
        painter.restore()

    def sizeHint(self, option, index):
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        doc = QTextDocument()
        doc.setHtml(opt.text);
        doc.setTextWidth(opt.rect.width())
        return QSize(doc.idealWidth(), doc.size().height())
