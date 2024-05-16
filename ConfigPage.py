import pandas as pd
import yaml
from PySide6 import QtWidgets
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox

from EmailResearchLab import EmailResearchLab


class ConfigPage(QtWidgets.QWidget):

    def __init__(self):
        super(ConfigPage, self).__init__()
        self.ui = QUiLoader().load('resources/UI_files/config_page2.ui')
        self.ui.resize(600, 600)
        # variables
        self.study = dict({
            'emailListLocation': '',
            'emailResourceLocation': '',
            'saveLocation': '',
            'sessions': {},

        })
        self.currentSession = 'session1'
        self.ui.sessionSelectDB.addItem(self.currentSession)

        # =================== file loading section =================================
        self.ui.BrowseBtn_E.clicked.connect(lambda: self.browseFile(self.ui.emailPath, self.study, 'emailListLocation'))
        self.ui.BrowseBtn_R.clicked.connect(
            lambda: self.browseFolder(self.ui.resourcePath, self.study, 'emailResourceLocation'))
        self.ui.BrowseBtn_S.clicked.connect(lambda: self.browseFolder(self.ui.savePath, self.study, 'saveLocation'))

        # test for text change
        self.ui.emailPath.editingFinished.connect(
            lambda: self.updateTextField(self.ui.emailPath, self.study, 'emailListLocation'))
        self.ui.resourcePath.editingFinished.connect(
            lambda: self.updateTextField(self.ui.resourcePath, self.study, 'emailResourceLocation'))
        self.ui.savePath.editingFinished.connect(
            lambda: self.updateTextField(self.ui.savePath, self.study, 'saveLocation'))


        # twice because one for update after click on button
        # self.ui.htmlPath.editingFinished.connect(
        #     lambda: self.updateTextField(self.ui.htmlPath, self.study, 'emailHtmlLocation'))
        # self.ui.attachmentPath.editingFinished.connect(
        #     lambda: self.updateTextField(self.ui.attachmentPath, self.study, 'attachmentsLocation'))
        # self.ui.pisPath.editingFinished.connect(
        #     lambda: self.updateTextField(self.ui.pisPath, self.study, 'PISLocation'))
        # self.ui.instructionPath.editingFinished.connect(
        #     lambda: self.updateTextField(self.ui.instructionPath, self.study, 'instructionLocation'))
        # self.ui.savePath.editingFinished.connect(
        #     lambda: self.updateTextField(self.ui.savePath, self.study, 'saveLocation'))
        # update if user choose to type manually

        # ========================== session tab ===================================

        self.ui.addSessionBtn.clicked.connect(self.addNewSession)
        self.ui.sessionSelectDB.currentTextChanged.connect(self.sessionSelectionDBUpdate)

        self.ui.sessionName.editingFinished.connect(
            lambda: self.updateSessionName(self.ui.sessionName, self.getCurrentSession(), 'name'))

        self.ui.sessionDuration.editingFinished.connect(
            lambda: self.updateTextField(self.ui.sessionDuration, self.getCurrentSession(), 'duration'))
        self.ui.audioNotifications.editingFinished.connect(
            lambda: self.updateTextField(self.ui.audioNotifications, self.getCurrentSession(), 'audioNotification'))
        self.ui.incomingInterval.editingFinished.connect(
            lambda: self.updateTextField(self.ui.incomingInterval, self.getCurrentSession(), 'incomingInterval'))
        self.ui.taskPath.editingFinished.connect(
            lambda: self.updateTextField(self.ui.taskPath, self.getCurrentSession(), 'primaryTaskHtml'))
        self.ui.endSessionPopup.editingFinished.connect(
            lambda: self.updateTextField(self.ui.endSessionPopup, self.getCurrentSession(), 'endSessionPopup'))

        self.ui.incomingCB.clicked.connect(self.updateCheckBoxRelatedFields)
        self.ui.incomingCB.clicked.connect(
            lambda: self.updateCheckBox(self.ui.incomingCB, self.getCurrentSession(), 'incomingEmails'))
        self.ui.phishEmailCB.clicked.connect(self.updateCheckBoxRelatedFields)
        self.ui.phishEmailCB.clicked.connect(
            lambda: self.updateCheckBox(self.ui.phishEmailCB, self.getCurrentSession(), 'hasPhishEmails'))
        self.ui.timeCountDownCB.clicked.connect(
            lambda: self.updateCheckBox(self.ui.timeCountDownCB, self.getCurrentSession(), 'timeCountDown'))

        self.ui.starBtn.clicked.connect(
            lambda: self.updateCheckBox(self.ui.starBtn, self.getCurrentSession(), 'starBtn'))
        self.ui.reportBtn.clicked.connect(
            lambda: self.updateCheckBox(self.ui.reportBtn, self.getCurrentSession(), 'reportBtn'))
        self.ui.deleteBtn.clicked.connect(
            lambda: self.updateCheckBox(self.ui.deleteBtn, self.getCurrentSession(), 'deleteBtn'))
        self.ui.unreadBtn.clicked.connect(
            lambda: self.updateCheckBox(self.ui.unreadBtn, self.getCurrentSession(), 'unreadBtn'))

        self.ui.BrowseBtn_T.clicked.connect(
            lambda: self.browseFile(self.ui.taskPath, self.getCurrentSession(), 'primaryTaskHtml'))

        # ========================= legit emails tab =================================

        self.ui.listStart_L.editingFinished.connect(
            lambda: self.updateTextField(self.ui.listStart_L, self.getCurrentLegit().get('emailListRange'), 'start'))
        self.ui.listEnd_L.editingFinished.connect(
            lambda: self.updateTextField(self.ui.listEnd_L, self.getCurrentLegit().get('emailListRange'), 'finish'))
        self.ui.incomingStart_L.editingFinished.connect(
            lambda: self.updateTextField(self.ui.incomingStart_L, self.getCurrentLegit().get('incomingRange'), 'start'))
        self.ui.incomingEnd_L.editingFinished.connect(
            lambda: self.updateTextField(self.ui.incomingEnd_L, self.getCurrentLegit().get('incomingRange'), 'finish'))

        self.ui.shuffleCB_L.clicked.connect(
            lambda: self.updateCheckBox(self.ui.shuffleCB_L, self.getCurrentLegit(), 'shuffleEmails'))

        # ========================= phish emails tab =================================

        self.ui.emailNum_P.editingFinished.connect(
            lambda: self.updateTextField(self.ui.emailNum_P, self.getCurrentPhish(), 'emailListNum'))
        self.ui.emailList_P.editingFinished.connect(
            lambda: self.updateTextField(self.ui.emailList_P, self.getCurrentPhish(), 'emailList'))
        self.ui.emailLoc_P.editingFinished.connect(
            lambda: self.updateTextField(self.ui.emailLoc_P, self.getCurrentPhish(), 'emailListLocations'))
        self.ui.incomingNum_P.editingFinished.connect(
            lambda: self.updateTextField(self.ui.incomingNum_P, self.getCurrentPhish(), 'incomingNum'))
        self.ui.incoming_P.editingFinished.connect(
            lambda: self.updateTextField(self.ui.incoming_P, self.getCurrentPhish(), 'incomingList'))
        self.ui.incomingLoc_P.editingFinished.connect(
            lambda: self.updateTextField(self.ui.incomingLoc_P, self.getCurrentPhish(), 'incomingLocations'))

        self.ui.shuffleCB_P.clicked.connect(self.updateCheckBoxRelatedFields)
        self.ui.shuffleCB_P.clicked.connect(
            lambda: self.updateCheckBox(self.ui.shuffleCB_P, self.getCurrentPhish(), 'shuffleEmails'))

        self.ui.RandomLocCB.clicked.connect(self.updateCheckBoxRelatedFields)
        self.ui.RandomLocCB.clicked.connect(
            lambda: self.updateCheckBox(self.ui.RandomLocCB, self.getCurrentPhish(), 'randomLoc'))

        # ========================= Question marks =================================

        self.ui.audioNotificationQ.clicked.connect(lambda: messageNotification('Information',
                                                                               'List of notification time to finish.\nFormat: integer separated by comma. \ne.g. 1, 5 means will have notification when there is 1 mins and 5 mins left.'))

        self.ui.invervalQ.clicked.connect(lambda: messageNotification('Information',
                                                                      'The time between two incoming emails.\nThe interval is consistant between all incoming emails\ne.g. interval of 2 minutes means the first incoming email will come 2 minutes into the session, then the second would come at 4 mins in, until all incoming emails have been sent. \nNote: 1) make sure the incoming emails are sent before the session ends, 2) both legit and phish incoming emails are included, 3) minumum input = 0.1, i.e. 6 seconds between two incoming emails .'))

        self.ui.emailRangeQ.clicked.connect(lambda: messageNotification('Information',
                                                                        'Enter the first and last email\'s id.\nThe email id should be interger.'))

        self.ui.emailListQ.clicked.connect(lambda: messageNotification('Information',
                                                                       'Enter the list of phish emails.\nFormat: integer (phish email id) separated by comma. \ne.g. 1, 2 means selecting phish email with id 1 and 2.'))

        self.ui.pLocQ.clicked.connect(lambda: messageNotification('Information',
                                                                  'Enter the corresponding location of phish emails in the list.\nFormat: integer separated by comma. \ne.g.'
                                                                  ' 1, 3 means the phish emails will be inserted at the 1st and 3rd position from top down. \nNote: please make'
                                                                  ' sure the number is less or equal to the number of phishing emails.e.g. phishing emails: 1,2,3,4; location 2,4,6,'
                                                                  ' means will display the first three phishing emails in the corresponding location, and skip the last one (because '
                                                                  'location is not given)'))

        self.ui.shufflePQ.clicked.connect(lambda: messageNotification('Information',
                                                                      'randomise the order of phishing emails, need to specify the number of phishing emails added to the '
                                                                      'inbox.\ne.g. out of the n phishing emails specified in "phishing emails", randomly select x of them '
                                                                      'and add into the inbox. When location of the phishing emails are specified, the number should be consistant.'
                                                                      '\nNote: shuffling emails would apply to both phishing emails in the inbox, and incoming phishing emails.'))


        self.ui.saveConfigBtn.clicked.connect(self.saveConfig)
        self.ui.openConfigBtn.clicked.connect(self.loadConfig)
        # self.ui.previewBtn.clicked.connect(self.getCurrentSession)

        self.ui.previewBtn.clicked.connect(self.previewStudy)

        self.addNewSession()

    # *********************************** Functions *********************************************

    # ======== load files and update fields ===================

    def browseFile(self, textField, target, field):
        fname = QFileDialog.getOpenFileName(self, 'open file', './')
        textField.setText(fname[0])
        self.updateTextField(textField, target, field)

    def browseFolder(self, textField, target, field):
        file = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        textField.setText(file)
        self.updateTextField(textField, target, field)

    def updateTextField(self, trigger, target, field):
        target.update({field: trigger.text()})

    def updateTextEditField(self, trigger, target, field):
        target.update({field: trigger.toPlainText()})

    def updateSessionName(self, trigger, target, field):
        index = self.ui.sessionSelectDB.currentIndex()
        self.updateTextField(trigger, target, field)
        self.ui.sessionSelectDB.setItemText(index, trigger.text())

    def updateCheckBox(self, trigger, target, field):
        target.update({field: trigger.isChecked()})

    def previewStudy(self):
        emailManagementApp = EmailResearchLab()
        emailManagementApp.setConfig(self.study)
        emailManagementApp.ui.show()

    # ========== getter ===============
    def getCurrentSession(self):
        return self.study.get('sessions').get(self.currentSession)

    def getCurrentLegit(self):
        return self.getCurrentSession().get('legitEmails')

    def getCurrentPhish(self):
        return self.getCurrentSession().get('phishEmails')

    # ======== session tab ===========
    def sessionSelectionDBUpdate(self):
        if self.ui.sessionSelectDB.currentText() != '':
            for session in self.study.get('sessions'):
                if self.study.get('sessions').get(session).get('name') == self.ui.sessionSelectDB.currentText():
                    self.currentSession = session

            self.updateSessionTabUI()
            self.updateLegitTabUI()
            self.updatePhishTabUI()
            self.updateCheckBoxRelatedFields()

    def addNewSession(self):
        print('clicked')
        sessions = self.study.get('sessions')
        print(len(sessions))
        session_name = 'session' + str(len(sessions) + 1)
        print(session_name)
        sessions[session_name] = {
            'name': session_name,
            'duration': '',
            'audioNotification': '',
            'timeCountDown': False,
            'hasPhishEmails': False,
            'incomingEmails': False,
            'incomingInterval': '',
            'primaryTaskHtml': '',
            'endSessionPopup': '',
            'starBtn': True,
            'reportBtn': True,
            'deleteBtn': True,
            'unreadBtn': True,
            'legitEmails': {},
            'phishEmails': {},

        }
        self.currentSession = session_name

        self.addNewLegit()
        self.addNewPhish()

        self.ui.sessionSelectDB.addItem(self.currentSession)
        self.ui.sessionSelectDB.setCurrentText(self.currentSession)
        # self.updateTable()

    def updateCheckBoxRelatedFields(self):

        if self.ui.phishEmailCB.isChecked():
            self.ui.phishEmailWidget.show()
        else:
            self.ui.phishEmailWidget.hide()

        # overall presence of incoming emails
        if self.ui.incomingCB.isChecked():
            self.ui.incomingIntervalBox.show()
            self.ui.incomingWidget_L.show()
            self.ui.incomingWidget_P.show()
        else:
            self.ui.incomingIntervalBox.hide()
            self.ui.incomingWidget_L.hide()
            self.ui.incomingWidget_P.hide()

        # phishing email shuffles
        if self.ui.shuffleCB_P.isChecked():
            self.ui.pListNumBox.show()
            self.ui.pIncomingNumBox.show()
        else:
            self.ui.pListNumBox.hide()
            self.ui.pIncomingNumBox.hide()

        # phishing email random locations
        if self.ui.RandomLocCB.isChecked():
            self.ui.pListLocBox.hide()
            self.ui.pIncomingLocBox.hide()
        else:
            self.ui.pListLocBox.show()
            self.ui.pIncomingLocBox.show()

    # ========================================================================
    def addNewLegit(self):

        self.getCurrentSession()['legitEmails'] = {
            'emailListRange': {'start': '', 'finish': ''},
            'shuffleEmails': False,
            'incomingRange': {'start': '', 'finish': ''},
        }

        print(self.getCurrentSession())

        self.updateCheckBoxRelatedFields()

        # update the dictionary

    def addNewPhish(self):

        self.getCurrentSession()['phishEmails'] = {
            'emailList': '',
            'randomLoc': False,
            'emailListLocations': '',
            'shuffleEmails': False,
            'emailListNum': '',

            'incomingList': '',
            'incomingLocations': '',
            'incomingNum': '',

        }
        self.updateCheckBoxRelatedFields()

    # ================================================================

    def checkDataType(self):
        integerField = [self.ui.sessionDuration, self.ui.listStart_L, self.ui.listEnd_L,
                        self.ui.incomingStart_L, self.ui.incomingEnd_L, self.ui.emailNum_P, self.ui.incomingNum_P]
        integerListField = [self.ui.audioNotifications, self.ui.emailList_P, self.ui.emailLoc_P, self.ui.incoming_P,
                            self.ui.incomingLoc_P]
        floatField = [self.ui.incomingInterval]

        for element in integerField:
            if (not element.text().isdigit()) and (element.text() != ''):
                messageNotification('Error', 'input field ' + element.objectName() + ' should be integers only')
                return False

        for element in floatField:
            if element.text() != '':
                try:
                    float(element.text())

                except ValueError:
                    messageNotification('Error', 'input field ' + element.objectName() + ' should be a number')
                    return False

            # if (not isinstance(element, (int,float))) and (element.text() != ''):
            #     print(element)
            #     print(ty)
            #     messageNotification('Error', 'input field ' + element.objectName() + ' should be a number')
            #     return False

        for element in integerListField:
            if not element.text() == '':
                list = element.text().split(',')
                print(element.objectName())
                print(list)
                for item in list:
                    if not item.strip().isdigit():
                        messageNotification('Error',
                                            'input field ' + element.objectName() + ' should be integers separated by comma')
                        return False

        if self.ui.sessionDuration.text() != '':
            for audioNotificationTime in self.ui.sessionDuration.text().split(','):
                if int(audioNotificationTime) > int(self.ui.sessionDuration.text()):
                    messageNotification('Error', 'All audio notification times most be less than the sesison duration')

                    return False

        return True

    def saveConfig(self):
        print(self.study)
        if self.checkDataType():

            options = QFileDialog.Options()
            options |= QFileDialog.DontUseNativeDialog
            fileName, _ = QFileDialog.getSaveFileName(self, "Save config file", "./",
                                                      "All Files (*);;", options=options)
            if fileName:
                if fileName[-5:] != '.yaml':
                    fileName = fileName + '.yaml'

                with open(fileName, 'w') as file:
                    yaml.dump(self.study, file, sort_keys=False)

    def loadConfig(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "",
                                                  "All Files (*);;Python Files (*.py)", options=options)
        if fileName:
            with open(fileName) as f:
                self.study = yaml.load(f, Loader=yaml.SafeLoader)

            self.currentSession = list(self.study.get('sessions').keys())[0]
            print('get current session')
            print(self.currentSession)

            self.updateLoadFileSection()
            self.updateSessionSelectDBText()
            self.ui.sessionSelectDB.setCurrentText(self.study.get('sessions').get(self.currentSession).get('name'))
            self.updateCheckBoxRelatedFields()
            # self.updateSessionTabUi()

    def updateLoadFileSection(self):

        self.ui.emailPath.setText(self.study.get('emailListLocation'))
        self.ui.resourcePath.setText(self.study.get('emailResourceLocation'))
        self.ui.savePath.setText(self.study.get('saveLocation'))

    def updateSessionSelectDBText(self):
        self.ui.sessionSelectDB.clear()
        for session in self.study.get('sessions'):
            self.ui.sessionSelectDB.addItem(self.study.get('sessions').get(session).get('name'))
        self.ui.sessionSelectDB.setCurrentText(self.currentSession)

    def updateSessionTabUI(self):
        self.ui.sessionName.setText(self.getCurrentSession().get('name'))
        self.ui.sessionDuration.setText(self.getCurrentSession().get('duration'))
        self.ui.audioNotifications.setText(self.getCurrentSession().get('audioNotification'))
        self.ui.incomingInterval.setText(self.getCurrentSession().get('incomingInterval'))
        self.ui.endSessionPopup.setText(self.getCurrentSession().get('endSessionPopup'))
        self.ui.taskPath.setText(self.getCurrentSession().get('primaryTaskHtml'))

        self.ui.phishEmailCB.setChecked(self.getCurrentSession().get('hasPhishEmails'))
        self.ui.incomingCB.setChecked(self.getCurrentSession().get('incomingEmails'))
        self.ui.timeCountDownCB.setChecked(self.getCurrentSession().get('timeCountDown'))

    def updateLegitTabUI(self):

        self.ui.listStart_L.setText(self.getCurrentLegit().get('emailListRange').get('start'))
        self.ui.listEnd_L.setText(self.getCurrentLegit().get('emailListRange').get('finish'))
        self.ui.incomingStart_L.setText(self.getCurrentLegit().get('incomingRange').get('start'))
        self.ui.incomingEnd_L.setText(self.getCurrentLegit().get('incomingRange').get('finish'))

        self.ui.shuffleCB_L.setChecked(self.getCurrentLegit().get('shuffleEmails'))

    def updatePhishTabUI(self):

        self.ui.emailNum_P.setText(self.getCurrentPhish().get('emailListNum'))
        self.ui.emailList_P.setText(self.getCurrentPhish().get('emailList'))
        self.ui.emailLoc_P.setText(self.getCurrentPhish().get('emailListLocations'))
        self.ui.incomingNum_P.setText(self.getCurrentPhish().get('incomingNum'))
        self.ui.incoming_P.setText(self.getCurrentPhish().get('incomingList'))
        self.ui.incomingLoc_P.setText(self.getCurrentPhish().get('incomingLocations'))

        self.ui.shuffleCB_P.setChecked(self.getCurrentPhish().get('shuffleEmails'))
        self.ui.RandomLocCB.setChecked(self.getCurrentPhish().get('randomLoc'))


def checkNumsBetween(num1, num2, data):
    return data.loc[(data['ID'] >= num1) & (data['ID'] <= num2)].shape[0]


def deleteItemsOfLayout(layout):
    if layout is not None:
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.setParent(None)
            else:
                deleteItemsOfLayout(item.layout())


def messageNotification(messageType, text):
    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Information)
    msgBox.setText(text)
    msgBox.setWindowTitle(messageType)
    msgBox.setStandardButtons(QMessageBox.Ok)
    returnValue = msgBox.exec()
    if returnValue == QMessageBox.Ok:
        print('OK clicked')


if __name__ == '__main__':
    app = QApplication([])
    mainWindow = ConfigPage()
    mainWindow.ui.show()
    app.exec_()
