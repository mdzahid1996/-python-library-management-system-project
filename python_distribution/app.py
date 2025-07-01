# app.py

# pip install mysqlclient
import sys
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
import MySQLdb as msd
import datetime
from xlrd import *
from xlsxwriter import *

ui, _ = loadUiType('app_design.ui')
loginUi, _ = loadUiType('login.ui')

class LoginCls(QMainWindow, loginUi):

    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.loginBtn.clicked.connect(self.handleLogin)
        self.main_window = None # To hold reference to the main window

    def handleLogin(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        loginUsername = self.loginUsername.text()
        loginUserPass = self.loginUserPass.text()

        sql = "SELECT * FROM users WHERE name = %s AND password = %s"
        self.cur.execute(sql, (loginUsername, loginUserPass))
        data = self.cur.fetchone()

        if data:
            self.main_window = Library() # Create an instance of the Library window
            self.main_window.show() # Show the main window
            self.hide() # Hide the login window
        else:
            self.loginError.setText('username or password is invalid...')

class Library(QMainWindow, ui):

    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.buttons()
        self.handleUiTab()
        self.show_categories()
        self.show_author()
        self.show_publisher()
        self.bookCategoryForEditDeleteTab()
        self.authorForEditDeleteTab()
        self.publisherForEditDeleteTab()
        self.bookCategoryForAddTab()
        self.authorForAddTab()
        self.publisherForAddTab()
        self.show_client()
        self.showAllBooks()
        self.show_day_operations()

    def buttons(self):
        self.dayOperationBtn.clicked.connect(self.day_operationsTab)
        self.booksBtn.clicked.connect(self.booksTab)
        self.userBtn.clicked.connect(self.usersTab)
        self.settingsBtn.clicked.connect(self.settingsTab)
        self.clientBtn.clicked.connect(self.clientTab)

        self.addBookBtn.clicked.connect(self.addBook)
        self.bookSearchBtn.clicked.connect(self.searchBook)
        self.updateBookInfo.clicked.connect(self.editBook)
        self.deleteBookBtn.clicked.connect(self.deleteBook)

        self.addCategoryBtn.clicked.connect(self.add_categories)
        self.addAuthorBtn.clicked.connect(self.add_author)
        self.addPublisherBtn.clicked.connect(self.add_publisher)

        self.addUserBtn.clicked.connect(self.addNewUser)
        self.userLogin.clicked.connect(self.UserLogin)
        self.userUpadteBtn.clicked.connect(self.updateUser)

        self.addClientBtn.clicked.connect(self.add_client)
        self.searchClientBtn.clicked.connect(self.search_client)
        self.updateClientBtn.clicked.connect(self.update_client)
        self.deleteClientBtn.clicked.connect(self.delete_client)

        self.dayOprBtn.clicked.connect(self.addDayOperations)

        self.dayoperationsExcel.clicked.connect(self.export_day_operations)
        self.booksExcel.clicked.connect(self.export_books)
        self.clientExcel.clicked.connect(self.export_clients)

    def handleUiTab(self):
        self.tabWidget.tabBar().setVisible(False)

    def day_operationsTab(self):
        self.tabWidget.setCurrentIndex(0)

    def booksTab(self):
        self.tabWidget.setCurrentIndex(1)

    def clientTab(self):
        self.tabWidget.setCurrentIndex(2)

    def usersTab(self):
        self.tabWidget.setCurrentIndex(3)

    def settingsTab(self):
        self.tabWidget.setCurrentIndex(4)

    ############# book operation #############
    def addBook(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        bookName = self.bookTitle.text()
        bookcode = self.bookCode.text()
        bookcategory = self.bookCategory.currentText()
        bookauthor = self.bookAuthor.currentText()
        bookpublisher = self.bookPublisher.currentText()
        bookPrice = self.bookPrice.text()
        bookDescription = self.addBookTabDescription.toPlainText()

        self.cur.execute('''
            INSERT INTO book(book_name, book_description, book_code, book_category, book_author, book_publisher, book_price)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        ''',(bookName, bookDescription, bookcode, bookcategory, bookauthor, bookpublisher, bookPrice))
        self.db.commit()
        self.statusBar().showMessage('new book is added successfully..')

        self.bookTitle.setText('')
        self.bookCode.setText('')
        self.bookCategory.setCurrentIndex(0)
        self.bookAuthor.setCurrentIndex(0)
        self.bookPublisher.setCurrentIndex(0)
        self.bookPrice.setText('')
        self.addBookTabDescription.setPlainText('')
        self.showAllBooks()

    def searchBook(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        searchedBookName = self.searchBookName.text()
        sql = "SELECT * FROM book WHERE book_name=%s"
        self.cur.execute(sql, [(searchedBookName)])
        data = self.cur.fetchone()
        if data:
            self.searchBookTitle.setText(data[1])
            self.searchBookCode.setText(data[3])
            self.searchBookPrice.setText(str(data[7]))
            self.editDelBookTabDescription.setPlainText(str(data[2]))
        else:
            self.statusBar().showMessage('Book not found.')

    def editBook(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        searchBookTitle = self.searchBookName.text()

        bookName = self.searchBookTitle.text()
        bookcode = self.searchBookCode.text()
        bookcategory = self.availableCategory.currentText()
        bookauthor = self.availablAuthor.currentText()
        bookpublisher = self.availablePublisher.currentText()
        bookPrice = self.searchBookPrice.text()
        bookDescription = self.editDelBookTabDescription.toPlainText()

        self.cur.execute('''
        UPDATE book SET book_name=%s, book_description=%s, book_code=%s, book_category=%s, book_author=%s, book_publisher=%s, book_price=%s WHERE book_name=%s
        ''', (bookName,bookDescription,bookcode,bookcategory,bookauthor,bookpublisher,bookPrice,searchBookTitle))
        self.db.commit()
        self.statusBar().showMessage('book is updated..')
        self.showAllBooks()

    def deleteBook(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        searchBookTitle = self.searchBookName.text()
        permission = QMessageBox.warning(self, 'delete book', 'do you want to delete the book', QMessageBox.Yes|QMessageBox.No)
        if permission == QMessageBox.Yes:
            delQuery = "DELETE FROM book WHERE book_name=%s"
            self.cur.execute(delQuery, [(searchBookTitle)])
            self.db.commit()
            self.statusBar().showMessage('book is deleted..')
            self.showAllBooks()

        self.searchBookName.setText('')
        self.searchBookTitle.setText('')
        self.searchBookCode.setText('')
        self.availableCategory.setCurrentIndex(0)
        self.availablAuthor.setCurrentIndex(0)
        self.availablePublisher.setCurrentIndex(0)
        self.searchBookPrice.setText('')
        self.editDelBookTabDescription.setPlainText('')

    def showAllBooks(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        self.cur.execute("SELECT book_code,book_name,book_author,book_publisher,book_category,book_price FROM book")
        data = self.cur.fetchall()
        if data:
            self.tableWidget_3.setRowCount(0) # Clear existing rows before populating
            self.tableWidget_3.setColumnCount(len(data[0])) # Set column count based on data
            for row_num, row_data in enumerate(data):
                self.tableWidget_3.insertRow(row_num)
                for col_num, item in enumerate(row_data):
                    self.tableWidget_3.setItem(row_num, col_num, QTableWidgetItem(str(item)))
        self.db.close()

    # *********** day operation ************
    # **************************************

    def addDayOperations(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        bookName = self.dayOprBook.text()
        clientName = self.dayOprClient.text()
        typ = self.dayOprTypSel.currentText()
        day = self.dayOprDaySel.currentIndex() + 1
        fromdate = datetime.date.today()
        todate = fromdate + datetime.timedelta(days=3)

        self.cur.execute("INSERT INTO dayoperations(bookname,type,days,fromDate,toDate,clientName) VALUES(%s, %s, %s, %s, %s, %s)",(bookName,typ,day,fromdate,todate,clientName))
        self.db.commit()
        self.statusBar().showMessage('new operation is added..')
        self.show_day_operations()

    def show_day_operations(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        self.cur.execute("SELECT bookname,clientName,type,fromDate,toDate FROM dayoperations")
        data = self.cur.fetchall()
        if data:
            self.tableWidget_2.setRowCount(0) # Clear existing rows before populating
            self.tableWidget_2.setColumnCount(len(data[0])) # Set column count based on data
            for row_num, row_data in enumerate(data):
                self.tableWidget_2.insertRow(row_num)
                for col_num, item in enumerate(row_data):
                    self.tableWidget_2.setItem(row_num, col_num, QTableWidgetItem(str(item)))
        self.db.close()


     ############# user operation #############
    def addNewUser(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        username = self.regUsername.text()
        mail = self.regEmail.text()
        password = self.regPass.text()
        rePassword = self.regPassAgain.text()

        if password == rePassword:
            self.cur.execute("INSERT INTO users(username, useremail, userspassword) VALUES(%s, %s, %s)", (username, mail, password))
            self.db.commit()
            self.statusBar().showMessage('new user is added successfully.....')
            self.errorMsg.setText('')
        else:
            self.errorMsg.setText('password is not matched....')
            self.statusBar().showMessage(' ')

    def UserLogin(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        loginUsername = self.userloginUsername.text()
        loginUserPass = self.userloginPass.text()

        sql = "SELECT * FROM users WHERE name = %s AND password = %s"
        self.cur.execute(sql, (loginUsername, loginUserPass))
        data = self.cur.fetchone()
        if data:
            self.edituserInfo.setEnabled(True)
            self.statusBar().showMessage('login successful...')

            self.editUserUsername.setText(data[1])
            self.editUserMail.setText(data[2])
            self.editUserPass.setText(data[3])
        else:
            self.statusBar().showMessage('username or password is invalid...')

    def updateUser(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        updateUsername = self.editUserUsername.text()
        updateUserMail = self.editUserMail.text()
        updateUserPass = self.editUserPass.text()
        updateUserPassAgain = self.lineEdit_17.text()
        beforeUpdateUsername = self.userloginUsername.text()

        if updateUserPass == updateUserPassAgain:
            self.cur.execute('''
            UPDATE users SET username=%s, useremail=%s, userspassword=%s WHERE username=%s
            ''',(updateUsername, updateUserMail, updateUserPass, beforeUpdateUsername))
            self.db.commit()
            self.statusBar().showMessage('user information is updated')
        else:
            self.statusBar().showMessage('password is not matched')

    ############# client operation #############

    def add_client(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        clientName = self.newClientName.text()
        clientEmail = self.newClientEmail.text()
        clientId = self.newClientId.text()

        self.cur.execute("INSERT INTO client(clientName, clientEmail, clientNid) VALUES(%s, %s, %s)", (clientName, clientEmail, clientId))
        self.db.commit()
        self.statusBar().showMessage('new client is added')
        self.show_client()
    def search_client(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        searchclient = self.searchClient.text()
        sql = "SELECT * FROM client WHERE clientNid=%s"
        self.cur.execute(sql, [(searchclient)])
        data = self.cur.fetchone()
        if data:
            self.updelClientName.setText(data[1])
            self.updelClientEmail.setText(data[2])
            self.updelClientId.setText(data[3])
        else:
            self.statusBar().showMessage('Client not found.')

    def update_client(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        upclientName = self.updelClientName.text()
        upclientEmail = self.updelClientEmail.text()
        upclientId = self.updelClientId.text()
        searchedclientId = self.searchClient.text()
        if upclientName and upclientEmail and upclientId:
             self.cur.execute("UPDATE client SET clientName=%s, clientEmail=%s, clientNid=%s WHERE clientNid=%s", (upclientName,upclientEmail,upclientId,searchedclientId))
             self.db.commit()
             self.statusBar().showMessage('client information is updated...')
             self.show_client()
        else:
            self.clientUperror.setText('fields are required....')

    def show_client(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        self.cur.execute("SELECT * FROM clients")
        data = self.cur.fetchall()
        if data:
            self.tableWidget.setRowCount(0) # Clear existing rows before populating
            self.tableWidget.setColumnCount(len(data[0])) # Set column count based on data
            for row_num, row_data in enumerate(data):
                self.tableWidget.insertRow(row_num)
                for col_num, item in enumerate(row_data):
                    self.tableWidget.setItem(row_num, col_num, QTableWidgetItem(str(item)))
        self.db.close()

    def delete_client(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        searchedclientId = self.searchClient.text()
        if searchedclientId:
            permission = QMessageBox.warning(self,'delete client', 'are you sure to delete the client', QMessageBox.Yes|QMessageBox.No)
            if permission == QMessageBox.Yes:
                sql = "DELETE FROM client WHERE clientNid=%s"
                self.cur.execute(sql, [(searchedclientId)])
                self.db.commit()
                self.statusBar().showMessage('information is deleted successfully...')
                self.show_client()
        else:
            self.clientUperror.setText('client id field is required...')


    ############# setting operation #############

    def add_author(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        authorname = self.addAuthor.text()
        self.cur.execute("INSERT INTO author(author_name) VALUES (%s)", (authorname,))
        self.db.commit()
        self.statusBar().showMessage('new author is added successfully...')
        self.show_author()
        self.authorForEditDeleteTab()
        self.authorForAddTab()
        print('author is added successfully...')

    def show_author(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        self.cur.execute("SELECT author_name FROM author")
        data = self.cur.fetchall()
        if data:
            self.allAuthors_2.setRowCount(0) # Clear existing rows before populating
            self.allAuthors_2.setColumnCount(len(data[0])) # Set column count based on data
            for row_num, row_data in enumerate(data):
                self.allAuthors_2.insertRow(row_num)
                for col_num, item in enumerate(row_data):
                    self.allAuthors_2.setItem(row_num, col_num, QTableWidgetItem(str(item)))

    def add_categories(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        categoryname = self.categoryName.text()
        self.cur.execute("INSERT INTO category(category_name) VALUES (%s)", (categoryname,))
        self.db.commit()
        self.statusBar().showMessage('new book category is added successfully...')
        self.show_categories()
        self.bookCategoryForEditDeleteTab()
        self.bookCategoryForAddTab()
        print('book category is added successfully...')

    def show_categories(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        self.cur.execute("SELECT category_name FROM category")
        data = self.cur.fetchall()
        if data:
            self.allCategories.setRowCount(0) # Clear existing rows before populating
            self.allCategories.setColumnCount(len(data[0])) # Set column count based on data
            for row_num, row_data in enumerate(data):
                self.allCategories.insertRow(row_num)
                for col_num, item in enumerate(row_data):
                    self.allCategories.setItem(row_num, col_num, QTableWidgetItem(str(item)))


    def add_publisher(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        publishername = self.addPublisher.text()
        self.cur.execute("INSERT INTO publisher(publisher_name) VALUES (%s)", (publishername,))
        self.db.commit()
        self.statusBar().showMessage('new publisher is added successfully...')
        self.show_publisher()
        self.publisherForEditDeleteTab()
        self.publisherForAddTab()
        print('publisher is added successfully...')

    def show_publisher(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        self.cur.execute("SELECT publisher_name FROM publisher")
        data = self.cur.fetchall()
        if data:
            self.allPublishers.setRowCount(0) # Clear existing rows before populating
            self.allPublishers.setColumnCount(len(data[0])) # Set column count based on data
            for row_num, row_data in enumerate(data):
                self.allPublishers.insertRow(row_num)
                for col_num, item in enumerate(row_data):
                    self.allPublishers.setItem(row_num, col_num, QTableWidgetItem(str(item)))

    def bookCategoryForAddTab(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        self.cur.execute("SELECT category_name FROM category")
        data = self.cur.fetchall()
        self.bookCategory.clear() # Clear existing items before repopulating
        for d in data:
            self.bookCategory.addItem(d[0])

    def authorForAddTab(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        self.cur.execute("SELECT author_name FROM author")
        data = self.cur.fetchall()
        self.bookAuthor.clear() # Clear existing items before repopulating
        for d in data:
            self.bookAuthor.addItem(d[0])

    def publisherForAddTab(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        self.cur.execute("SELECT publisher_name FROM publisher")
        data = self.cur.fetchall()
        self.bookPublisher.clear() # Clear existing items before repopulating
        for d in data:
            self.bookPublisher.addItem(d[0])

    def bookCategoryForEditDeleteTab(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        self.cur.execute("SELECT category_name FROM category")
        data = self.cur.fetchall()
        self.availableCategory.clear() # Clear existing items before repopulating
        for d in data:
            self.availableCategory.addItem(d[0])

    def authorForEditDeleteTab(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        self.cur.execute("SELECT author_name FROM author")
        data = self.cur.fetchall()
        self.availablAuthor.clear() # Clear existing items before repopulating
        for d in data:
            self.availablAuthor.addItem(d[0])

    def publisherForEditDeleteTab(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        self.cur.execute("SELECT publisher_name FROM publisher")
        data = self.cur.fetchall()
        self.availablePublisher.clear() # Clear existing items before repopulating
        for d in data:
            self.availablePublisher.addItem(d[0])

    # *********** Export ***********
    # *****************************
    def export_day_operations(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        self.cur.execute("SELECT bookname,clientName,type,fromDate,toDate FROM dayoperations")
        data = self.cur.fetchall()
        wb = Workbook('day_operations.xlsx')
        sheet1 = wb.add_worksheet()
        sheet1.write(0,0, 'bookname')
        sheet1.write(0,1, 'clientName')
        sheet1.write(0,2, 'type')
        sheet1.write(0,3, 'fromDate')
        sheet1.write(0,4, 'toDate')
        row_num = 1
        for row in data:
            col_num = 0
            for item in row:
                sheet1.write(row_num, col_num,str(item))
                col_num +=1
            row_num += 1
        wb.close()
        self.statusBar().showMessage('data is downloaded successfully....')
    def export_books(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()

        self.cur.execute("SELECT book_code,book_name,book_author,book_publisher,book_category,book_price FROM book")
        data = self.cur.fetchall()
        wb = Workbook('allBooks.xlsx')
        sheet1 = wb.add_worksheet()
        sheet1.write(0,0, 'book_code')
        sheet1.write(0,1, 'book_name')
        sheet1.write(0,2, 'book_author')
        sheet1.write(0,3, 'book_publisher')
        sheet1.write(0,4, 'book_category')
        sheet1.write(0,5, 'book_price')
        row_num = 1
        for row in data:
            col_num = 0
            for item in row:
                sheet1.write(row_num, col_num,str(item))
                col_num +=1
            row_num += 1
        wb.close()
        self.statusBar().showMessage('data is downloaded successfully....')
    def export_clients(self):
        self.db = msd.connect(host='localhost', user='admin', password='12345678', db='library_management')
        self.cur = self.db.cursor()
        self.cur.execute("SELECT clientName,clientEmail,clientNid FROM client")
        data = self.cur.fetchall()
        wb = Workbook('allClients.xlsx')
        sheet1 = wb.add_worksheet()
        sheet1.write(0,0, 'clientName')
        sheet1.write(0,1, 'clientEmail')
        sheet1.write(0,2, 'clientNid')
        row_num = 1
        for row in data:
            col_num = 0
            for item in row:
                # Removed the backslash from the end of the previous line
                sheet1.write(row_num, col_num,str(item))
                col_num +=1
            row_num += 1
        wb.close()
        self.statusBar().showMessage('data is downloaded successfully....')

def main():
    app = QApplication(sys.argv)
    login_window = LoginCls() # Create the login window instance
    login_window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
