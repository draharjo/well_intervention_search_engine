'''
- inisialisasi data frame
- input location of the folder
- mencari file apa saja pada folder tersebut
- membuka tiap excel pada file tersebut
- return CSV file

1. get_contents
2. list_file
3. identify new_file
'''
import os
import csv
import datetime, xlrd
from itertools import chain 
import pandas as pd
import time
import matplotlib.pyplot as plt
import win32com.client as win32
import pythoncom

class LessonLearn():
    '''
    1. List a file --> get_contents(loct)
    '''
    def __init__(self):
        #initiate pandas dataframe
        self.lesson_learn = {"directory":[], "summary":[]}
        self.df = pd.DataFrame(self.lesson_learn)
        
        #set allowable extension
        self.allowable_extension = ['xlsx', 'xls']
        
        #count lesson learn total
        self.total_lesson_learn = 0 
        
    def get_contents(self, loct):
        #initiate loct
        self.loct = loct
        #reinitialize pandas dataframe
        self.lesson_learn = {"directory":[], "summary":[]}
        self.df = pd.DataFrame(self.lesson_learn)
        
        filenames = {file:root.split("\\")[-1] for root, dirs, files in os.walk(self.loct) for file in files if file.split(".")[-1] in self.allowable_extension}
        self.total_lesson_learn = len(filenames)
        for ll in filenames:
            #membuat path file yang akan dibukan tersebut
            excel_path = self.loct + "\\"+ ll

            #membuka file excel
            wb = xlrd.open_workbook(excel_path)
            sheet = wb.sheet_by_index(0)

            #flattening 2D array jadi 1D array agar bisa digabung
            cells = list(chain.from_iterable(sheet.__dict__["_cell_values"]))

            #gabungin seluruh data cells jadi satu string
            a =' '.join(str(cell) for cell in cells)

            #tambahin datanya ke pd dataframe
            self.df = self.df.append({'directory': ll, 'summary': a}, ignore_index=True)

            #convert string to lower case
            self.df["clean_summary"] = self.df.summary.apply(to_lower)
        return self.df
    
    def list_file(self, loct):
        self.loct = loct
        filenames = []
        filenames = [file for root, dirs, files in os.walk(self.loct) for file in files if file.split(".")[-1] in self.allowable_extension]
        return filenames
        
    def identify_new_file(self, old_list, new_list):
        self.new_doc = []
        for doc in new_list:
            if doc not in old_list:
                self.new_doc.append(doc)
        return self.new_doc
    
    def get_contributor_data(self, loct_path):
        '''
        Mengambil data nama dan tanggal pembuatan dari dokumen lesson learned yang dibuat
        '''
        self.loct = loct_path
        self.name = ''
        self.date = ''
        self.title = ''
        self.inputdata = []

        excel_path = self.loct
        wb = xlrd.open_workbook(excel_path)
        sheet = wb.sheet_by_index(0)

        # row, column
        #name = sheet.cell_value(6,7)
        #date = sheet.cell_value(7,7)
    
        name = sheet.cell_value(6,3)
        date = sheet.cell_value(7,3)
       
        date_as_dt = datetime.datetime(*xlrd.xldate_as_tuple(date, wb.datemode))
        date = date_as_dt.strftime('%m/%d/%Y')
        # data["date"] = datetime.date.today().strftime('%m/%d/%Y')

        title = self.loct.split("\\")[-1]
        
        self.inputdata.append(name)
        self.inputdata.append(date)
        self.inputdata.append(title)

        return self.inputdata

    def update_log_submission(self, log_file, data):
        '''
        log_file = location of the log file
        data = ['name', 'date', 'title']
        '''
        self.data = {'name': data[0], 'date': data[1], 'title': data[2]}
        df = pd.read_csv(log_file, index_col = None)
        df = df.append(self.data, ignore_index = True)
        #index = False biar enggak nambah banyak index
        df.to_csv(log_file, index = False)
    
    def generate_statistic(self, df, loct ="submission_by_month.png"):
        '''
        df berisi name, title, 
        '''
        plt.figure(figsize = (20,6))
        df.resample('M').count()['title'].plot()
        plt.title("Lesson Learned Historical Submission", fontsize =  14)
        plt.xlabel("TOTAL LL SUBMISSION")
        plt.savefig(loct)

    def generate_bar(self, df, loct= "submission_by_people.png"):
        df = df.groupby(["name"]).count().reset_index()
        x = df.name
        y = df.title
        plt.figure(figsize = (20,6))
        plt.barh(x, y, color= 'b')
        plt.title("Lesson Learned Historical Submission", fontsize =  20)
        plt.axvline(y.mean(), c='r' ,ls='--')
        plt.text(y.mean(), 0.5, "MEAN = " + "{:.2f}".format(y.mean()), ha='center', color ='r')
        plt.xlabel("TOTAL LL SUBMISSION")
        for index, value in enumerate(y):
            plt.text(value, index, str(value))

        plt.savefig(loct)
    
    def send_email(self, files):
        '''
        list of file name
        '''
        #CREATING THE MESSAGE
        text = ["<h4> Dear All, </h4> Someone just recently uploaded new lesson learned, click the link to search the documents"]
        for index, file in enumerate(files):
            text.append(f"<br> {str(index)}. {file} </br>")
        text.append("<br></br><br><b>http://192.168.1.105:5000/lessonlearn</br>")
        text.append("<br></br><br><b>Thank you </br>")
        body = ''.join(text)
        
        pythoncom.CoInitialize() #harus diinisialisasikan
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'dharmawan.r@pertamina.com'
        mail.Subject = 'Automatic Email'
        mail.Body = 'Message body'
        mail.HTMLBody = body

        # To attach a file to the email (optional):
        #attachment  = "F:\\My Photo\\Operasi WOWS\\20200319_002000.jpg"
        #mail.Attachments.Add(attachment)
        mail.Send()
        
def to_lower(x):
    return str(x.lower())
