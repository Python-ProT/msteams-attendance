import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill

from flask import Flask, flash, request, redirect, url_for, render_template,send_file
# import urllib.request
import os
from os.path import join, dirname, realpath
from werkzeug.utils import secure_filename
 
app = Flask(__name__)
 
# UPLOADS_PATH = 'static/uploads/'
UPLOADS_PATH = join(dirname(realpath(__file__)), 'static/uploads')
 
app.secret_key = "secret key"
app.config['UPLOADS_PATH'] = UPLOADS_PATH
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
 
# ALLOWED_EXTENSIONS = set(['png', 'xlsx'])

# ALLOWED_EXTENSIONS = set(['txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif', 'csv','xlsx'])
ALLOWED_EXTENSIONS = set(['csv','xlsx'])

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
     
 
@app.route('/')
def home():
    return render_template('index.html')



# @app.route('/', methods=['POST'])
# def upload_image():
    
           # return render_template('main.html')

@app.route('/submit',methods=['POST','GET'])
def submit():
    if 'file' not in request.files:
        flash('No file Selected') 
        return redirect(request.url)
    file = request.files['file']
    file2 = request.files['file2']
    if file.filename == '' or file2.filename == '':
        flash('No File selected for uploading')
        return redirect(request.url)
    if (file and allowed_file(file.filename)) and (file2 and allowed_file(file2.filename)):
        filename = secure_filename(file.filename)
        global filename2
        filename2 = secure_filename(file2.filename)
        
        file.save(os.path.join(app.config['UPLOADS_PATH'], filename))
        file2.save(os.path.join(app.config['UPLOADS_PATH'], filename2))
        
       
        compute(filename,filename2)
        # download(filename2)
        
        return render_template('index.html', filename=filename,filename2=filename2)
    else:
        return redirect(request.url) 

@app.route('/download')
def download():
    global filename2
    uploads=os.path.join(app.config['UPLOADS_PATH'], filename2)
    print(uploads)
    # uploads = "/home/yashita/Desktop/upload/env/static/uploads/try1.xlsx"
    return send_file(uploads,as_attachment=True) 



def difftime(a, b):
    a = a.split(":")
    b = b.split(":")
    if b[-1][-2].upper() == "P" and (int)(b[0]) != 12:
        b[0] = str(int(b[0]) + 12)
    if a[-1][-2].upper() == "P" and (int)(a[0]) != 12:
        a[0] = str(int(a[0]) + 12)
    return (int(b[0]) - int(a[0]))*3600 + (int(b[1]) - int(a[1]))*60 + int(b[2][:2]) - int(a[2][:2])

    # Minimum time calculation for setting start and end time for individual students


def minitime(a, b):
    a = a.split(":")
    b = b.split(":")
    if b[-1][-2].upper() == "P" and a[-1][-2].upper() == "A":
        return True
    elif (b[-1][-2].upper() == "P" and a[-1][-2].upper() == "P") or (b[-1][-2].upper() == "A" and a[-1][-2].upper() == "A"):
        if int(a[0]) < int(b[0]):
            return True
        elif int(a[0]) == int(b[0]):
            if int(a[1]) < int(b[1]):
                return True
            elif int(a[1]) == int(a[1]):
                if int(a[2][:2]) < int(b[2][:2]) or int(a[2][:2]) == int(b[2][:2]):
                    return True
                else:
                    return False
            else:
                return False
        else:
            return False
    else:
        return False
def compute(filename,filename2):
    totol = 0
    file2 = os.path.join(app.config['UPLOADS_PATH'], filename)
    file1 = os.path.join(app.config['UPLOADS_PATH'], filename2)

    df1 = pd.read_excel(file1)
    print(df1)
    # totol += 1
    path = os.path.join(app.config['UPLOADS_PATH'], filename2)

    df2 = pd.read_csv(file2, encoding="utf-16", sep='\t')
    print(df2)
    newdate = df2["Timestamp"][0].split(',')[0]
    date = newdate.split('/')[1] + "/" + \
        newdate.split('/')[0] + "/" + newdate.split('/')[2]
    print(date)
    print(df2["Timestamp"][0].split(',')[0])
    start=request.form['start']
    end=request.form['end']
    start = start.split(":")      
    end = end.split(":")
    
    t ="AM"
    if start[0]>"12":
        start[0]=int(start[0])-12
        t="PM"
    

    h =str(start[0])
    
    m =start[1]
    s ="00"
    # t =request.form['clicked4']
    if len(h) == 1:
        h = "0" + str(h)
    if len(m) == 1:
        m = "0" + str(m)
    if len(s) == 1:
        s = "0" + str(s)
    start_time = h + ":" + m + ":" + s + " " + t
    print(start_time)

    if end[0]>"12":
        end[0]=int(end[0])-12
        t="PM" 
    h =str(end[0])
    m = end[1]
    s = "00"
    # t = request.form['clicked8']
    if len(h) == 1:
        h = "0" + str(h)
    if len(m) == 1:
        m = "0" + str(m)
    if len(s) == 1:
        s = "0" + str(s)
    end_time = h + ":" + m + ":" + s + " " + t
    print(end_time)

    thers = "25"
    # print(thers)
    # print(date)
    print(type(df2["Full Name"]))
    stud_attend = {}.fromkeys(df2["Full Name"])
    # print(stud_attend)
    for i in stud_attend:
        i.strip(" ").upper()
        print(i.strip(" ").upper())
    # Timestamp extraction for individual students
    l = []
    # ye ek ka time bta rha h ek jagah pe
    for i in stud_attend:
        for j in range(0, len(df2["Full Name"])):
            if i == df2["Full Name"][j]:
                l.append(df2["Timestamp"][j].split(",")[1].lstrip(" "))
                # print(l)
        stud_attend[i] = l
        l = []
        # print(l)
    # Calculate difference between every joining and left time and calculate whether the person is present for given threshold
    # True if he is present and False, if he is absent
    tot_time = difftime(start_time, end_time)
    print(tot_time)
    # If the person not attended meeting itself means set False for that person
    for i in stud_attend:
        sumtime = 0
        x = stud_attend[i]
        if minitime(x[0], start_time):
            x[0] = start_time
        # if minitime(temp, x[0]):
        #     temp = x[0]
        if (len(x) % 2) == 0 and minitime(end_time, x[-1]):
            x[-1] = end_time
        if (len(x) % 2) != 0:
            x.append(end_time)
            # print(x.append(end_time))
        for j in range(0, len(x), 2):
            sum_time = difftime(x[j], x[j+1])
            if sum_time >= eval(thers)*tot_time / 100:
                stud_attend[i] = "P"
            else:
                stud_attend[i] = "A"
    # df1 = pd.DataFrame(columns= ['Full Name"'])
    # for column in df1.columns:
    #     if ((column != "Total") and (column != "Percentage")):
    #         df1['Total'] = 0
    #         df1['Percentage'] = 0
    #         print(df1)
    if "Total" not in df1:
        df1["Total"] = 0

    if "Percentage" not in df1:
        df1["Percentage"] = 0
    # Empty column creation for the person to append the attendance list created
    for i in df1["Full Name"]:
        print(i.strip(" ").upper())
        if i.strip(" ").upper() not in stud_attend:
            stud_attend[i.strip(" ").upper()] = "A"
    # Empty column creation for the person to append the attendance list created
    df1[date] = list(range(0, len(df1["Full Name"])))
    # For each person put the value True or False according to rules defined above
    d = list(df1["Full Name"])
    # print(d)
    for i in d:
        df1[date][d.index(i)] = stud_attend[i.strip(" ").upper()]
    #   print(stud_attend[i.strip("s ").upper()])
    df1 = df1.set_index("Scholar No")
    # print(df1)
    df1 = df1.sort_values("Scholar No")
    # print(df1)
    # df3 = df3.sort_index(ascending=True, axis=1)
    #df3 = df3.sort_values("Scholar No")

    for column in df1.columns:
        if ((column != "Full Name") and (column != "Total") and (column != "Percentage")):
            totol += 1
# #nump int64
# np_in = np.int64(0)
# printtype(np_int))
# # <clss 'numpy.int64'>
# #Convrt to python int
# py_in = np_int.item()
# printtype(py_int))
# # <clss 'int'>

    for j in range(0, len(df1["Full Name"])):
        # print(df3[date][j]," ",df3["Tota"][j])
        py_int = np.int64(df1["Total"][j]).item()
        if df1[date][j] == "P":
            # print(type(np.int64(df3["Total"][j]).item()))
            # print(py_int)
            py_int += 1
            # print(py_int)
            df1["Total"][j] = py_int
        df1["Percentage"][j] = py_int*100/totol

        df3 = df1[["Full Name"]].copy()
    for column in df1.columns:
        if ((column != "Scholar No") or (column != "Full Name") or (column != "Total") or (column != "Percentage")):
            df3[column] = df1[column]
    df3.to_excel(path)
    wb = openpyxl.load_workbook(path)
    ws = wb['Sheet1']
    fill_pattern = PatternFill(patternType='solid', fgColor='C64747')
    for j in range(0, len(df1["Full Name"])):

        if(df3["Percentage"][j] < 75):
            my_list = list(df3)
            index = my_list.index("Percentage")
            col = chr(index+65+1)
            ws[col+str(j+2)].fill = fill_pattern
            wb.save(path)
    print(df3)
    print(file1)
    print('\nDone!Check your Excel')
    # download(file1)





if __name__=="__main__":
    app.run(debug=True)
