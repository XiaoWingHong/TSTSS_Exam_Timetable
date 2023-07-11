import pandas as pd
import class_def as cd
import re
import random
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles.borders import Border, Side
import numpy as np

def printErrorMsg(fileName):
    print(fileName)
    print('Press any key to continue ...')
    input()
    exit()

print('Reading data ...')

try:
    df= pd.read_excel('Input/Specific Examer.xlsx')
except:
    printErrorMsg('Can\'t find file \'Specific Examer.xlsx\'!')

MAIN_EXAMER_OF_ENG_SPEAKING = [x for x in df['English Speaking\n主考官'].tolist() if x == x]
ORAL_EXAMER_OF_ENG_SPEAKING = [x for x in df['English Speaking\nOral 考官'].tolist() if x == x]
MAIN_EXAMER_OF_ENG_LISTENING = [x for x in df['English Listening\n主考官'].tolist() if x == x]
MAIN_EXAMER_OF_CHIN_SPEAKING = [x for x in df['中文説話\n主考官'].tolist() if x == x]
ORAL_EXAMER_OF_CHIN_SPEAKING = [x for x in df['中文説話\nOral 考官'].tolist() if x == x]
MAIN_EXAMER_OF_CHIN_LISTENING = [x for x in df['中文聆聽\n主考官'].tolist() if x == x]
MAIN_EXAMER_OF_PTH = [x for x in df['普通話\n主考官'].tolist() if x == x]
MAIN_EXAMER_OF_VA = [x for x in df['VA\n主考官'].tolist() if x == x]
FOREIGN_TEACHER = [x for x in df['外籍老師'].tolist() if x == x]

tmp = {}
for examer in MAIN_EXAMER_OF_VA:
    examer = examer.replace(' ', '')
    tmp[int(examer[0])] = examer[2:]
MAIN_EXAMER_OF_VA = tmp

CANT_BE_EXAMER = [x for x in (df['不能監考\n(校長)'].tolist() + df['外籍老師'].tolist()) if x == x]

TA_DATA = []
for ta in [x for x in df['TA'].tolist() if x == x]:
    TA_DATA.append(cd.TA(name=ta))

try:
    df= pd.read_excel('Input/Exam Timetable.xlsx', skiprows=[0], usecols=lambda x: 'Unnamed' not in x)
except:
    printErrorMsg('Can\'t find file \'Exam Timetable.xlsx\'!')

ET_DATA = []
for date in df.columns:
    if date[-2:-1] in ['一', '二', '三', '四', '五', '六', '日']:
        ET_DATA.append(cd.exam(examDate=date))
    else:
        print('Exam Timetable: Date Formate Error!')
    
for exam in ET_DATA:
    exam.subjects = []
    exam.noExam = []
    listedColum = df[exam.examDate].tolist()
    listedColum.insert(0, exam.examDate)
    form = 0
    for i in range(len(listedColum)-1):
        if listedColum[i] == exam.examDate:
            form += 1
            if listedColum[i+1] == '上課':
                exam.noExam.append(form)
    listedColum = [x for x in listedColum if x == x] #remove nan
    listedColum = list(filter(lambda i: i != '上課', listedColum)) #remove '上課'
    i = 0
    form = 0
    while i < len(listedColum):
        if listedColum[i] == exam.examDate:
            form += 1
            i += 1
        else:
            exam.subjects.append(cd.subject(name = listedColum[i], timeLimit = listedColum[i+1], room = listedColum[i+2], period = listedColum[i+3], form=form, parent=exam))
            i += 4

def findForm(lessonName):
    if re.match('[1-9]', lessonName[0]):
        return lessonName[0]
    else:
       return 'all'
    
def getClass(lessonName):
    listedClass = []
    pattern = re.compile(r'[1-9]+[0-9]+[0-9]')
    listedClass = pattern.findall(lessonName)
    return listedClass

try:
    sheets = pd.ExcelFile('Input/Teacher Timetable.xlsx')
    timeSlot = pd.read_excel('Input/Teacher Timetable.xlsx', skiprows=[0,1,3,6,10,13,16])['Unnamed: 0'].tolist()
except:
    printErrorMsg('Can\'t find file \'Teacher Timetable.xlsx\'!')

TT_DATA = []
dateDict = {'Mon' : '一', 'Tue' : '二', 'Wed' : '三', 'Thu' : '四', 'Fri' : '五'}
for sheetName in sheets.sheet_names:
    if sheetName not in CANT_BE_EXAMER:
        TT_DATA.append(cd.teacher(sheetName))
for teacher in TT_DATA:
    teacher.lessons = {}
    teacher.totalTime = 0
    teacher.classes = []
    teacher.exams = {}
    df = pd.read_excel('Input/Teacher Timetable.xlsx', skiprows=[0,1,3,6,10,13,16], sheet_name=teacher.name, usecols=lambda x: 'Unnamed' not in x)
    for date in df.columns:
        noLesson = df[date].isnull().tolist()
        listedColum = df[date].tolist()
        teacher.lessons[dateDict[date]] = {}
        for i in range(len(noLesson)):
            if noLesson[i] is False:
                teacher.lessons[dateDict[date]][timeSlot[i]] = findForm(listedColum[i])
                teacher.classes += getClass(listedColum[i])
    teacher.classes = [*set(teacher.classes)]

for teacher in TT_DATA:
    teacher.totalTime = 0
    for exam in ET_DATA:
        for needLessonForms in exam.noExam:
            for key in teacher.lessons[exam.examDate[-2:-1]]:
                if teacher.lessons[exam.examDate[-2:-1]][key] == str(needLessonForms) or teacher.lessons[exam.examDate[-2:-1]][key] == 'all':
                    teacher.totalTime += 35
                    
# AVG_TIME = 0
# for exam in ET_DATA:
#     for subject in exam.subjects:
#         if 'eaking' not in subject.name:
#             if subject.room[0] == 'HALL':
#                 AVG_TIME += subject.timeLimit * (len(subject.room) - 1)
#             else:
#                 AVG_TIME += subject.timeLimit * len(subject.room)

# for teacher in TT_DATA:
#     AVG_TIME += teacher.totalTime
# AVG_TIME /= len(TT_DATA)

def checkTime(examTime, lessonTime):
    time1 = []
    time2 = []
    numPattern = re.compile(r'\d+')
    time1.append(int(numPattern.findall(examTime)[0])*60+int(numPattern.findall(examTime)[1]))
    time1.append((int(numPattern.findall(examTime)[-2]) + (12 if (re.search( r'p', examTime, re.I) and len(numPattern.findall(examTime)[-2]) < 2) else 0))*60+int(numPattern.findall(examTime)[-1]))
    time2.append(int(numPattern.findall(lessonTime)[0])*60+int(numPattern.findall(lessonTime)[1]))
    time2.append((int(numPattern.findall(lessonTime)[-2]) + (12 if (re.search( r'p', lessonTime, re.I) and len(numPattern.findall(lessonTime)[-2]) < 2) else 0))*60+int(numPattern.findall(lessonTime)[-1]))
    if (time1[0] > time2[1]) or (time1[1] < time2[0]):
        return False
    else:
        return True


def findAvalibleTeachers(subject, specificExamer=None):
    avalibleTeachersList = []
    teacherData = TT_DATA
    if specificExamer != None:
        teacherData = []
        for teacherNames in specificExamer:
            teacherData.append(findParentObj(TT_DATA, teacherNames))
    for teacher in teacherData:
        isBussy = False
        if len(subject.parent.noExam) > 0:
            for value in teacher.lessons[subject.parent.examDate[-2:-1]].values(): 
                if value in subject.parent.noExam or value == 'all':
                    for lessonTime in [key for key in teacher.lessons[subject.parent.examDate[-2:-1]] if (teacher.lessons[subject.parent.examDate[-2:-1]][key] == value or teacher.lessons[subject.parent.examDate[-2:-1]][key] == 'all')]:
                        if not isBussy:
                            isBussy = checkTime(subject.period, lessonTime)
                        else:
                            break
                if isBussy:
                    break
        if subject.parent.examDate in [key for key in teacher.exams]:
            for examTime in teacher.exams[subject.parent.examDate]:
                if not isBussy:
                    isBussy = checkTime(subject.period, examTime)
                else:
                    break

        if not isBussy:
            avalibleTeachersList.append(teacher)
            
    avalibleTeachersList.sort(key=lambda x: x.totalTime, reverse=False)
    return avalibleTeachersList[0]
    
    

def findParentObj(data, name):
    return data[list(map(lambda x : x.name == name, data)).index(True)]

def appendTeachers(i, subject, avalibleTeacher, ignore=False):
    if subject.teachers[i] != '':
        return
    subject.teachers[i] = avalibleTeacher.name
    if not ignore:
        avalibleTeacher.totalTime += subject.timeLimit
    if subject.parent.examDate not in [key for key in avalibleTeacher.exams]:
        avalibleTeacher.exams[subject.parent.examDate] = []
    avalibleTeacher.exams[subject.parent.examDate].append(subject.period)

def appendTA(i, subject):
    avalibleTAList = TA_DATA
    avalibleTAList.sort(key=lambda x: x.totalTime, reverse=False)
    avalibleTA = avalibleTAList[0]
    subject.teachers[i] = avalibleTA.name
    avalibleTA.totalTime += subject.timeLimit

print('Processing ...')

for exam in ET_DATA:
    for subject in exam.subjects:
        if 'peaking' in subject.name:
            appendTeachers(0, subject, findAvalibleTeachers(subject, MAIN_EXAMER_OF_ENG_SPEAKING), ignore=True)
            for i in range(1,4):
                if subject.room[i] == 'HALL' or subject.room[i][-1] == '1':
                    subject.teachers[i] = TA_DATA[i-1].name
                    findParentObj(TA_DATA, TA_DATA[i-1].name).totalTime += subject.timeLimit
            subject.teachers[subject.teachers.index('')] = FOREIGN_TEACHER[0]
            for i in range(subject.teachers.index(''),len(subject.room)):
                appendTeachers(i, subject, findAvalibleTeachers(subject, ORAL_EXAMER_OF_ENG_SPEAKING), ignore=True)
        elif 'istening' in subject.name and 'TSA' not in subject.name:
            appendTeachers(0, subject, findAvalibleTeachers(subject, MAIN_EXAMER_OF_ENG_LISTENING))
            appendTeachers(1, subject, findAvalibleTeachers(subject))
            for i in range(1,len(subject.room)):
                appendTA(i, subject)
        elif '說話' in subject.name or '説話' in subject.name:
            appendTeachers(0, subject, findAvalibleTeachers(subject, MAIN_EXAMER_OF_CHIN_SPEAKING))
            for i in range(1,3):
                subject.teachers[i] = TA_DATA[i-1].name
                findParentObj(TA_DATA, TA_DATA[i-1].name).totalTime += subject.timeLimit
            for i in range(3,len(subject.room)):
                appendTeachers(i, subject, findAvalibleTeachers(subject, ORAL_EXAMER_OF_CHIN_SPEAKING))
        elif '普通話' in subject.name:
            appendTeachers(0, subject, findAvalibleTeachers(subject, MAIN_EXAMER_OF_PTH))
            for i in range(1,len(subject.room)):
                appendTA(i, subject)
        elif '聆聽' in subject.name and 'TSA' not in subject.name and '普通話' not in subject.name:
            appendTeachers(0, subject, findAvalibleTeachers(subject, MAIN_EXAMER_OF_CHIN_LISTENING))
            for i in range(1,len(subject.room)):
                appendTA(i, subject)
        elif '視覺藝術' in subject.name:
            appendTeachers(0, subject, findAvalibleTeachers(subject, [MAIN_EXAMER_OF_VA[subject.form]]))

for subject in sorted(list(filter(lambda x: '' in x.teachers, list(np.concatenate(list(map(lambda x: x.subjects, ET_DATA))).flat))), key=lambda x: x.timeLimit, reverse=True):
    if 'HALL' in subject.room:
        for i in range(0,2):
            appendTeachers(i, subject, findAvalibleTeachers(subject))
        for i in range(2,len(subject.room)):
            appendTA(i, subject)
    else:
        for i in range(0,len(subject.room)):
            appendTeachers(i, subject, findAvalibleTeachers(subject))
        
workbook = openpyxl.Workbook()
sheet = workbook.worksheets[0]
formDict = { 1 : '中一級', 2 : '中二級', 3 : '中三級', 4 : '中四級', 5 : '中五級', 6 : '中六級'}

orangeFill = PatternFill(patternType='solid', fgColor=Color(rgb='FFC000'))
yellowFill = PatternFill(patternType='solid', fgColor=Color(rgb='FFFF00'))
greyFill = PatternFill(patternType='solid', fgColor=Color(rgb='D9D9D9'))
cellborder = Border(left=Side(style='medium'), 
                     right=Side(style='medium'), 
                     top=Side(style='medium'), 
                     bottom=Side(style='medium'))


for i in range(1, ET_DATA[0].subjects[-1].form + 1):
    sheet.cell(row = sheet.max_row+2, column = 1).value = formDict[i]
    top = sheet.max_row+1
    for col, exam in enumerate(ET_DATA,start=1):
        sheet.cell(row = top, column = col).value = exam.examDate
        sheet.cell(row = top, column = col).border = cellborder
        sheet.cell(row = top, column = col).font = Font(bold=True)
        sheet.column_dimensions[get_column_letter(col)].width = 17
        current_row = top+1
        for subject in list(filter(lambda x: x.form == i, exam.subjects)):
            sheet.cell(row = current_row, column = col).value = subject.name
            sheet.cell(row = current_row, column = col).fill = orangeFill
            sheet.cell(row = current_row, column = col).font = Font(bold=True)

            sheet.cell(row = current_row+1, column = col).value = subject.timeLimit

            sheet.cell(row = current_row+2, column = col).value = subject.period

            for j in range(current_row, current_row+3):
                sheet.cell(row = j, column = col).border = cellborder
                sheet.cell(row = j, column = col).alignment = Alignment(horizontal='center', wrapText=True, vertical = 'center')

            current_row += 3
            for j in range(len(subject.room)):
                sheet.cell(row = current_row, column = col).value = subject.room[j] + ': ' + subject.teachers[j]
                sheet.cell(row = current_row, column = col).alignment = Alignment(horizontal='center', wrapText=True, vertical = 'center')
                sheet.cell(row = current_row, column = col).border = cellborder
                sheet.cell(row = current_row, column = col).fill = yellowFill
                current_row += 1
    
    for y in range(1, sheet.max_column+1):
        for x in range(top, sheet.max_row+1):
            if sheet.cell(row = x, column = y).value == None:
                sheet.cell(row = x, column = y).fill = greyFill

workbook.create_sheet('Total Time')
sheet2 = workbook.worksheets[1]
sheet2.cell(row = 1, column = 1).value = 'Teacher'
sheet2.cell(row = 1, column = 2).value = 'Minutes'
for i, teacher in enumerate(TT_DATA, start=2):
    sheet2.cell(row = i, column = 1).value = teacher.name
    sheet2.cell(row = i, column = 2).value = teacher.totalTime

workbook.save('監考時間表.xlsx')
