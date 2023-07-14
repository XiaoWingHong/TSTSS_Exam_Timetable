import re

def printErrorMsg(fileName):
    print(fileName)
    print('Press any key to continue ...')
    input()
    exit()

class exam:
    def __init__(self, examDate):
        self.examDate = examDate
        self.subjects = []
        self.noExam = []

class subject:
    def __init__(self, name, timeLimit, period, room, form, parent):
        self.name = name
        self.period = period
        self.form = form
        self.parent = parent

        room = room.replace(' ','').split(',')
        for i in room:
            if '<' in i and '>' in i:
                n = int(i[i.index('<')+1:i.index('>')])
                str_room = i[:i.index('<')]
                for j in range(n):
                    room.insert(room.index(i), str_room)
                room.remove(i)
        self.room = room
        self.teachers = ['']*len(self.room)

        timeLimit = re.compile(r'\d+').findall(str(timeLimit))
        self.timeLimit = list(map(lambda x: int(x), timeLimit))

class teacher:
    def __init__(self, name):
        self.name = name
        self.lessons = {}
        self.exams = {}
        self.totalTime = 0
        self.lessonTime = 0
        self.teachedSubjectsAndClasses = {}

class TA:
    def __init__(self, name):
        self.name = name
        self.totalTime = 0
        self.exams = {}

class examDetails:
    def __init__(self, name, period, room):
        self.name = name
        self.period = period
        self.room = room