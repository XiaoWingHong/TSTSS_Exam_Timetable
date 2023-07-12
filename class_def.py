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

        try:
            self.timeLimit = int(timeLimit)
        except:
            printErrorMsg('Exam Timetable: Time Limit is not a number!')

class teacher:
    def __init__(self, name):
        self.name = name
        self.lessons = {}
        self.exams = {}
        self.classes = []
        self.totalTime = 0
        self.lessonTime = 0

class TA:
    def __init__(self, name):
        self.name = name
        self.totalTime = 0