class exam:
    def __init__(self, examDate):
        self.examDate = examDate
        self.subjects = []
        self.noExam = []

class subject:
    def __init__(self, name, timeLimit, period, room, form):
        self.name = name
        self.timeLimit = timeLimit
        self.period = period
        self.form = form

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

class teacher:
    def __init__(self, name):
        self.name = name
        self.lessons = {}
        self.exams = {}
        self.classes = []
        self.totalTime = 0

class TA:
    def __init__(self, name):
        self.name = name
        self.totalTime = 0