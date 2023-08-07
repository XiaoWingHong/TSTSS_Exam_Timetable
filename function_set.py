import re

def printErrorMsg(fileName):
    print(fileName)
    print('Press any key to continue ...')
    input()
    exit()

def findParentObj(data, name):
    return data[list(map(lambda x : x.name == name, data)).index(True)]

def transferTimeFormat(inputTime):
    tmp = re.compile(r'\d+').findall(inputTime)
    if len(re.compile(r'p', re.I).findall(inputTime)) > 0:
        if len(tmp[2]) == 1:
            tmp[2] = str(int(tmp[2]) + 12)
        if len(re.compile(r'p', re.I).findall(inputTime)) > 1 and len(tmp[0]) == 1:
            tmp[0] = str(int(tmp[0]) + 12)

    return tmp[0]+':'+tmp[1]+'-'+tmp[2]+':'+tmp[3]

def checkTime(examTime, lessonTime):
    time1 = []
    time2 = []
    numPattern = re.compile(r'\d+')
    time1.append(int(numPattern.findall(examTime)[0])*60+int(numPattern.findall(examTime)[1]))
    time1.append(int(numPattern.findall(examTime)[-2])*60+int(numPattern.findall(examTime)[-1]))
    time2.append(int(numPattern.findall(lessonTime)[0])*60+int(numPattern.findall(lessonTime)[1]))
    time2.append(int(numPattern.findall(lessonTime)[-2])*60+int(numPattern.findall(lessonTime)[-1]))
    if (time1[0] > time2[1]) or (time1[1] < time2[0]):
        return True
    else:
        return False