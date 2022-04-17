class notPassInfo(object):
    semester = ""
    courseName = ""
    achievement = 0
    credit = 0.0

    def _init_(self):
        self.semester = ""
        self.courseName = ""
        self.achievement = 0
        self.credit = 0.0

    def cleanData(self):
        self.semester = ""
        self.courseName = ""
        self.achievement = 0
        self.credit = 0.0

    def setsemester(self, insemester):
        self.semester = insemester

    def getsemester(self):
        return self.semester

    def setcourseName(self, incourseName):
            self.courseName = incourseName

    def getcourseName(self):
            return self.courseName

    def setachievement(self, inachievement):
            self.achievement = inachievement

    def getachievement(self):
            return self.achievement

    def setcredit(self, incredit):
            self.credit = incredit

    def getcredit(self):
            return self.credit


class alarmNotifyInfo(object):
    className = ""
    studentName = ""
    studentId = 0
    totalCredit = 0
    notPassInfo = []

    def _init_(self):
        self.className = ""
        self.studentName = ""
        self.studentId = 0
        self.totalCredit = 0
        self.notPassInfo = []

    def cleanData(self):
        self.className = ""
        self.studentName = ""
        self.studentId = 0
        self.totalCredit = 0
        self.notPassInfo = []

    def setclassName(self, inclassName):
            self.className = inclassName

    def getclassName(self):
            return self.className

    def setstudentName(self, instudentName):
            self.studentName = instudentName

    def getstudentName(self):
            return self.studentName

    def setstudentId(self, instudentId):
            self.studentId = instudentId

    def getstudentId(self):
            return self.studentId

    def addnotPassInfo(self, innotPassInfo):
            self.notPassInfo.append(copy.deepcopy(innotPassInfo))

    def getnotPassInfo(self):
            return self.notPassInfo

    def addtotalCredit(self, intotalCredit):
            self.totalCredit += intotalCredit

    def gettotalCredit(self):
            return self.totalCredit