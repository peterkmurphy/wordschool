# Used for writing school reports.

from docx import Document

class ReportWriter:
    ''' This is for writing reports using Word. '''

    def __init__(self, filename):
        ''' We initialise this using a existing file as a template. '''
        self.document = Document(filename)
        self.heading = self.document.tables[0]
        self.marks = self.document.tables[1]

    def assigntexttoheading(self, row, col, parnum, text):
        runstocount = len(self.heading.cell(row,col).paragraphs[parnum].runs)
        self.heading.cell(row,col).paragraphs[parnum].runs[0].text = text
        for i in range(runstocount - 1):
            self.heading.cell(row,col).paragraphs[parnum].runs[i + 1].clear()

    def assigntexttomarks(self, row, col, parnum, text):
        runstocount = len(self.marks.cell(row,col).paragraphs[parnum].runs)
        self.marks.cell(row,col).paragraphs[parnum].runs[0].text = text
        for i in range(runstocount - 1):
            self.marks.cell(row,col).paragraphs[parnum].runs[i + 1].clear()


    def writeinitial(self, initid):
        ''' This writes the report using a dict (initid) with the
        following possible keys:
        id: the student id
        name: the name of the student
        course: the course (like "General English")
        cl: the class (like "Pre-Intermediate")
        sd: the starting date
        ed: the end date.
        '''

        if initid.has_key("id"):
            self.assigntexttoheading(0, 0, 2, initid["id"])
        if initid.has_key("name"):
            self.assigntexttoheading(0, 1, 2, initid["name"])
        if initid.has_key("course"):
            self.assigntexttoheading(0, 2, 1, initid["course"])
        if initid.has_key("cl"):
            self.assigntexttoheading(1, 2, 1, initid["cl"])
        if initid.has_key("sd"):
            self.assigntexttoheading(0, 3, 2, initid["sd"])
        if initid.has_key("ed"):
            self.assigntexttoheading(0, 4, 2, initid["ed"])

    def writedates(self, weekdates):
        ''' Writes the starting dates for each week of class to the report.
        Since there are 10 weeks, there should be ten values in weekdates.
        By convention, they should be the starting Monday.
        '''
        for i in range(5):
            self.assigntexttomarks(i + 1, 1, 0, weekdates[i])
        for i in range(5):
            self.assigntexttomarks(i + 7, 1, 0, weekdates[5 + i])

    def writemarks(self, firstweek, lastweek, marks):
        numweeks = lastweek - firstweek + 1
        if firstweek <= 5 and lastweek >= 5:
            numweeks += 1
        if lastweek == 10:
            numweeks += 1
        if firstweek <= 5:
            startindex = firstweek
        else:
            startindex = firstweek + 1
        for i in range(len(marks)):
            for j in range(len(marks[i])):
                self.assigntexttomarks(startindex + i, j + 2, 0, marks[i][j])



    def save(self, newfilename):
        self.document.save(newfilename)

print ("Hello")
document = Document('Sample.docx')
thecore = document.core_properties



tables = document.tables
headingtable = tables[0]
print [i.text for i in headingtable.cell(0,0).paragraphs[2].runs]
print [i.text for i in headingtable.cell(0,1).paragraphs[2].runs]
print [i.text for i in headingtable.cell(0,2).paragraphs]
print [i.text for i in headingtable.cell(1,2).paragraphs]
print [i.text for i in headingtable.cell(0,3).paragraphs]
print [i.text for i in headingtable.cell(0,4).paragraphs]
markstable = tables[1]
for i in range(len(markstable.rows)):
    for j in range(len(markstable.columns)):
        print [k.text for k in markstable.cell(i,j).paragraphs]
thecore.author
document.save('demo.docx')

R = ReportWriter('Sample.docx')
R.writeinitial({"id":"666", "name": "Biggles", "course": "DDD", "cl":"Int",
    "sd":"01/01/2000", "ed":"31/12/2012"})
R.writedates([unichr(i + 64) for i in range(10)])
R.writemarks(1, 10,  [[unichr(j + 75 + i) for j in range(6)] for i in range(11)])
R.save('Other.docx')
