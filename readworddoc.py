# Used for writing student reports.

from docx import Document
import yaml
import os

class ReportWriter:
    ''' This is for writing student reports using Word. '''

    def __init__(self, filename):
        ''' We initialise this using a existing file as a template. '''
        self.filename = filename
        self.document = Document(filename)
        self.heading = self.document.tables[0]
        self.marks = self.document.tables[1]

    def assigntexttotable(self, table, row, col, parnum, text):
        ''' Writes text to a cell of a table in a Word doc. Arguments:
        table: the table in a Word document
        row: the row number (zero based)
        col: the col number (zero based)
        parnum: the number of the paragraph.
        text: the text to write.

        Note: if there is existing text at the destination paragraph, then the
        method overwrites the text while keeping the existing font. However,
        if there is no text in the paragraph, then the method just adds text
        to the paragraph; the font defaults to the "Normal" style, which is
        perhaps not what you want. This is probably a quirk of python-docx
        which needs to be worked around.
        '''
        runstocount = len(table.cell(row,col).paragraphs[parnum].runs)
        if (runstocount > 0):
            table.cell(row,col).paragraphs[parnum].runs[0].text = text
            for i in range(runstocount - 1):
                table.cell(row,col).paragraphs[parnum].runs[i + 1].clear()
        else:
            table.cell(row,col).paragraphs[parnum].add_run(text)


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
            self.assigntexttotable(self.heading, 0, 0, 2, initid["id"])
        if initid.has_key("name"):
            self.assigntexttotable(self.heading, 0, 1, 2, initid["name"])
        if initid.has_key("course"):
            self.assigntexttotable(self.heading, 0, 2, 1, initid["course"])
        if initid.has_key("cl"):
            self.assigntexttotable(self.heading, 1, 2, 1, initid["cl"])
        if initid.has_key("sd"):
            self.assigntexttotable(self.heading, 0, 3, 2, initid["sd"])
        if initid.has_key("ed"):
            self.assigntexttotable(self.heading, 0, 4, 2, initid["ed"])

    def writedates(self, weekdates):
        ''' Writes the starting dates for each week of class to the report.
        Since there are 10 weeks, there should be ten values in weekdates.
        By convention, they should be the starting Monday.
        '''
        for i in range(5):
            self.assigntexttotable(self.marks, i + 1, 1, 0, weekdates[i])
        for i in range(5):
            self.assigntexttotable(self.marks, i + 7, 1, 0, weekdates[5 + i])

    def writemarks(self, firstweek, lastweek, marks):
        ''' This writes marks to a table of marks, under the assumption that:
        (a) There are 10 weeks to write to, and 6 criteria to mark by;
        each row is for a week, and each column is a criterion.
        (b) There is a progress report for week 5 (with 6 criteria to mark by);
        if you write marks for week 5, you need to mark the subsequent
        progress report row.
        (c) There is a progress report for week 10 (with 6 criteria to mark by);
        if you write marks for week 10, you need to mark the subsequent
        progress report row.
        There is a firstweek and a lastweek parameter, with the condition:
        1 <= firstweek <= lastweek <= 10
        There is also a marks table, which contains the marks for each week
        (and progress reports as necessary). For each element marks[i][j],
        i represents row indices, and j represents criteria in columns.
        '''
        numweeks = lastweek - firstweek + 1
        if firstweek <= 5 and lastweek <= 5:
            numweeks += 1
        if lastweek == 10:
            numweeks += 1
        if firstweek <= 5:
            startindex = firstweek
        else:
            startindex = firstweek + 1
        for i in range(len(marks)):
            for j in range(len(marks[i])):
                self.assigntexttotable(self.marks, startindex + i, j + 2, 0,
                    marks[i][j])

    def writecomment(self, comment):
        ''' This writes a comment for a student report. '''
        self.assigntexttotable(self.marks, 13, 0, 1, "Comment: " + comment)

    def save(self, newfilename):
        ''' This saves the document to a new file (newfilename). '''
        self.document.save(newfilename)

    @staticmethod
    def WriteReports(template, inputyaml, outputdir):
        ''' Takes a inputyaml file, which contains report data, and writes
        it to an output directory
        '''
        reportdata = yaml.load_all(file(inputyaml, 'r'))
        for report in reportdata:
            R = ReportWriter(template)
            R.writeinitial(report[0])
            if report[1]["comment"]:
                R.writecomment(report[1]["comment"])
            R.writemarks(report[2]["start"], report[2]["end"], report[3:])
            R.save(os.path.join(outputdir, report[0]["name"]+".docx"))

# Commented out, but can be uncommented as necessary.
#
#document = Document('Sample.docx')
#thecore = document.core_properties
#
#tables = document.tables
#headingtable = tables[0]
#print [i.text for i in headingtable.cell(0,0).paragraphs[2].runs]
#print [i.text for i in headingtable.cell(0,1).paragraphs[2].runs]
#print [i.text for i in headingtable.cell(0,2).paragraphs]
#print [i.text for i in headingtable.cell(1,2).paragraphs]
#print [i.text for i in headingtable.cell(0,3).paragraphs]
#print [i.text for i in headingtable.cell(0,4).paragraphs]
#markstable = tables[1]
#for i in range(len(markstable.rows)):
#    for j in range(len(markstable.columns)):
#        print [k.text for k in markstable.cell(i,j).paragraphs]
#thecore.author
#document.save('demo.docx')
#
#R = ReportWriter('Sample.docx')
#R.writeinitial({"id":"667", "name": "Biggles", "course": "DDD", "cl":"Int",
#    "sd":"01/01/2000", "ed":"31/12/2012"})
#R.writedates([unichr(i + 64) for i in range(10)])
#R.writemarks(1, 10,  [[unichr(j + 75 + i) for j in range(6)]
#    for i in range(12)])
#R.writecomment("Blah blah blah")
#R.save('Other.docx')
#
#data = yaml.load_all(file('report.yml', 'r'))
#for item in data:
#    print item

ReportWriter.WriteReports("Sample.docx", "report.yml", "test")
