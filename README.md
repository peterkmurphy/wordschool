# Wordschool

The **wordschool** project consist of Python scripts designed to automate the creation
of student progress reports (in Word .docx format) - for students.   

## Rationale

I'm an ESL teacher. There are many things I need to do for my work, and one of them
is writing progress reports for students. These are Word .docx files that should conform to the template provided by the school. For each week, I need to write down how well a particular student does in such criteria as "Grammar", "Listening" and "Participation".
As each student gets their own report, I also need to write down their name and their
student id. When a student finishes at the school, I need to write down a comment on
their own progress. Writing a progress report requires a lot of data entry, but it is (*mostly*) necessary data entry - students deserve to know how they are going in their study.

The big problem is when I am using Microsoft Word to write multiple reports, especially when I am doing it simultaneously. I find it ***utterly not ergonomic*** to handle all
these files, especially with class sizes of 18 or so. The program creates giant menus to
handle everything. Yes, I could work on one document at a time, and close all other files I am not writing, but it doesn't work in practice; it is common to compare how two or more students are doing when writing reports. So I find it slow and frustrating to use Word to open one doc, close another doc, re-open a third, and shift between viewing the first and a fourth and a fifth. It gets a little confusing.

I had the big idea one day. Let's not enter write multiple student progress reports.
Instead, I write all the student data in *one* place - preferably *one* text file. That will make it easier to keep track of all students in a class. If I organise the data in
that text file in a particular way, I can use my Python skills to write a script to read file and generate all the student reports file automatically. That saves me from using Word to write any student progress report. Thus came the **wordschool** project to fruition.

## Installation

This is how you install the project.

* Install Python 2.*x* (if you haven't installed it already). The project is not tested
    with Python 3.*x*, so be cautious trying with this.
* Set up your own virtual environment using **venv** or similar, and activate it.
* There are two dependencies. One is [PyYaml](https://github.com/yaml/pyyaml), as YAML
    is the format I desired for my text files. The other is [python-docx](https://python-docx.readthedocs.io/en/latest/), which is a great
    Python library for interacting with Word docx files. You can install them as so:

```
pip install pyyaml
pip install python-docx
```

* Finally, clone this project somewhere in the virtual environment folder. You are
   good to go.

## The scripts

### The readworddoc.py script

The readworddoc.py script is the main script for the **wordschool** project.
The idea is to read a text file that contains the student report data, and
write student reports to a given directory. The text file is meant to be a
[YAML](http://yaml.org/); the actual format is described below.

The syntax for readworddoc.py is:

```
usage: readworddoc.py [-h] [--gradsub] [--no-gradsub] [--graddir GRADDIR]
                      [--template TEMPLATE]
                      mode reportyaml outputdir
```
The positional arguments are:

* mode: what is the main goal of using readworddoc.py? The modes are:
    - "N": takes a template *TEMPLATE*, reads the report data file *reportyaml*,
    and generate student report files in the directory *outputdir*. This mode
    may overwrite existing reports, so use with caution. However, this mode is
    probably suited for a new class of students, where reports need to created
    from scratch.
    - "E": takes a template *TEMPLATE*, reads the report data file *reportyaml*,
    but then adds data to existing report files in the directory *outputdir*.
    This mode is probably suited for an existing class with existing
    class records.
    - "T": Reads the dates from *reportyaml*, and sets the dates in all reports
    in *outputdir*. This mode can be used when there are existing reports that
    need to be reused for a new term.
* reportyaml: a text file that contains all student report file for a class
    (in YAML format).
* outputdir: a directory where all data is written to (either by creating
    student reports, or appending data to existing ones).

The Optional arguments are:

* -h, --help: show a help message and exit.
* --gradsub: The reports for "graduate" students (see below) are written to
    a subdirectory of *outputdir*. This is the default mode.
* --no-gradsub: The reports for "graduate" students are written directly to
    *outputdir*.
* --graddir *GRADDIR*: The name of the subdirectory where graduate student
    reports are written. If this argument is not present, the default is
    "Graduated".
* --template *TEMPLATE*: The name of the template file (in docx format) used for
    generating reports.

Examples of calling this program.

* Reading report.yml, and creating *new* reports in the Test1 directory, using
    TemplatePublicHyphen.docx as a template; graduate reports go into the
    TheGrad subdirectory.

```    
python readworddoc.py --graddir TheGrad --template TemplatePublicHyphen.docx N report.yml "Test1"
```

* Reading report.yml, and creating *new* reports in the Test2 directory, using
    TemplatePublicHyphen.docx as a template. Graduate reports go into the
    same directory.

```
python readworddoc.py --no-gradsub  --template TemplatePublicHyphen.docx N report.yml "Test2"
```

* Updating the reports in the Test2 directory with data from reportwithdates.yml.

```
python readworddoc.py --no-gradsub  --template TemplatePublicHyphen.docx E reportwithdates.yml "Test2"
```

* Updating the *dates* in the reports in Test2, otherwise leaving data untouched.

```
python readworddoc.py T sampledate.yml "Test2"
```

### The report files

Report files are YAML text files that consist of documents like the following
(taken from report.yml):


```
---
- {name: "King Gizzard", id: "5678", sd: "01/02/2001", ed: "11/09/2001"}
- comment:
    "While King Gizzard had a bad habit of rocking out in class and distracting
    people, he was able to restrain his Krautrock flavored beats in order to
    study!"
- {start: 1, end: 4}
- [A, B, B, A, A, Very Good]
- [C, D, D, E, F, Ordinary]
- [B, C, B, C, D, Getting Better]
- [C, A, B, A, A, Good]
...

```

Let's break it down:

* Each student has its own YAML document (starting with "---", and ending with
    "...") in the text file. Student data is expressed as a sequence of items.
* The first item is a YAML map with keys "name", "id", "sd" and "ed". These are
    meant to be the student name, their id, their start date and end date.
* The next item is a map for comments. At the school, you need to provide
    comments for students who are graduates. Otherwise, they can be left blank.
    At school, students reports for Graduates are generally put in a subdirectory
    (called "Graduates").
* The next item is a map which indicates the start week and the end week. Terms
    go from week 1 to week 10; students can start at any time through the cycle.
* The final items are the progress for the weeks that the students are at
    the school (here, week 1 to 4). Generally, there are six criteria for
    each week.

What generally happens is that readworddoc.py takes an existing template (such as
TemplatePublicHyphen.docx) fills the data in the fields, and produces Word
documents. Sometimes you have input files such as reportwithdates.yml, where the
first document looks like this:

```
# Some sample reports. The first document sets the term (of 10 weeks) for the
# reports. Each item is the date of the starting Monday for that week.
---
- 23/10/17
- 30/10/17
- 06/11/17
- 13/11/17
- 20/11/17
- 27/11/17
- 04/12/17
- 11/12/17
- 18/12/17
- 01/01/18
...
```    
When readworddoc.py encounters YAML like this, it looks at the first document.
If it sees that the first document consists of sequence of items of the form
DD/MM/YY, it sets the *dates* in the word document as well.

The final example file provided is sampledate.yml - this only contains date
data rather than student data. This sort of file is only used when readworddoc.py
is running in "T" mode.


### The makerepfromtsv.py script

It is a little time consuming to generate the YAML student report file from
scratch. I found that some time would be saved if one starts with a tab
separated file with the following format.

```
Student_ID_1[TAB]Student_Name_1[TAB]StartDate_1[TAB]EndDate_1
Student_ID_2[TAB]Student_Name_2[TAB]StartDate_2[TAB]EndDate_2
...
```
This utility can then generate YAML documents with the form:

```
---
- {name: "Student_Name_n", id: "Student_ID_n", sd: "StartDate_n", ed: "EndDate_n"}
- comment:
- {start: 1, end: 10}
...
```
I found Microsoft Excel useful for contain the results of grammar and other tests of my students - and copying the content of an Excel file to a text editor results in a tab separated file. Having a
YAML file generated from this saved me lots of time.

The syntax for makerepfromtsv.py is:

```
usage: makerepfromtsv.py [-h] input_file output_file
```

The mandatory arguments are:

* input_file: A tab separated value file with student data
* output_file: A YAML file representing student data (with blank class scores)


The Optional arguments are:

* -h, --help: show a help message and exit.

An example of using this to create a YAML file from reporttsv.txt:

```
python makerepfromtsv.py reporttsv.txt reporttsv.yml
```

## Copyright

Copyright Peter Murphy 2018. 
