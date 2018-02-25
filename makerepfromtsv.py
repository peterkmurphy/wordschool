# makerepfromtsv.py. Written by Peter Murphy
# Goal: To turn files consisting of lines of the following TSV form:
#
# ID[TAB]NAME[TAB]START_DATE[TAB]END_DATE
#
# Into YAML consisting of blank student reports documents of the following form:
#
# ---
# - {name: "NAME", id: "ID", sd: "START_DATE", ed: "END_DATE"}
# - comment:
# - {start: 1, end: 10}
# ...
#
# Where ID, NAME, START_DATE and END_DATE are the id, name, starting date and
# endind date of students, and [TAB] representing tab characters.

import yaml
import argparse
parser = argparse.ArgumentParser(
    description='Generate blank student report YAML from tab separated data')
parser.add_argument('input_file',
    help='A tab separated value file with student data')
parser.add_argument('output_file',
    help='A YAML file representing student data (with blank class scores)')
args = parser.parse_args()
with open(args.input_file, "r") as ins:
    with open(args.output_file, "w") as ous:
        for line in ins:
            tabsplit = line.split("\t")
            ous.write("---\n")
            NAME = tabsplit[1]
            ID = tabsplit[0]
            START_DATE = tabsplit[2]
            tabsplit[3] = tabsplit[3].split("\n")[0]
            END_DATE = tabsplit[3]
            mapinternal = \
                'name: "{1}", id: "{0}", sd: "{2}", ed: "{3}"'.format(*tabsplit)
            ous.write("- {" + mapinternal + "}\n")
            ous.write("- comment:\n")
            ous.write("- {start: 1, end: 10}\n")
            ous.write("...\n")

# To check that the output works.

reportdata = list(yaml.load_all(open(args.output_file, 'r')))
