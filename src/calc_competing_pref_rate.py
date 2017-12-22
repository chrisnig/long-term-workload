import argparse
import os
import util.io

parser = argparse.ArgumentParser(description="Calculates preference and conflict rate for a given set of data")
parser.add_argument("directory", help="The directory containing subdirectories with CDAT files")
parser.add_argument("--filename", help="Filename of the CDAT files to parse", default="transformed.cdat")
args = parser.parse_args()

total_pref = 0
conflict_pref = 0
total_days = 0

parent_dir = os.listdir(args.directory)

for directory in filter(lambda x: os.path.isdir(x),
                        sorted(os.path.join(args.directory, y) for y in os.listdir(args.directory))):
    data = util.io.CmplCdatReader.read(os.path.join(directory, args.filename))
    total_days += (len(data.get_set("J")) * len(data.get_set("W")) * len(data.get_set("D"))) - len(
        data.get_values("D_off"))
    for w in data.get_set("W"):
        for d in data.get_set("D"):
            for i in data.get_set("I"):
                prefs = data.count_if_exists("g_req_on", (None, i, w, d))
                total_pref += prefs
                if prefs > 1:
                    conflict_pref += prefs

print("total preferences: {:18}".format(total_pref))
print("conflicting preferences: {:12}".format(conflict_pref))
print("preference rate: {:19.2f}%".format((total_pref / total_days) * 100))
print("conflict percentage: {:15.2f}%".format((conflict_pref / total_pref) * 100))
print("total days: {:25}".format(total_days))
