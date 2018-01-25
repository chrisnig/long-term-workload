import argparse
import os
import sys

parser = argparse.ArgumentParser()
parser.add_argument("directory", help="Directory containing subdirectories with .out.log files.")
args = parser.parse_args()

main_dir = args.directory

modes = ["unfair", "fair-linear", "fair-nonlinear"]
results = {mode: {"min": sys.maxsize, "max": 0} for mode in modes}

for directory in filter(lambda x: os.path.isdir(os.path.join(main_dir, x)),
                        sorted(os.listdir(main_dir))):
    for mode in modes:
        with open(os.path.join(main_dir, directory, mode + ".out.log"), "r") as logfile:
            for line in logfile:
                if line.startswith("CMPL: Time used for solving the model:"):
                    time = int(line.split()[7])
                    if time > results[mode]["max"]:
                        results[mode]["max"] = time
                    if time < results[mode]["min"]:
                        results[mode]["min"] = time
                    break

for mode in modes:
    print("{} max: {}s, min: {}s".format(mode, results[mode]["max"], results[mode]["min"]))
