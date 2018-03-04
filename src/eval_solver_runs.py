import argparse
import util.satisfaction_evaluator

parser = argparse.ArgumentParser()
parser.add_argument("--unfair", action="store_const", const=True, default=False,
                    help="Evaluate unfair models instead of fair models")
parser.add_argument("directory", help="Directory containing subdirectories with CDAT files.")
parser.add_argument("solution_directory", help="Directory containing subdirectories with SOL files.")
parser.add_argument("outfile", help="XLSX file to write evaluation data to.")
args = parser.parse_args()

data_dir = args.directory
solution_dir = args.solution_directory
outfile = args.outfile
include_unfair = args.unfair

evaluator = util.satisfaction_evaluator.SatisfactionEvaluator(data_dir, solution_dir, include_unfair)
evaluator.evaluate_to_file(outfile)
