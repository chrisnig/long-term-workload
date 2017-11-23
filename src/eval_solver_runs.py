import argparse
import util.satisfaction_evaluator

parser = argparse.ArgumentParser()
parser.add_argument("directory", help="Directory containing subdirectories with CDAT files.")
parser.add_argument("solution_directory", help="Directory containing subdirectories with SOL files.")
parser.add_argument("outfile", help="XLSX file to write evaluation data to.")
args = parser.parse_args()

data_dir = args.directory
solution_dir = args.solution_directory
outfile = args.outfile

evaluator = util.satisfaction_evaluator.SatisfactionEvaluator(data_dir, solution_dir)
evaluator.evaluate_to_file(outfile)
