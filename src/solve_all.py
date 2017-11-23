import argparse
import util.solver_runner

parser = argparse.ArgumentParser()
parser.add_argument("directory", help="The directory which contains all subdirectories with solver data")
parser.add_argument("outdir", help="The directory in which to store the output. If not given, output is stored in "
                                   "the same directory as input data.")
parser.add_argument("--solver", help="Solver parameter to pass to CMPL.", default="cplex")

args = parser.parse_args()

for runner_class in [util.solver_runner.UnfairSolverRunner,
                     util.solver_runner.FairSolverRunner,
                     util.solver_runner.FairNonLinearSolverRunner]:
    runner = runner_class(args.solver)
    runner.run_in_directory(args.directory, args.outdir)
