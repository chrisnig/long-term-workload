import argparse
import util.solver_runner

parser = argparse.ArgumentParser()
parser.add_argument("--unfair", action="store_const", const=True, default=False,
                    help="Run unfair models instead of fair models")
parser.add_argument("directory", help="The directory which contains all subdirectories with solver data")
parser.add_argument("outdir", help="The directory in which to store the output. If not given, output is stored in "
                                   "the same directory as input data.")
parser.add_argument("--solver", help="Solver parameter to pass to CMPL.", default="cplex")

args = parser.parse_args()
unfair = args.unfair

if unfair:
    runners = [util.solver_runner.UnfairUnequalSolverRunner,
               util.solver_runner.UnfairEqualSolverRunner]
else:
    runners = [util.solver_runner.FairNonLinearUnequalSolverRunner,
               util.solver_runner.FairNonLinearEqualSolverRunner]

for runner_class in runners:
    runner = runner_class(args.solver)
    runner.run_in_directory(args.directory, args.outdir)
