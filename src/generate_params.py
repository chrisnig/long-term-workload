import argparse
import util.generator


parser = argparse.ArgumentParser(description="Takes a directory which contains subdirectories with CDAT files. Reads "
                                             "the content of the files, replaces the physicians, duties, demand for "
                                             "duties, on and off requests with generated data and saves a new file "
                                             "with the respective data replaced.")
parser.add_argument("directory", help="Directory containing subdirectories with CDAT files")
parser.add_argument("output_directory", help="Directory to write result files to")
parser.add_argument("number_physicians", help="Number of physicians to generate", default=85, type=int)
parser.add_argument("number_duties", help="Number of duties to generate", default=6, type=int)
parser.add_argument("multiskill_probability", help="Probability that a physician has more than one skill", default=0.2,
                    type=float)
parser.add_argument("--filename", help="Filename of the CDAT files to parse", default="transformed.cdat")
parser.add_argument("--on_probability", help="Likelihood in percent that a user has a duty request on a date",
                    type=float, default=0.07)
parser.add_argument("--off_probability", help="Likelihood in percent that a user has a request for no duty on a date",
                    type=float, default=0.1)
parser.add_argument("--conflict_probability", help="Likelihood in percent that a duty request is in conflict with at "
                                                   "least one other request", type=float, default=None)
args = parser.parse_args()

request_generator = util.generator.RequestGenerator(args.on_probability, args.off_probability,
                                                    args.conflict_probability)
param_generator = util.generator.ParameterGenerator(args.number_physicians, args.multiskill_probability,
                                                    args.number_duties, request_generator)
generator = util.generator.DirectoryGenerator(param_generator, args.filename)
generator.run_in_directory(args.directory, args.output_directory)
