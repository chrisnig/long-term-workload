from abc import abstractmethod
from math import ceil
from shutil import copyfile
from tempfile import TemporaryDirectory
import os
import subprocess
import util.io
import util.data


class SolverRunner:
    cdat_filename = "transformed.cdat"
    cmpl_prefix = "ltw-"

    def __init__(self, solver: str):
        self.l_hat = None
        self.solver = solver

    @abstractmethod
    def get_file_prefix(self):
        raise NotImplementedError()

    def pre_solve(self, directory: str, outdir: str):
        pass

    def post_solve(self, directory: str, outdir: str):
        pass

    def get_data_filename(self, directory: str) -> str:
        return os.path.join(directory, self.cdat_filename)

    def get_solution_filename(self, directory: str) -> str:
        return os.path.join(directory, self.get_file_prefix() + ".sol")

    def get_cmpl_filename(self, directory: str) -> str:
        return os.path.join(directory, self.cmpl_prefix + self.get_file_prefix() + ".cmpl")

    def run_in_directory(self, d: str, outdir: str):
        if not os.path.exists(outdir):
            os.mkdir(outdir)

        for directory in filter(lambda x: os.path.isdir(os.path.join(d, x)), sorted(os.listdir(d))):
            fulldir = os.path.join(d, directory)
            fulloutdir = os.path.join(outdir, directory)

            if not os.path.exists(fulloutdir):
                os.mkdir(fulloutdir)

            with TemporaryDirectory() as working_directory:
                orig_cmpl_file = self.get_cmpl_filename(d)
                cmpl_file = os.path.join(working_directory, os.path.basename(orig_cmpl_file))
                copyfile(orig_cmpl_file, cmpl_file)

                orig_data_file = self.get_data_filename(fulldir)
                data_file = os.path.join(working_directory, os.path.basename(orig_data_file))
                copyfile(orig_data_file, data_file)

                fileoutprefix = os.path.join(fulloutdir, self.get_file_prefix())
                self.pre_solve(working_directory, fulloutdir)
                outfilename = fileoutprefix + ".out.log"
                errorfilename = fileoutprefix + ".err.log"
                cdatfilename = fileoutprefix + ".cdat"
                copyfile(data_file, cdatfilename)
                with open(outfilename, "w") as outfile, open(errorfilename, "w") as errfile:
                    subprocess.check_call(
                        [
                            "cmpl",
                            cmpl_file,
                            "-solutionAscii",
                            self.get_solution_filename(fulloutdir),
                            "-solver",
                            self.solver
                        ], stdout=outfile, stderr=errfile)
                self.check_file_empty(errorfilename)
                self.post_solve(working_directory, fulloutdir)

    @staticmethod
    def check_file_empty(file: str):
        if os.path.getsize(file):
            raise RuntimeError("Solver failed with error!")

    def replace_param_in_cdat(self, param: str, values: dict, directory: str):
        data_file = os.path.join(directory, self.cdat_filename)
        tmpfilename = data_file + "_tmp"
        with open(data_file, "r") as infile, open(tmpfilename, "w") as outfile:
            searching = True
            skipping = False
            for line in infile:
                if not skipping and (not searching or not line.startswith("%" + param + "[")):
                    outfile.write(line)
                    continue

                if not skipping and line.startswith("%" + param + "["):
                    outfile.write(line)
                    searching = False
                    for key, value in values.items():
                        if isinstance(key, str):
                            # output string keys as they are. note that this check is required because strings are
                            # iterable, so we can't just check for iterable
                            key_out = key
                        else:
                            try:
                                # tuple keys
                                key_out = "\t".join(str(entry) for entry in key)
                            except TypeError:
                                # not a tuple, convert directly to str
                                key_out = str(key)
                        outfile.write("\t{}\t{}\n".format(key_out, value))
                    if not values:
                        # find the parameter dimensions and the default value to write at least one line
                        # if we don't do this there will be an empty data element and the cdat file is not valid
                        dimensions = len(line.split("[")[1].split("]")[0].split(","))
                        default = line.split("=")[1].split("indices")[0].strip()
                        outfile.write("\t0" * dimensions + "\t" + default + "\n")
                    skipping = True
                    continue

                if skipping and not line.startswith(">"):
                    continue

                if skipping and line.startswith(">"):
                    outfile.write(line)
                    skipping = False
                    continue

                raise RuntimeError("We should never get here! Problem was with the following line: {}".format(line))

            if searching:
                raise RuntimeError("Couldn't find parameter {} in file {}!".format(param, data_file))

        os.unlink(data_file)
        os.rename(tmpfilename, data_file)


class FairNonLinearEqualSolverRunner(SolverRunner):
    gamma_1 = 0.8
    gamma_2 = 0.8
    alpha_2 = 10
    alpha_3 = 1

    def __init__(self, solver: str):
        super().__init__(solver)
        self.s_hat = None
        self.l_hat = None
        self.parameters = None

    def get_file_prefix(self):
        return "equal"

    def post_solve(self, directory: str, outdir: str):
        self.s_hat = dict()
        self.l_hat = dict()
        with open(os.path.join(outdir, self.get_file_prefix() + ".sol"), "r") as solfile:
            for line in solfile:
                if line.startswith("s["):
                    phys = line.split("[")[1].split("]")[0]
                    value = line.split()[2]
                    self.s_hat[phys] = value

                if line.startswith("l["):
                    phys = line.split("[")[1].split("]")[0]
                    value = line.split()[2]
                    self.l_hat[phys] = value

    def pre_solve(self, directory: str, outdir: str):
        self.parameters = util.io.CmplCdatReader.read(self.get_data_filename(directory))

        if self.s_hat is None:
            self.s_hat = dict()
            for phys in self.parameters.get_set("J"):
                self.s_hat[phys] = 1

        self.write_s_hat(directory)

        if self.l_hat is None:
            self.l_hat = dict()
            for phys in self.parameters.get_set("J"):
                self.l_hat[phys] = 0

        self.replace_param_in_cdat("l_hat", self.l_hat, directory)

        c = dict()

        weeks = len(self.parameters.get_set("W"))
        days = len(self.parameters.get_set("D"))

        for phys in self.parameters.get_set("J"):
            req_on = self.parameters.count_if_exists("g_req_on", (phys, None, None, None))
            req_off = self.parameters.count_if_exists("g_req_off", (phys, None, None))

            total_requests = req_on + req_off

            for violations in range(total_requests + 1):
                current_sat = (total_requests - violations) / (weeks * days)
                vio_sat = self.gamma_1 * current_sat + (1 - self.gamma_1) * self.parameters.get("s_hat", (phys,))
                c[(phys, violations)] = self.alpha_2 * violations * (2 - vio_sat)

        self.replace_param_in_cdat("c", c, directory)

        c_work = dict()

        for phys in self.parameters.get_set("J"):
            max_assignments = int(ceil(weeks * days / 2))
            for assignments in range(max_assignments + 1):
                l_j = (self.gamma_2 * assignments) / (weeks * days) + (1 - self.gamma_2) * \
                      self.parameters.get("l_hat", (phys,))
                c_work[(phys, assignments)] = self.alpha_3 * assignments * l_j
        self.replace_param_in_cdat("c_work", c_work, directory)

    def write_s_hat(self, directory: str):
        self.replace_param_in_cdat("s_hat", self.s_hat, directory)


class FairNonLinearUnequalSolverRunner(FairNonLinearEqualSolverRunner):
    def get_file_prefix(self):
        return "unequal"


class UnfairEqualSolverRunner(FairNonLinearEqualSolverRunner):
    def get_file_prefix(self):
        return "equal-unfair"


class UnfairUnequalSolverRunner(FairNonLinearEqualSolverRunner):
    def get_file_prefix(self):
        return "unequal-unfair"
