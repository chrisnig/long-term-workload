from collections import defaultdict
import os
import random
from util.days_helper import DaysHelper
import util.data
import util.io


class DataGenerator:
    def generate(self, data: util.data.DataSet) -> None:
        raise NotImplementedError


class RequestGenerator(DataGenerator):
    def __init__(self, on_probability: float, off_probability: float, conflict_probability: float =None) -> None:
        self.request_on_probability = on_probability
        self.request_off_probability = off_probability
        self.request_conflict_probability = conflict_probability

    def generate(self, data: util.data.DataSet) -> None:
        self.clear_request_vars(data)
        physicians = data.get_set("J")
        total_days = int(data.get_scalar("W_max")) * 7
        for week, day in [DaysHelper.int_to_day_and_week(x) for x in range(1, total_days+1)]:
            for phys in physicians:
                if not data.get("D_off", (phys, week, day)):
                    rand = random.random()
                    if rand < self.request_on_probability:
                        if not data.get_matching_keys("g_req_on",
                                                      (phys, None, *DaysHelper.get_previous_day(week, day))):
                            possible_duties = tuple(x[1] for x
                                                    in data.get_matching_keys("E_pos", (phys, None, week, day)))
                            if not possible_duties:
                                continue
                            duty = random.choice(possible_duties)
                            data.set("g_req_on", (phys, duty, week, day), 1)
                    elif rand < self.request_off_probability + self.request_on_probability:
                        data.set("g_req_off", (phys, week, day), 1)

        if self.request_conflict_probability is not None:
            conflicts_cnt = 0
            conflicts = defaultdict(lambda: set())
            total_requests = 0
            for week, day in [DaysHelper.int_to_day_and_week(x) for x in range(1, total_days + 1)]:
                for duty in data.get_set("I"):
                    requests_on_day = data.count_if_exists("g_req_on", (None, duty, week, day))
                    if requests_on_day > 1:
                        conflicts_cnt += requests_on_day
                        conflicts[(duty, week, day)].update(
                            set(data.get_matching_keys("g_req_on", (None, duty, week, day))))
                    total_requests += requests_on_day

            conflict_prob = conflicts_cnt / total_requests
            if conflict_prob > self.request_conflict_probability:
                while conflict_prob > self.request_conflict_probability and conflicts_cnt:
                    conflict_key, conflict_set = random.choice(list(conflicts.items()))
                    conflict_value = random.choice(list(conflict_set))
                    conflict_phys, conflict_duty, conflict_week, conflict_day = conflict_value

                    data.set("g_req_on", conflict_value, 0)
                    conflicts_cnt -= 1
                    conflicts[conflict_key].remove(conflict_value)
                    total_requests -= 1

                    # try to find alternative day for this request
                    random_days = [DaysHelper.int_to_day_and_week(x) for x in range(1, total_days + 1)]
                    random.shuffle(random_days)
                    for week, day in random_days:
                        if (week, day) == (conflict_week, conflict_day):
                            # do not try to assign the same day again
                            continue
                        if not data.get("E_pos", (conflict_phys, conflict_duty, week, day)):
                            # do not assign if the physician is not qualified for the duty on this date
                            continue
                        if data.count_if_exists("g_req_on", (conflict_phys, None, week, day)):
                            # do not assign if the physician already has a preference for a duty on this date
                            continue
                        if data.get("g_req_off", (conflict_phys, week, day)):
                            # do not assign if the physician already has a preference for no duty on this date
                            continue
                        if data.count_if_exists("g_req_on", (None, conflict_duty, week, day)):
                            # do not assign if there is already a preference by someone else for this duty on this date
                            continue
                        if data.get("D_off", (conflict_phys, week, day)):
                            # do not assign if physician is absent on this date
                            continue

                        data.set("g_req_on", (conflict_phys, conflict_duty, week, day), 1)
                        total_requests += 1
                        break

                    if len(conflicts[conflict_key]) == 1:
                        conflicts_cnt -= 1
                        del conflicts[conflict_key]

                    conflict_prob = conflicts_cnt / total_requests
            elif conflict_prob < self.request_conflict_probability:
                non_conflicts = list(set(data.get_values("g_req_on").keys()).difference(
                    set().union(*conflicts.values())))
                while conflict_prob < self.request_conflict_probability and non_conflicts:
                    my_non_conflict = random.choice(non_conflicts)
                    non_conflict_phys, non_conflict_duty, non_conflict_week, non_conflict_day = my_non_conflict

                    available_phys = [x[0] for x in data.get_matching_keys("E_pos", (None,
                                                                                     non_conflict_duty,
                                                                                     non_conflict_week,
                                                                                     non_conflict_day))]

                    if len(available_phys) == 1:
                        # only one physician is qualified for this request, so we can never create a conflict here
                        non_conflicts.remove(my_non_conflict)
                        continue

                    random.shuffle(available_phys)

                    # try to find a physician who can request a conflicting assignment
                    for candidate in available_phys:
                        if candidate == non_conflict_phys:
                            # do not assign the request to the same physician twice
                            continue
                        if data.get("g_req_off", (candidate, non_conflict_week, non_conflict_day)):
                            # do not assign if the physician already has a no-duty preference on this date
                            continue
                        if data.count_if_exists("g_req_on", (candidate, None, non_conflict_week, non_conflict_day)):
                            # do not assign if the physician already has a duty preference on this date
                            continue
                        if data.count_if_exists("D_off", (candidate, non_conflict_week, non_conflict_day)):
                            # do not assign a physician who is absent
                            continue

                        data.set("g_req_on", (candidate, non_conflict_duty, non_conflict_week, non_conflict_day), 1)
                        total_requests += 1
                        non_conflicts.remove(my_non_conflict)
                        conflicts_cnt += 2

                        # try to delete another request to keep total number of requests constant
                        if non_conflicts:
                            drop_assignment = random.choice(non_conflicts)
                            non_conflicts.remove(drop_assignment)
                            data.set("g_req_on", drop_assignment, 0)
                            total_requests -= 1

                        break

                    conflict_prob = conflicts_cnt / total_requests

    @staticmethod
    def clear_request_vars(data: util.data.DataSet) -> None:
        req_on_dimensions = data.get_dimensions("g_req_on")
        req_off_dimensions = data.get_dimensions("g_req_off")
        data.remove("g_req_on")
        data.remove("g_req_off")
        data.create("g_req_on", 0, req_on_dimensions)
        data.create("g_req_off", 0, req_off_dimensions)


class DirectoryGenerator:
    def __init__(self, base_generator: DataGenerator, cdat_filename: str) -> None:
        self.generator = base_generator
        self.cdat_filename = cdat_filename

    def run_in_directory(self, indir: str, outdir: str) -> None:
        if not os.path.exists(outdir):
            os.mkdir(outdir)

        for directory in filter(lambda x: os.path.isdir(os.path.join(indir, x)), sorted(os.listdir(indir))):
            fulldir = os.path.join(indir, directory)
            fulloutdir = os.path.join(outdir, directory)

            if not os.path.exists(fulloutdir):
                os.mkdir(fulloutdir)

            data = util.io.CmplCdatReader.read(os.path.join(fulldir, self.cdat_filename))
            self.generator.generate(data)
            util.io.CmplCdatWriter.write(data, os.path.join(fulloutdir, self.cdat_filename))


class ParameterGenerator(DataGenerator):
    def __init__(self, number_physicians: int, multiskill_probability: float, number_duties: int,
                 request_generator: RequestGenerator):
        self.request_generator = request_generator
        self.number_physicians = number_physicians
        self.multiskill_probability = multiskill_probability
        self.number_duties = number_duties
        self.duties = None
        self.physicians = None
        self.skills = None

    def initialize(self):
        if self.duties is not None:
            return

        self.duties = set(range(1, self.number_duties + 1))
        self.physicians = set(range(1, self.number_physicians + 1))
        self.skills = dict()

        if self.duties:
            for phys_cnt, physician in enumerate(self.physicians):
                current_duty = (phys_cnt % self.number_duties) + 1
                self.skills.setdefault(physician, [])
                self.skills[physician].append(current_duty)

                if self.number_duties > 1:
                    rand = random.random()
                    if rand <= self.multiskill_probability:
                        # we have a multi-skilled physician, add another duty
                        additional_duty = random.choice(tuple(self.duties))
                        while additional_duty == current_duty:
                            # do not add the same skill twice, get a different duty
                            additional_duty = random.choice(tuple(self.duties))
                        self.skills[physician].append(additional_duty)

    def generate(self, data: util.data.DataSet) -> None:
        self.initialize()
        self.clear_data_vars(data)
        data.create_set("I", self.duties)
        data.create_set("J", self.physicians)
        if self.duties:
            for duty in self.duties:
                for day in data.get_set("D"):
                    data.set("d_bar_duty", (duty, day), 1)

            for physician in self.skills:
                for duty in self.skills[physician]:
                    self._add_skill(data, physician, duty)

            if self.request_generator:
                self.request_generator.generate(data)

    @staticmethod
    def _add_skill(data, physician, duty):
        for week in range(1, data.get_scalar("W_max") + 1):
            for day in data.get_set("D"):
                data.set("E_pos", (physician, duty, week, day), 1)

    @staticmethod
    def clear_data_vars(data: util.data.DataSet) -> None:
        data.remove("I")
        data.remove("J")
        d_bar_duty_dimensions = data.get_dimensions("d_bar_duty")
        data.remove("d_bar_duty")
        data.create("d_bar_duty", 0, d_bar_duty_dimensions)
        e_pos_dimensions = data.get_dimensions("E_pos")
        data.remove("E_pos")
        data.create("E_pos", 0, e_pos_dimensions)
        d_off_dimensions = data.get_dimensions("D_off")
        data.remove("D_off")
        data.create("D_off", 0, d_off_dimensions)
        s_hat_dimensions = data.get_dimensions("s_hat")
        data.remove("s_hat")
        data.create("s_hat", 0, s_hat_dimensions)
