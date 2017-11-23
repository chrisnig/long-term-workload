import math
from typing import Tuple


class DaysHelper:
    @staticmethod
    def get_next_day(w: int, d: int) -> Tuple[int, int]:
        if d < 7:
            return w, d + 1
        else:
            return w + 1, 1

    @staticmethod
    def get_previous_day(w: int, d: int) -> Tuple[int, int]:
        if d > 1:
            return w, d - 1
        else:
            return w - 1, 7

    @staticmethod
    def int_to_day_and_week(d: int) -> Tuple[int, int]:
        week = int(math.ceil(d / 7))
        day = d - (week * 7) + 7
        return week, day
