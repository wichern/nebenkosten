#!/usr/bin/env python3
# -*- coding: utf-8 -*-

''' Managing meter values '''

import bisect
from typing import Tuple

from nebenkosten import Date, MeterValueException

# pylint: disable=too-few-public-methods
class MeterManager:
    ''' MeterValue helper

    This class filter out all meter values of name given in the constructor.

    The most important part of this class is to replace dates with cell values in the formuale.
    Unfortunately, we cannot know the cell values before we write them into the result sheet.
    '''

    def __init__(self, meter_values, meter_name):
        # Filter out meter values with different meter name.
        # Turn meter values into dictionary. The meter value date is its key.
        self.values = { mv.date: mv for mv in meter_values if mv.name == meter_name }

        # For looking up surrounding dates, we need a list of all measured dates.
        self._measured_dates = sorted(list(self.values.keys()))

        self._meter_name = meter_name

    def get_surrounding_dates(self, date: Date) -> Tuple[Date, Date]:
        ''' Return the latest date before `date` and the first date after `date` '''

        idx = bisect.bisect_left(self._measured_dates, date)
        if 0 == idx:
            raise MeterValueException(
                f'Kein Zählerstand für "{self._meter_name}" vor dem {date} gefunden.')
        if idx >= len(self._measured_dates):
            raise MeterValueException(
                f'Kein Zählerstand für "{self._meter_name}" nach dem {date} gefunden.')
        return self._measured_dates[idx - 1], self._measured_dates[idx]
