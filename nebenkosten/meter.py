#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import bisect
from typing import Tuple

from nebenkosten import MeterValue, Date, MeterValueException

class MeterManager(object):
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

    def set_row(self, date: Date, row: int):
        ''' Map date to row in cell '''
        self.values[date].row = row

    def get_row(self, date: Date) -> int:
        return self.values[date].row

    def get_surrounding_dates(self, date: Date) -> Tuple[Date, Date]:
        ''' Return the latest date before `date` and the first date after `date` '''

        idx = bisect.bisect_left(self._measured_dates, date)
        if 0 == idx:
            raise MeterValueException('Kein Zählerstand für "{0}" vor dem {1} gefunden.'.format(self._meter_name, str(date)))
        if idx >= len(self._measured_dates):
            raise MeterValueException('Kein Zählerstand für "{0}" nach dem {1} gefunden.'.format(self._meter_name, str(date)))
        return self._measured_dates[idx - 1], self._measured_dates[idx]

