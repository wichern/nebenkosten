#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import bisect
from typing import Tuple

from nebenkosten import MeterValue, DateRange, Date, MeterValueException

class MeterManager(object):
    ''' MeterValue helper 
    
    This class filter out all meter values of name given in the constructor.
    It can then calculate 'Consumption' values for dates.

    All calculated meter values are stored as well.

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

    def add_meter_value(self, date: Date):
        ''' Make sure that a meter value for given date exists. '''

        # When we have a measured meter value for this date we can return it immediately.
        if date in self.values:
            return self.values[date]

        self.values[date] = MeterValue(
            self._meter_name,
            None,  # empty count (formula added later).
            date,
            'Berechnet'
        )

    def get_count_formula(self, date) -> str:
        ''' Create a formula, calculating the meter value in the result sheet. '''

        before, after = self.__get_surrounding_dates(date)
        before_row = self.get_row(before)
        after_row = self.get_row(after)
        
        # Calculate the total consumption between measured values.
        delta_value = f'C{after_row}-C{before_row}'

        # Calculate the days between measured values.
        delta_days_total = f'_xlfn.days(B{after_row},B{before_row})'

        # Calculate the days between before date and date.
        delta_days_new = f'_xlfn.days(B{self.get_row(date)},B{before_row})'

        return f'=C{before_row}+({delta_value})/{delta_days_total}*{delta_days_new}'

    def set_row(self, date: Date, row: int):
        ''' Map date to row in cell '''
        self.values[date].row = row

    def get_row(self, date: Date) -> int:
        return self.values[date].row

    def __get_surrounding_dates(self, date: Date) -> Tuple[Date, Date]:
        ''' Return the latest date before `date` and the first date after `date` '''

        idx = bisect.bisect_left(self._measured_dates, date)
        if 0 == idx:
            raise MeterValueException('Kein Zählerstand für "{0}" vor dem {1} gefunden.'.format(self._meter_name, str(date)))
        if idx >= len(self._measured_dates):
            raise MeterValueException('Kein Zählerstand für "{0}" nach dem {1} gefunden.'.format(self._meter_name, str(date)))
        return self._measured_dates[idx - 1], self._measured_dates[idx]

