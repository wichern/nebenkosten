#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from typing import List, Tuple
import copy
from dataclasses import dataclass
import datetime

class MeterValueException(Exception):
    ''' Raised, when a meter value could not be read nor estimated. '''
    pass

@dataclass
class Date:
    ''' A date (Converts from string representation(s)) '''
    date: datetime.date

    def __init__(self, date: datetime.date):
        self.date = date

    @classmethod
    def from_str(cls, date_str: str):
        return Date(datetime.datetime.strptime(date_str, '%d.%m.%Y').date())

    def yesterday(self):
        return Date(self.date - datetime.timedelta(days=1))

    def tomorrow(self):
        return Date(self.date + datetime.timedelta(days=1))

    def __le__(self, other) -> bool:
        return self.date <= other.date

    def __lt__(self, other) -> bool:
        return self.date < other.date

    def __gt__(self, other) -> bool:
        return self.date > other.date

    def __str__(self):
        return self.date.strftime('%d.%m.%Y')

@dataclass
class DateRange:
    ''' A date range '''
    begin: Date
    end: Date

    def overlaps(self, other) -> bool:
        ''' Check if this date overlaps with given date '''
        return self.begin in other or self.end in other

    def __contains__(self, date: Date) -> bool:
        ''' Check if date is in range '''
        return date >= self.begin and date <= self.end

@dataclass
class Invoice:
    ''' Invoice structure '''
    type: str
    supplier: str
    invoice_number: str
    date: Date
    notes: str
    range: DateRange
    net: str
    amount: str
    tax: str

    def split(self, split_dates: List[Tuple['Invoice', int]]) -> List[Tuple['Invoice', int]]:
        ''' Split this invoices at dates, where the total people count over all tenants changed. '''

        ret: List[Tuple[Invoice, int]] = []

        people_count_before: int = 0
        invoice_begin: 'Date' = self.range.begin
        end_of_invoice: bool = False
        for split_date, people_count in split_dates:
            if split_date == self.range.begin:
                people_count_before = people_count
                continue

            if split_date >= self.range.begin:
                invoice_split = copy.deepcopy(self)
                invoice_split.range.begin = max(self.range.begin, invoice_begin)
                invoice_split.range.end = min(self.range.end, split_date.yesterday())
                ret.append((invoice_split, people_count_before))

                invoice_begin = split_date

            people_count_before = people_count

            if split_date > self.range.end:
                end_of_invoice = True
                break

        if not end_of_invoice:
            invoice_split = copy.deepcopy(self)
            invoice_split.range.begin = invoice_begin
            ret.append((invoice_split, people_count_before))
        
        return ret

@dataclass
class Appartement:
    ''' Appartement structure '''
    name: str
    size: str

@dataclass
class Tenant:
    ''' Tenant structure '''
    name: str
    appartement: str
    moving_in: Date
    moving_out: Date
    people: int

    def __contains__(self, date: Date) -> bool:
        ''' Check if date is in range '''
        if date < self.moving_in:
            return False
        if self.moving_out:
            return date <= self.moving_out
        return True

@dataclass
class MeterValue:
    ''' MeterValue structure '''
    name: str
    count: str
    date: Date
    notes: str

@dataclass
class Meter:
    ''' MeterValue structure '''
    name: str
    number: str
    unit: str

@dataclass
class BillCalculationItem:
    ''' BillCalculationItem structure '''
    appartement: str
    split: str
    unit: str
    invoice_type: str
    meter: str
