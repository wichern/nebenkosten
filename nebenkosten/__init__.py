#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from typing import List, Tuple

from nebenkosten.types import Tenant, DateRange, Date
from nebenkosten.excel import InputSheet

def get_people_count_change_dates(tenants: List[Tenant], range: DateRange) -> List[Tuple[Date, int]]:
    ret: List[Date, int] = []
    date: Date = range.begin
    idx: int = 0
    people_count_before: int = -1
    while date <= range.end:
        people_count: int = sum([t.people for t in tenants if date in t])

        if people_count != people_count_before:
            ret.append((date, people_count))
        
        people_count_before = people_count
        date = date.tomorrow()
        idx += 1

    return ret
