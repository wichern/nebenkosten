#!/usr/bin/env python3
# -*- coding: utf-8 -*-

''' Package for nebenkosten helper '''

from typing import List, Tuple

# Move everything into this namespace
from nebenkosten.types import *
from nebenkosten.excel import *
from nebenkosten.meter import *

def get_people_count_changes(range, tenants) -> List[Tuple[Date, int]]:
    ''' Calculate a list of dates and their new people count '''

    ret = []

    date = range.begin
    idx: int = 0
    people_count_before: int = -1

    while date <= range.end:
        people_count: int = sum(t.people for t in tenants if date in t)

        if people_count != people_count_before:
            ret.append((date, people_count))

        people_count_before = people_count
        date = date.tomorrow()
        idx += 1

    return ret
