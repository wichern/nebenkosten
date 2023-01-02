#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from typing import List

import openpyxl

from nebenkosten.types import Invoice, Appartement, Tenant, MeterValue, Meter, Date, DateRange, BillCalculationItem
from nebenkosten.types import MeterValueException

class InputSheet(object):
    '''
    Parse the input sheet.
    '''

    def __init__(self, path, appartement_name: str, range: DateRange):
        self._path = path
        self._appartement_name = appartement_name
        self._range = range

        self.invoices = []
        self.appartements = []
        self.appartement = None
        self.tenants = []
        self.tenant = None
        self.meter_values = []
        self.meter = []
        self.bcis = []

    def __enter__(self):
        workbook = openpyxl.load_workbook(filename=str(self._path))

        for row in self.__get_rows(workbook, 'Rechnungen'):
            invoice_range = DateRange(Date.from_str(row[5]), Date.from_str(row[6]))
            if invoice_range.overlaps(self._range):
                invoice = Invoice(
                    row[0],
                    row[1],
                    row[2],
                    Date.from_str(row[3]),
                    row[4],
                    invoice_range,
                    row[7],
                    row[8],
                    row[9])
                self.invoices.append(invoice)

        for row in self.__get_rows(workbook, 'Wohnungen'):
            self.appartements.append(Appartement(row[0], row[1]))
        self.appartement = next(a for a in self.appartements if a.name == self._appartement_name)

        for row in self.__get_rows(workbook, 'Mieter'):
            # For 'moving out' we also accept None
            moving_out = None
            if row[3]:
                moving_out = Date.from_str(row[3])
            self.tenants.append(Tenant(row[0], row[1], Date.from_str(row[2]), moving_out, int(row[4])))
        self.tenant = next(t for t in self.tenants if self._range.begin in t and self._range.end in t and t.appartement == self._appartement_name)

        for row in self.__get_rows(workbook, 'Zählerstände'):
            self.meter_values.append(MeterValue(row[0], row[1], Date.from_str(row[2]), row[3]))

        for row in self.__get_rows(workbook, 'Zähler'):
            self.meter.append(Meter(row[0], row[1], row[2]))

        for row in self.__get_rows(workbook, 'Abrechnungseinstellungen'):
            bci = BillCalculationItem(row[0], row[1], row[2], row[3], row[4])
            # Skip BCIs that are not relevant for this appartement
            if bci.appartement == self._appartement_name:
                self.bcis.append(bci)

        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def __get_rows(self, workbook, sheet_name: str):
        sheet = workbook[sheet_name]
        # We need to filter out the rows without content, because we will receive those as well.
        return filter(lambda row: row[0], sheet.iter_rows(min_row=2, values_only=True))

    def get_meter(self, meter_name) -> Meter:
        return next(m for m in self.meter if m.name == meter_name)

class ResultSheet(object):
    '''
    Write the result Excel sheet.
    '''

    template_filepath = 'example-bill.xlsx'

    def __init__(self, filepath: str, input_sheet: InputSheet, bill_range: DateRange):
        self._filepath = filepath
        self._input_sheet = input_sheet
        self._bill_range = bill_range
        self._bill_item_row = 2
        self._meter_value_row = 2

    def add_bill_item(self, invoice: Invoice, bci: BillCalculationItem, total_people_count: int = 0, comment: str = None, consumption: str = None):
        row = self._bill_item_row

        self.__write_date(max(self._bill_range.begin, invoice.range.begin), sheet=self._sheet_details, row=row, column=1)
        self.__write_date(min(self._bill_range.end, invoice.range.end), sheet=self._sheet_details,     row=row, column=2)
        self.__write_number('=_xlfn.days(B{0}, A{0})'.format(row), sheet=self._sheet_details, row=row, column=3, precision=0)
        self.__write(invoice.type, sheet=self._sheet_details, row=row, column=4)
        self.__write(invoice.notes, sheet=self._sheet_details, row=row, column=5)
        self.__write(bci.meter, sheet=self._sheet_details, row=row, column=6)
        self.__write_currency(invoice.net, sheet=self._sheet_details, row=row, column=7)
        self.__write_number(invoice.amount, sheet=self._sheet_details, row=row, column=8)
        self.__write_percentage(invoice.tax, sheet=self._sheet_details, row=row, column=9)
        self.__write_currency('=G{0}*H{0}*(1+I{0})'.format(row), sheet=self._sheet_details, row=row, column=10)
        self.__write_number(f'=_xlfn.days("{invoice.range.end}", "{invoice.range.begin}")', sheet=self._sheet_details, row=row, column=11, unit='Tage', precision=0)
        self.__write(bci.split, sheet=self._sheet_details, row=row, column=12)
        
        if bci.split == 'Pro Wohnung':
            self.__write_percentage(f'=1/{len(self._input_sheet.appartements)}', sheet=self._sheet_details, row=row, column=13)
        elif bci.split == 'Pro Person':
            self.__write_percentage(f'={self._input_sheet.tenant.people}/{total_people_count}', sheet=self._sheet_details, row=row, column=13)
        elif bci.split == 'Pro Quadratmeter':
            self.__write_percentage(f'={self._input_sheet.appartement.size}/{sum([a.size for a in self._input_sheet.appartements])}', sheet=self._sheet_details, row=row, column=13)
        elif bci.split == 'Nach Verbrauch':
            self.__write_number(consumption, sheet=self._sheet_details, row=row, column=13, unit=bci.unit)
        elif bci.split == 'Hälfte':
            self.__write_percentage('=1/2', sheet=self._sheet_details, row=row, column=13)
        elif bci.split == 'Drittel':
            self.__write_percentage('=1/3', sheet=self._sheet_details, row=row, column=13)
        elif bci.split == 'Viertel':
            self.__write_percentage('=1/4', sheet=self._sheet_details, row=row, column=13)
        elif bci.split == 'Komplett':
            self.__write_percentage('1', sheet=self._sheet_details, row=row, column=13)
        else:
            raise InvalidCellValue('Unknown bill split: "%s"' % split)

        if bci.split == 'Nach Verbrauch':
            self.__write_currency('=G{0}*(1+I{0})*M{0}'.format(row), sheet=self._sheet_details, row=row, column=14)
        else:
            self.__write_currency('=J{0}/K{0}*C{0}*M{0}'.format(row), sheet=self._sheet_details, row=row, column=14)

        if comment:
            self.__write(comment, sheet=self._sheet_details, row=row, column=15)
        
        self._bill_item_row += 1

    def add_meter_value(self, meter_value: MeterValue):
        row = self._meter_value_row
        self.__write(meter_value.name, sheet=self._sheet_meter_values, row=row, column=1)
        self.__write_date(meter_value.date, sheet=self._sheet_meter_values, row=row, column=2)
        if meter_value.count:
            meter = self._input_sheet.get_meter(meter_value.name)
            self.__write_number(meter_value.count, sheet=self._sheet_meter_values, row=row, column=3, unit=meter.unit)
            self.__write('Gemessen', sheet=self._sheet_meter_values, row=row, column=4)
        self._meter_value_row += 1
        return row

    def update_meter_value_formula(self, meter_value: MeterValue, formula: str):
        meter = self._input_sheet.get_meter(meter_value.name)
        self.__write_number(formula, sheet=self._sheet_meter_values, row=meter_value.row, column=3, unit=meter.unit)
        self.__write('Berechnet', sheet=self._sheet_meter_values, row=meter_value.row, column=4)

    def __write_date(self, date: Date, sheet, row: int, column: int):
        self.__write(date.date, sheet=sheet, row=row, column=column, number_format='DD.MM.YY')

    def __write_number(self, number: str, sheet, row: int, column: int, unit: str = None, precision: int = 2):
        precision_format = '0'
        if precision > 0:
            precision_format += '.' + (precision * '0')

        if unit:
            self.__write(number, sheet=sheet, row=row, column=column, number_format=precision_format + '" ' + unit + '"')
        else:
            self.__write(number, sheet=sheet, row=row, column=column, number_format=precision_format)

    def __write_currency(self, number: str, sheet, row: int, column: int):
        self.__write(number, sheet=sheet, row=row, column=column, number_format='0.00" "€')

    def __write_percentage(self, number: str, sheet, row: int, column: int):
        self.__write(number, sheet=sheet, row=row, column=column, number_format='0.00" "%')

    def __write(self, content, sheet, row: int, column: int, number_format = None):
        cell = sheet.cell(row=row, column=column)
        cell.style = 'content'
        cell.value = content
        if number_format:
            cell.number_format = number_format

    def __enter__(self):
        ''' Open the template sheet '''
        self._wb = openpyxl.load_workbook(self.template_filepath)
        self._sheet_details = self._wb['Details']
        self._sheet_meter_values = self._wb['Zählerstände']

        header = openpyxl.styles.NamedStyle(name='header')
        header.font = openpyxl.styles.Font(name='Calibri', bold=True, size=11)
        self._wb.add_named_style(header)

        content = openpyxl.styles.NamedStyle(name='content')
        content.font = openpyxl.styles.Font(name='Calibri', size=11)
        self._wb.add_named_style(content)

        # Apply styles
        for cell in self._sheet_details['1']:
            cell.style = 'header'

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        ''' Write the result '''

        self._wb.save(self._filepath)