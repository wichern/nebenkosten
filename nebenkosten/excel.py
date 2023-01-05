#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import enum
from typing import List

import openpyxl

from nebenkosten.types import Invoice, Appartement, Tenant, MeterValue, Meter, Date, DateRange, BillCalculationItem

class InputSheet(enum.Enum):
    ''' Name of the sheets in the input file '''
    INVOICES = 'Rechnungen'
    APPARTEMENTS = 'Wohnungen'
    TENANTS = 'Mieter'
    METER_VALUES = 'Zählerstände'
    METERS = 'Zähler'
    BILL_CALCULATION_ITEMS = 'Abrechnungseinstellungen'

class InputSheetReader:
    '''
    Parse the input sheet.
    '''

    def __init__(self, path, appartement_name: str, bill_range: DateRange):
        '''
        :param path:                Path to the input file
        :param appartement_name:    Name of the appartement which we are creating a bill for.
        :param bill_range:          Bill range
        '''

        workbook = openpyxl.load_workbook(filename=str(path))

        self.invoices = []
        for row in self.__get_rows(workbook, InputSheet.INVOICES):
            invoice_range = DateRange(Date.from_str(row[5]), Date.from_str(row[6]))
            if invoice_range.overlaps(bill_range):
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

        self.appartements = []
        for row in self.__get_rows(workbook, InputSheet.APPARTEMENTS):
            self.appartements.append(Appartement(row[0], row[1]))

        self.appartement = None
        self.appartement = next(a for a in self.appartements if a.name == appartement_name)

        self.tenants = []
        for row in self.__get_rows(workbook, InputSheet.TENANTS):
            # For 'moving out' we also accept None
            moving_out = None
            if row[3]:
                moving_out = Date.from_str(row[3])

            self.tenants.append(
                Tenant(
                    row[0],
                    row[1],
                    Date.from_str(row[2]),
                    moving_out,
                    int(row[4])
                )
            )

        self.tenant = None
        self.tenant = next(t for t in self.tenants if bill_range.begin in t and bill_range.end in t and t.appartement == appartement_name)

        self.meter_values = []
        for row in self.__get_rows(workbook, InputSheet.METER_VALUES):
            self.meter_values.append(MeterValue(row[0], row[1], Date.from_str(row[2]), row[3]))

        self.meter = []
        for row in self.__get_rows(workbook, InputSheet.METERS):
            self.meter.append(Meter(row[0], row[1], row[2]))

        self.bcis = []
        for row in self.__get_rows(workbook, InputSheet.BILL_CALCULATION_ITEMS):
            bci = BillCalculationItem(row[0], row[1], row[2], row[3], row[4])
            # Skip BCIs that are not relevant for this appartement
            if bci.appartement == appartement_name:
                self.bcis.append(bci)

    def __get_rows(self, workbook, sheet: InputSheet):
        ''' Get all data rows of given sheet. '''
        sheet = workbook[sheet.value]
        # We need to filter out the rows without content, because we will receive those as well.
        return filter(lambda row: row[0], sheet.iter_rows(min_row=2, values_only=True))

    def get_meter(self, meter_name) -> Meter:
        ''' Get Meter by name. '''
        return next(m for m in self.meter if m.name == meter_name)

    def get_invoices(self, bci: BillCalculationItem, date_range: DateRange) -> List[Invoice]:
        ''' List all invoices related to given bci in given date range. '''
        return filter(lambda i: i.type == bci.invoice_type and i.range.overlaps(date_range), self.invoices)

class ResultSheet(enum.Enum):
    ''' Sheet names in the result sheet. '''
    OVERVIEW = 'Zusammenfassung'
    DETAILS = 'Details'
    METER_VALUES = 'Zählerstände'

class CellWriter:
    ''' Helper class to write into a cell of a sheet '''

    def __init__(self, sheet, row: int, column: int):
        self._sheet = sheet
        self._row = row
        self._column = column

    def write_date(self, date: Date):
        ''' Write a date '''
        self.write(date.date, number_format='DD.MM.YY')

    def write_number(self, number: str, unit: str = None, precision: int = 2):
        ''' Write a nunmber '''
        precision_format = '0'
        if precision > 0:
            precision_format += '.' + (precision * '0')

        if unit:
            self.write(number, number_format=f'{precision_format}" {unit}"')
        else:
            self.write(number, number_format=precision_format)

    def write_currency(self, number: str):
        ''' Write a currency '''
        self.write(number, number_format='0.00" "€')

    def write_percentage(self, number: str):
        ''' Write a percentage '''
        self.write(number, number_format='0.00" "%')

    def write(self, content, number_format = None):
        ''' Write into a cell '''
        cell = self._sheet.cell(row=self._row, column=self._column)
        cell.style = 'content'
        cell.value = content
        if number_format:
            cell.number_format = number_format

    def row(self):
        ''' Get the row of this writer '''
        return self._row

class RowWriter(CellWriter):
    ''' Helper class to write a row into a sheet '''

    def write(self, content, number_format = None):
        ''' Write into a cell '''
        super().write(content, number_format)
        self._column += 1

class ResultSheetWriter:
    '''
    Write the result Excel sheet.
    '''

    template_filepath = 'example-bill.xlsx'

    def __init__(self):
        self._current_row = {
            ResultSheet.METER_VALUES: 2,
            ResultSheet.DETAILS: 2
        }

        self._wb = openpyxl.load_workbook(self.template_filepath)

        header = openpyxl.styles.NamedStyle(name='header')
        header.font = openpyxl.styles.Font(name='Calibri', bold=True, size=11)
        self._wb.add_named_style(header)

        content = openpyxl.styles.NamedStyle(name='content')
        content.font = openpyxl.styles.Font(name='Calibri', size=11)
        self._wb.add_named_style(content)

        # Apply styles
        for cell in self._wb[ResultSheet.DETAILS.value]['1']:
            cell.style = 'header'

    def row_writer(self, sheet: ResultSheet):
        ''' Create a row writer for the next row in given sheet. '''
        writer = RowWriter(self._wb[sheet.value], self._current_row[sheet], 1)
        self._current_row[sheet] += 1
        return writer

    def cell_writer(self, sheet: ResultSheet, row: int, column: int):
        ''' Create a cell writer for the given row and column in given sheet. '''
        return CellWriter(self._wb[sheet.value], row, column)

    def save(self, filepath):
        ''' Write the result '''

        self._wb.save(filepath)