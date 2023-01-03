#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import enum

import openpyxl

from nebenkosten.types import Invoice, Appartement, Tenant, MeterValue, Meter, Date, DateRange, BillCalculationItem, SplitType

class InputSheet(enum.Enum):
    ''' Name of the sheets in the input file '''
    INVOICES = 'Rechnungen'
    APPARTEMENTS = 'Wohnungen'
    TENANTS = 'Mieter'
    METER_VALUES = 'Zählerstände'
    METERS = 'Zähler'
    BILL_CALCULATION_ITEMS = 'Abrechnungseinstellungen'

class InputSheetReader():
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
        
class ResultSheet(enum.Enum):
    ''' Sheet names in the result sheet. '''
    OVERVIEW = 'Zusammenfassung'
    DETAILS = 'Details'
    METER_VALUES = 'Zählerstände'

class RowWriter():
    ''' Helper class to write a row into a sheet '''

    def __init__(self, sheet, row: int, column: int = 1):
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
        self._column += 1

class ResultSheetWriter():
    '''
    Write the result Excel sheet.
    '''

    template_filepath = 'example-bill.xlsx'

    def __init__(self, filepath: str, input_sheet: InputSheetReader, bill_range: DateRange):
        self._filepath = filepath
        self._input_sheet = input_sheet
        self._bill_range = bill_range
        self._bill_item_row = 2
        self._meter_value_row = 2

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

    def add_bill_item(self, invoice: Invoice, bci: BillCalculationItem, total_people_count: int = 0, comment: str = None, consumption: str = None):
        row_writer = RowWriter(self._wb[ResultSheet.DETAILS.value], self._bill_item_row)
        row_writer.write_date(max(self._bill_range.begin, invoice.range.begin))
        row_writer.write_date(min(self._bill_range.end, invoice.range.end))
        row_writer.write_number('=_xlfn.days(B{0}, A{0})'.format(self._bill_item_row), precision=0)
        row_writer.write(invoice.type)
        row_writer.write(invoice.notes)
        row_writer.write(bci.meter)
        row_writer.write_currency(invoice.net)
        row_writer.write_number(invoice.amount)
        row_writer.write_percentage(invoice.tax)
        row_writer.write_currency('=G{0}*H{0}*(1+I{0})'.format(self._bill_item_row))
        row_writer.write_number(f'=_xlfn.days("{invoice.range.end}", "{invoice.range.begin}")', unit='Tage', precision=0)
        row_writer.write(bci.split)
        
        if bci.split == SplitType.PER_APPARTEMENT.value:
            row_writer.write_percentage(f'=1/{len(self._input_sheet.appartements)}')
        elif bci.split == SplitType.PER_PERSON.value:
            row_writer.write_percentage(f'={self._input_sheet.tenant.people}/{total_people_count}')
        elif bci.split == SplitType.PER_SQUAREMETER.value:
            row_writer.write_percentage(f'={self._input_sheet.appartement.size}/{sum([a.size for a in self._input_sheet.appartements])}')
        elif bci.split == SplitType.PER_CONSUMPTION.value:
            row_writer.write_number(consumption, unit=bci.unit)
        elif bci.split == SplitType.HALF.value:
            row_writer.write_percentage('=1/2')
        elif bci.split == SplitType.THIRD.value:
            row_writer.write_percentage('=1/3')
        elif bci.split == SplitType.QUARTER.value:
            row_writer.write_percentage('=1/4')
        elif bci.split == SplitType.COMPLETE.value:
            row_writer.write_percentage('1')
        else:
            raise InvalidCellValue('Unknown bill split: "%s"' % split)

        if bci.split == 'Nach Verbrauch':
            row_writer.write_currency('=G{0}*(1+I{0})*M{0}'.format(self._bill_item_row))
        else:
            row_writer.write_currency('=J{0}/K{0}*C{0}*M{0}'.format(self._bill_item_row))

        if comment:
            row_writer.write(comment)

        self._bill_item_row += 1

    def add_meter_value(self, meter_value: MeterValue):
        row_writer = RowWriter(self._wb[ResultSheet.METER_VALUES.value], self._meter_value_row)
        row_writer.write(meter_value.name)
        row_writer.write_date(meter_value.date)
        if meter_value.count:
            meter = self._input_sheet.get_meter(meter_value.name)
            row_writer.write_number(meter_value.count, unit=meter.unit)
            row_writer.write('Gemessen')
        self._meter_value_row += 1
        return self._meter_value_row - 1

    def update_meter_value_formula(self, meter_value: MeterValue, formula: str):
        meter = self._input_sheet.get_meter(meter_value.name)
        row_writer = RowWriter(self._wb[ResultSheet.METER_VALUES.value], meter_value.row, column=3)
        row_writer.write_number(formula, unit=meter.unit)
        row_writer.write('Berechnet', )

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        ''' Write the result '''

        self._wb.save(self._filepath)