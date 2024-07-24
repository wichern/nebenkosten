#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
Bill calculation
'''

import argparse
from dataclasses import dataclass
import logging
from typing import Dict
import sys
import shutil
import tempfile

from nebenkosten import Date, DateRange, DateCoverage, Invoice, BillCalculationItem, SplitType
from nebenkosten import RowWriter, ResultSheet, ResultSheetWriter, InputSheetReader
from nebenkosten import MeterManager
from nebenkosten import InvalidCellValue, InputFileError
from nebenkosten import get_people_count_changes

__author__ = "Paul Wichern"
__license__ = "MIT"
__version__ = "1.0.0"

WATER_TEMPERATURE_COLD = 10
WATER_TEMPERATURE_WARM = 43

# 2.0
# TODO: Sort BCI by which is limiting bill creation the most.
#       e.g.:   Strom bis 31.12.2020, Wasser bis 31.03.2021, Müll bis 15.04.2021, ...
# TODO: Create sheet that calculates costs for each month individually, in order to show changes over time
# TODO: Abrechnungseinstellungen haben ein Start und ein Enddatum
# TODO: Vermieteranteil an CO2 Steuer berücksichtigen
# TODO: Create PDF
# TODO: Simple GUI
# TODO: Add validate() to every type for self-check and print cell that contained ill-formated data
# TODO: Print consumptions in ranges
# TODO: Set page view to be able to create pdf faster
# TODO: Neue Abrechnungseinstellung: "Rest" (nimmt Menge aus Rechnung und subtrahiert alle anderweitig in Wohnungen berechneten Verbräuche)
#       Damit können wir NK für Hauptwohnung berechnen

@dataclass
class BillItem:
    ''' Temporary storage for bill item (DETAILS sheet). '''
    invoice: Invoice
    bci: BillCalculationItem
    billed_range: DateRange
    split_percentage: str
    comment: str = None

    def write(self, row_writer: RowWriter):
        ''' Write this bill item into the result sheet. '''

        row = row_writer.row()

        row_writer.write_date(self.billed_range.begin)
        row_writer.write_date(self.billed_range.end)
        row_writer.write_number(f'=_xlfn.days(B{row}, A{row})+1', precision=0)
        row_writer.write(self.invoice.type)
        if self.comment:
            row_writer.write(self.comment)
        else:
            row_writer.write(self.invoice.notes)
        row_writer.write(self.bci.meter)
        row_writer.write_currency(self.invoice.net)
        row_writer.write_number(self.invoice.amount)
        row_writer.write_percentage(self.invoice.tax)
        row_writer.write_currency(f'=G{row}*H{row}*(1+I{row})')
        row_writer.write_number(
            f'=_xlfn.days("{self.invoice.range.end}", "{self.invoice.range.begin}")+1',
            unit='Tage',
            precision=0)
        row_writer.write(self.bci.split)
        if self.bci.split == SplitType.PER_CONSUMPTION.value:
            row_writer.write_number(self.split_percentage, unit=self.bci.unit)
            row_writer.write_currency(f'=G{row}*(1+I{row})*M{row}')
        else:
            row_writer.write_percentage(self.split_percentage)
            row_writer.write_currency(f'=J{row}/K{row}*C{row}*M{row}')
        row_writer.write(self.invoice.path)

# pylint: disable=too-few-public-methods
class BillCreator:
    ''' Helper class to create a bill '''

    def __init__(self, input_sheet: InputSheetReader, bill_range: DateRange, out_path: str):
        self._input = input_sheet
        self._range = bill_range
        self._out = ResultSheetWriter()
        self._out_path = out_path
        self._bill_items = []
        self._split_dates = get_people_count_changes(bill_range, input_sheet.tenants)
        self._coverages = {bci.invoice_type: DateCoverage(bill_range) for bci in input_sheet.bcis}
        self._receipts = set()

    def create(self, receipts_dir : str):
        ''' Create the bill '''

        for bci in self._input.bcis:
            logging.debug(bci)
            if bci.split == SplitType.PER_CONSUMPTION.value:
                self.__per_consumption(bci)
            elif bci.split == SplitType.PER_PERSON.value:
                self.__per_person(bci)
            else:
                self.__per_percentage(bci)

        # Analyze bill coverage
        for bci in self._input.bcis:
            for uncovered_range in self._coverages[bci.invoice_type].ranges:
                logging.warning('Für "%s" wurde von %s bis %s nichts in Rechnung gestellt.',
                    bci.invoice_type, uncovered_range.begin, uncovered_range.end)

        # We write the overview after calculating the bill.
        # This way we already know how many rows where added to the DETAILS sheet.
        self._out.write_overview(
            self._input.appartement.name,
            self._input.tenant.name,
            self._range,
            self._input.bcis)
        self._out.save(self._out_path)

        # Create zip file with receipts
        logging.info('Erstelle Archiv mit Rechnungen ...')
        with tempfile.TemporaryDirectory() as temp_dir:
            for receipt in self._receipts:
                if receipt:
                    shutil.copy(receipts_dir + '/' + receipt, temp_dir + '/' + receipt)
            shutil.make_archive(self._out_path.replace('.xlsx', ''), 'zip', temp_dir)

    # pylint: disable=too-many-locals
    def __per_consumption(self, bci):
        ''' Handle all BCI that are based on consumption '''

        meter = self._input.get_meter(bci.meter)
        if not meter:
            logging.error('Zähler "%s" nicht in Eingabedatei.', bci.meter)
            raise InputFileError

        # 1. Create all meter values that do not exist already and store all meter values we
        #    want to see in our result sheet into a set.
        dates = set()
        meter_manager = MeterManager(self._input.meter_values, bci.meter)
        for invoice in self.__invoices(bci):
            logging.debug(invoice)

            consumption_range = DateRange(
                max(invoice.range.begin, self._range.begin),
                min(invoice.range.end, self._range.end))

            logging.debug('Verbrauchszeitraum: %s bis %s',
                consumption_range.begin, consumption_range.end)

            dates.add(consumption_range.begin)
            dates.add(consumption_range.end)

            if consumption_range.begin not in meter_manager.values:
                date_before, date_after = meter_manager.get_surrounding_dates(
                    consumption_range.begin)
                dates.add(date_before)
                dates.add(date_after)

            if consumption_range.end not in meter_manager.values:
                date_before, date_after = meter_manager.get_surrounding_dates(
                    consumption_range.end)
                dates.add(date_before)
                dates.add(date_after)

        # 2. Save all meter values to result sheet.
        rows: Dict[Date, int] = {}
        for mv_date in sorted(dates):
            row = self._out.row_writer(ResultSheet.METER_VALUES)
            row.write(bci.meter)
            row.write_date(mv_date)

            if mv_date in meter_manager.values:
                row.write_number(meter_manager.values[mv_date].count, unit=meter.unit)
                row.write('Gemessen')

            rows[mv_date] = row.row()

        # 3. Update all meter value formulas, now that we know their rows in the result sheet.
        for mv_date in sorted(dates):
            if mv_date not in meter_manager.values:
                row = rows[mv_date]
                count = self.__count_formula(mv_date, meter_manager, rows)
                self._out.cell_writer(ResultSheet.METER_VALUES, row, 3)\
                    .write_number(count, unit=meter.unit)
                self._out.cell_writer(ResultSheet.METER_VALUES, row, 4)\
                    .write('Berechnet')

        # 4. Save all bill items to the result sheet.
        for invoice in self.__invoices(bci):
            consumption_range = DateRange(
                max(invoice.range.begin, self._range.begin),
                min(invoice.range.end, self._range.end))

            consumption = f'{ResultSheet.METER_VALUES.value}!C{rows[consumption_range.end]}'\
                f'-{ResultSheet.METER_VALUES.value}!C{rows[consumption_range.begin]}'

            # Check if unit types match and add conversion and comment otherwise.
            comment = None
            if bci.unit != meter.unit:
                consumption = self.__convert_units(meter.unit, bci.unit, consumption)
                comment = f'{meter.unit} in {bci.unit} umgerechnet.'

            bill_item = BillItem(
                invoice, bci, consumption_range, f'={consumption}', comment=comment)
            bill_item.write(self._out.row_writer(ResultSheet.DETAILS))
            self._receipts.add(invoice.path)

            # Update bill coverage for this BCI
            self._coverages[bci.invoice_type].cover(bill_item.billed_range)

    def __per_person(self, bci):
        ''' Handle all BCI that are based on tenant count '''

        for invoice in self.__invoices(bci):
            logging.debug(invoice)

            comment = None  # In order to only append the comment to the second invoice
            for invoice_part, people_count in invoice.split(self._split_dates):
                logging.debug('Rechnungsteil (%d Personen): %s', people_count, str(invoice_part))

                billed_range = DateRange(
                    max(invoice_part.range.begin, self._range.begin),
                    min(invoice_part.range.end, self._range.end))

                split_percentage = f'={self._input.tenant.people}/{people_count}'
                bill_item = BillItem(invoice, bci, billed_range, split_percentage, comment=comment)
                bill_item.write(self._out.row_writer(ResultSheet.DETAILS))
                self._receipts.add(invoice.path)

                # Update bill coverage for this BCI
                self._coverages[bci.invoice_type].cover(bill_item.billed_range)

                # Comment will be added to the second invoice, if it was split.
                comment = 'Gesamtbewohnerzahl geändert'

    def __per_percentage(self, bci):
        ''' Handle all BCI that are based on percentage '''

        for invoice in self.__invoices(bci):
            logging.debug(invoice)

            split_percentage = ''
            if bci.split == SplitType.PER_APPARTEMENT.value:
                split_percentage = f'=1/{len(self._input.appartements)}'
            elif bci.split == SplitType.PER_SQUAREMETER.value:
                split_percentage = f'={self._input.appartement.size}'\
                    f'/{sum(a.size for a in self._input.appartements)}'
            elif bci.split == SplitType.HALF.value:
                split_percentage = '=1/2'
            elif bci.split == SplitType.THIRD.value:
                split_percentage = '=1/3'
            elif bci.split == SplitType.QUARTER.value:
                split_percentage = '=1/4'
            elif bci.split == SplitType.COMPLETE.value:
                split_percentage = '1'
            else:
                raise InvalidCellValue(f'Unknown bill split: "{bci.split}"')

            billed_range = DateRange(
                max(invoice.range.begin, self._range.begin),
                min(invoice.range.end, self._range.end))
            bill_item = BillItem(invoice, bci, billed_range, split_percentage)
            bill_item.write(self._out.row_writer(ResultSheet.DETAILS))
            self._receipts.add(invoice.path)

            # Update bill coverage for this BCI
            self._coverages[bci.invoice_type].cover(bill_item.billed_range)

    def __count_formula(self,
                        date: Date,
                        meter_manager: MeterManager,
                        rows: Dict[Date, int]) -> str:
        ''' Create a formula, that calculates the meter value. '''

        before, after = meter_manager.get_surrounding_dates(date)
        before_row = rows[before]
        after_row = rows[after]

        # Calculate the total consumption between measured values.
        delta_value = f'C{after_row}-C{before_row}'

        # Calculate the days between measured values.
        delta_days_total = f'_xlfn.days(B{after_row},B{before_row})'

        # Calculate the days between before date and date.
        delta_days_new = f'_xlfn.days(B{rows[date]},B{before_row})'

        return f'=C{before_row}+({delta_value})/{delta_days_total}*{delta_days_new}'

    def __invoices(self, bci):
        ''' Get all invoices related to this BCI. '''
        return filter(lambda i: i.type == bci.invoice_type \
            and i.range.overlaps(self._range), self._input.invoices)

    def __convert_units(self, unit_from: str, unit_to: str, value: str) -> str:
        ''' Add Excel formula to convert `value` '''
        if unit_from == 'm³' and unit_to == 'kWh':
            # Wärmeverbrauch für Warmwasser (kWh) = Warmwasserverbrauch (m³) x
            #   (Warmwassertemperatur (K) – Kaltwassertemperatur (K)) x 2,5
            return f'({value})*({WATER_TEMPERATURE_WARM}-{WATER_TEMPERATURE_COLD})*2.5'
        logging.error('Kann "%s" nicht in "%s" konvertieren.', unit_from, unit_to)
        raise InputFileError

def main():
    ''' Main '''
    # Parse command line arguments.
    parser = argparse.ArgumentParser(description='Nebenkosten Abrechner')
    parser.add_argument('invoices', help='Pfad zu Nebenkostentabelle', nargs='?')
    parser.add_argument('begin', help='Startdatum', nargs='?', type=Date.from_str)
    parser.add_argument('end', help='Enddatum', nargs='?', type=Date.from_str)
    parser.add_argument('appartement', help='Wohnung', nargs='?')
    parser.add_argument('receipts', help='Ordner mit gescannten Rechnungen', nargs='?')
    args = parser.parse_args()

    filename = f'{args.appartement}-{args.begin}-{args.end}'\
        .replace(' ', '_')\
        .replace('.', '_')

    # Setup logging to file
    logging.basicConfig(
        filename=filename + '.log',
        filemode='w',  # Overwrite existing file
        level=logging.DEBUG,
        encoding='utf-8',
        format= '[%(asctime)s] {%(pathname)s:%(lineno)-3d} %(levelname)-8s - %(message)s',
        datefmt='%H:%M:%S')

    # Add logging to stdout
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    formatter = logging.Formatter('%(levelname)-8s %(message)s')
    console.setFormatter(formatter)
    logging.getLogger('').addHandler(console)

    logging.info('Erstelle Nebenkostenabrechnung für "%s" von %s bis %s',
        args.appartement, args.begin, args.end)

    # Create bill
    bill_range = DateRange(args.begin, args.end)
    input_sheet = InputSheetReader(args.invoices, args.appartement, bill_range)

    bill = BillCreator(input_sheet, bill_range, filename + '.xlsx')
    bill.create(args.receipts)

if __name__ == '__main__':
    try:
        main()
    except InputFileError:
        logging.fatal('Parameter oder Eingabedatei nicht korrekt.')
        sys.exit(1)
