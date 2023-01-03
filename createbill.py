#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import logging

from nebenkosten import Date, DateRange
import nebenkosten.excel
import nebenkosten.meter

__author__ = "Paul Wichern"
__license__ = "MIT"
__version__ = "0.1.0"

# Setup logging to file
logging.basicConfig(
    filename='nebenkosten.log',
    level=logging.INFO,
    encoding='utf-8',
    format= '[%(asctime)s] {%(pathname)s:%(lineno)-3d} %(levelname)-8s - %(message)s',
    datefmt='%H:%M:%S')

# Add logging to stdout
console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter('%(levelname)-8s %(message)s')
console.setFormatter(formatter)
logging.getLogger('').addHandler(console)

# TODO: Write categories to overview page
# TODO: Print report, showing (logging);
#           - missing files for invoices
#           - missing meter values
#           - missing bills for a 
# TODO: Do not print meter values that are not important for this bill range
# TODO: Handle conversion of m³ into kWh (detect unit mismatch between invoice and bci?)
# TODO: Bundle Excel Sheet with all invoices into a zip
# TODO: Mention invoice filename in BillItem Row
# TODO: clean up ResultSheet row writing

def main(args):
    bill_range = DateRange(args.begin, args.end)

    input_sheet = nebenkosten.excel.InputSheetReader(args.invoices, args.appartement, bill_range)
    split_dates = nebenkosten.get_people_count_change_dates(input_sheet.tenants, bill_range)

    # TODO: log appartement size and tenant name
    
    with nebenkosten.excel.ResultSheetWriter(args.out, input_sheet, bill_range) as result_sheet:
        # Get all meter values important linked to this bci.
        meter_values = filter(lambda mv: mv.name == bci.meter, input_sheet.meter_values)

        for bci in input_sheet.bcis:
            # We cannot directly write invoices that rely on calculated meter values, because we do not
            # yet know the row to reference.
            meter_manager = nebenkosten.meter.MeterManager(input_sheet.meter_values, bci.meter)
            invoices_by_consumption = []

            # Write all invoices related to this bci.
            for invoice in filter(lambda i: i.type == bci.invoice_type and i.range.overlaps(bill_range), input_sheet.invoices):
                # TODO: Make an enum of split types
                if bci.split == 'Pro Person':
                    for invoice_split, people_count in invoice.split(split_dates):
                        result_sheet.add_bill_item(invoice_split, bci, people_count, comment='Personenzahl geändert')
                elif bci.split == 'Nach Verbrauch':
                    consumption_range = DateRange(
                        max(invoice.range.begin, bill_range.begin),
                        min(invoice.range.end, bill_range.end))
                    meter_manager.add_meter_value(consumption_range.begin)
                    meter_manager.add_meter_value(consumption_range.end)
                    invoices_by_consumption.append((invoice, consumption_range))
                else:
                    result_sheet.add_bill_item(invoice, bci)
            
            # Write meter values to result sheet first, so that we know which meter value is in which row.
            # We have to loop twice:
            # 1. Write all meter values in order to get their row.
            # 2. Update the value field of all calculated meter values with actual rows
            # TODO: ResultSheet should contain all the logic for how to write rows and how to store date->row mapping.
            for mv_date in sorted(meter_manager.values):
                row = result_sheet.add_meter_value(meter_manager.values[mv_date])
                meter_manager.set_row(mv_date, row)
            for mv_date in sorted(meter_manager.values):
                if meter_manager.values[mv_date].count == None:
                    result_sheet.update_meter_value_formula(meter_manager.values[mv_date], meter_manager.get_count_formula(mv_date))

            for invoice, consumption_range in invoices_by_consumption:
                row_before = meter_manager.get_row(consumption_range.begin)
                row_after = meter_manager.get_row(consumption_range.end)
                # TODO: Only ResultSheet should know about Sheet names and columns.
                consumption = f'=Zählerstände!C{row_after}-Zählerstände!C{row_before}'
                result_sheet.add_bill_item(invoice, bci, consumption=str(consumption))

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Nebenkosten Abrechner')

    parser.add_argument('invoices', help='Pfad zu Nebenkostentabelle', nargs='?')
    parser.add_argument('begin', help='Startdatum', nargs='?', 
        type=lambda s: Date.from_str(s))
    parser.add_argument('end', help='Enddatum', nargs='?',
        type=lambda s: Date.from_str(s))
    parser.add_argument('appartement', help='Wohnung', nargs='?')
    parser.add_argument('out', help='Ausgabedatei', nargs='?')

    args = parser.parse_args()
    main(args)