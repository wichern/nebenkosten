#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse

from nebenkosten import Date, DateRange
import nebenkosten.excel
import nebenkosten.meter

__author__ = "Paul Wichern"
__license__ = "MIT"
__version__ = "0.1.0"

# TODO: Concept:
#       Caluclated meter values should be added to Zählerstände table.
#       Then, the bill calculation items should reference those
#       How could a calculated meter value reference the value in a cell?
#
#       1. We fetche all meter values for this BCI and turn it into a dict (UUID -> MeterValue)
#       2. We extend the meter values whenever we have to guess a meter value and reference to the UUID of the meter values in the formula
#       3. We sort the meter values by date.
#       4. We write the meter values to the result sheet, replace the UUID with the respective rows
#       5. we write the invoices to the result sheet (meter value formula now references "Zählerstände" cells)
#

#=(2610,300049-1622,75)/TAGE("14.05.2021"; "31.12.2019")*TAGE("05.05.2020"; "31.12.2019")-(2610,300049-1622,75)/TAGE("14.05.2021"; "31.12.2019")*TAGE("11.03.2020"; "31.12.2019")
# Verbrauch von 11.03.2020 bis 05.05.2020:
# = Gesamtverbrauch/Gesamtzeitraum*(31.12.2019 bis 05.05.2020)
# - Gesamtverbrauch/Gesamtzeitraum*(31.12.2019 bis 11.03.2020)
# TODO: "Gesamtverbrauch/Gesamtzeitraum*(11.03.2020 bis 05.05.2020)" wäre ausreichend. Ist das mit get_consumption() möglich?
#       Was, wenn es bekannte Zählerstände dazwischen gibt?


# TODO: Write categories to overview page
# TODO: Print report, showing (logging);
#           - missing files for invoices
#           - missing meter values
#           - missing bills for a 
# TODO: Handle conversion of m³ into kWh (detect unit mismatch between invoice and bci?)
# TODO: beautify example bill
# TODO: Bundle Excel Sheet with all invoices into a zip
# TODO: Guess NK for 2023
# TODO: Mention invoice filename in BillItem Row
# TODO: Make output file a parameter
# TODO: clean up ResultSheet row writing

def main(args):
    bill_range = DateRange(args.begin, args.end)

    with nebenkosten.excel.InputSheet(args.invoices, args.appartement, bill_range) as input_sheet:
        split_dates = nebenkosten.get_people_count_change_dates(input_sheet.tenants, bill_range)

        # TODO: log appartement size and tenant name
        
        with nebenkosten.excel.ResultSheet('test.xlsx', input_sheet, bill_range) as result_sheet:
            # Get all meter values important linked to this bci.
            meter_values = filter(lambda mv: mv.name == bci.meter, input_sheet.meter_values)

            for bci in input_sheet.bcis:
                # We cannot directly write invoices that rely on calculated meter values, because we do not
                # yet know the row to reference.
                meter_manager = nebenkosten.meter.MeterManager(input_sheet.meter_values, bci.meter)
                invoices_by_consumption = []

                # Write all invoices related to this bci.
                for invoice in filter(lambda i: i.type == bci.invoice_type and i.range.overlaps(bill_range), input_sheet.invoices):
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
                for mv_date in sorted(meter_manager.values):
                    row = result_sheet.add_meter_value(meter_manager.values[mv_date])
                    meter_manager.set_row(mv_date, row)
                for mv_date in sorted(meter_manager.values):
                    if meter_manager.values[mv_date].count == None:
                        result_sheet.update_meter_value_formula(meter_manager.values[mv_date], meter_manager.get_count_formula(mv_date))

                for invoice, consumption_range in invoices_by_consumption:
                    row_before = meter_manager.get_row(consumption_range.begin)
                    row_after = meter_manager.get_row(consumption_range.end)
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

    args = parser.parse_args()
    main(args)