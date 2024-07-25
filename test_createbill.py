#!/usr/bin/env python3

import unittest
import datetime
import logging

from nebenkosten import Tenant, Date, DateRange, DateCoverage
from nebenkosten import Invoice
from nebenkosten import get_people_count_changes

class CreateBill(unittest.TestCase):
    def testDateRangeContains(self):
        range = DateRange(Date.from_str('01.01.2020'), Date.from_str('31.01.2021'))
        assert Date.from_str('01.01.2020') in range
        assert Date.from_str('31.12.2019') not in range
        assert Date.from_str('15.06.2020') in range
        assert Date.from_str('30.01.2021') in range
        assert Date.from_str('31.01.2021') in range
        assert Date.from_str('01.02.2021') not in range

    def testMaxDate(self):
        assert Date.from_str('01.01.2020') > Date.from_str('31.12.2019')
        assert not (Date.from_str('01.01.2020') < Date.from_str('31.12.2019'))
        assert Date.from_str('01.01.2020') < Date.from_str('31.12.2020')
        assert not (Date.from_str('01.01.2020') > Date.from_str('31.12.2020'))
        assert not (Date.from_str('01.01.2020') > Date.from_str('01.01.2020'))
        assert not (Date.from_str('01.01.2020') < Date.from_str('01.01.2020'))

    def testYesterday(self):
        assert Date.from_str('02.01.2020').yesterday() == Date.from_str('01.01.2020')
        assert Date.from_str('01.01.2020').yesterday() == Date.from_str('31.12.2019')

    def testTomorrow(self):
        assert Date.from_str('02.01.2020').tomorrow() == Date.from_str('03.01.2020')
        assert Date.from_str('31.12.2020').tomorrow() == Date.from_str('01.01.2021')

    def testTenantDefault(self):
        tenant = Tenant('Tenant Name', 'Appartement Name', Date.from_str('15.01.2019'), Date.from_str('12.05.2021'), 1, 100, 50)
        assert tenant.name == 'Tenant Name'
        assert tenant.appartement == 'Appartement Name'
        assert tenant.moving_in == Date.from_str('15.01.2019')
        assert tenant.moving_out == Date.from_str('12.05.2021')
        assert tenant.people == 1
        assert tenant.rent == 100
        assert tenant.advance == 50

        assert Date.from_str('14.01.2019') not in tenant
        assert Date.from_str('15.01.2019') in tenant
        assert Date.from_str('12.05.2021') in tenant
        assert Date.from_str('13.05.2021') not in tenant

    def testTenantNotMoveOut(self):
        tenant = Tenant('Tenant Name', 'Appartement Name', Date.from_str('15.01.2019'), None, 1, 100, 50)
        assert tenant.moving_out == None

        assert Date.from_str('14.01.2019') not in tenant
        assert Date.from_str('15.01.2019') in tenant
        assert Date.from_str('13.05.2021') in tenant
        assert Date.from_str('01.01.2080') in tenant

    def testDateRangeOverlaps(self):
        range_a = DateRange(
            Date.from_str('01.01.2020'),
            Date.from_str('15.01.2020')
        )

        range_b = DateRange(
            Date.from_str('17.01.2019'),
            Date.from_str('18.01.2020')
        )

        assert range_b.overlaps(range_a)
        assert range_a.overlaps(range_b)

    def testDateInRange(self):
        range = DateRange(
            Date.from_str('01.01.2020'),
            Date.from_str('15.01.2020')
        )

        assert Date.from_str('01.01.2020') in range
        assert Date.from_str('02.01.2020') in range
        assert Date.from_str('15.01.2020') in range
        assert Date.from_str('31.12.2019') not in range
        assert Date.from_str('16.01.2020') not in range

    def testRangeInRange(self):
        range = DateRange(
            Date.from_str('01.01.2020'),
            Date.from_str('15.01.2020')
        )

        assert DateRange(Date.from_str('01.01.2020'), Date.from_str('15.01.2020')) in range
        assert DateRange(Date.from_str('31.12.2019'), Date.from_str('15.01.2020')) not in range
        assert DateRange(Date.from_str('31.12.2019'), Date.from_str('16.01.2020')) not in range
        assert DateRange(Date.from_str('05.01.2020'), Date.from_str('16.01.2020')) not in range
        assert DateRange(Date.from_str('05.01.2020'), Date.from_str('10.01.2020')) in range

    def testDateCoverage(self):
        initial_range = DateRange(
            Date.from_str('01.01.2020'),
            Date.from_str('31.12.2020')
        )
        coverage = DateCoverage(initial_range)
        #print(coverage)
        assert len(coverage.ranges) == 1

        covered_range = DateRange(Date.from_str('01.02.2020'), Date.from_str('29.02.2020'))
        coverage.cover(covered_range)
        #print(coverage)
        for uncovered_range in coverage.ranges:
            assert not uncovered_range.overlaps(covered_range)
        assert len(coverage.ranges) == 2

        covered_range = DateRange(Date.from_str('15.08.2020'), Date.from_str('16.08.2020'))
        coverage.cover(covered_range)
        #print(coverage)
        for uncovered_range in coverage.ranges:
            assert not uncovered_range.overlaps(covered_range)

        covered_range = DateRange(Date.from_str('15.08.2020'), Date.from_str('16.08.2020'))
        coverage.cover(covered_range)
        #print(coverage)
        for uncovered_range in coverage.ranges:
            assert not uncovered_range.overlaps(covered_range)
        assert len(coverage.ranges) == 3

        covered_range = DateRange(Date.from_str('01.01.2020'), Date.from_str('31.07.2020'))
        coverage.cover(covered_range)
        #print(coverage)
        for uncovered_range in coverage.ranges:
            assert not uncovered_range.overlaps(covered_range)
        assert len(coverage.ranges) == 2

    def test_get_people_count_change_dates(self):
        tenants = [
            Tenant('T1', 'A1', Date.from_str('01.01.2020'), Date.from_str('31.01.2020'), 1, 100, 50),
            Tenant('T2', 'A1', Date.from_str('01.02.2020'), Date.from_str('31.12.2020'), 2, 100, 50),
            Tenant('T3', 'A2', Date.from_str('01.01.2020'), Date.from_str('31.08.2020'), 3, 100, 50),
            Tenant('T3', 'A2', Date.from_str('01.09.2020'), Date.from_str('31.12.2020'), 1, 100, 50),
        ]
        split_dates = get_people_count_changes(DateRange(Date.from_str('01.01.2020'), Date.from_str('31.12.2020')), tenants)

        assert len(split_dates) == 3

        assert split_dates[0][0] == Date.from_str('01.01.2020')
        assert split_dates[0][1] == 4

        assert split_dates[1][0] == Date.from_str('01.02.2020')
        assert split_dates[1][1] == 5

        assert split_dates[2][0] == Date.from_str('01.09.2020')
        assert split_dates[2][1] == 3

    def test_split_dates(self):
        invoice = Invoice('I', 'S', 'N', Date(datetime.date(2020, 1, 24)), None, DateRange(Date(datetime.date(2020, 1, 1)), Date(datetime.date(2020, 12, 31))), 138, 1, 0, '')

        # Empty split dates
        invoices = invoice.split([])
        assert len(invoices) == 0

        # People count does not change
        split_dates = [(Date(datetime.date(2020, 9, 1)), 7)]
        invoices = invoice.split(split_dates)
        assert len(invoices) == 2
        assert invoices[0][1] == 7
        assert invoices[1][1] == 7

        # People count does change
        split_dates = [(Date(datetime.date(2020, 9, 1)), 7), (Date(datetime.date(2020, 9, 1)), 6)]
        invoices = invoice.split(split_dates)
        assert len(invoices) == 3
        assert invoices[0][1] == 7
        assert invoices[1][1] == 7
        assert invoices[2][1] == 6

    def test_split_invoice(self):
        invoice = Invoice(type='Strom', supplier='eon', invoice_number='211162337852', date=Date(date=datetime.date(2024, 3, 19)), notes=None, range=DateRange(begin=Date(date=datetime.date(2023, 3, 6)), end=Date(date=datetime.date(2023, 8, 31))), net=0.4199, amount=1768, tax=0.19, path='2024_03_19_eon.pdf')
        splits = invoice.split([(Date(date=datetime.date(2023, 4, 1)), 5)])
        for invoice_part, people_count in splits:
            print(invoice_part)
            print(people_count)
        

            billed_range = DateRange(
                max(invoice_part.range.begin, Date(datetime.date(2023, 1, 1))),
                min(invoice_part.range.end, Date(datetime.date(2023, 12, 31))))
            print(billed_range)


    # def test_split_invoice_where_person_count_changes(self):
    #     tenants = [
    #         Tenant.from_row(['T1', 'A1', '01.01.2020', '31.01.2020', 1]),
    #         Tenant.from_row(['T2', 'A1', '01.02.2020', '31.12.2020', 2]),
    #         Tenant.from_row(['T3', 'A2', '01.01.2020', '31.08.2020', 3]),
    #         Tenant.from_row(['T3', 'A2', '01.09.2020', '31.12.2020', 1]),
    #     ]
    #     bill_range = DateRange(Date.from_str('01.01.2020'), Date.from_str('31.12.2020'))
    #     split_dates = get_people_count_change_dates(tenants, bill_range)
    #     invoice = Invoice.from_row(['Type', 'Supplier', 'Number', '01.01.2021', 'Notes', '01.12.2019', '31.12.2020', '10', '1', '19'])
    #     invoices = split_invoice_where_person_count_changes(split_dates, invoice, bill_range)

    #     assert len(invoices) == 3
    #     assert invoices[0][0].range.begin == Date.from_str('01.12.2019')
    #     assert invoices[0][0].range.end == Date.from_str('31.01.2020')
    #     assert invoices[0][1] == 4
    #     assert invoices[1][0].range.begin == Date.from_str('01.02.2020')
    #     assert invoices[1][0].range.end == Date.from_str('31.08.2020')
    #     assert invoices[1][1] == 5
    #     assert invoices[2][0].range.begin == Date.from_str('01.09.2020')
    #     assert invoices[2][0].range.end == Date.from_str('31.12.2020')
    #     assert invoices[2][1] == 3
        
    #     # Edge case: invoice on date of a split
    #     invoice = Invoice.from_row(['Type', 'Supplier', 'Number', '01.01.2021', 'Notes', '01.12.2019', '01.02.2020', '10', '1', '19'])
    #     invoices = split_invoice_where_person_count_changes(split_dates, invoice, bill_range)

    #     assert len(invoices) == 2
    #     assert invoices[0][0].range.begin == Date.from_str('01.12.2019')
    #     assert invoices[0][0].range.end == Date.from_str('31.01.2020')
    #     assert invoices[0][1] == 4
    #     assert invoices[1][0].range.begin == Date.from_str('01.02.2020')
    #     assert invoices[1][0].range.end == Date.from_str('01.02.2020')
    #     assert invoices[1][1] == 5
        
    #     # Edge case: invoice before date of a split
    #     invoice = Invoice.from_row(['Type', 'Supplier', 'Number', '01.01.2021', 'Notes', '01.12.2019', '31.01.2020', '10', '1', '19'])
    #     invoices = split_invoice_where_person_count_changes(split_dates, invoice, bill_range)

    #     assert len(invoices) == 1
    #     assert invoices[0][0].range.begin == Date.from_str('01.12.2019')
    #     assert invoices[0][0].range.end == Date.from_str('31.01.2020')
    #     assert invoices[0][1] == 4
        
    #     # Edge case: invoice after date of last split
    #     invoice = Invoice.from_row(['Type', 'Supplier', 'Number', '01.01.2021', 'Notes', '15.10.2020', '31.12.2020', '10', '1', '19'])
    #     invoices = split_invoice_where_person_count_changes(split_dates, invoice, bill_range)

    #     assert len(invoices) == 1
    #     assert invoices[0][0].range.begin == Date.from_str('15.10.2020')
    #     assert invoices[0][0].range.end == Date.from_str('31.12.2020')
    #     assert invoices[0][1] == 3
        
    #     # Edge case: no splits at all
    #     tenants_single = [
    #         Tenant.from_row(['T1', 'A1', '01.01.2020', '31.12.2020', 1]),
    #     ]
    #     split_dates = get_people_count_change_dates(tenants_single, bill_range)
    #     invoice = Invoice.from_row(['Type', 'Supplier', 'Number', '01.01.2021', 'Notes', '01.01.2020', '31.12.2020', '10', '1', '19'])
    #     invoices = split_invoice_where_person_count_changes(split_dates, invoice, bill_range)

    #     # print('')
    #     # for i, c in invoices:
    #     #     print(i)

    #     assert len(invoices) == 1
    #     assert invoices[0][0].range.begin == Date.from_str('01.01.2020')
    #     assert invoices[0][0].range.end == Date.from_str('31.12.2020')
    #     assert invoices[0][1] == 1

    # def test_get_meter_value(self):
    #     meter_values = [
    #         MeterValue.from_row(['Name', '15,6', Date.from_str('01.01.2020'), '']),
    #         MeterValue.from_row(['Name', '55,6', Date.from_str('01.02.2020'), '']),
    #         MeterValue.from_row(['Name', '100', Date.from_str('01.03.2020'), ''])
    #     ]

    #     mv = get_meter_value(meter_values, 'Name', Date.from_str('01.01.2020'))
    #     assert mv == meter_values[0]
    #     mv = get_meter_value(meter_values, 'Name', Date.from_str('01.02.2020'))
    #     assert mv == meter_values[1]
    #     mv = get_meter_value(meter_values, 'Name', Date.from_str('01.03.2020'))
    #     assert mv == meter_values[2]

    #     with pytest.raises(MeterValueException):
    #         mv = get_meter_value(meter_values, 'Name', Date.from_str('01.01.2019'))
    #     with pytest.raises(MeterValueException):
    #         mv = get_meter_value(meter_values, 'Name', Date.from_str('02.03.2020'))

    #     mv = get_meter_value(meter_values, 'Name', Date.from_str('15.01.2020'))
    #     assert mv.date == Date.from_str('15.01.2020')
    #     assert mv.count == '(55,6-15,6)/DAYS("01.02.2020", "01.01.2020")*DAYS("15.01.2020", "01.01.2020")'
    #     assert mv.name == 'Name'
    #     assert mv.notes == 'Geschätzt'

    # def test_get_consumption(self):
    #     meter_values = [
    #         MeterValue.from_row(['Name', '15,6', Date.from_str('01.01.2020'), '']),
    #         MeterValue.from_row(['Name', '55,6', Date.from_str('31.01.2020'), '']),  # Tenant moved out
    #         MeterValue.from_row(['Name', '55,6', Date.from_str('01.02.2020'), '']),  # Tenant moved in
    #         MeterValue.from_row(['Name', '100', Date.from_str('31.03.2020'), '']),
    #         MeterValue.from_row(['Other Name', '1000', Date.from_str('15.02.2020'), ''])
    #     ]
    #     consumption_old_tenant = get_consumption('Name', meter_values, DateRange(Date.from_str('01.01.2020'), Date.from_str('31.01.2020')))
    #     consumption_new_tenant = get_consumption('Name', meter_values, DateRange(Date.from_str('01.02.2020'), Date.from_str('31.03.2020')))

    #     assert consumption_old_tenant == '55,6-15,6'
    #     assert consumption_new_tenant == '100-55,6'

    #     consumption_complete = get_consumption('Name', meter_values, DateRange(Date.from_str('01.01.2020'), Date.from_str('31.03.2020')))
    #     assert consumption_complete == '100-15,6'

    #     consumption_between = get_consumption('Name', meter_values, DateRange(Date.from_str('15.01.2020'), Date.from_str('15.03.2020')))
    #     assert consumption_between == '=(100-55,6)/DAYS("31.03.2020", "01.02.2020")*DAYS("15.03.2020", "01.02.2020")-(55,6-15,6)/DAYS("31.01.2020", "01.01.2020")*DAYS("15.01.2020", "01.01.2020")'
        
    #     with pytest.raises(MeterValueException):
    #         get_consumption('', meter_values, DateRange(Date.from_str('15.01.2020'), Date.from_str('15.03.2020')))
    #     with pytest.raises(MeterValueException):
    #         get_consumption(None, meter_values, DateRange(Date.from_str('15.01.2020'), Date.from_str('15.03.2020')))

    # def test_bill_item(self):
    #     invoice = Invoice.from_row(['Strom', 'Supplier', 'Number', '01.01.2020', 'Notes', '01.01.2020', '31.12.2020', '100,00', '1', '19'])
    #     bci = BillCalculationItem.from_row(['Wohnung #1', 'Nach Verbrauch', 'kWh', 'Strom', 'Zählername'])
    #     appartements = [
    #         Appartement.from_row(['Wohnung #1', '100']),
    #         Appartement.from_row(['Wohnung #2', '50'])
    #     ]
    #     tenant = Tenant.from_row(['Tenant Name', 'Wohnung #1', '01.01.2019', None, 2])
    #     bill_item = BillItem(invoice, bci, appartements, DateRange(Date.from_str('01.01.2020'), Date.from_str('31.03.2020')), 0, 5, '50', '15.0', tenant)

    #     assert bill_item.days == '=DAYS(B0, A0)'
    #     assert bill_item.type == 'Strom'
    #     assert bill_item.description == 'Notes'
    #     assert bill_item.meter == 'Zählername'
    #     assert bill_item.net == '100,00'
    #     assert bill_item.amount == '1'
    #     assert bill_item.tax == '19'
    #     assert bill_item.gross == '=G0*H0*(1+I0)'
    #     assert bill_item.invoice_range_days == '=DAYS(31.12.2020, 01.01.2020)'
    #     assert bill_item.share_name == 'Nach Verbrauch'
    #     assert bill_item.share_percentage == '15.0'
    #     assert bill_item.sum == '=J0/K0*C0*M0'

if __name__ == '__main__':
    unittest.main()