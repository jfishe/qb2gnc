#!/usr/bin/python
# Taken from Template suggested by 
# http://current.workingdirectory.net/posts/2011/gnucash-python-bindings/
# /gnucash-2.6.5/src/optional/python-bindings/example_scripts/rest-api/
import argparse
import sys
import csv
import os
import datetime
from decimal import Decimal
from gnucash import Session, Account, GncNumeric, Query
from gnucash.gnucash_business import Customer, Vendor, Invoice, Entry, TaxTable, TaxTableEntry, Transaction, Split, Bill
from gnucash.gnucash_core_c import  ACCT_TYPE_LIABILITY

from lxml import etree as et

outfieldnames = (
    'id', 'company', 'name', 'addr1', 'addr2', 'addr3', 'addr4', 'phone', 'fax', 'email', 'notes', 'shipname',
    'shipaddr1', 'shipaddr2', 'shipaddr3', 'shipaddr4', 'shiphone', 'shipfax', 'shipmail')

# Format for merging QuickBooks fields when gnucash doesn't have an equivalent
# Combine with spaces, comma and space or use QuickBook label in notes
# E.g. 123 Broadway, Kennewick, wa, 143 Georgetown, wa
# Mr. George Smith
# Contact: George L. Smith
# If not in these lists, overwrite with last occurrence in map.
# Change Bill to 5 from notes to addr4 if want to use commakey.
# commakey = ('addr4', 'shipaddr4')
commakey = ()
# noinspection PyRedundantParentheses
spacekey = ('notes')
linekey = ()

# QuickBooks column head : gnucash columnhead
custmap = {'Active Status': 'active',
           'Customer': 'company',
           'Balance': 'not mapped',
           'Balance Total': 'not mapped',
           'Company': 'not mapped',
           'Mr./Ms./...': 'not mapped',
           'First Name': 'not mapped',
           'M.I.': 'not mapped',
           'Last Name': 'not mapped',
           'Contact': 'notes',
           'Phone': 'phone',
           'Fax': 'fax',
           'Alt. Phone': 'shiphone',
           'Alt. Contact': 'notes',
           'Email': 'email',
           'Bill to 1': 'name',
           'Bill to 2': 'addr1',
           'Bill to 3': 'addr2',
           'Bill to 4': 'addr3',
           'Bill to 5': 'addr4',
           'Ship to 1': 'shipname',
           'Ship to 2': 'shipaddr1',
           'Ship to 3': 'shipaddr2',
           'Ship to 4': 'shipaddr3',
           'Ship to 5': 'shipaddr4',
           'Customer Type': 'not mapped',
           'Terms': 'not mapped',
           'Rep': 'not mapped',
           'Sales Tax Code': 'sales tax code',
           'Tax item': 'tax item',
           'Resale Num': 'notes',
           'Account No.': 'id',
           'Credit Limit': 'not mapped',
           'Job Status': 'not mapped',
           'Job Type': 'not mapped',
           'Job Description': 'not mapped',
           'Start Date': 'not mapped',
           'Projected End': 'not mapped',
           'End Date': 'not mapped',
           'Note': 'notes'}

itemap = {'Active Status': 'not mapped',
          'Type': 'type',
          'Item': 'item',
          'Description': 'not mapped',
          'Sales Tax Code': 'not mapped',
          'Account': 'account',
          'COGS Account': 'not mapped',
          'Asset Account': 'not mapped',
          'Accumulated Depreciation': 'not mapped',
          'Purchase Description': 'not mapped',
          'Quantity On Hand': 'not mapped',
          'Cost': 'not mapped',
          'Preferred Vendor': 'not mapped',
          'Tax Agency': 'not mapped',
          'Price': 'rate',
          'Reorder Point': 'not mapped',
          'MPN': 'not mapped'}

vendmap = {'Active Status': 'active',
           'Vendor': 'company',
           'Balance': 'not mapped',
           'Balance Total': 'not mapped',
           'Company': 'not mapped',
           'Mr./Ms./...': 'not mapped',
           'First Name': 'not mapped',
           'M.I.': 'not mapped',
           'Last Name': 'not mapped',
           'Bill from 1': 'name',
           'Bill from 2': 'addr1',
           'Bill from 3': 'addr2',
           'Bill from 4': 'addr3',
           'Bill from 5': 'addr4',
           'Ship from 1': 'not mapped',
           'Ship from 2': 'not mapped',
           'Ship from 3': 'not mapped',
           'Ship from 4': 'not mapped',
           'Ship from 5': 'not mapped',
           'Contact': 'notes',
           'Phone': 'phone',
           'Fax': 'fax',
           'Alt. Phone': 'notes',
           'Alt. Contact': 'notes',
           'Email': 'email'}

transmap = {'Type': 'type',
            'Date': 'date_opened',
            'Num': 'num',
            'Name': 'owner',
            'Memo': 'notes',
            'Paid': 'paid',
            'Item': 'item',
            'Item Description': 'description',
            'Account': 'account',
            'Sales Tax Code': 'taxable',
            'Qty': 'quantity',
            'Sales Price': 'price',
            'Split': 'split',
            'Amount': 'amount'}


def main():
    parser = argparse.ArgumentParser(description='Convert QuickBooks vendor, customer or tax_items csv file to gnucash')
    parser.add_argument('outfile', metavar='file.gnucash', nargs=1,
                        help='gnucash file to add Customers or Vendors from QuickBooks')
    parser.add_argument('--test',
                        action='store_false',
                        help='Process the input but do not update the gnucash')

    group = parser.add_mutually_exclusive_group()
    group.add_argument('-c', '--customer',
                       metavar='customer.csv',
                       help='Customer QuickBooks input.csv file')
    group.add_argument('-v', '--vendor',
                       metavar='vendor.csv',
                       help='Vendor QuickBooks input.csv file')
    group.add_argument('-i', '--items',
                       metavar='tax_items.csv',
                       help='Tax Items QuickBooks QuickBooks input.csv file \
        gnucash import file')
    group.add_argument(
        '-t', '--transactions',
        metavar='transactions.csv',
        help='Invoices, Bills, Journal, etc. QuickBooks QuickBooks input.csv \
              file gnucash import file')

    parser.parse_args('--items tax_items.csv outfile'.split())
    parser.parse_args('--customer customer.csv outfile'.split())
    parser.parse_args('--vendor vendor.csv outfile'.split())
    parser.parse_args('--transactions transactions.csv outfile'.split())
    args = parser.parse_args()

    items = args.items
    customer = args.customer
    vendor = args.vendor
    transaction = args.transactions
    path = args.outfile[0]
    GncFile.nosave = args.test

    try:
        GncFile.gnc_open(path)
        book = GncFile.book
        root = GncFile.root
        # noinspection PyPep8Naming
        USD = GncFile.USD

        # Process the input file
        reader = csv.DictReader(open(items, 'r'))
        if items is not None:
            out = mapqb2gnc(reader, itemap)
            for row in out:
                new_tax(root, book, USD, row)

        if customer is not None:
            out = mapqb2gnc(reader, custmap)
            for row in out:
                new_customer(book, row, USD)

        if vendor is not None:
            out = mapqb2gnc(reader, vendmap)
            for row in out:
                new_vendor(book, row, USD)

        if transaction is not None:
            out = mapqb2gnc(reader, transmap)
            new_transaction(root, book, out, USD)

        GncFile.gnc_save()
        GncFile.gnc_end()

    except:
        GncFile.gnc_end()
        raise

    return


class TaxRates(object):
    ratetable = {}

    @classmethod
    def israte(cls, tablename, rate):
        if tablename in TaxRates.ratetable.keys():
            oldrate = TaxRates.ratetable[str(tablename)]
            if rate.to_string() != oldrate.to_string():
                TaxRates.ratetable[str(tablename)] = rate
                if GncFile.nosave:
                    print "Current TaxTable %s rate: %s" % (tablename, oldrate)
                    print "Update TaxTable %s rate: %s" % (tablename, rate)
                    GncFile.gnc_save()
                    GncFile.gnc_end()
                    raw_input('After you save & exit GnuCash, press Enter')
                    GncFile.gnc_open()
                else:
                    print("Check TaxTable %s rate: %s, before running " +
                          "without --test") % (tablename, oldrate)
            return True
        else:
            TaxRates.ratetable[str(tablename)] = rate
            return False


def new_tax(root, book, USD, row):
    if row['type'] == 'Sales Tax Item':
        tablename = row['item']
        parent = root.lookup_by_name(row['account'])
        account = Account(book)
        parent.append_child(account)
        account.SetName(tablename)
        account.SetType(ACCT_TYPE_LIABILITY)
        account.SetCommodity(USD)
        rate = gnc_numeric_from(row['rate'])
        # Skip existing tax table and create new ones
        try:
            assert (
                not isinstance(book.TaxTableLookupByName(tablename), TaxTable)
            )
            TaxTable(book, tablename, TaxTableEntry(account, True, rate))
        except AssertionError:
            print '"%s" tax table already exists, skipping' \
                  % tablename
            # Add method to update rate
            return 0
    return 1


def new_customer(book, row, USD):
    # Assume customer exists. Check id and company for new customer.
    cid = ''
    testcompany = None

    if 'company' in row.keys():
        company = row['company']
    else:
        print "Company missing in %s" % row
        return 1

        # Get company and ID from existing Customers
    try:
        testcompany = GetCustomers.iscustomer(company)
        cid = testcompany.GetID()
    except AttributeError:
        pass

    if 'id' in row.keys():
        if row['id'] != cid:
            cid = row['id']
            testid = book.CustomerLookupByID(cid)
        else:
            testid = testcompany
    else:
        testid = None

    if testid is None and testcompany is None:
        # If ID missing, create one
        if not cid:
            cid = book.CustomerNextID()
        # Customer not found, create.
        cust_acct = Customer(book, cid, USD, company)
    elif testid == testcompany:
        # ID and Company match, use.
        cust_acct = testid
    elif testid is not None and testcompany is None:
        # Customer found by ID, update Company
        cust_acct = testid
        cust_acct.SetCompany(company)
#    elif testid is None and testcompany is not None:
    else:
        if not cid:
            # Customer found by Company, ID missing, use.
            cust_acct = testcompany
        else:
            # Customer found by Company, update ID
            cust_acct = testcompany
            cust_acct.SetID(cid)

    try:
        assert (isinstance(cust_acct, Customer))
    except AssertionError:
        print cust_acct, " didn't work"
        return 2

    # Update the rest if they're in the row
    address = cust_acct.GetAddr()
    if 'name' in row.keys():
        address.SetName(row['name'])
    if 'addr1' in row.keys():
        address.SetAddr1(row['addr1'])
    if 'addr2' in row.keys():
        address.SetAddr2(row['addr2'])
    if 'addr3' in row.keys():
        address.SetAddr3(row['addr3'])
    if 'addr4' in row.keys():
        address.SetAddr4(row['addr4'])
    if 'phone' in row.keys():
        address.SetPhone(row['phone'])
    if 'fax' in row.keys():
        address.SetFax(row['fax'])
    if 'email' in row.keys():
        address.SetEmail(row['email'])
    if 'notes' in row.keys():
        cust_acct.SetNotes(str(row['notes']))

    address = cust_acct.GetShipAddr()
    if 'shipname' in row.keys():
        address.SetName(row['shipname'])
    if 'shipaddr1' in row.keys():
        address.SetAddr1(row['shipaddr1'])
    if 'shipaddr2' in row.keys():
        address.SetAddr2(row['shipaddr2'])
    if 'shipaddr3' in row.keys():
        address.SetAddr3(row['shipaddr3'])
    if 'shipaddr4' in row.keys():
        address.SetAddr4(row['shipaddr4'])
    if 'shipphone' in row.keys():
        address.SetPhone(row['shipphone'])

    if 'tax item' in row.keys():
        tablename = row['tax item']
        try:
            cust_taxtable = book.TaxTableLookupByName(tablename)
            assert (isinstance(cust_taxtable, TaxTable))
            if 'sales tax code' in row.keys():
                sales_tax_code = row['sales tax code']
                if sales_tax_code == 'Tax':
                    cust_acct.SetTaxTable(cust_taxtable)
                    cust_acct.SetTaxTableOverride(True)
                elif sales_tax_code == 'Non':
                    cust_acct.SetTaxTable(cust_taxtable)
                    cust_acct.SetTaxTableOverride(False)
                else:
                    print "%s sales tax code %s not recognized assume Tax \
                        and use TaxTable %s" \
                          % (company, sales_tax_code, tablename)
                    cust_acct.SetTaxTable(cust_taxtable)
                    cust_acct.SetTaxTableOverride(True)
        except:
            print "%s TaxTable %s does not exist. Customer Tax not updated" \
                  % (company, tablename)
            raise
    return 0


def new_vendor(book, row, USD):
    # Assume vendomer exists. Check id and company for new vendor.
    cid = ''
    testcompany = None

    if 'company' in row.keys():
        company = row['company']
    else:
        print "Company missing in %s" % row
        return 1
        # Get company and ID from existing Vendors
    try:
        testcompany = GetVendors.isvendor(company)
        cid = testcompany.GetID()
    except AttributeError:
        pass

    if 'id' in row.keys():
        if row['id'] != cid:
            cid = row['id']
            testid = book.VendorLookupByID(cid)
        else:
            testid = testcompany
    else:
        testid = None

    if testid is None and testcompany is None:
        # If ID missing, create one
        if not cid:
            cid = book.VendorNextID()
        # Vendor not found, create.
        vend_acct = Vendor(book, cid, USD, company)
    elif testid == testcompany:
        # ID and Company match, use.
        vend_acct = testid
    elif testid is not None and testcompany is None:
        # Vendor found by ID, update Company
        vend_acct = testid
        vend_acct.SetCompany(company)
    # elif testid is None and testcompany is not None:
    else:
        if not cid:
            # Vendor found by Company, ID missing, use.
            vend_acct = testcompany
        else:
            # Vendor found by Company, update ID
            vend_acct = testcompany
            vend_acct.SetID(cid)

    try:
        assert (isinstance(vend_acct, Vendor))
    except AssertionError:
        print vend_acct, " didn't work"
        return 2

    # Update the rest if they're in the row
    address = vend_acct.GetAddr()
    if 'name' in row.keys():
        address.SetName(row['name'])
    if 'addr1' in row.keys():
        address.SetAddr1(row['addr1'])
    if 'addr2' in row.keys():
        address.SetAddr2(row['addr2'])
    if 'addr3' in row.keys():
        address.SetAddr3(row['addr3'])
    if 'addr4' in row.keys():
        address.SetAddr4(row['addr4'])
    if 'phone' in row.keys():
        address.SetPhone(row['phone'])
    if 'fax' in row.keys():
        address.SetFax(row['fax'])
    if 'email' in row.keys():
        address.SetEmail(row['email'])
    if 'notes' in row.keys():
        vend_acct.SetNotes(str(row['notes']))

    if 'tax item' in row.keys():
        tablename = row['tax item']
        try:
            vend_taxtable = book.TaxTableLookupByName(tablename)
            assert (isinstance(vend_taxtable, TaxTable))
            if 'sales tax code' in row.keys():
                sales_tax_code = row['sales tax code']
                if sales_tax_code == 'Tax':
                    vend_acct.SetTaxTable(vend_taxtable)
                    vend_acct.SetTaxTableOverride(True)
                elif sales_tax_code == 'Non':
                    vend_acct.SetTaxTable(vend_taxtable)
                    vend_acct.SetTaxTableOverride(False)
                else:
                    print "%s sales tax code %s not recognized assume Tax \
                        and use TaxTable %s" \
                          % (company, sales_tax_code, tablename)
                    vend_acct.SetTaxTable(vend_taxtable)
                    vend_acct.SetTaxTableOverride(True)
        except:
            print "%s TaxTable %s does not exist. Vendor Tax not updated" \
                  % (company, tablename)
            raise
    return 0


def new_transaction(root, book, out, USD):
    # global existing_customers, existing_vendors
    # Assemble the invoice dictionary
    isinvoice = False
    isbill = False
    isinvpayment = False
    isbillpayment = False
    isentry = False

    for row in out:
        if 'type' in row.keys():
            rtype = row['type']
        else:
            rtype = ''

        if rtype:
            new_rtype, date_opened = get_rtype(row)
            new_rtype['entries'] = []

            if rtype == 'Invoice':
                isinvoice = True
            elif rtype == 'Payment':
                isinvpayment = True
            elif rtype == 'Sales Tax Payment':
                isentry = True
            elif rtype == 'Paycheck':
                continue
            elif rtype == 'Bill':
                isbill = True
            elif rtype == 'Credit':
                isbill = True
            elif rtype == 'Bill Pmt -CCard':
                isbillpayment = True
            else:
                isentry = True

        elif 'account' in row.keys():
            if row['account']:
                test, new_entry = get_entries(row, date_opened)
                if test == 'tax_table':
                    new_rtype['tax_table'] = new_entry['tax_table']
                    new_rtype['tax_rate'] = new_entry['price']
                elif test == 'entry':
                    new_rtype['entries'].append(new_entry)

                    # No account in total row, so process entries
        elif isentry:
            trans1 = Transaction(book)
            trans1.BeginEdit()
            trans1.SetCurrency(USD)
            if 'owner' in new_rtype.keys():
                trans1.SetDescription(new_rtype['owner'])
            trans1.SetDateEnteredTS(
                new_rtype['date_opened'] + datetime.timedelta(microseconds=1))
            trans1.SetDatePostedTS(
                new_rtype['date_opened'] + datetime.timedelta(microseconds=1))
            if 'num' in new_rtype.keys():
                trans1.SetNum(new_rtype['num'])
            if 'notes' in new_rtype.keys():
                trans1.SetNotes = new_rtype['notes']

            if new_rtype['account'] != '-SPLIT-':
                split1 = Split(book)
                split1.SetParent(trans1)
                # if new_rtype['type'] == 'Deposit':
                # new_rtype['amount'] = new_rtype['amount'].neg()
                split1.SetAccount(root.lookup_by_name(new_rtype['account']))
                # if split1.GetAccount() == ACCT_TYPE_EQUITY:
                # isequity = True
                # new_rtype['amount'] = new_rtype['amount'].neg()
                # else:
                # isequity = False
                split1.SetValue(new_rtype['amount'])
                if 'owner' in new_rtype.keys():
                    split1.SetMemo(new_rtype['owner'])
                    # split1.SetAction(get_action(new_rtype['type']))

            for entry in new_rtype['entries']:
                if 'amount' in entry.keys():
                    split1 = Split(book)
                    split1.SetParent(trans1)
                    # if isequity:
                    # entry['amount'] = entry['amount'].neg()
                    split1.SetValue(entry['amount'])
                    split1.SetAccount(root.lookup_by_name(entry['account']))
                    if 'description' in entry.keys():
                        split1.SetMemo(entry['description'])
            # split1.SetAction(get_action(new_rtype['type']))
            trans1.CommitEdit()

            isentry = False

        elif isinvpayment:
            try:
                owner = GetCustomers.iscustomer(new_rtype['owner'])
                assert (isinstance(owner, Customer))
            except AssertionError:
                print 'Customer %s does not exist; skipping' % \
                      new_rtype['owner']
                continue

            xfer_acc = root.lookup_by_name(new_rtype['account'])
            date_opened = new_rtype['date_opened']
            if 'notes' in new_rtype.keys():
                notes = new_rtype['notes']
            else:
                notes = ''
            if 'num' in new_rtype.keys():
                num = new_rtype['num']
            else:
                num = ''
            for entry in new_rtype['entries']:
                posted_acc = root.lookup_by_name(entry['account'])

                owner.ApplyPayment(None, None, posted_acc, xfer_acc,
                                   new_rtype['amount'], entry['amount'],
                                   date_opened, notes, num, False)
            isinvpayment = False

        elif isbillpayment:
            try:
                owner = GetVendors.isvendor(new_rtype['owner'])
                assert (isinstance(owner, Vendor))
            except AssertionError:
                print 'Vendor %s does not exist; skipping' % \
                      new_rtype['owner']
                continue

            xfer_acc = root.lookup_by_name(new_rtype['account'])
            date_opened = new_rtype['date_opened']
            if 'notes' in new_rtype.keys():
                notes = new_rtype['notes']
            else:
                notes = ''
            if 'num' in new_rtype.keys():
                num = new_rtype['num']
            else:
                num = ''
            for entry in new_rtype['entries']:
                posted_acc = root.lookup_by_name(entry['account'])

                owner.ApplyPayment(None, None, posted_acc, xfer_acc,
                                   new_rtype['amount'], entry['amount'],
                                   date_opened, notes, num, False)
            isbillpayment = False

        # new_customer.ApplyPayment(self, invoice, posted_acc, xfer_acc, amount,
        # exch, date, memo, num)
        # new_customer.ApplyPayment(None, None, a2, a6, GncNumeric(100,100),
        # GncNumeric(1), datetime.date.today(), "", "", False)

        # invoice_customer.ApplyPayment(None, a6, GncNumeric(7,100),
        # GncNumeric(1), datetime.date.today(), "", "")

        elif isbill:
            # put item on entries!!!
            # Accumulate splits
            # QuickBooks Journal has a total row after splits,
            # which is used to detect the end of splits.
            try:
                owner = GetVendors.isvendor(new_rtype['owner'])
                assert (isinstance(owner, Vendor))
            except AssertionError:
                print 'Vendor %s does not exist; skipping' % \
                      new_rtype['owner']
                continue

            try:
                cid = book.BillNextID(owner)
            # save Bill ID and tax rate for xml overlay.
            # ReplaceTax.bill(cid, new_rtype['tax_rate'])
            except:
                raise

            bill_vendor = Bill(book, cid, USD, owner)
            vendor_extract = bill_vendor.GetOwner()
            assert (isinstance(vendor_extract, Vendor))
            assert (vendor_extract.GetName() == owner.GetName())

            if new_rtype['type'] == 'Credit':
                bill_vendor.SetIsCreditNote(True)

            bill_vendor.SetDateOpened(new_rtype['date_opened'])

            if 'notes' in new_rtype.keys():
                bill_vendor.SetNotes(new_rtype['notes'])

            if 'num' in new_rtype.keys():
                bill_vendor.SetBillingID(new_rtype['num'])

            if 'tax_table' in new_rtype.keys():
                tax_table = book.TaxTableLookupByName(new_rtype['tax_table'])
                assert (isinstance(tax_table, TaxTable))

            # Add the entries
            for entry in new_rtype['entries']:
                # skip entries that link COGS and Billentory
                if 'quantity' not in entry.keys():
                    continue
                bill_entry = Entry(book, bill_vendor)

                account = root.lookup_by_name(entry['account'])
                assert (isinstance(account, Account))
                bill_entry.SetBillAccount(account)

                if 'tax_table' in new_rtype.keys():
                    bill_entry.SetBillTaxTable(tax_table)
                    bill_entry.SetBillTaxIncluded(False)
                else:
                    bill_entry.SetBillTaxable(False)

                if 'description' in entry.keys():
                    bill_entry.SetDescription(entry['description'])
                bill_entry.SetQuantity(entry['quantity'])
                bill_entry.SetBillPrice(entry['price'])
                bill_entry.SetDateEntered(entry['date'])
                bill_entry.SetDate(entry['date'])
                if 'notes' in entry.keys():
                    bill_entry.SetNotes(entry['notes'])

            isbill = False

            # Post bill
            account = root.lookup_by_name(new_rtype['account'])
            assert (isinstance(account, Account))
            bill_vendor.PostToAccount(account, new_rtype['date_opened'], new_rtype['date_opened'],
                                      str(new_rtype['owner']), True, False)

        elif isinvoice:
            # put item on entries!!!
            # Accumulate splits
            # QuickBooks Journal has a total row after splits,
            # which is used to detect the end of splits.
            try:
                owner = GetCustomers.iscustomer(new_rtype['owner'])
                assert (isinstance(owner, Customer))
            except AssertionError:
                print 'Customer %s does not exist; skipping' % \
                      new_rtype['owner']
                continue

            try:
                cid = book.InvoiceNextID(owner)
            # save Invoice ID and tax rate for xml overlay.
            # ReplaceTax.invoice(cid, new_rtype['tax_rate'])
            except:
                raise

            invoice_customer = Invoice(book, cid, USD, owner)
            customer_extract = invoice_customer.GetOwner()
            assert (isinstance(customer_extract, Customer))
            assert (customer_extract.GetName() == owner.GetName())

            invoice_customer.SetDateOpened(new_rtype['date_opened'])

            if 'notes' in new_rtype.keys():
                invoice_customer.SetNotes(new_rtype['notes'])

            if 'num' in new_rtype.keys():
                invoice_customer.SetBillingID(new_rtype['num'])

            if 'tax_table' in new_rtype.keys():
                tax_table = book.TaxTableLookupByName(new_rtype['tax_table'])
                assert (isinstance(tax_table, TaxTable))

            # assert( not isinstance( \
            # book.InvoiceLookupByID(new_rtype['id']), Invoice))

            # Add the entries
            for entry in new_rtype['entries']:
                invoice_entry = Entry(book, invoice_customer)

                account = root.lookup_by_name(entry['account'])
                assert (isinstance(account, Account))
                invoice_entry.SetInvAccount(account)

                if 'tax_table' in new_rtype.keys():
                    invoice_entry.SetInvTaxTable(tax_table)
                    invoice_entry.SetInvTaxIncluded(False)
                else:
                    invoice_entry.SetInvTaxable(False)

                invoice_entry.SetDescription(entry['description'])
                invoice_entry.SetQuantity(entry['quantity'])
                invoice_entry.SetInvPrice(entry['price'])
                invoice_entry.SetDateEntered(entry['date'])
                invoice_entry.SetDate(entry['date'])
                if 'notes' in entry.keys():
                    invoice_entry.SetNotes(entry['notes'])

            isinvoice = False

            # Post invoice
            account = root.lookup_by_name(new_rtype['account'])
            assert (isinstance(account, Account))
            invoice_customer.PostToAccount(account, new_rtype['date_opened'], new_rtype['date_opened'],
                                           str(new_rtype['owner']), True, False)
            # ReplaceTax.replace(gnc_file.path)

    return 0


def pad_number(num):
    return "%(number)06d" % {'number': num}


def mapqb2gnc(reader, usemap):
    # map QuickBooks csv data to GnuCash csv data
    out = []
    for row in reader:
        outrow = {}
        for field in usemap:
            test = row[str(field)]
            if usemap[str(field)] == 'not mapped' or \
                    not test:
                continue
            elif str((usemap[str(field)])) in outrow:
                if usemap[str(field)] in commakey:
                    outrow[usemap[str(field)]] += ', ' + \
                                                  row[str(field)]
                elif usemap[str(field)] in spacekey:
                    outrow[usemap[str(field)]] += ' ' + row[str(field)]
                elif usemap[str(field)] in linekey:
                    outrow[usemap[str(field)]] += '\n' + field + \
                                                  ': ' + row[str(field)]
                elif usemap[str(field)] in linekey:
                    outrow[usemap[str(field)]] = field + ': ' + row[str(field)]
            else:
                outrow[usemap[str(field)]] = row[str(field)]
        out.append(outrow)
    return out


class GetCustomers(object):
    @staticmethod
    def iscustomer(company):
        book = GncFile.book
        query = Query()
        query.search_for('gncCustomer')
        query.set_book(book)

        for result in query.run():
            customer = Customer(instance=result)
            if customer.GetName() == company:
                query.destroy()
                return customer
        query.destroy()
        return None


class GetVendors(object):
    @staticmethod
    def isvendor(company):
        book = GncFile.book
        query = Query()
        query.search_for('gncVendor')
        query.set_book(book)

        for result in query.run():
            vendor = Vendor(instance=result)
            if vendor.GetName() == company:
                query.destroy()
                return vendor
        query.destroy()
        return None


def gnc_numeric_from(any_value):
    if str(any_value).find('%') != -1:
        decimal_value = Decimal(any_value.strip('%'))
        ispercent = True
    else:
        decimal_value = Decimal(any_value)
        ispercent = False

    sign, digits, exponent = decimal_value.as_tuple()

    # convert decimal digits to a fractional numerator
    # equivlent to
    # numerator = int(''.join(digits))
    # but without the wated conversion to string and back,
    # this is probably the same algorithm int() uses
    numerator = 0
    TEN = int(Decimal(0).radix())  # this is always 10
    numerator_place_value = 1
    # add each digit to the final value multiplied by the place value
    # from least significant to most sigificant
    for i in xrange(len(digits) - 1, -1, -1):
        numerator += digits[i] * numerator_place_value
        numerator_place_value *= TEN

    if decimal_value.is_signed():
        numerator = -numerator

    # if the exponent is negative, we use it to set the denominator
    if exponent < 0:
        denominator = TEN ** (-exponent)
    # if the exponent isn't negative, we bump up the numerator
    # and set the denominator to 1
    else:
        numerator *= TEN ** exponent
        denominator = 1
    if ispercent:
        x = 100000 / denominator
        numerator *= x
        denominator *= x

    return GncNumeric(numerator, denominator)


def gnc_numeric_from_decimal(decimal_value):
    sign, digits, exponent = decimal_value.as_tuple()

    # convert decimal digits to a fractional numerator
    # equivlent to
    # numerator = int(''.join(digits))
    # but without the wated conversion to string and back,
    # this is probably the same algorithm int() uses
    numerator = 0
    TEN = int(Decimal(0).radix())  # this is always 10
    numerator_place_value = 1
    # add each digit to the final value multiplied by the place value
    # from least significant to most sigificant
    for i in xrange(len(digits) - 1, -1, -1):
        numerator += digits[i] * numerator_place_value
        numerator_place_value *= TEN

    if decimal_value.is_signed():
        numerator = -numerator

    # if the exponent is negative, we use it to set the denominator
    if exponent < 0:
        denominator = TEN ** (-exponent)
    # if the exponent isn't negative, we bump up the numerator
    # and set the denominator to 1
    else:
        numerator *= TEN ** exponent
        denominator = 1

    return GncNumeric(numerator, denominator)


def get_rtype(row):
    new_rtype = {}
    if 'num' in row.keys():
        new_rtype['num'] = row['num']
    dt_str = '%m/%d/%Y'
    date_opened = datetime.datetime.strptime(row['date_opened'], dt_str)
    new_rtype['date_opened'] = date_opened
    if 'owner' in row.keys():
        new_rtype['owner'] = row['owner']
    elif 'notes' in row.keys():
        new_rtype['owner'] = row['notes']

    new_rtype['account'] = row['account']

    if 'notes' in row.keys():
        new_rtype['notes'] = row['notes']
    if 'paid' in row.keys():
        if row['paid'] == 'Paid':
            new_rtype['paid'] = True
    else:
        new_rtype['paid'] = False

    if 'amount' in row.keys():
        new_rtype['amount'] = gnc_numeric_from(row['amount'])
    new_rtype['type'] = row['type']

    return new_rtype, date_opened


def get_entries(row, date_opened):
    entry = {}
    if row['account'] == 'Sales Tax Payable' and 'price' in row.keys():
        entry['tax_table'] = row['item']
        entry['price'] = gnc_numeric_from(row['price'])
        TaxRates.israte(entry['tax_table'], entry['price'])
        return 'tax_table', entry

    if 'description' in row.keys():
        entry['description'] = row['description']
    elif 'owner' in row.keys():
        entry['description'] = row['owner']
    elif 'notes' in row.keys():
        entry['description'] = row['notes']

    if 'quantity' in row.keys():
        entry['quantity'] = \
            gnc_numeric_from(abs(Decimal(row['quantity'])))
        if 'price' in row.keys():
            entry['price'] = gnc_numeric_from(row['price'])

    entry['date'] = date_opened
    if row['account'] == 'Sales Tax Payable' and 'item' in row.keys():
        entry['account'] = row['item']
    else:
        entry['account'] = row['account']

    if 'amount' in row.keys():
        entry['amount'] = gnc_numeric_from(row['amount'])

    if 'notes' in row.keys():
        entry['notes'] = row['notes']

    return 'entry', entry


def get_action(rtype):
    actionmap = {"Deposit": "Deposit",
                 # "Withdraw" ,
                 "Check": "Check",
                 # "Interest" ,
                 # "ATM Deposit" ,
                 # "ATM Draw" ,
                 # "Teller" ,
                 "Credit Card Charge": "Charge",
                 "Payment": "Payment",
                 "Sales Tax Payment": "Payment",
                 "Credit Card Credit": "Credit",
                 "Credit": "Credit",
                 "Transfer": "Transfer",
                 "Invoice": "Invoice",
                 "Bill": "Bill",
                 "Bill Pmt -CCard": "Payment",
                 "General Journal": "Credit",
                 "Inventory Adjust": 'Equity'}

    if str(rtype) in actionmap.keys():
        return list(actionmap[str(rtype)])
    else:
        print rtype
        raise Exception('Action Type does not exist')


class GncFile(object):
    s = ''
    book = ''
    root = ''
    commod_table = ''
    USD = ''
    path = ''
    status = False
    nosave = True

    @classmethod
    def gnc_open(cls, path=None):

        if GncFile.status:
            return GncFile.status
        else:
            if path is not None:
                GncFile.path = path
            else:
                path = GncFile.path
            if not os.path.exists(path):
                print """GnuCash file (%s) does not exist. """ % path
                raise Exception('GnuCash file missing')
            if os.path.exists(path + '.LCK'):
                print """Lock file exists. Is GNUCash running?\n"""
                raise Exception('GnuCash locked')

            try:
                GncFile.s = Session(GncFile.path, is_new=False)
            except:
                raise Exception('Could not open GnuCash file')

            GncFile.book = GncFile.s.book
            GncFile.root = GncFile.book.get_root_account()
            GncFile.commod_table = GncFile.book.get_table()
            GncFile.USD = GncFile.commod_table.lookup('CURRENCY', 'USD')

            GncFile.status = True
            return GncFile.status

    @classmethod
    def gnc_save(cls):
        if GncFile.status and GncFile.nosave:
            GncFile.s.save()

    @classmethod
    def gnc_end(cls):
        GncFile.s.end()
        GncFile.status = False


class ReplaceTax(object):
    invoice_list = {}

    @classmethod
    def invoice(cls, invoiceid, rate):
        ReplaceTax.invoice_list[str(invoiceid)] = rate.to_string()

    @classmethod
    def replace(cls, outfile):
        GncFile.gnc_save()
        GncFile.gnc_end()

        tree = et.parse(outfile)
        root = tree.getroot()

        xmlgnc = root.nsmap['gnc']
        xmlgncbook = '{' + xmlgnc + '}book'
        xmlgnctaxtable = '{' + xmlgnc + '}GncTaxTable'
        xmlgnctaxtableentry = '{' + xmlgnc + '}GncTaxTableEntry'
        xmlgncinvoice = '{' + xmlgnc + '}GncInvoice'
        xmlgncentry = '{' + xmlgnc + '}GncEntry'

        xmlinvoice = root.nsmap['invoice']
        xmlinvoice_id = '{' + xmlinvoice + '}id'
        xmlinvoice_guid = '{' + xmlinvoice + '}guid'

        xmltaxtable = root.nsmap['taxtable']
        xmltaxtable_name = '{' + xmltaxtable + '}name'
        xmltaxtable_guid = '{' + xmltaxtable + '}guid'
        xmltaxtable_entries = '{' + xmltaxtable + '}entries'

        xmltte_amount = '{' + root.nsmap['tte'] + '}amount'

        xmlentry = root.nsmap['entry']
        xmlentry_invoice = '{' + xmlentry + '}invoice'
        xmlentry_itaxtable = '{' + xmlentry + '}i-taxtable'

        book = root.find(xmlgncbook)
        for cid in ReplaceTax.invoice_list:
            for child in book.iter(xmlgncinvoice):
                if child.find(xmlinvoice_id).text == str(cid):
                    guid = child.find(xmlinvoice_guid).text

            for child in book.iter(xmlgncentry):
                if child.find(xmlentry_invoice).text == guid:
                    itaxtable = child.find(xmlentry_itaxtable).text

            for child in book.iter(xmlgnctaxtable):
                if child.find(xmltaxtable_guid).text == itaxtable:
                    name = child.find(xmltaxtable_name)
                    entries = child.find(xmltaxtable_entries)
                    for entry in entries.findall(xmlgnctaxtableentry):
                        print name.text, itaxtable
                        amount = entry.find(xmltte_amount)
                        print amount.text
                        amount.text = str(ReplaceTax.invoice_list[str(cid)])
                        print amount.text

                        # out = ET.tostring(tree, xml_declaration=True, pretty_print=True,
                        # encoding='utf-8')
                        # print out
        tree.write(outfile,
                   xml_declaration=True,
                   pretty_print=True,
                   encoding='utf-8')

        GncFile.gnc_open()


if __name__ == "__main__":

    sys.exit(main())
