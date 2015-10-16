#!/usr/bin/python

import argparse
import csv

from operator import itemgetter
import sys


toplevelaccounts = ['Asset', 'Equity', 'Expense', 'Income', 'Liability']

missingaccounts = ['Fixed Asset', 'Other Current Liability', 'Other Income', 'Other Expense', 'Non-Posting',
                   'Other Current Asset', 'Cost of Goods Sold']

infieldnames = ['Active Status', 'Account', 'Type', 'Balance Total', 'Description', 'Tax Line']
outfieldnames = ['type', 'full_name', 'name', 'code', 'description', 'color', 'notes', 'commoditym', 'commodityn',
                 'hidden', 'tax', 'place_holder']

fieldmap = {str(outfieldnames[0]): infieldnames[2],
            str(outfieldnames[1]): infieldnames[1],
            str(outfieldnames[2]): 'oname',
            str(outfieldnames[3]): 'not mapped',
            str(outfieldnames[4]): infieldnames[4],
            str(outfieldnames[5]): 'not mapped',
            str(outfieldnames[6]): 'not mapped',
            str(outfieldnames[7]): 'USD',
            str(outfieldnames[8]): 'CURRENCY',
            str(outfieldnames[9]): infieldnames[0],
            str(outfieldnames[10]): 'F',
            str(outfieldnames[11]): 'F'}

accountmap = {'Asset': 'ASSET',
              'Accounts Payable': 'PAYABLE',
              'Accounts Receivable': 'RECEIVABLE',
              'Bank': 'BANK',
              'Cost of Goods Sold': 'EXPENSE',
              'Credit Card': 'CREDIT',
              'Equity': 'EQUITY',
              'Expense': 'EXPENSE',
              'Fixed Asset': 'ASSET',
              'Income': 'INCOME',
              'Liability': 'LIABILITY',
              'Non-Posting': 'ASSET',
              'Other Current Asset': 'ASSET',
              'Other Current Liability': 'LIABILITY',
              'Other Income': 'INCOME',
              'Other Expense': 'EXPENSE'
              }

useaccount = {'Asset': 'no',
              'Accounts Payable': 'no',
              'Accounts Receivable': 'no',
              'Bank': 'no',
              'Cost of Goods Sold': 'yes',
              'Credit Card': 'no',
              'Equity': 'no',
              'Expense': 'no',
              'Fixed Asset': 'yes',
              'Income': 'no',
              'Liability': 'no',
              'Non-Posting': 'yes',
              'Other Current Asset': 'yes',
              'Other Current Liability': 'yes',
              'Other Income': 'yes',
              'Other Expense': 'yes'}

parentmap = {'Asset': 'Assets:',
             'Accounts Payable': 'Liabilities:',
             'Accounts Receivable': 'Assets:',
             'Bank': 'Assets:',
             'Cost of Goods Sold': 'Expenses:',
             'Credit Card': 'Liabilities:',
             'Equity': 'Equity:',
             'Expense': 'Expenses:',
             'Fixed Asset': 'Assets:',
             'Income': 'Income:',
             'Liability': 'Liabilities:',
             'Non-Posting': 'Assets:',
             'Other Current Asset': 'Assets:',
             'Other Current Liability': 'Liabilities:',
             'Other Income': 'Income:',
             'Other Expense': 'Expenses:'}


def getname(atype, usetype, name, parent):
    if atype == name or usetype == 'no':
        fullname = parent + name
    else:
        fullname = parent + atype + ':' + name
    if name.find(':') != -1:
        rename = name.split(':')
        nlen = len(rename) - 1
        name = rename[nlen]
    return fullname, name


def lfind(f, seq):
    """Return first item in sequence where f(item) == True."""
    for item in seq:
        if f(item):
            return item


def main():
    # get infile and outfile from arguments
    parser = argparse.ArgumentParser(description='Convert QuickBooks accounts, \
        to gnucash importable accounts csv')
    parser.add_argument('infile',
                        metavar='infile.csv',
                        nargs=1,
                        help='QuickBooks exported account csv file')
    parser.add_argument('outfile',
                        metavar='outfile.csv',
                        nargs=1,
                        help='gnucash account import file')
    args = parser.parse_args()
    infile = args.infile[0]
    outfile = args.outfile[0]

    # Add top level accounts to output list
    out = []
    for iout in toplevelaccounts:
        name = parentmap[str(iout)]
        name = name.split(':')[0]
        out.append(dict(name=name, commodityn='CURRENCY', commoditym='USD', tax='F', full_name=name, hidden='F',
                        type=accountmap[str(iout)], place_holder='T', description=''))
    # Add missing mid-level accounts to output list
    for iout in missingaccounts:
        parent = parentmap[str(iout)]
        name = iout
        fullname = parent + name
        out.append(dict(name=name, commodityn='CURRENCY', commoditym='USD', tax='F', full_name=fullname, hidden='F',
                        type=accountmap[str(iout)], place_holder='F', description=''))

    # Initialize the output list and write the header
    with open(outfile, 'w') as csvout:
        writer = csv.DictWriter(csvout, fieldnames=outfieldnames, dialect='excel')
        writer.writeheader()

        # Process the input file
        with open(infile) as csvfile:
            # with open('emerald_accounts.CSV') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:

                # reserve variables for loop tests in fieldmap
                atype = ' '
                oname = ' '
                outrow = {}

                # Map the output field names using fieldmap
                for iout in outfieldnames:
                    field = fieldmap[str(iout)]

                    # Determine account type for use in later fields
                    # and assign re-mapped account type to output dictionary
                    if iout == outfieldnames[0]:
                        atype = row[str(field)]
                        outrow[str(iout)] = accountmap[str(atype)]

                    # Determine name, full name, parent account and whether type added to full name.
                    # Assign generated full name to output dictionary.
                    elif iout == outfieldnames[1] and atype != ' ':
                        usetype = useaccount[str(atype)]
                        name = row[str(field)]
                        # account = accountmap[ str(type) ]
                        parent = parentmap[str(atype)]
                        ofull_name, oname = getname(atype, usetype, name, parent)
                        outrow[str(iout)] = ofull_name

                    # Assign account name to output dictionary
                    elif iout == outfieldnames[2] and oname != ' ':
                        outrow[str(iout)] = oname

                    # Assign Description to output dictionary
                    elif iout == outfieldnames[4] and oname != ' ':
                        outrow[str(iout)] = row[str(field)]

                    # commoditym commodityn hidden tax place_holder
                    elif iout == outfieldnames[9]:
                        if row[str(field)] == 'Active':
                            outrow[str(iout)] = 'F'
                        else:
                            outrow[str(iout)] = 'T'
                    elif field == 'not mapped':
                        continue
                    else:
                        outrow[str(iout)] = field

                        # Append row to output list
                out.append(outrow)
        sortedlist = sorted(out, key=itemgetter('full_name'))
        writer.writerows(sortedlist)


if __name__ == '__main__':
    sys.exit(main())