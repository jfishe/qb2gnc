#!/usr/bin/python

import sys
import argparse
import csv
import re


def main():
    parser = argparse.ArgumentParser(description='Remove numbers from accounts')
    parser.add_argument('infile',
                        metavar='infile.csv',
                        nargs=1,
                        help='QuickBooks export file')
    parser.add_argument('outfile',
                        metavar='outfile.csv',
                        nargs=1,
                        help='Export file with leading accoun numbers removed')
    args = parser.parse_args()

    reader = csv.DictReader(open(args.infile[0], 'r'), dialect='excel')
    outfieldnames = reader.fieldnames

    out = []
    outrow = {}
    rem1 = '^\d+ '
    rem2 = ':\d+ '
    for row in reader:
        for name in outfieldnames:
            if name == 'Account':
                outrow[str(name)] = re.sub(rem2, ':',
                                           re.sub(rem1, '', row[str(name)]))
            elif name == 'Split' and row[str(name)] != '-SPLIT-':
                outrow[str(name)] = re.sub(rem2, ':',
                                           re.sub(rem1, '', row[str(name)]))
            elif name == 'COGS Account':
                outrow[str(name)] = re.sub(rem2, ':',
                                           re.sub(rem1, '', row[str(name)]))
            elif name == 'Memo':
                outrow[str(name)] = re.sub(rem2, ':',
                                           re.sub(rem1, '', row[str(name)]))
            else:
                outrow[str(name)] = row[str(name)]
        out.append(dict(outrow))

    writer = csv.DictWriter(open(args.outfile[0], 'w'),
                            fieldnames=outfieldnames,
                            dialect='excel')
    writer.writeheader()
    writer.writerows(out)


if __name__ == "__main__":
    sys.exit(main())
