#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
UE4 Localization Tools.

Usage:
    ue4-localization-tools.py po2excel <in_file> <out_file> [-f | --full_export]
    ue4-localization-tools.py excel2po <in_file> <out_file> [-i | --ignore_length_check]
    ue4-localization-tools.py (-h | --help)
    ue4-localization-tools.py (-v | --version)

Options:
    -h --help                   # Show help.
    -v --version                # Show version.
    -f --full_export            # Export all text to excel.
    -i --ignore_length_check    # Ignore length check, just quietly export it.
'''  # nopep8


from docopt import docopt
from excel2po import excel2po
from po2excel import po2excel


def exec_excel2po(arguments):
    in_file = arguments['<in_file>']
    out_file = arguments['<out_file>']
    ignore_length_check = arguments['--ignore_length_check']
    excel2po(in_file, out_file, ignore_length_check)


def exec_po2excel(arguments):
    in_file = arguments['<in_file>']
    out_file = arguments['<out_file>']
    full_export = arguments['--full_export']
    po2excel(in_file, out_file, full_export)


if __name__ == '__main__':
    arguments = docopt(__doc__, version='UE4 Localization Tools 1.0')
    if arguments['po2excel']:
        exec_po2excel(arguments)
    elif arguments['excel2po']:
        exec_excel2po(arguments)
