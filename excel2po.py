#!/usr/bin/python
# coding=utf-8

import xlrd

EXCEL_SHEET_NAME = 'sheel1'

COLUMN_INDEX_MSGCTXT = 0
COLUMN_INDEX_MSGID = 1
COLUMN_INDEX_MSGSTR = 2

PO_VALUE_FORMAT = 'msgctxt "%s\nmsgid "%s"\nmsgstr "%s"\n\n'


def write_entry_to_po(po_file, po_entry, ignore_length_check):
    '''
    将一个entry写入po
    '''
    msgctxt = po_entry[COLUMN_INDEX_MSGCTXT]
    msgid = po_entry[COLUMN_INDEX_MSGID]
    msgstr = po_entry[COLUMN_INDEX_MSGSTR]
    po_file.write(PO_VALUE_FORMAT % (msgctxt, msgid, msgstr))

    if (len(str(msgid)) != len(str(msgstr))) and (not ignore_length_check):
        print("[Warning] Length is not equal. msgctxt =", msgctxt)


def excel2po(in_file, out_file, ignore_length_check):
    '''
    Eecel文件转换成po文件
    '''
    with open(out_file, 'w', encoding='UTF-8') as po_file:
        workbook = xlrd.open_workbook(in_file)
        sheet = workbook.sheet_by_name(EXCEL_SHEET_NAME)
        for nrow in range(sheet.nrows):
            if nrow > 0:
                po_entry = sheet.row_values(nrow)
                write_entry_to_po(po_file, po_entry, ignore_length_check)
