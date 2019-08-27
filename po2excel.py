#!/usr/bin/python
# coding=utf-8

import re
import xlwt

EXCEL_SHEET_NAME = 'sheel1'

COLUMN_INDEX_MSGCTXT = 0
COLUMN_INDEX_MSGID = 1
COLUMN_INDEX_MSGSTR = 2

COLUMN_TITLE_MSGCTXT = 'msgctxt'
COLUMN_TITLE_MSGID = 'msgid'
COLUMN_TITLE_MSGSTR = 'msgstr'

TITLE_HEADER = {
    COLUMN_TITLE_MSGCTXT: COLUMN_TITLE_MSGCTXT,
    COLUMN_TITLE_MSGID: COLUMN_TITLE_MSGID,
    COLUMN_TITLE_MSGSTR: COLUMN_TITLE_MSGSTR,
}

PO_PATTERN = r'^(.*) "(.*)"'


def __is_valid_po_entry(po_entry, full_export):
    '''
    判断一个entry是否有效
    '''
    if (not full_export) and (len(po_entry.get(COLUMN_TITLE_MSGSTR)) > 0):
        return False
    for value in po_entry.values():
        if len(value) > 0:
            return True
    return False


def __write_entry_to_excel(sheet, po_entry):
    '''
    将一个entry写入excel
    '''
    nrow = len(sheet.get_rows())
    sheet.write(nrow, COLUMN_INDEX_MSGCTXT, po_entry.get(COLUMN_TITLE_MSGCTXT))
    sheet.write(nrow, COLUMN_INDEX_MSGID, po_entry.get(COLUMN_TITLE_MSGID))
    sheet.write(nrow, COLUMN_INDEX_MSGSTR, po_entry.get(COLUMN_TITLE_MSGSTR))


def __parse_po_file(sheet, po_file, full_export):
    '''
    解析po文件
    '''
    po_entry = dict()
    for line_text in po_file:
        result = re.match(PO_PATTERN, line_text)
        if result:
            key = result.group(1)
            value = result.group(2)
            po_entry[key] = value

            # 每遇见msgstr写入一次po_entry
            if key == COLUMN_TITLE_MSGSTR:
                if __is_valid_po_entry(po_entry, full_export):
                    __write_entry_to_excel(sheet, po_entry)
                po_entry = dict()


def po2excel(in_file, out_file, full_export):
    '''
    Po文件转换成Excel文件
    '''
    with open(in_file, 'r', encoding='UTF-8') as po_file:
        work_book = xlwt.Workbook()
        sheet = work_book.add_sheet(EXCEL_SHEET_NAME)
        __write_entry_to_excel(sheet, TITLE_HEADER)
        __parse_po_file(sheet, po_file, full_export)
        work_book.save(out_file)
