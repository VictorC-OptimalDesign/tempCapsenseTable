# === IMPORTS ==================================================================
import enum
import glob
import math
import operator
import os
import shutil
import string
import typing
import xlsxwriter

# === GLOBAL CONSTANTS =========================================================

NAME_KEYWORDS: typing.List[str] = (
    'const',
    'char',
    'bool',
    'unsigned',
    'int',
    'uint8_t',
    'uint16_t',
    'uint32_t',
    'uint64_t',
    'int8_t',
    'int16_t',
    'int32_t',
    'int64_t',
    '=',
)


# === FUNCTIONS ================================================================

def __containsNameKeyword(line: str) -> bool:
    words = line.split(' ')
    for word in words:
        if word in NAME_KEYWORDS:
            return True
    return False
    
def __filterName(line: str) -> str:
    words = line.split(' ')
    for word in words:
        if word not in NAME_KEYWORDS:
            return word
    return ''

def __parseData(line: str) -> typing.List[int]:
    words = line.split(',')
    words = [word.strip() for word in words]
    for i, word in enumerate(words):
        try:
            words[i] = int(word)
        except ValueError:
            if not word or word == '}':
                del words[i]
                continue
            else:
                return []
    return words

def __writeDataToWorksheet(ws: xlsxwriter.Workbook.worksheet_class, col: int, name: str, data: typing.List[int]):
    row: int = 0
    ws.write(row, col, name)
    row += 1
    for i, datum in enumerate(data):
        ws.write_number(row + i, col, datum)
    

def __process():
    wb: xlsxwriter.Workbook = xlsxwriter.Workbook('tempCapsenseTable.xlsx')
    for fileName in glob.glob('*.txt'):
        ws: xlsxwriter.Workbook.worksheet_class = wb.add_worksheet(fileName)
        print('processing {0}...'.format(fileName))
        filePath = os.path.join(os.getcwd(), fileName)
        with open(filePath, 'r') as file:
            readLines = file.readlines()
        file.close()
        processingData: bool = False
        col: int = 0
        name: str = ''
        data: typing.List[int] = []
        for line in readLines:
            if not processingData:
                if __containsNameKeyword(line):
                    name = __filterName(line)
                    processingData = True
            else:
                if __containsNameKeyword(line):
                    __writeDataToWorksheet(ws, col, name, data)
                    name = __filterName(line)
                    processingData = True
                    col += 1
                    data = []
                else:
                    data.extend(__parseData(line))
        __writeDataToWorksheet(ws, col, name, data)
    wb.close()


# === MAIN =====================================================================

if __name__ == "__main__":
    __process()
else:
    print("ERROR: tempCapsenseTable needs to be the calling python module!")
    