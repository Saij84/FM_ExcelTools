from openpyxl import load_workbook

filePath = r"C:\tools\FM_ExcelTools\testData\SampleData.xlsx"

wb = load_workbook(filePath)
salesOrderSheet = wb["SalesOrders"]


def getRow(workbook, collumRange="abcd", rowRange=(1, 1)):
    rowKey = list()
    for digit in range(rowRange[0], rowRange[1]+1):
        for letter in collumRange:
            cell = "{}{}".format(letter, digit)
            rowKey.append(workbook[cell].value)
    print(rowKey)
    return ("".join(rowKey))


def conformKey(key1, key2):
    lenKey1 = len(key1)
    lenKey2 = len(key2)

    if lenKey1 > lenKey2:
        addSpaceMult = lenKey1 - lenKey2
        newKey2 = "{}{}".format(key2, (" " * addSpaceMult))
        return zip(key1, newKey2)
    else:
        addSpaceMult = lenKey2 - lenKey1
        newKey1 = "{}{}".format(key1, (" " * addSpaceMult))
        return zip(newKey1, key2)


def compareKeys(key1, key2):
    if key1 == key2:
        print("same")
    else:
        keyZipList = conformKey(key1, key2)
        keyDiff = ""

        for letter1, letter2 in keyZipList:
            if not letter1 == letter2:
                keyDiff += letter2
        print(keyDiff)


key1 = getRow(salesOrderSheet, "fffF", rowRange=(1, 1))
key2 = getRow(salesOrderSheet, "fbcaf", rowRange=(1, 1))
compareKeys(key1, key2)
