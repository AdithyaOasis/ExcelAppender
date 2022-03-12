
import openpyxl


def readModifiedRow(sheet_obj, row, colList):
    new_val = [0]*len(colList)
    for i in range(len(colList)):
        if colList[i] != 100:
            new_val[i] = (sheet_obj.cell(row, colList[i]+1).value)
        else:
            new_val[i] = ' '
    new_val[8] = float(sheet_obj.cell(row, 82).value) + \
        float(sheet_obj.cell(row, 84).value)
    new_val[7] = getPurchasePrice(sheet_obj, row)
    new_val[12] = getPurchaseDate(sheet_obj, row)
    new_val[13] = getExpiryDate(sheet_obj, row)
    new_val[22] = (float(new_val[10]) - new_val[7]) / \
        float(new_val[10])  # ((MRP-Purchase.Price)/MRP)
    # (MRP - PurchasePrice)
    new_val[23] = float(new_val[10]) - new_val[7]
    new_val[22] = str(new_val[22])[:7]
    new_val[23] = str(new_val[23])[:7]
    return new_val


def getPurchasePrice(sheet_obj, row):
    val = float(sheet_obj.cell(row, 51).value) * \
        (100 - float(sheet_obj.cell(row, 50).value)) / 10000
    val = val * (100 + float(sheet_obj.cell(row, 82).value) +
                 float(sheet_obj.cell(row, 84).value))
    return val

# (in d-monthname-yy format) InvDate[5] or InvDay[6] + enum[InvMonth][7] + InvYear(only tens and units)[8]


def getPurchaseDate(sheet_obj, row):
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
              'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    val = sheet_obj.cell(
        row, 7).value + '-' + months[int(sheet_obj.cell(row, 8).value)-1] + '-' + str(int(sheet_obj.cell(row, 9).value))[2:]
    return val

# (in dd/mm/yyyy format)ExpDate[41]   (replace - with /) or expDay[42] + Expmonth[43] + Expyear[44]


def getExpiryDate(sheet_obj, row):
    val = sheet_obj.cell(row, 43).value + '/' + str(int(sheet_obj.cell(row,
                                                                       44).value)) + '/' + str(int(sheet_obj.cell(row, 45).value))
    return val


def printRows(rows):
    headings = ['Invoice number-', 'Code-\t', 'Name-\t', 'Type-\t', 'Packing Number-', 'Quantity-', 'Free-\t', 'PurchacePrice-',
                'GST-\t', 'Discount-', 'MRP-\t', 'BatchNo-\t', 'PurchaseDate-', 'ExpDate\t', 'Vendor Name- ', 'Manufacturer Name-',
                'Compostitions-', 'Rack No-', 'HsnCode-', 'Schedule Type-', 'GST includedIn Rt-', 'department-\t', 'Margin %-',
                'Margin-\t']
    print("TABLE:\n")
    # for cell in headings:
    #     print(cell, end=" ")
    # print('\n')
    # for row in rows:
    #     for cell in row:
    #         if(cell == ' '):
    #             print('blank')
    #         else:
    #             print(cell, end=" ")
    #     print('\n')
    # Print this only if no. of items in bill < 4?
    for i in range(len(headings)):
        print(headings[i], end="\t")
        for j in range(len(rows)):
            if(rows[j][i] == ' '):
                print("blank", end="\t")
            else:
                print(rows[j][i], end="\t")
        print("\n")
    return


def updateRows(main_sheet, modified_rows):
    # To DO
    return


def updateSheet(main_sheet, sheet_obj, colList):  # Update rows from sheet_obj to main_sheet
    modified_rows = []
    for row in range(2, sheet_obj.max_row+1):
        modified_row = readModifiedRow(sheet_obj, row, colList)
        modified_rows.append(modified_row)
    printRows(modified_rows)
    print("Continue with this?\n")
    opt = input("Y/N ")
    if(opt == 'Y'):
        print("Updating...")
        updateRows(main_sheet, modified_rows)
    else:
        print("Aborting...")
    return


def openSheet(path):
    wb_obj = main_wb = openpyxl.load_workbook(path)
    sheet_obj = wb_obj.active
    return sheet_obj


# path_main = "./trial-main.xlsx"
path_main = "WellnessVibes-  pharmacy  data.xlsx"
path_test = "test.csv"
path_test_xlsx = "bill.xlsx"

if __name__ == '__main__':
    print("Hello user")
    colList = [4, 9, 37, 39, 39, 45, 46, 100, 100, 100, 58, 40,
               100, 100, 100, 65, 100, 100, 80, 100, 100, 100, 100, 100]
    main_sheet = openSheet(path_main)
    new_bill_sheet = openSheet(path_test_xlsx)
    updateSheet(main_sheet, new_bill_sheet, colList)
