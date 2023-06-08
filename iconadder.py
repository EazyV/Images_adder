import openpyxl as op
import os
from openpyxl.drawing.image import Image


def main():
    path = os.getcwd()
    names = os.listdir(path + '\\' + 'icon')
    wb = op.Workbook()
    for i in range(len(names)):
        text = names[i]
        sheet = wb.active
        sheet.cell(row=i + 1, column=4).value = text
        img = Image(path + '\\' + 'icon' + '\\' + text)
        img.width, img.height = 80, 80
        sheet.add_image(img, 'C%d' % (i + 1))
        sheet.column_dimensions['C'].width = 20
        sheet.column_dimensions['D'].width = 20
        sheet.row_dimensions[i + 1].height = 60
    wb.save('iconadd.xlsx')


if __name__ == '__main__':
    main()
