import random
import seaborn as sns
import openpyxl
import pandas as pd

from openpyxl.styles import Border, Side, Alignment


def export_excel(i):
    return


def generate_colors(number):

    color = set()
    while len(color) < number:
        c = []
        for j in range(6):
            c.append(random.choice('0123456789ABCDEF'))
        b = ''.join(c)
        color.add("#"+b)

    return sns.color_palette(list(color))


def style_excel(df_entropy, file,sheet_, index=0,top=5):

    wb = openpyxl.load_workbook(file)
    with pd.ExcelWriter(file, engine="openpyxl") as writer:

        workbook = writer.book
        writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)

        wb = openpyxl.load_workbook(file)
        ws = wb[sheet_]
        wb.active = ws
        print("ws:", ws)
        row_range = list()
        start = 1
        for k in range(int(df_entropy.shape[0]/5) + 1):
            a = "A" + str(start) + ":M" + str(start)
            start = start + top
            row_range.append(a)

        a = "A{0}:B{1}".format(df_entropy.shape[0],df_entropy.shape[0])
        b = "C{0}:J{1}".format(df_entropy.shape[0],df_entropy.shape[0])
        c = "K{0}:M{1}".format(df_entropy.shape[0],df_entropy.shape[0])
        cols = {a: 10, b: 15, c: 35}

        wb.save(file)