import numpy as np
import pandas as pd
import xlsxwriter
import argparse

# parse command line arguments
parser = argparse.ArgumentParser(description="Generate JC bingo sheets")
parser.add_argument("-n", type=int, default=8)
parser.add_argument("-o", type=str, default="bingo_jc.xlsx")
args = parser.parse_args()

bingo = pd.read_csv('bingo_jc.csv', header=None)

writer = pd.ExcelWriter(path=args.o, engine='xlsxwriter')
workbook = writer.book
cell_format = workbook.add_format({'text_wrap': True,\
                                   'align': 'center',\
                                   'valign': 'vcenter',\
                                   'border': 1,\
                                   'font_name': 'Arial'})
# make 8 copies
for cycle in range(args.n):
    # shuffle the input values
    np.random.shuffle(bingo[0])
    # select first 25 and put in 5x5 grid
    board = bingo[0][:25].values.reshape(5,5)
    board = pd.DataFrame(board)
    board.rename(columns={0: 'B', 1: 'I', 2: 'N', 3: 'G', 4: '0'}, inplace=True)
    # write to xlsx
    board.to_excel(writer, sheet_name=str(cycle), index=False)
    worksheet = writer.sheets[str(cycle)]
    worksheet.set_column('A:E', 13, cell_format)
    # set row heights
    for i in range(1,6):
        worksheet.set_row(i, 80)

workbook.close()
