import pandas as pd
from pathlib import Path


def excel_diff(path_OLD, path_NEW):

    df_OLD = pd.read_excel(path_OLD).fillna(0)
    print (df_OLD)
    df_NEW = pd.read_excel(path_NEW).fillna(0)
    print(df_NEW)
    # Perform Diff
    dfDiff = df_OLD.copy()

    for row in range(dfDiff.shape[0]):
       # print ( dfDiff.shape[0])
        for col in range(dfDiff.shape[1]):
          # print( dfDiff.shape[1])
            value_OLD = df_OLD.iloc[row,col]
            #print(df_OLD.iloc[row,col])
            try:
                value_NEW = df_NEW.iloc[row,col]
             #   print('new' + df_NEW.iloc[row, col])
                if value_OLD==value_NEW:
                    dfDiff.iloc[row,col] = df_NEW.iloc[row,col]
                else:
                    dfDiff.iloc[row,col] = ('{}-->{}').format(value_OLD,value_NEW)
            except:
                dfDiff.iloc[row,col] = ('{}-->{}').format(value_OLD, 'NaN')

    # Save output and format

    fname = '{} vs {}.xls'.format(path_OLD.stem,path_NEW.stem)
    print(fname)
    writer = pd.ExcelWriter(fname, engine='xlsxwriter')
    print(path_NEW)
    print(path_NEW.stem)
    dfDiff.to_excel(writer, sheet_name='DIFF', index=False)
    df_NEW.to_excel(writer, sheet_name=path_NEW.stem, index=False)
    df_OLD.to_excel(writer, sheet_name=path_OLD.stem, index=False)

    # get xlsxwriter objects
    workbook  = writer.book
    worksheet = writer.sheets['DIFF']
    worksheet.hide_gridlines(2)

    # define formats
    date_fmt = workbook.add_format({'align': 'center', 'num_format': 'yyyy-mm-dd'})
    center_fmt = workbook.add_format({'align': 'center'})
    number_fmt = workbook.add_format({'align': 'center', 'num_format': '#,##0.00'})
    cur_fmt = workbook.add_format({'align': 'center', 'num_format': '$#,##0.00'})
    perc_fmt = workbook.add_format({'align': 'center', 'num_format': '0%'})
    grey_fmt = workbook.add_format({'font_color': '#E0E0E0'})
    #highlight_fmt = workbook.add_format({'font_color': '#FF0000', 'bg_color':'#B1B3B3'})
    highlight_fmt = workbook.add_format({'font_color': '#FF0000'})
    # set column width and format over columns
    # worksheet.set_column('J:AX', 5, number_fmt)

    # set format over range
    ## highlight changed cells
    worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
                                            'criteria': 'containing',
                                            'value':'-->',
                                            'format': highlight_fmt})
    ## highlight unchanged cells
    worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
                                            'criteria': 'not containing',
                                            'value':'-->',
                                            'format': grey_fmt})

    # save
    writer.save()
    print('Done.')


def main():
    path_OLD = Path('xls/OldVersion/AccountStatement1.xls')
    path_NEW = Path('xls/NewVersion/AccountStatement1.xls')



    excel_diff(path_OLD, path_NEW)

if __name__ == '__main__':
    main()