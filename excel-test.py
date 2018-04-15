from excelquery import xlquery


#workbook = load_excel('sample-full-nomerge.xlsx')
#sheet = get_sheet(workbook, 'financials')
#workbook = load_excel('/mnt/host/Book1.xlsx')
#sheet = get_sheet(workbook, 'sheet 1')
#workbook = load_excel('/mnt/host/sample.xlsx')

xl = xlquery('/mnt/host/sample.xlsx')

sheet = xl.get_sheet('financials')

if sheet is not None:
    print 'nrows:', sheet.nrows;
    print 'ncols:', sheet.ncols;

    tables = xl.get_tables(sheet)
    print tables

    print xl.find_sum_in_row(sheet, tables['1_1'], 'revenue', 'jan', 'mar',\
            {'or': {'cp': {'gt': 0.28}, 'sga': {'ne': 0}}})
    print xl.find_count_in_row(sheet, tables['1_1'], 'revenue', 'jan', 'mar',\
            {'or': {'cp': {'gt': 0.28}, 'sga': {'ne': 0}}})
    print xl.find_max_in_row(sheet, tables['1_1'], 'revenue', 'jan', 'mar',\
            {'or': {'cp': {'gt': 0.28}, 'sga': {'ne': 0}}})
    print xl.find_min_in_row(sheet, tables['1_1'], 'revenue', 'jan', 'mar',\
            {'or': {'cp': {'gt': 0.28}, 'sga': {'ne': 0}}})
    print xl.find_min_in_row(sheet, tables['1_1'], 'revenue', 'jan', 'mar')
    print xl.find_sum_in_col(sheet, tables['1_1'], 'jan')
    print xl.find_avg_in_col(sheet, tables['1_1'], 'jan', 'revenue', 'sga')
    print xl.find_avg_in_col(sheet, tables['1_1'], 'jan')

    print xl.get_col_or_row(sheet, tables['1_1'], 'feb')
    print xl.get_col_or_row(sheet, tables['1_1'], 'sga')

else:
    print 'sheet not found'


