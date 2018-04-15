import xlrd

def is_number(val):
    return type(val) is int or type(val) is float

def is_nan(val):
    return not is_number(val)

def normalize_text(text):
    return text.replace('_', '').\
                replace('-','').\
                replace(',','').\
                replace(':','').\
                replace("'",'').\
                replace(' ', '').\
                replace('\t', '').\
                replace('%', 'percentage').\
                lower()


def compare_text(ltext, rtext, lineant=True):
    ltext = normalize_text(ltext)
    rtext = normalize_text(rtext)

    return (ltext in rtext or rtext in ltext) if lineant \
                else (ltext == rtext)


def compare_date(ldate, rtext, lineant=True):
    dt = xlrd.xldate.xldate_as_datetime(xldate=ldate,datemode=0)

    return compare_text(dt.strftime('%d%b%y'),rtext) \
                or compare_text(dt.strftime('%b%y'),rtext) \
                or compare_text(dt.strftime('%b%Y'),rtext) \
                or compare_text(dt.strftime('%B%y'),rtext) \
                or compare_text(dt.strftime('%B%Y'),rtext)


def get_date_value(xldt):
    dt = xlrd.xldate.xldate_as_datetime(xldate=xldt,datemode=0)

    return dt.strftime('%b,%y' if dt.strftime('%d') == '01' else '%d %b,%y')


def evaluate(lvalue, condition):
    if lvalue is None:
        return False

    result = None

    for oper in condition.keys():
        if oper == 'eq':
            result = lvalue == condition[oper] if is_number(lvalue) \
                    else compare_text(str(lvalue), str(condition[oper]))

        if oper == 'ne':
            result = lvalue != condition[oper] if is_number(lvalue) \
                    else not compare_text(str(lvalue), str(condition[oper]))

        elif oper == 'lt' or oper == 'le' or oper == 'gt' or oper == 'ge':
            rvalue = condition[oper]
            if is_nan(lvalue) or is_nan(rvalue):
                result = False
            else:
                result = lvalue < rvalue if oper == 'lt' \
                            else lvalue <= rvalue if oper == 'le' \
                            else lvalue > rvalue if oper == 'gt' \
                            else lvalue >= rvalue

        elif oper == 'or':
            item_result = evaluate(lvalue, condition[oper])
            result = result or item_result if result is not None \
                        else item_result

        elif oper == 'and':
            item_result = evaluate(lvalue, condition[oper])
            result = result and item_result if result is not None \
                        else item_result

    return result


def load_excel(filename):
    workbook = xlrd.open_workbook(filename, on_demand=True)
    return workbook


def get_sheet(workbook, name):

    sheet_names = workbook.sheet_names();

    for sheet in sheet_names:
        if compare_text(sheet, name):
            return workbook.sheet_by_name(sheet)

    return None


def get_row_no(sheet, table, row_name):

    header_row = table['rowlo']
    header_col = table['collo']
    row_count = table['rowhi'] - table['rowlo'] + 1

    rowno = -1

    if row_name is None or row_name == '':
        return rowno

    rowlo = header_row + 1
    rowhi = header_row + row_count

    for ri in range(rowlo, rowhi + 1):
        rowname = sheet.cell_value(rowx=ri,colx=header_col)
        if(compare_text(rowname, row_name) if type(rowname) is not float \
                else compare_date(rowname, row_name)):
            rowno = ri
            break

    return rowno


def get_col_no(sheet, table, col_name):

    header_row = table['rowlo']
    header_col = table['collo']
    col_count = table['colhi'] - table['collo'] + 1

    colno = -1

    if col_name is None or col_name == '':
        return colno

    collo = header_col + 1
    colhi = header_col + col_count

    for ci in range(collo, colhi + 1):
        colname = sheet.cell_value(rowx=header_row,colx=ci)
        if(compare_text(colname, col_name) if type(colname) is not float \
                else compare_date(colname, col_name)):
            colno = ci
            break

    return colno


def get_cell_value(sheet, table, row_name, col_name):

    rowno = get_row_no(sheet, table, row_name)
    colno = get_col_no(sheet, table, col_name)

    cellval = None

    if rowno != -1 and colno != -1:
        cellval = sheet.cell_value(rowx=rowno, colx=colno)

    return cellval


def sum_of_cells(sheet, row_start, row_end, col_start, col_end):
    sum = 0
    for ri in range(row_start, row_end + 1):
        for ci in range(col_start, col_end + 1):
            cellval = sheet.cell_value(rowx=ri, colx=ci)

            if cellval is None or not is_number(cellval):
                continue

            sum += cellval

    return sum

def apply_filter(filter, sheet, table, rowno, colno, direction):
    result = None
    if filter is not None:
        first_key = filter.keys()[0] 
        if first_key == 'or' or first_key == 'and':
            for key, condition in filter[first_key].items():
                cellno = get_row_no(sheet, table, key) if direction == "row" \
                            else get_col_no(sheet, table, key)
                if cellno == -1:
                    continue

                lvalue = \
                    sheet.cell_value(rowx=cellno, colx=colno) \
                    if direction == "row" \
                    else sheet.cell_value(rowx=rowno, colx=cellno)

                eval_result = evaluate(lvalue, condition)  
                result = \
                    eval_result if result is None \
                    else result and eval_result if first_key == 'and' \
                    else result or eval_result if first_key == 'or' \
                    else None

        else:
            condition = filter[first_key]
            key = first_key

            cellno = get_row_no(sheet, table, key) if direction == "row" \
                        else get_col_no(sheet, table, key)
            if cellno != -1:

                lvalue = \
                    sheet.cell_value(rowx=cellno, colx=colno) \
                    if direction == "row" \
                    else sheet.cell_value(rowx=rowno, colx=cellno)


                result = evaluate(lvalue, condition)  

    return result


def find_sum_in_row(sheet, table, row_name, \
        start_col_name = None, end_col_name = None, filter=None):

    header_col = table['collo']
    col_count = table['colhi'] - table['collo']

    rowno = get_row_no(sheet, table, row_name)
    if rowno == -1:
        return None

    colno_start = \
        get_col_no(sheet, table, start_col_name)
    if colno_start  == -1:
        colno_start = header_col + 1

    colno_end = \
        get_col_no(sheet, table, end_col_name)
    if colno_end  == -1:
        colno_end = header_col + col_count

    if colno_end < colno_start:
        tmp = colno_end
        colno_end = colno_start
        colno_start = tmp

    #sum = sum_of_cells(sheet, rowno, rowno, colno_start, colno_end)
    sum = 0

    for ci in range(colno_start, colno_end + 1):
        result = apply_filter(filter, sheet, table, rowno, ci, "row")

        if filter is None or result:
            cell = sheet.cell_value(rowx=rowno, colx=ci)
            if cell is not None and is_number(cell):
                sum += cell

    return sum



def find_avg_in_row(sheet, table, row_name, \
        start_col_name = None, end_col_name = None):

    header_col = table['collo']
    col_count = table['colhi'] - table['collo']

    rowno = get_row_no(sheet, table, row_name)
    if rowno == -1:
        return None

    colno_start = \
        get_col_no(sheet, table, start_col_name)
    if colno_start  == -1:
        colno_start = header_col + 1

    colno_end = \
        get_col_no(sheet, table, end_col_name)
    if colno_end  == -1:
        colno_end = header_col + col_count

    if colno_end < colno_start:
        tmp = colno_end
        colno_end = colno_start
        colno_start = tmp

    sum = sum_of_cells(sheet, rowno, rowno, colno_start, colno_end)
    
    return sum / (colno_end - colno_start + 1)


def find_max_in_row(sheet, table, row_name, \
        start_col_name=None, end_col_name=None, filter=None):

    header_col = table['collo']
    col_count = table['colhi'] - table['collo']
    rowlo = table['rowlo']

    rowno = get_row_no(sheet, table, row_name)
    if rowno == -1:
        return None

    colno_start = \
        get_col_no(sheet, table, start_col_name)
    if colno_start  == -1:
        colno_start = header_col + 1

    colno_end = \
        get_col_no(sheet, table, end_col_name)
    if colno_end  == -1:
        colno_end = header_col + col_count

    if colno_end < colno_start:
        tmp = colno_end
        colno_end = colno_start
        colno_start = tmp

    max = None
    colname = None

    for ci in range(colno_start, colno_end + 1):
        result = apply_filter(filter, sheet, table, rowno, ci, "row")

        if filter is None or result:
            cell = sheet.cell_value(rowx=rowno, colx=ci)
            if cell is not None and is_number(cell):
                cname = get_date_value(sheet.cell_value(rowx=rowlo, colx=ci))
                colname = cname if max is None or cell > max else colname
                max = cell if max is None or cell > max else max

    return max, colname


def find_min_in_row(sheet, table, row_name, \
        start_col_name=None, end_col_name=None, filter=None):

    header_col = table['collo']
    col_count = table['colhi'] - table['collo']
    rowlo = table['rowlo']

    rowno = get_row_no(sheet, table, row_name)
    if rowno == -1:
        return None

    colno_start = \
        get_col_no(sheet, table, start_col_name)
    if colno_start  == -1:
        colno_start = header_col + 1

    colno_end = \
        get_col_no(sheet, table, end_col_name)
    if colno_end  == -1:
        colno_end = header_col + col_count

    if colno_end < colno_start:
        tmp = colno_end
        colno_end = colno_start
        colno_start = tmp

    min = None
    colname = None

    for ci in range(colno_start, colno_end + 1):
        result = apply_filter(filter, sheet, table, rowno, ci, "row")

        if filter is None or result:
            cell = sheet.cell_value(rowx=rowno, colx=ci)
            if cell is not None and is_number(cell):
                cname = get_date_value(sheet.cell_value(rowx=rowlo, colx=ci))
                colname = cname if min is None or cell < min else colname
                min = cell if min is None or cell < min else min

    return min, colname


def find_count_in_row(sheet, table, row_name, \
        start_col_name=None, end_col_name=None, filter=None):

    header_col = table['collo']
    col_count = table['colhi'] - table['collo']

    rowno = get_row_no(sheet, table, row_name)
    if rowno == -1:
        return None

    colno_start = \
        get_col_no(sheet, table, start_col_name)
    if colno_start  == -1:
        colno_start = header_col + 1

    colno_end = \
        get_col_no(sheet, table, end_col_name)
    if colno_end  == -1:
        colno_end = header_col + col_count

    if colno_end < colno_start:
        tmp = colno_end
        colno_end = colno_start
        colno_start = tmp

    count = 0

    for ci in range(colno_start, colno_end + 1):
        result = apply_filter(filter, sheet, table, rowno, ci, "row")

        if filter is None or result:
            cell = sheet.cell_value(rowx=rowno, colx=ci)
            if cell is not None:
                count += 1

    return count


def find_sum_in_col(sheet, table, col_name, \
        start_row_name = None, end_row_name = None, filter=None):

    header_row = table['rowlo']
    row_count = table['rowhi'] - table['rowlo']

    colno = get_col_no(sheet, table, col_name)
    if colno == -1:
        return None

    rowno_start = \
        get_row_no(sheet, table, start_row_name)
    if rowno_start  == -1:
        rowno_start = header_row + 1

    rowno_end = \
        get_row_no(sheet, table, end_row_name)
    if rowno_end  == -1:
        rowno_end = header_row + row_count

    if rowno_end < rowno_start:
        tmp = rowno_end
        rowno_end = rowno_start
        rowno_start = tmp

    #sum = sum_of_cells(sheet, rowno_start, rowno_end, colno, colno)
    sum = 0

    for ri in range(rowno_start, rowno_end + 1):
        result = apply_filter(filter, sheet, table, ri, colno, "col")

        if filter is None or result:
            cell = sheet.cell_value(rowx=ri, colx=colno)
            if cell is not None and is_number(cell):
                sum += cell


    return sum


def find_avg_in_col(sheet, table, col_name, \
        start_row_name = None, end_row_name = None):

    header_row = table['rowlo']
    row_count = table['rowhi'] - table['rowlo']

    colno = get_col_no(sheet, table, col_name)
    if colno == -1:
        return None

    rowno_start = \
        get_row_no(sheet, table, start_row_name)
    if rowno_start  == -1:
        rowno_start = header_row + 1

    rowno_end = \
        get_row_no(sheet, table, end_row_name)
    if rowno_end  == -1:
        rowno_end = header_row + row_count

    if rowno_end < rowno_start:
        tmp = rowno_end
        rowno_end = rowno_start
        rowno_start = tmp

    sum = sum_of_cells(sheet, rowno_start, rowno_end, colno, colno)

    return sum / (rowno_end - rowno_start + 1)


def find_max_in_col(sheet, table, col_name, \
        start_row_name = None, end_row_name = None, filter=None):

    header_row = table['rowlo']
    row_count = table['rowhi'] - table['rowlo']

    colno = get_col_no(sheet, table, col_name)
    if colno == -1:
        return None

    rowno_start = \
        get_row_no(sheet, table, start_row_name)
    if rowno_start  == -1:
        rowno_start = header_row + 1

    rowno_end = \
        get_row_no(sheet, table, end_row_name)
    if rowno_end  == -1:
        rowno_end = header_row + row_count

    if rowno_end < rowno_start:
        tmp = rowno_end
        rowno_end = rowno_start
        rowno_start = tmp

    max = None

    for ri in range(rowno_start, rowno_end + 1):
        result = apply_filter(filter, sheet, table, ri, colno, "col")

        if filter is None or result:
            cell = sheet.cell_value(rowx=ri, colx=colno)
            if cell is not None and is_number(cell):
                max = cell if max is None or cell > max else max

    return max


def find_min_in_col(sheet, table, col_name, \
        start_row_name = None, end_row_name = None, filter=None):

    header_row = table['rowlo']
    row_count = table['rowhi'] - table['rowlo']

    colno = get_col_no(sheet, table, col_name)
    if colno == -1:
        return None

    rowno_start = \
        get_row_no(sheet, table, start_row_name)
    if rowno_start  == -1:
        rowno_start = header_row + 1

    rowno_end = \
        get_row_no(sheet, table, end_row_name)
    if rowno_end  == -1:
        rowno_end = header_row + row_count

    if rowno_end < rowno_start:
        tmp = rowno_end
        rowno_end = rowno_start
        rowno_start = tmp

    min = None

    for ri in range(rowno_start, rowno_end + 1):
        result = apply_filter(filter, sheet, table, ri, colno, "col")

        if filter is None or result:
            cell = sheet.cell_value(rowx=ri, colx=colno)
            if cell is not None and is_number(cell):
                min = cell if min is None or cell < min else min

    return min


def find_count_in_col(sheet, table, col_name, \
        start_row_name = None, end_row_name = None, filter=None):

    header_row = table['rowlo']
    row_count = table['rowhi'] - table['rowlo']

    colno = get_col_no(sheet, table, col_name)
    if colno == -1:
        return None

    rowno_start = \
        get_row_no(sheet, table, start_row_name)
    if rowno_start  == -1:
        rowno_start = header_row + 1

    rowno_end = \
        get_row_no(sheet, table, end_row_name)
    if rowno_end  == -1:
        rowno_end = header_row + row_count

    if rowno_end < rowno_start:
        tmp = rowno_end
        rowno_end = rowno_start
        rowno_start = tmp

    count = 0

    for ri in range(rowno_start, rowno_end + 1):
        result = apply_filter(filter, sheet, table, ri, colno, "col")

        if filter is None or result:
            cell = sheet.cell_value(rowx=ri, colx=colno)
            if cell is not None:
                count += 1

    return count


def find_diff_in_row(sheet, table, row_name, \
        lcol_name, rcol_name):

    rowno = get_row_no(sheet, table, row_name)
    if rowno == -1:
        return None

    lcolno = \
        get_col_no(sheet, table, lcol_name)
    if lcolno == -1:
        return None

    rcolno = \
        get_col_no(sheet, table, rcol_name)
    if rcolno == -1:
        return None

    lcellval = sheet.cell_value(rowx=rowno, colx=lcolno)
    rcellval = sheet.cell_value(rowx=rowno, colx=rcolno)

    return (rcellval - lcellval) if \
                is_number(lcellval) and is_number(rcellval) \
                else None


def find_diff_in_col(sheet, table, col_name, \
        lrow_name, rrow_name):

    colno = get_col_no(sheet, table, col_name)
    if colno == -1:
        return None

    lrowno = \
        get_row_no(sheet, table, lrow_name)
    if lrowno == -1:
        return None

    rrowno = \
        get_row_no(sheet, table, rrow_name)
    if rrowno == -1:
        return None

    lcellval = sheet.cell_value(rowx=lrowno, colx=colno)
    rcellval = sheet.cell_value(rowx=rrowno, colx=colno)

    return (rcellval - lcellval) if \
                is_number(lcellval) and is_number(rcellval) \
                else None



def has_valid_value(val):
    return val is not None and (is_number(val) or len(val) > 0)


def find_table(tables, rowno, colno):
    for key, value in tables.items():
        if value['collo'] <= colno \
                and value['colhi'] >= colno \
                and value['rowhi'] >= rowno:
            return value

    return None


def find_table_name(sheet, rowlo, collo, colhi):

    if rowlo > 0:
        for ci in range(collo - 1 if collo > 0 else collo, colhi):
            cell = sheet.cell_value(rowx=rowlo-1, colx=ci)
            if has_valid_value(cell):
                return cell

    return '%d_%d'%(rowlo, collo)


def get_tables(sheet):
    tables = {}
    for ri in range(sheet.nrows):
        start_table = False
        non_blank_cnt = 0
        for ci in range(sheet.ncols):
            cell = sheet.cell_value(rowx=ri, colx=ci)

            table = find_table(tables, ri, ci)

            if has_valid_value(cell):
                if table is not None:
                    continue

                non_blank_cnt += 1

                if not start_table:
                    start_table = True
                    collo = ci
                    rowlo = ri
                else:
                    if ci == sheet.ncols - 1:
                        # end of table
                        colhi = ci

                        for r2 in range(rowlo + 1, sheet.nrows):
                            tblcellval = sheet.cell_value(r2, collo)

                            if not has_valid_value(tblcellval) \
                                    or r2 == sheet.nrows - 1:

                                tablename = find_table_name(\
                                        sheet, rowlo, collo, colhi)
                                tables[tablename] = {\
                                        'rowlo':rowlo,\
                                        'collo':collo,\
                                        'rowhi':r2 if r2 == sheet.nrows - 1\
                                                    else r2 - 1,\
                                        'colhi':colhi}
                                break
                        
            else:
                if start_table:
                    # end of table
                    start_table = False
                    if non_blank_cnt == 1:
                        non_blank_cnt = 0
                        continue

                    colhi = ci - 1

                    for r2 in range(rowlo + 1, sheet.nrows):
                        tblcellval = sheet.cell_value(r2, collo)


                        #if has_valid_value(tblcellval):
                        #    continue

                        if not has_valid_value(tblcellval) \
                                or r2 == sheet.nrows - 1:

                            tablename = find_table_name(\
                                    sheet, rowlo, collo, colhi)
                            tables[tablename] = {\
                                    'rowlo':rowlo,\
                                    'collo':collo,\
                                    'rowhi':r2 if has_valid_value(tblcellval)\
                                                    and r2 == sheet.nrows - 1\
                                                else r2 - 1,\
                                    'colhi':colhi}
                            break


    return tables



#workbook = load_excel('sample-full-nomerge.xlsx')
#sheet = get_sheet(workbook, 'financials')
#workbook = load_excel('/mnt/host/Book1.xlsx')
#sheet = get_sheet(workbook, 'sheet 1')
workbook = load_excel('/mnt/host/sample.xlsx')
sheet = get_sheet(workbook, 'financials')

if sheet is not None:
    print 'nrows:', sheet.nrows;
    print 'ncols:', sheet.ncols;

    tables = get_tables(sheet)
    print tables

    #print find(sheet, 0, 0, 17, 3, 'cp', 'march')
    print find_sum_in_row(sheet, tables['1_1'], 'revenue', 'jan', 'mar',\
            {'or': {'cp': {'gt': 0.28}, 'sga': {'ne': 0}}})
    print find_count_in_row(sheet, tables['1_1'], 'revenue', 'jan', 'mar',\
            {'or': {'cp': {'gt': 0.28}, 'sga': {'ne': 0}}})
    print find_max_in_row(sheet, tables['1_1'], 'revenue', 'jan', 'mar',\
            {'or': {'cp': {'gt': 0.28}, 'sga': {'ne': 0}}})
    print find_min_in_row(sheet, tables['1_1'], 'revenue', 'jan', 'mar',\
            {'or': {'cp': {'gt': 0.28}, 'sga': {'ne': 0}}})
    print find_min_in_row(sheet, tables['1_1'], 'revenue', 'jan', 'mar')
    #print find_avg_in_row(sheet, 0, 0, 17, 3, 'revenue', 'jan', 'mar')
    print find_sum_in_col(sheet, tables['1_1'], 'jan')
    print find_avg_in_col(sheet, tables['1_1'], 'jan', 'revenue', 'sga')
    print find_avg_in_col(sheet, tables['1_1'], 'jan')
    #print find_diff_in_row(sheet, 0, 0, 17, 3, 'revenue', 'q3', 'q4')
    #print find_diff_in_col(sheet, 0, 0, 17, 3, 'jan', 'revenue', 'sga')

    print sheet

else:
    print 'sheet not found'

#print evaluate(5, {'or': {'eq': 6, 'eq': 4}})

