import xlrd
import difflib
from fuzzymatch import best_match
    
class xlquery:
    def __init__(self, path):
        self.workbook = self._load_excel(path)
    
    def _is_number(self, val):
        return type(val) is int or type(val) is float
    
    def _is_nan(self, val):
        return not self._is_number(val)

    def _compare_text(self, lefttext, righttext, lineant=True):
        normalize = lambda t: t.replace('_', '').\
                        replace('-','').\
                        replace(',','').\
                        replace(':','').\
                        replace("'",'').\
                        replace(' ', '').\
                        replace('\t', '').\
                        replace('%', 'percentage').\
                        lower()

        ltext = normalize(lefttext)
        rtext = normalize(righttext)
    
        return (rtext in ltext) if lineant else (ltext == rtext)    

        
    def _format_date(self, dt, format = '%d%b%y'):
        return \
            xlrd.xldate.xldate_as_datetime(\
                xldate=dt,datemode=0).strftime(format)

    def _compare_date(self, ldate, rtext, lineant=True):
        dt = xlrd.xldate.xldate_as_datetime(xldate=ldate,datemode=0)
    
        return self._compare_text(dt.strftime('%d%b%y'),rtext) \
                    or self._compare_text(dt.strftime('%b%y'),rtext) \
                    or self._compare_text(dt.strftime('%b%Y'),rtext) \
                    or self._compare_text(dt.strftime('%B%y'),rtext) \
                    or self._compare_text(dt.strftime('%B%Y'),rtext)
    
    
    def _has_valid_value(self, val):
        return val is not None and (self._is_number(val) or len(val) > 0)
    
    
    def _get_date_value(self, xldt):
        dt = xlrd.xldate.xldate_as_datetime(xldate=xldt,datemode=0)
    
        return dt.strftime('%b,%y' if dt.strftime('%d') == '01' else '%d %b,%y')
    
    
    def _evaluate(self, lvalue, condition):
        if lvalue is None:
            return False
    
        result = None
    
        for oper in condition.keys():
            if oper == 'eq':
                result = lvalue == condition[oper] if self._is_number(lvalue) \
                        else self._compare_text(str(lvalue), \
                                str(condition[oper]))
    
            if oper == 'ne':
                result = lvalue != condition[oper] if self._is_number(lvalue) \
                        else not self._compare_text(str(lvalue), \
                                str(condition[oper]))
    
            elif oper == 'lt' or oper == 'le' or oper == 'gt' or oper == 'ge':
                rvalue = condition[oper]
                if self._is_nan(lvalue) or self._is_nan(rvalue):
                    result = False
                else:
                    result = lvalue < rvalue if oper == 'lt' \
                                else lvalue <= rvalue if oper == 'le' \
                                else lvalue > rvalue if oper == 'gt' \
                                else lvalue >= rvalue
    
            elif oper == 'or':
                item_result = self._evaluate(lvalue, condition[oper])
                result = result or item_result if result is not None \
                            else item_result
    
            elif oper == 'and':
                item_result = self._evaluate(lvalue, condition[oper])
                result = result and item_result if result is not None \
                            else item_result
    
        return result
    
    
    def _load_excel(self, filename):
        return xlrd.open_workbook(filename, on_demand=True)
    
    
    def get_sheet(self, name):
        match = best_match(self.workbook.sheet_names(), name)
        return self.workbook.sheet_by_name(match) \
                    if match is not None else None
    
    
    def _get_row_no(self, sheet, table, row_name):
    
        header_row = table['rowlo']
        header_col = table['collo']
        
        if row_name is None or row_name == '':
            return -1
    
        rowlo = header_row + 1
        rowhi = table['rowhi']
        
        rownames = [v if type(v) is not float else self._format_date(v) \
                        for v in (sheet.cell_value(rowx=ri,colx=header_col) \
                            for ri in range(rowlo, rowhi + 1))]
                    
        match = best_match(rownames, row_name)
        
        return rowlo + rownames.index(match) if match is not None else -1
    
    
    def _get_col_no(self, sheet, table, col_name):
    
        header_row = table['rowlo']
        header_col = table['collo']
    
        if col_name is None or col_name == '':
            return colno
    
        collo = header_col + 1
        colhi = table['colhi']
        
        colnames = [v if type(v) is not float else self._format_date(v) \
                        for v in (sheet.cell_value(rowx=header_row,colx=ci) \
                            for ci in range(collo, colhi + 1))]

        match = best_match(colnames, col_name)

        return collo + colnames.index(match) if match is not None else -1
    
    
    def get_col_or_row(self, sheet, table, name):
        rowno = self._get_row_no(sheet, table, name)
        if rowno == -1:
            colno = self._get_col_no(sheet, table, name)
            if colno == -1:
                return None, None
            else:
                return 'col', colno
        else:
            return 'row', rowno
    

    def get_cell_value(self, sheet, table, row_name, col_name):
    
        rowno = self._get_row_no(sheet, table, row_name)
        colno = self._get_col_no(sheet, table, col_name)
    
        cellval = None
    
        if rowno != -1 and colno != -1:
            cellval = sheet.cell_value(rowx=rowno, colx=colno)
    
        return cellval
    
    
    def sum_of_cells(self, sheet, row_start, row_end, col_start, col_end):
        sum = 0
        for ri in range(row_start, row_end + 1):
            for ci in range(col_start, col_end + 1):
                cellval = sheet.cell_value(rowx=ri, colx=ci)
    
                if cellval is None or not self._is_number(cellval):
                    continue
    
                sum += cellval
    
        return sum

    
    def _apply_filter(self, filter, sheet, table, rowno, colno, direction):
        result = None
        if filter is not None:
            first_key = filter.keys()[0] 
            if first_key == 'or' or first_key == 'and':
                for key, condition in filter[first_key].items():
                    cellno = self._get_row_no(sheet, table, key) \
                                if direction == "row" \
                                else self._get_col_no(sheet, table, key)
                    if cellno == -1:
                        continue
    
                    lvalue = \
                        sheet.cell_value(rowx=cellno, colx=colno) \
                        if direction == "row" \
                        else sheet.cell_value(rowx=rowno, colx=cellno)
    
                    eval_result = self._evaluate(lvalue, condition)  
                    result = \
                        eval_result if result is None \
                        else result and eval_result if first_key == 'and' \
                        else result or eval_result if first_key == 'or' \
                        else None
    
            else:
                condition = filter[first_key]
                key = first_key
    
                cellno = self._get_row_no(sheet, table, key) \
                            if direction == "row" \
                            else self._get_col_no(sheet, table, key)
                if cellno != -1:
    
                    lvalue = \
                        sheet.cell_value(rowx=cellno, colx=colno) \
                        if direction == "row" \
                        else sheet.cell_value(rowx=rowno, colx=cellno)
    
    
                    result = self._evaluate(lvalue, condition)  
    
        return result
    
    def _get_list_of_row_or_col(self, sheet, table, row_or_col, \
            start = None, end = None, list = None):
        
        returnlst = []
        
        collo = table['collo']
        colhi = table['colhi']
        rowlo = table['rowlo']
        rowhi = table['rowhi']
        
        if list is not None:
            for name in list:
                n = self._get_col_no(sheet, table, name) \
                        if row_or_col == 'col' \
                        else self._get_row_no(sheet, table, name)
                if n != -1:
                    returnlst.append(n)
        else:
            n_start = self._get_col_no(sheet, table, start) \
                        if row_or_col == 'col' \
                        else self._get_row_no(sheet, table, start)
            if n_start == -1:
                n_start = collo + 1 if row_or_col == 'col' else rowlo + 1
            
            n_end = self._get_col_no(sheet, table, end) \
                        if row_or_col == 'col' \
                        else self._get_row_no(sheet, table, end)
            if n_end == -1:
                n_end = colhi if row_or_col == 'col' else rowhi
            
            returnlst.extend(range(n_start, n_end + 1))
        
        return returnlst

    
    def find_sum_in_row(self, sheet, table, row_name, \
            start_col_name = None, end_col_name = None, col_list = None, \
            filter = None):
    
        rowno = self._get_row_no(sheet, table, row_name)
        if rowno == -1:
            return None
    
        clist = self._get_list_of_row_or_col(sheet, table, 'col', \
                    start_col_name, end_col_name, col_list)
    
        #sum = sum_of_cells(sheet, rowno, rowno, colno_start, colno_end)
        sum = 0
    
        for ci in clist:
            result = self._apply_filter(filter, sheet, table, rowno, ci, "row")
    
            if filter is None or result:
                cell = sheet.cell_value(rowx=rowno, colx=ci)
                if cell is not None and self._is_number(cell):
                    sum += cell
    
        return sum
    
    
    
    def find_avg_in_row(self, sheet, table, row_name, \
            start_col_name = None, end_col_name = None, col_list = None):
    
        rowno = self._get_row_no(sheet, table, row_name)
        if rowno == -1:
            return None
    
        clist = self._get_list_of_row_or_col(sheet, table, 'col', \
                    start_col_name, end_col_name, col_list)
    
        sum = 0
    
        for ci in clist:
            result = self._apply_filter(filter, sheet, table, rowno, ci, "row")
    
            if filter is None or result:
                cell = sheet.cell_value(rowx=rowno, colx=ci)
                if cell is not None and self._is_number(cell):
                    sum += cell
        
        return sum / len(clist) if len(clist) > 0 else None
    
    
    def find_max_in_row(self, sheet, table, row_name, \
            start_col_name = None, end_col_name = None, col_list = None, \
            filter = None):
    
        rowlo = table['rowlo']
    
        rowno = self._get_row_no(sheet, table, row_name)
        if rowno == -1:
            return None
    
        clist = self._get_list_of_row_or_col(sheet, table, 'col', \
                    start_col_name, end_col_name, col_list)
            
        max = None
        colname = None
    
        for ci in clist:
            result = self._apply_filter(filter, sheet, table, rowno, ci, "row")
    
            if filter is None or result:
                cell = sheet.cell_value(rowx=rowno, colx=ci)
                if cell is not None and self._is_number(cell):
                    cval = sheet.cell_value(rowx=rowlo, colx=ci)
                    cname = self._get_date_value(cval) \
                                if type(cval) is float else cval
                    colname = cname if max is None or cell > max else colname
                    max = cell if max is None or cell > max else max
    
        return max, colname
    
    
    def find_min_in_row(self, sheet, table, row_name, \
            start_col_name = None, end_col_name = None, col_list = None, \
            filter = None):
    
        rowlo = table['rowlo']
    
        rowno = self._get_row_no(sheet, table, row_name)
        if rowno == -1:
            return None
    
        clist = self._get_list_of_row_or_col(sheet, table, 'col', \
                    start_col_name, end_col_name, col_list)
    
        min = None
        colname = None
    
        for ci in clist:
            result = self._apply_filter(filter, sheet, table, rowno, ci, "row")
    
            if filter is None or result:
                cell = sheet.cell_value(rowx=rowno, colx=ci)
                if cell is not None and self._is_number(cell):
                    cval = sheet.cell_value(rowx=rowlo, colx=ci)
                    cname = self._get_date_value(cval) \
                                if type(cval) is float else cval
                    colname = cname if min is None or cell < min else colname
                    min = cell if min is None or cell < min else min
    
        return min, colname
    
    
    def find_count_in_row(self, sheet, table, row_name, \
            start_col_name = None, end_col_name = None, col_list = None, \
            filter = None):
    
        rowno = self._get_row_no(sheet, table, row_name)
        if rowno == -1:
            return None
    
        clist = self._get_list_of_row_or_col(sheet, table, 'col', \
                    start_col_name, end_col_name, col_list)
    
        count = 0
        
        cnt = len([sheet.cell_value(rowx=rowno, colx=ci) \
            for ci in clist \
                if self._apply_filter(filter, sheet, table, rowno, ci, "row") \
                    and sheet.cell_value(rowx=rowno, colx=ci) is not None])
    
        for ci in clist:
            result = self._apply_filter(filter, sheet, table, rowno, ci, "row")
    
            if filter is None or result:
                cell = sheet.cell_value(rowx=rowno, colx=ci)
                if cell is not None:
                    count += 1
    
        return count
    
    
    def find_sum_in_col(self, sheet, table, col_name, \
            start_row_name = None, end_row_name = None, row_list = None, \
            filter = None):
        
        colno = self._get_col_no(sheet, table, col_name)
        if colno == -1:
            return None
    
        rlist = self._get_list_of_row_or_col(sheet, table, 'row', \
                    start_row_name, end_row_name, row_list)
    
        #sum = sum_of_cells(sheet, rowno_start, rowno_end, colno, colno)
        sum = 0
    
        for ri in rlist:
            result = self._apply_filter(filter, sheet, table, ri, colno, "col")
    
            if filter is None or result:
                cell = sheet.cell_value(rowx=ri, colx=colno)
                if cell is not None and self._is_number(cell):
                    sum += cell
    
    
        return sum
    
    
    def find_avg_in_col(self, sheet, table, col_name, \
            start_row_name = None, end_row_name = None, row_list = None):
    
        colno = self._get_col_no(sheet, table, col_name)
        if colno == -1:
            return None
    
        rlist = self._get_list_of_row_or_col(sheet, table, 'row', \
                    start_row_name, end_row_name, row_list)
                    
        sum = 0
    
        for ri in rlist:
            result = self._apply_filter(filter, sheet, table, ri, colno, "col")
    
            if filter is None or result:
                cell = sheet.cell_value(rowx=ri, colx=colno)
                if cell is not None and self._is_number(cell):
                    sum += cell
    
        return sum / len(rlist) if len(rlist) > 0 else None
    
    
    def find_max_in_col(self, sheet, table, col_name, \
            start_row_name = None, end_row_name = None, row_list = None, \
            filter = None):
    
        collo = table['collo']
    
        colno = self._get_col_no(sheet, table, col_name)
        if colno == -1:
            return None
    
        rlist = self._get_list_of_row_or_col(sheet, table, 'row', \
                    start_row_name, end_row_name, row_list)
    
        max = None
        rowname = None
    
        for ri in rlist:
            result = self._apply_filter(filter, sheet, table, ri, colno, "col")
    
            if filter is None or result:
                cell = sheet.cell_value(rowx=ri, colx=colno)
                if cell is not None and self._is_number(cell):
                    rval = sheet.cell_value(rowx=ri, colx=collo)
                    rname = self._get_date_value(rval) \
                                if type(rval) is float else rval
                    rowname = rname if max is None or cell > max else rowname
                    max = cell if max is None or cell > max else max
    
        return max, rowname
    
    
    def find_min_in_col(self, sheet, table, col_name, \
            start_row_name = None, end_row_name = None, row_list = None, \
            filter = None):
    
        collo = table['collo']
    
        colno = self._get_col_no(sheet, table, col_name)
        if colno == -1:
            return None
    
        rlist = self._get_list_of_row_or_col(sheet, table, 'row', \
                    start_row_name, end_row_name, row_list)

        min = None
        rowname = None
    
        for ri in rlist:
            result = self._apply_filter(filter, sheet, table, ri, colno, "col")
    
            if filter is None or result:
                cell = sheet.cell_value(rowx=ri, colx=colno)
                if cell is not None and self._is_number(cell):
                    rval = sheet.cell_value(rowx=ri, colx=collo)
                    rname = self._get_date_value(rval) \
                                if type(rval) is float else rval
                    rowname = rname if min is None or cell < min else rowname
                    min = cell if min is None or cell < min else min
    
        return min, rowname
    
    
    def find_count_in_col(self, sheet, table, col_name, \
            start_row_name = None, end_row_name = None, row_list = None, \
            filter = None):
        
        colno = self._get_col_no(sheet, table, col_name)
        if colno == -1:
            return None
    
        rlist = self._get_list_of_row_or_col(sheet, table, 'row', \
                    start_row_name, end_row_name, row_list)
    
        count = 0
    
        for ri in rlist:
            result = self._apply_filter(filter, sheet, table, ri, colno, "col")
    
            if filter is None or result:
                cell = sheet.cell_value(rowx=ri, colx=colno)
                if cell is not None:
                    count += 1
    
        return count
    
    
    def find_diff_in_row(self, sheet, table, row_name, \
            lcol_name, rcol_name):
    
        rowno = self._get_row_no(sheet, table, row_name)
        if rowno == -1:
            return None
    
        lcolno = \
            self._get_col_no(sheet, table, lcol_name)
        if lcolno == -1:
            return None
    
        rcolno = \
            self._get_col_no(sheet, table, rcol_name)
        if rcolno == -1:
            return None
    
        lcellval = sheet.cell_value(rowx=rowno, colx=lcolno)
        rcellval = sheet.cell_value(rowx=rowno, colx=rcolno)
    
        return (rcellval - lcellval) if \
                    self._is_number(lcellval) and self._is_number(rcellval) \
                    else None
    
    
    def find_diff_in_col(self, sheet, table, col_name, \
            lrow_name, rrow_name):
    
        colno = self._get_col_no(sheet, table, col_name)
        if colno == -1:
            return None
    
        lrowno = \
            self._get_row_no(sheet, table, lrow_name)
        if lrowno == -1:
            return None
    
        rrowno = \
            self._get_row_no(sheet, table, rrow_name)
        if rrowno == -1:
            return None
    
        lcellval = sheet.cell_value(rowx=lrowno, colx=colno)
        rcellval = sheet.cell_value(rowx=rrowno, colx=colno)
    
        return (rcellval - lcellval) if \
                    self._is_number(lcellval) and self._is_number(rcellval) \
                    else None
    
    
    
    def _find_table(self, tables, rowno, colno):
        for key, value in tables.items():
            if value['collo'] <= colno \
                    and value['colhi'] >= colno \
                    and value['rowhi'] >= rowno:
                return value
    
        return None
    
    
    def _find_table_name(self, sheet, rowlo, collo, colhi):
    
        if rowlo > 0:
            for ci in range(collo - 1 if collo > 0 else collo, colhi):
                cell = sheet.cell_value(rowx=rowlo-1, colx=ci)
                if self._has_valid_value(cell):
                    return cell
    
        return '%d_%d'%(rowlo, collo)
    
    
    def get_tables(self, sheet):
        tables = {}
        for ri in range(sheet.nrows):
            start_table = False
            non_blank_cnt = 0
            for ci in range(sheet.ncols):
                cell = sheet.cell_value(rowx=ri, colx=ci)
    
                table = self._find_table(tables, ri, ci)
    
                if self._has_valid_value(cell):
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
    
                                if not self._has_valid_value(tblcellval) \
                                        or r2 == sheet.nrows - 1:
    
                                    tablename = self._find_table_name(\
                                            sheet, rowlo, collo, colhi)
                                    tables[tablename] = {\
                                            'rowlo':rowlo,\
                                            'collo':collo,\
                                            'rowhi':r2 if \
                                                r2 == sheet.nrows - 1 \
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
    
                            if not self._has_valid_value(tblcellval) \
                                    or r2 == sheet.nrows - 1:
    
                                tablename = self._find_table_name(\
                                        sheet, rowlo, collo, colhi)
                                tables[tablename] = {\
                                        'rowlo':rowlo,\
                                        'collo':collo,\
                                        'rowhi':r2 if \
                                            self._has_valid_value(tblcellval) \
                                                and r2 == sheet.nrows - 1\
                                            else r2 - 1,\
                                        'colhi':colhi}
                                break
    
    
        return tables


    def find_table_for(self, sheet, tables, name, type):
        for table in tables:
            rowlo = table['rowlo']
            rowhi = table['rowhi']
            collo = table['collo']
            colhi = table['colhi']

            if type is None or type == 'col':
            
                cnames = [v if type(v) is not float else self._format_date(v)\
                            for v in (sheet.cell_value(rowx=rowlo,colx=ci) \
                                for ci in range(collo, colhi + 1))]

                if best_match(cnames, name) is not None:
                    return table

            if type is None or type == 'row':
                rnames = [v if type(v) is not float else self._format_date(v)\
                            for v in (sheet.cell_value(rowx=ri,colx=collo) \
                                for ci in range(collo, colhi + 1))]

                if best_match(rnames, name) is not None:
                    return table

        return None
    
    def find_table_by_name(self, tables, name):
        match = best_match(tables.keys(), name)
        return tables[match] if match is not None and match in tables else None

    def release(self):
        self.workbook.release_resources()
    