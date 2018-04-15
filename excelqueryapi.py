from excelquery import xlquery
from flask import Flask, jsonify, request, abort, make_response
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)

app.config['UPLOAD_DIR'] = '.'
app.config['ALLOWED_EXT'] = set(['xls', 'xlsx', 'csv'])

@app.errorhandler(400)
def bad_request(error):
    return make_response(jsonify(\
        {'error': 'invalid request: %s'%error.description}\
        ), 400)



@app.route('/xlquery/api', methods=['POST'])
def query_excel():
    if not request.json:
        abort(400, 'empty request')

    qry = request.json
    
    file = qry['file']
    filepath = os.path.join(app.config['UPLOAD_DIR'], secure_filename(file))

    try:
        xl = xlquery(filepath)
    except:
        abort(500)

    sheet = xl.get_sheet(qry['sheet'])
    if not sheet:
        abort(400, 'sheet not found')

    tables = xl.get_tables(sheet)
    table = xl.find_table_by_name(tables, qry['table']) \
                if 'table' in qry else None

    req = qry['do']

    for_name = qry['for']['name'] \
        if 'for' in qry and 'name' in qry['for'] else None
    for_type = qry['for']['type'] \
        if 'for' in qry and 'type' in qry['for'] else None

    range_start = qry['range']['start'] \
        if 'range' in qry and 'start' in qry['range'] else None
    range_end = qry['range']['end'] \
        if 'range' in qry and 'end' in qry['range'] else None
    
    in_list = qry['in'] if 'in' in qry else None

    filter = qry['select'] if 'select' in qry else None

    if not for_type:
        for_type, n = xl.get_col_or_row(sheet, table, for_name)
    elif for_type == 'column':
        for_type = 'col'

    if for_type != 'row' and for_type != 'col':
        abort(400, "invalid for type. valid values: 'row', 'col' or 'column")

    if table is None:
        table = xl.find_table_for(sheet, tables, for_name, for_type)

        if table is None:
            abort(400, 'table not found')

    if req == 'sum':
        result = \
            xl.find_sum_in_row(sheet, table, for_name, \
                    range_start, range_end, in_list, filter) \
                    if for_type == 'row' else \
                    xl.find_sum_in_col(sheet, table, for_name, \
                            range_start, range_end, in_list, filter)

    elif req == 'avg' or req == 'average':
        result = \
            xl.find_avg_in_row(sheet, table, for_name, \
                range_start, range_end, in_list) if for_type == 'row' else \
                xl.find_avg_in_col(sheet, table, for_name, \
                        range_start, range_end, in_list)

    elif req == 'max' or req == 'maximum' or req == 'top':
        result = \
            xl.find_max_in_row(sheet, table, for_name, \
                range_start, range_end, in_list, filter) \
                if for_type == 'row' else \
                xl.find_max_in_col(sheet, table, for_name, \
                        range_start, range_end, in_list, filter)

    elif req == 'min' or req == 'minimum' or req == 'bottom':
        result = \
            xl.find_min_in_row(sheet, table, for_name, \
                    range_start, range_end, in_list, filter) \
                    if for_type == 'row' else \
                    xl.find_min_in_col(sheet, table, for_name, \
                        range_start, range_end, in_list, filter)

    elif req == 'count':
        result = \
            xl.find_count_in_row(sheet, table, for_name, \
                    range_start, range_end, in_list, filter) \
                    if for_type == 'row' else \
                    xl.find_count_in_col(sheet, table, for_name, \
                        range_start, range_end, in_list, filter)

    elif req == 'cmp' or req == 'compare':
        left = qry['compare']['left'] \
            if 'compare' in qry and 'left' in qry['compare'] else None
        right = qry['compare']['right'] \
            if 'compare' in qry and 'right' in qry['compare'] else None

        result = \
            xl.find_diff_in_row(sheet, table, for_name, left, right) \
            if for_type == 'row' else \
            xl.find_diff_in_col(sheet, table, for_name, left, right)

    elif req == 'get':
        at = qry['at'] if 'at' in qry else None
        if at is None:
            abort(400, " 'at' parameter is missing")
        
        rowname = for_name if for_type == 'row' else at
        colname = at if for_type == 'row' else for_name

        result = xl.get_cell_value(sheet, table, rowname, colname)

    else:
        abort(400, "invalid do: valid values: 'sum', 'average', 'avg', \
                    'max', 'maximum', 'top', 'min', 'minimum', 'bottom', \
                    'count', 'cmp', 'compare', 'get'")

    if result is None:
        return jsonify({'result': 'no result found'}), 200

    return jsonify({'result': result}), 200



def is_allowed(filename):
    return '.' in filename \
                and filename.rsplit('.', 1)[1].lower() \
                    in app.config['ALLOWED_EXT']


@app.route('/upload', methods=['POST'])
def upload_excel():

    file = request.files['file']

    if file and is_allowed(file.filename):
        fname = secure_filename(file.filename)

        try:
            file.save(os.path.join(app.config['UPLOAD_DIR'], fname))
        except:
            abort(500)
        else:
            return jsonify({'status': '%s uploaded successfully'%fname}), 201

    else:
        abort(400, 'no valid file to upload')


if __name__ == '__main__':
    app.run(host='localhost', port='21000')

