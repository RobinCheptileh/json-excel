import pandas as pd
from flask import Flask, jsonify, request, make_response
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.writer.excel import save_virtual_workbook

app = Flask(__name__)


@app.route("/api/v1/to-spreadsheet", methods=['POST'])
def to_spreadsheet():
    if not request.json or 'rows' not in request.json or not isinstance(request.json['rows'], list):
        return jsonify({
            'message': 'Bad Request'
        }), 400

    try:
        data = request.json
        header = 'header' in data
        data_frame = pd.DataFrame(data['rows'], columns=data['header']) if header else pd.DataFrame(data['rows'])

        print(data_frame)

        wb = Workbook()
        ws = wb.active
        for r in dataframe_to_rows(data_frame, index=False, header=header):
            ws.append(r)
        raw_data = save_virtual_workbook(wb)

        response = make_response(raw_data)
        response.headers['Content-Type'] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        response.headers['Content-Disposition'] = "inline; filename=spreadsheet.xlsx"
        return response
    except ValueError as error:
        return jsonify({
            'message': str(error)
        }), 400


@app.route("/api/v1/to-json", methods=['POST'])
def to_json():
    return jsonify({
        'message': 'Coming Soon!'
    })


if __name__ == "__main__":
    app.run(debug=True)
