from flask import Flask, jsonify, request
from flask_cors import CORS

from BackEnd.Database.Queries.Insert.Reports.insert_historical import InsertHistorical
from BackEnd.Database.Queries.Select.Reports.MainCommand import MainCommand
from BackEnd.Processes.Reporting.ReportingProcess import ReportingProcess

app = Flask(__name__)
CORS(app)

report_service = MainCommand()

current_batch_id = None
report_data_cache = {}
current_report_data = None




@app.route('/api/generate-report/', methods=['POST'])

def generate_report():
    
    global current_report_data
    
    try:
        
        body = request.get_json()

        if not body or 'batch_id' not in body:
            
            return jsonify({
                
                'success': False,
                
                'error': 'Work order required'
                
            }), 400
            
        batch_id = int(body['batch_id'])


        data = report_service.caller(batch_id)
        
        current_report_data = {
            
            'batch_id': batch_id,
            
            'data': data,
            
            'timestamp': None
            
            
        }
        
            # Trigger -> Inserts a register with the report data (work order, user, client)
    
        ReportingProcess().insert_historical_record({'created_by': 'Julian Criollo', 'work_order': current_report_data['batch_id']})

        return jsonify({
            
            'success': True,
            
            'batch_id': batch_id,
            
            'message': 'Report data ready'
            
        }), 200

    except Exception as e:
        
        return jsonify({
            
            'success': False,
            
            'error': str(e)
            
        }), 500


@app.route('/api/get-report/', methods=['GET'])

def get_report():
    
    global current_report_data
    
    if current_report_data is None:
        
        return jsonify({
            
            'success': False,
            
            'error': 'No report data available'
            
        }), 404
        
    

        
    return jsonify({
        
        'success': True,
        'batch_id': current_report_data['batch_id'],
        'project_data': current_report_data['data']['project_data'],
        'samples_data': current_report_data['data']['samples_data'],
        'samples_tw': current_report_data['data']['samples_tw'],
        'quality_controls': current_report_data['data']['quality_controls']
        
    })
    



def run_flask():
    app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)
