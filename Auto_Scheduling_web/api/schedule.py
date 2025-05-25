from flask import Flask, request, jsonify
from flask_cors import CORS
from datetime import datetime
from pulp import LpMaximize, LpProblem, LpVariable, lpSum
from flask_talisman import Talisman
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
import logging
from flask import Flask, request, jsonify
from flask_talisman import Talisman
app = Flask(__name__)
talisman = Talisman(app)
talisman.init_app(app, content_security_policy=None)

limiter = Limiter(app, key_func=get_remote_address, default_limits=["200 per day", "50 per hour"])
logging.basicConfig(level=logging.INFO)

@app.before_request
def log_request():
    logging.info(f"Request: {request.method} {request.url} - Data: {request.json}")

@app.errorhandler(Exception)
def handle_exception(e):
    """全局錯誤處理"""
    logging.error(f"Error: {str(e)}")
    return jsonify({'error': 'Internal Server Error', 'details': str(e)}), 500

@app.route('/api/schedule', methods=['POST'])
def schedule():
    try:
        data = request.json
        result = calculate_schedule(data)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def calculate_schedule(data):
    personnel = data['personnel']
    stations_data = data['stations']
    employees = [p['id'] for p in personnel if not p.get('onLeave', False)] + ["overpeople"]
    stations = [s['id'] for s in stations_data]
    capabilities = {}
    
    for s in stations:
        capabilities[("overpeople", s)] = 1  
    
    for p in personnel:
        for s in stations:
            ability = 0
            if s in p['stationSettings']:
                ability = 1 if p['stationSettings'][s]['ability'] else 0
            capabilities[(p['id'], s)] = ability
    
    preferences = {}
    
    for s in stations:
        preferences[("overpeople", s)] = -100  
    preference_values = {'low': 0, 'normal': 10, 'high': 20}
    
    for p in personnel:
        for s in stations:
            pref = 0
            if s in p['stationSettings']:
                pref = preference_values[p['stationSettings'][s]['preference']]
            preferences[(p['id'], s)] = pref
    
    unlimited_stations = [s['id'] for s in stations_data if s.get('highPriority', False)]
    priority_stations = {s['id']: 30 for s in stations_data if s.get('priorityAssign', False)}

    model = LpProblem(name="scheduling", sense=LpMaximize)
    x = LpVariable.dicts("assign", [(e, s) for e in employees for s in stations], cat="Binary")
    
    model += (lpSum(preferences[e, s] * x[e, s] for e in employees for s in stations) + 
             lpSum(priority_stations.get(s, 0) * lpSum(x[e, s] for e in employees if e != "overpeople") 
                  for s in stations))
    
    for e in employees:
        if e != "overpeople":
            model += lpSum(x[e, s] for s in stations) == 1
    
    for s in stations:
        if s not in unlimited_stations:
            model += lpSum(x[e, s] for e in employees) == 1
    
    for e in employees:
        for s in stations:
            model += x[e, s] <= capabilities[e, s]
    
    model.solve()
    
    assignments = []
    for e in employees:
        if e != "overpeople":  
            for s in stations:
                if x[e, s].value() == 1:
                    person = next(p for p in personnel if p['id'] == e)
                    station = next(s_data for s_data in stations_data if s_data['id'] == s)
                    line = next(l for l in data['productionLines'] if l['id'] == station['lineId'])
                    
                    assignments.append({
                        'person': person,
                        'station': station,
                        'line': line
                    })
    
    return {
        'assignments': assignments,
        'timestamp': datetime.now().isoformat()
    }


