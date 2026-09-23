from flask import Flask, render_template_string, request, redirect, url_for, send_file
import sqlite3
import psycopg2
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

DB_URL = os.environ.get('DATABASE_URL')

def get_db():
    if DB_URL:
        url = DB_URL.replace("postgres://", "postgresql://", 1) if DB_URL.startswith("postgres://") else DB_URL
        return psycopg2.connect(url)
    return sqlite3.connect("debug_traceability.db")

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    
    if DB_URL:
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS debug_logs (
                id SERIAL PRIMARY KEY,
                project_number TEXT NOT NULL,
                serial_number TEXT NOT NULL,
                timestamp TEXT NOT NULL,
                error_code TEXT NOT NULL,
                action_type TEXT NOT NULL,
                replaced_components TEXT,
                action_taken TEXT,
                debug_technician TEXT,
                final_status TEXT NOT NULL
            );
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS project_batches (
                project_number TEXT PRIMARY KEY,
                batch_size INTEGER NOT NULL
            );
        ''')
    else:
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS debug_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_number TEXT NOT NULL,
                serial_number TEXT NOT NULL,
                timestamp TEXT NOT NULL,
                error_code TEXT NOT NULL,
                action_type TEXT NOT NULL,
                replaced_components TEXT,
                action_taken TEXT,
                debug_technician TEXT,
                final_status TEXT NOT NULL
            );
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS project_batches (
                project_number TEXT PRIMARY KEY,
                batch_size INTEGER NOT NULL
            );
        ''')
    conn.commit()
    conn.close()

init_db()

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>SAI50 Traceability & Yield Analytics</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f4f6f9; }
        .container { max-width: 1150px; margin: 0 auto; }
        .card { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-bottom: 20px; }
        h2, h3, h4 { color: #102C57; margin-top: 0; }
        
        .grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; }
        .full-width { grid-column: span 3; }
        
        label { font-weight: bold; font-size: 0.85em; color: #333; display: block; margin-bottom: 5px; }
        input, select, textarea { width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; font-size: 14px; }
        
        button { background-color: #102C57; color: white; border: none; padding: 10px 20px; font-size: 15px; font-weight: bold; border-radius: 4px; cursor: pointer; width: 100%; }
        button:hover { background-color: #0b1f3f; }
        
        .btn-export { background-color: #28a745; width: auto; float: right; padding: 8px 15px; font-size: 14px; text-decoration: none; color: white; border-radius: 4px; font-weight: bold; }
        
        .stats { display: flex; justify-content: space-between; margin-bottom: 20px; gap: 10px; }
        .stat-box { background: white; padding: 15px; border-radius: 6px; text-align: center; flex: 1; box-shadow: 0 1px 5px rgba(0,0,0,0.05); }
        .stat-number { font-size: 22px; font-weight: bold; color: #102C57; }
        .stat-pct { font-size: 13px; color: #666; font-weight: normal; margin-top: 4px; }
        
        .filter-bar { background: #e9ecef; padding: 15px; border-radius: 6px; margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between; }
        
        table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #102C57; color: white; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        
        .badge-pass { background-color: #28a745; color: white; padding: 3px 8px; border-radius: 3px; font-weight: bold; }
        .badge-retest { background-color: #17a2b8; color: white; padding: 3px 8px; border-radius: 3px; font-weight: bold; }
        .badge-fail { background-color: #dc3545; color: white; padding: 3px 8px; border-radius: 3px; font-weight: bold; }
        .proj-badge { background-color: #007bff; color: white; padding: 3px 8px; border-radius: 4px; font-weight: bold; }
    </style>
</head>
<body>
<div class="container">
    <div style="overflow: hidden; margin-bottom: 10px;">
        <h2 style="float: left;">SAI50 Program - Traceability & Yield Analytics</h2>
        <a href="/export" class="btn-export">📊 Export Full Excel Analytics</a>
    </div>

    <!-- Filter Bar -->
    <div class="filter-bar">
        <form method="GET" action="/" style="display: flex; gap: 15px; align-items: center; width: 100%;">
            <label style="margin: 0; white-space: nowrap; font-size: 1rem;">🔍 <strong>Dashboard by Project:</strong></label>
            <select name="filter_project" onchange="this.form.submit()" style="max-width: 250px;">
                <option value="ALL" {% if selected_project == 'ALL' %}selected{% endif %}>-- All Projects Combined --</option>
                {% for proj in available_projects %}
                <option value="{{ proj }}" {% if selected_project == proj %}selected{% endif %}>Project {{ proj }}</option>
                {% endfor %}
            </select>
            {% if selected_project != 'ALL' %}
            <a href="/" style="font-size: 13px; color: #007bff; text-decoration: none;">Clear Filter</a>
            {% endif %}
        </form>
    </div>

    <!-- Dashboard Metrics -->
    <div class="stats">
        <div class="stat-box">
            <div class="stat-number">{{ target_batch_size }}</div>
            <div>Planned Batch Size</div>
            <div class="stat-pct">Total Target Boards</div>
        </div>
        <div class="stat-box">
            <div class="stat-number">{{ stats['total'] }}</div>
            <div>Total Debugged</div>
            <div class="stat-pct">Unique Defective Boards</div>
        </div>
        <div class="stat-box">
            <div class="stat-number" style="color: #17a2b8;">{{ stats['retest_only'] }}</div>
            <div>Direct Retest Pass</div>
            <div class="stat-pct"><strong>{{ stats['retest_pct'] }}%</strong> of Debug</div>
        </div>
        <div class="stat-box">
            <div class="stat-number" style="color: #28a745;">{{ stats['repaired'] }}</div>
            <div>Repaired & Passed</div>
            <div class="stat-pct"><strong>{{ stats['repaired_pct'] }}%</strong> of Debug</div>
        </div>
        <div class="stat-box">
            <div class="stat-number" style="color: #dc3545;">{{ stats['scrapped'] }}</div>
            <div>Scrapped / Failed</div>
            <div class="stat-pct"><strong>{{ stats['scrapped_pct'] }}%</strong> of Debug</div>
        </div>
        <div class="stat-box" style="border: 2px solid #28a745;">
            <div class="stat-number" style="color: #28a745;">{{ overall_yield }}%</div>
            <div>True Production Yield</div>
            <div class="stat-pct">(Batch - Scrap) / Batch</div>
        </div>
    </div>

    <!-- Update Project Batch Size -->
    <div class="card" style="padding: 15px; background: #f0f4f8;">
        <h4 style="margin-bottom: 10px;">⚙️ Update Project Batch Size</h4>
        <form action="/set_batch" method="POST" style="display: flex; gap: 10px; align-items: center;">
            <input type="text" name="project_number" placeholder="Project # (e.g. EN107577)" value="{{ selected_project if selected_project != 'ALL' else '' }}" required style="flex: 1;">
            <input type="number" name="batch_size" placeholder="Total Planned Boards (e.g. 400)" value="{{ target_batch_size if target_batch_size > 0 else '' }}" required style="flex: 1;">
            <button type="submit" style="width: auto;">Save Batch Size</button>
        </form>
    </div>

    <!-- Failure Type Breakdown Cards -->
    <div class="card" style="padding: 15px;">
        <h4 style="margin-top:0;">Initial Failure Breakdown (%)</h4>
        <div style="display: flex; gap: 10px; font-size: 13px; flex-wrap: wrap;">
            {% for code, count in err_counts.items() %}
            <div style="background: #f0f4f8; padding: 8px 12px; border-radius: 5px; flex: 1; min-width: 120px; text-align: center;">
                <strong>{{ code }}</strong><br>
                <span style="font-size: 16px; font-weight: bold; color: #102C57;">{{ count }}</span> 
                <span style="color: #555;">({{ "%.1f"|format(count / stats['total'] * 100) if stats['total'] > 0 else 0 }}%)</span>
            </div>
            {% endfor %}
        </div>
    </div>

    <!-- Horizontal Entry Form -->
    <div class="card">
        <h3>Log Board Failure / Retest Entry</h3>
        <form action="/add" method="POST">
            <div class="grid">
                <div>
                    <label for="project">1. Project Number (e.g., EN107577):</label>
                    <input type="text" id="project" name="project_number" placeholder="e.g. EN107577" value="{{ selected_project if selected_project != 'ALL' else '' }}" required>
                </div>

                <div>
                    <label for="sn">2. Board Serial Number (Scan Barcode):</label>
                    <input type="text" id="sn" name="serial_number" placeholder="e.g., SN408123" pattern="SN\d{6}" title="Serial number must start with 'SN' followed by 6 digits" autofocus required>
                </div>

                <div>
                    <label for="error_code">3. Initial Failure / Error Code:</label>
                    <select id="error_code" name="error_code" required>
                        <option value="">-- Select Error Code --</option>
                        <option value="Configuration_Troot">Configuration_Troot</option>
                        <option value="Production Troot Programming">Production Troot Programming</option>
                        <option value="Linux Login Fail">Linux Login Fail</option>
                        <option value="Extended eMMc Failure">Extended eMMc Failure</option>
                        <option value="Serial Communication">Serial Communication(Error Connecting to USB1)</option>
                        <option value="No Power">No Power</option>
                        <option value="OTHER">OTHER (Specify in notes)</option>
                    </select>
                </div>

                <div>
                    <label for="action_type">4. How Was It Solved? (Action Category):</label>
                    <select id="action_type" name="action_type" required>
                        <option value="Direct Retest">Direct Retest (No Repair - Retry_Commands)</option>
                        <option value="Rework / Component Replacement">Rework / Component Replacement</option>
                        <option value="Scrap / Unrepairable">Scrap / Unrepairable</option>
                    </select>
                </div>

                <div>
                    <label for="status">5. Final Status after Retest:</label>
                    <select id="status" name="final_status" required>
                        <option value="PASSED">PASSED</option>
                        <option value="FAILED">FAILED</option>
                        <option value="SCRAPPED">SCRAPPED</option>
                    </select>
                </div>

                <div>
                    <label for="components">6. Replaced IC / Components (if any):</label>
                    <input type="text" id="components" name="replaced_components" placeholder="e.g. Replaced U4, R12">
                </div>

                <div class="full-width">
                    <label for="tech">7. Debug Technician Name / ID:</label>
                    <input type="text" id="tech" name="debug_technician" placeholder="e.g. Tech_01" required>
                </div>

                <div class="full-width">
                    <label for="notes">8. Detailed Action Taken & Root Cause Notes:</label>
                    <textarea id="notes" name="action_taken" rows="2" placeholder="Describe root cause and resolution..."></textarea>
                </div>

                <div class="full-width">
                    <button type="submit">💾 Save Log & Calculate Metrics</button>
                </div>
            </div>
        </form>
    </div>

    <!-- Recent Activity Table -->
    <div class="card">
        <h3>Recent Debug Activity {% if selected_project != 'ALL' %}(Project {{ selected_project }}){% endif %}</h3>
        <table>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>Timestamp</th>
                    <th>Project #</th>
                    <th>Serial Number</th>
                    <th>Initial Failure</th>
                    <th>How Solved</th>
                    <th>Replaced ICs</th>
                    <th>Status</th>
                    <th>Tech</th>
                </tr>
            </thead>
            <tbody>
                {% for row in logs %}
                <tr>
                    <td><strong>#{{ row[0] }}</strong></td>
                    <td>{{ row[3] }}</td>
                    <td><span class="proj-badge">{{ row[1] }}</span></td>
                    <td><strong>{{ row[2] }}</strong></td>
                    <td>{{ row[4] }}</td>
                    <td>{{ row[5] }}</td>
                    <td>{{ row[6] if row[6] else '-' }}</td>
                    <td>
                        {% if row[9] == 'PASSED' and 'Direct Retest' in row[5] %}
                            <span class="badge-retest">RETEST PASS</span>
                        {% elif row[9] == 'PASSED' %}
                            <span class="badge-pass">REPAIRED PASS</span>
                        {% else %}
                            <span class="badge-fail">{{ row[9] }}</span>
                        {% endif %}
                    </td>
                    <td>{{ row[8] }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>
</body>
</html>
'''

@app.route('/', methods=['GET'])
def index():
    selected_project = request.args.get('filter_project', 'ALL').strip()

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("SELECT DISTINCT project_number FROM debug_logs UNION SELECT project_number FROM project_batches")
    available_projects = [row[0] for row in cursor.fetchall() if row[0]]

    placeholder = "%s" if DB_URL else "?"

    if selected_project != 'ALL':
        logs_query = f"SELECT * FROM debug_logs WHERE project_number = {placeholder} ORDER BY id DESC LIMIT 15"
        cursor.execute(logs_query, (selected_project,))
    else:
        logs_query = "SELECT * FROM debug_logs ORDER BY id DESC LIMIT 15"
        cursor.execute(logs_query)
        
    logs = cursor.fetchall()

    if selected_project != 'ALL':
        cursor.execute(f"SELECT COUNT(DISTINCT serial_number) FROM debug_logs WHERE project_number = {placeholder}", (selected_project,))
        total_unique = cursor.fetchone()[0] or 0

        cursor.execute(f"SELECT COUNT(DISTINCT serial_number) FROM debug_logs WHERE project_number = {placeholder} AND action_type LIKE 'Direct Retest%' AND final_status = 'PASSED'", (selected_project,))
        retest_only = cursor.fetchone()[0] or 0

        cursor.execute(f"SELECT COUNT(DISTINCT serial_number) FROM debug_logs WHERE project_number = {placeholder} AND action_type NOT LIKE 'Direct Retest%' AND final_status = 'PASSED'", (selected_project,))
        repaired = cursor.fetchone()[0] or 0

        cursor.execute(f"SELECT COUNT(DISTINCT serial_number) FROM debug_logs WHERE project_number = {placeholder} AND final_status IN ('SCRAPPED', 'FAILED')", (selected_project,))
        scrapped = cursor.fetchone()[0] or 0

        cursor.execute(f"SELECT error_code, COUNT(*) FROM debug_logs WHERE project_number = {placeholder} GROUP BY error_code", (selected_project,))
        err_counts = dict(cursor.fetchall())

        cursor.execute(f"SELECT batch_size FROM project_batches WHERE project_number = {placeholder}", (selected_project,))
        batch_row = cursor.fetchone()
        target_batch_size = batch_row[0] if batch_row else 0
    else:
        cursor.execute("SELECT COUNT(DISTINCT serial_number) FROM debug_logs")
        total_unique = cursor.fetchone()[0] or 0

        cursor.execute("SELECT COUNT(DISTINCT serial_number) FROM debug_logs WHERE action_type LIKE 'Direct Retest%' AND final_status = 'PASSED'")
        retest_only = cursor.fetchone()[0] or 0

        cursor.execute("SELECT COUNT(DISTINCT serial_number) FROM debug_logs WHERE action_type NOT LIKE 'Direct Retest%' AND final_status = 'PASSED'")
        repaired = cursor.fetchone()[0] or 0

        cursor.execute("SELECT COUNT(DISTINCT serial_number) FROM debug_logs WHERE final_status IN ('SCRAPPED', 'FAILED')")
        scrapped = cursor.fetchone()[0] or 0

        cursor.execute("SELECT error_code, COUNT(*) FROM debug_logs GROUP BY error_code")
        err_counts = dict(cursor.fetchall())

        cursor.execute("SELECT SUM(batch_size) FROM project_batches")
        batch_sum = cursor.fetchone()[0]
        target_batch_size = batch_sum if batch_sum else 0

    conn.close()

    overall_yield = round(((target_batch_size - scrapped) / target_batch_size) * 100, 2) if target_batch_size > 0 else 0.0

    stats = {
        'total': total_unique,
        'retest_only': retest_only,
        'retest_pct': round((retest_only / total_unique) * 100, 1) if total_unique > 0 else 0,
        'repaired': repaired,
        'repaired_pct': round((repaired / total_unique) * 100, 1) if total_unique > 0 else 0,
        'scrapped': scrapped,
        'scrapped_pct': round((scrapped / total_unique) * 100, 1) if total_unique > 0 else 0,
    }

    return render_template_string(
        HTML_TEMPLATE, 
        logs=logs, 
        stats=stats, 
        err_counts=err_counts, 
        available_projects=available_projects, 
        selected_project=selected_project,
        target_batch_size=target_batch_size,
        overall_yield=overall_yield
    )

@app.route('/set_batch', methods=['POST'])
def set_batch():
    project_number = request.form['project_number'].strip()
    batch_size = int(request.form['batch_size'])

    conn = get_db()
    cursor = conn.cursor()
    placeholder = "%s" if DB_URL else "?"
    
    if DB_URL:
        cursor.execute(f"INSERT INTO project_batches (project_number, batch_size) VALUES ({placeholder}, {placeholder}) ON CONFLICT (project_number) DO UPDATE SET batch_size = EXCLUDED.batch_size", (project_number, batch_size))
    else:
        cursor.execute(f"INSERT OR REPLACE INTO project_batches (project_number, batch_size) VALUES ({placeholder}, {placeholder})", (project_number, batch_size))
        
    conn.commit()
    conn.close()
    return redirect(url_for('index', filter_project=project_number))

@app.route('/add', methods=['POST'])
def add_log():
    project_number = request.form['project_number'].strip()
    sn = request.form['serial_number'].strip()
    error_code = request.form['error_code']
    action_type = request.form['action_type']
    replaced_components = request.form['replaced_components'].strip()
    action_taken = request.form['action_taken'].strip()
    tech = request.form['debug_technician'].strip()
    final_status = request.form['final_status']
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = get_db()
    cursor = conn.cursor()
    placeholder = "%s" if DB_URL else "?"
    
    query = f'''
        INSERT INTO debug_logs (project_number, serial_number, timestamp, error_code, action_type, replaced_components, action_taken, debug_technician, final_status)
        VALUES ({placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder}, {placeholder})
    '''
    
    cursor.execute(query, (project_number, sn, timestamp, error_code, action_type, replaced_components, action_taken, tech, final_status))
    conn.commit()
    conn.close()

    return redirect(url_for('index', filter_project=project_number))

@app.route('/export')
def export():
    conn = get_db()
    df_logs = pd.read_sql_query("SELECT * FROM debug_logs ORDER BY id DESC", conn)
    df_batches = pd.read_sql_query("SELECT * FROM project_batches", conn)
    conn.close()

    export_path = "Debug_Traceability_Analytics.xlsx"

    with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
        df_logs.to_excel(writer, sheet_name='Full Retest Audit Trail', index=False)
        df_batches.to_excel(writer, sheet_name='Project Batch Sizes', index=False)

    return send_file(export_path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
