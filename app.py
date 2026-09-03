from flask import Flask, render_template_string, request, redirect, url_for, send_file
import sqlite3
import psycopg2
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

# Fetch database connection URL from Render environment variable
DB_URL = os.environ.get('DATABASE_URL')

def get_db():
    """Dynamically connects to PostgreSQL on Render or SQLite locally."""
    if DB_URL:
        # Render provides 'postgres://', psycopg2 requires 'postgresql://'
        url = DB_URL.replace("postgres://", "postgresql://", 1) if DB_URL.startswith("postgres://") else DB_URL
        return psycopg2.connect(url)
    return sqlite3.connect("debug_traceability.db")

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    
    if DB_URL:
        # PostgreSQL syntax
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
            )
        ''')
    else:
        # SQLite syntax
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
            )
        ''')
    conn.commit()
    conn.close()

init_db()

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Production Traceability & Failure Analytics</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f4f6f9; }
        .container { max-width: 1150px; margin: 0 auto; }
        .card { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-bottom: 20px; }
        h2 { margin-top: 0; color: #102C57; }
        .grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; }
        .full-width { grid-column: span 3; }
        label { font-weight: bold; font-size: 0.85em; color: #333; display: block; margin-bottom: 5px; }
        input, select, textarea { width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; font-size: 14px; }
        button { background-color: #102C57; color: white; border: none; padding: 12px 20px; font-size: 16px; font-weight: bold; border-radius: 4px; cursor: pointer; width: 100%; }
        button:hover { background-color: #0b1f3f; }
        .btn-export { background-color: #28a745; width: auto; float: right; padding: 8px 15px; font-size: 14px; }
        .stats { display: flex; justify-content: space-between; margin-bottom: 20px; }
        .stat-box { background: white; padding: 15px; border-radius: 6px; text-align: center; flex: 1; margin: 0 5px; box-shadow: 0 1px 5px rgba(0,0,0,0.05); }
        .stat-number { font-size: 22px; font-weight: bold; color: #102C57; }
        .stat-pct { font-size: 13px; color: #666; font-weight: normal; margin-top: 4px; }
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
        <h2 style="float: left;">Traceability & Failure Analysis Portal</h2>
        <a href="/export"><button class="btn-export">📊 Export Full Excel Analytics</button></a>
    </div>

    <!-- Live Percentage Dashboard -->
    <div class="stats">
        <div class="stat-box">
            <div class="stat-number">{{ stats['total'] }}</div>
            <div>Total Debugged</div>
            <div class="stat-pct">100% of Logged Units</div>
        </div>
        <div class="stat-box">
            <div class="stat-number" style="color: #17a2b8;">{{ stats['retest_only'] }}</div>
            <div>Direct Retest Passed</div>
            <div class="stat-pct"><strong>{{ stats['retest_pct'] }}%</strong> of Total</div>
        </div>
        <div class="stat-box">
            <div class="stat-number" style="color: #28a745;">{{ stats['repaired'] }}</div>
            <div>Repaired & Passed</div>
            <div class="stat-pct"><strong>{{ stats['repaired_pct'] }}%</strong> of Total</div>
        </div>
        <div class="stat-box">
            <div class="stat-number" style="color: #dc3545;">{{ stats['scrapped'] }}</div>
            <div>Scrapped / Failed</div>
            <div class="stat-pct"><strong>{{ stats['scrapped_pct'] }}%</strong> of Total</div>
        </div>
    </div>

    <!-- Failure Type Breakdown Cards -->
    <div class="card" style="padding: 15px;">
        <h4 style="margin-top:0; color:#102C57;">Initial Failure Breakdown (%)</h4>
        <div style="display: flex; gap: 10px; font-size: 13px;">
            {% for code, count in err_counts.items() %}
            <div style="background: #f0f4f8; padding: 8px 12px; border-radius: 5px; flex: 1; text-align: center;">
                <strong>{{ code }}</strong><br>
                <span style="font-size: 16px; font-weight: bold; color: #102C57;">{{ count }}</span> 
                <span style="color: #555;">({{ "%.1f"|format(count / stats['total'] * 100) if stats['total'] > 0 else 0 }}%)</span>
            </div>
            {% endfor %}
        </div>
    </div>

    <!-- Entry Form -->
    <div class="card">
        <form action="/add" method="POST">
            <div class="grid">
                <div>
                    <label for="project">1. Project Number (e.g., 107, 108, 109):</label>
                    <input type="text" id="project" name="project_number" placeholder="e.g. 107" required>
                </div>
                <div>
                    <label for="sn">2. Board Serial Number (Scan Barcode):</label>
                    <input type="text" id="sn" name="serial_number" placeholder="Scan unit SN..." autofocus required>
                </div>
                <div>
                    <label for="error_code">3. Initial Failure / Error Code:</label>
                    <select id="error_code" name="error_code" required>
                        <option value="">-- Select Error Code --</option>
                        <option value="ERR_PMIC_VOLT">ERR_PMIC_VOLT (Power Rail Failure)</option>
                        <option value="ERR_CLOCK_GEN">ERR_CLOCK_GEN (Clock Flashing Fail)</option>
                        <option value="ERR_DDR_READ">ERR_DDR_READ (DDR Test Fail)</option>
                        <option value="ERR_ETH_COMM">ERR_ETH_COMM (Ethernet Fail)</option>
                        <option value="ERR_FPGA_CONFIG">ERR_FPGA_CONFIG (FPGA Boot Fail)</option>
                        <option value="ERR_NO_BOOT">ERR_NO_BOOT (No Power / No Boot)</option>
                        <option value="OTHER">OTHER (Specify in notes)</option>
                    </select>
                </div>
                <div>
                    <label for="action_type">4. How Was It Solved? (Action Category):</label>
                    <select id="action_type" name="action_type" required>
                        <option value="Direct Retest (No Repair)">Direct Retest (No Repair - Socket/Contact Issue)</option>
                        <option value="Rework / Component Replacement">Rework / Component Replacement</option>
                        <option value="Reflow Soldering">Reflow / Solder Bridge Clean</option>
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
        <h3>Recent Debug Activity</h3>
        <table>
            <thead>
                <tr>
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
                    <td>{{ row[3] }}</td>
                    <td><span class="proj-badge">Project {{ row[1] }}</span></td>
                    <td><strong>{{ row[2] }}</strong></td>
                    <td>{{ row[4] }}</td>
                    <td>{{ row[5] }}</td>
                    <td>{{ row[6] if row[6] else '-' }}</td>
                    <td>
                        {% if row[9] == 'PASSED' and row[5] == 'Direct Retest (No Repair)' %}
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

@app.route('/')
def index():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM debug_logs ORDER BY id DESC LIMIT 15")
    logs = cursor.fetchall()

    cursor.execute("SELECT COUNT(*) FROM debug_logs")
    total = cursor.fetchone()[0] or 1  # Avoid div by zero

    cursor.execute("SELECT COUNT(*) FROM debug_logs WHERE action_type = 'Direct Retest (No Repair)' AND final_status = 'PASSED'")
    retest_only = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM debug_logs WHERE action_type != 'Direct Retest (No Repair)' AND final_status = 'PASSED'")
    repaired = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM debug_logs WHERE final_status IN ('SCRAPPED', 'FAILED')")
    scrapped = cursor.fetchone()[0]

    # Failure counts
    cursor.execute("SELECT error_code, COUNT(*) FROM debug_logs GROUP BY error_code")
    err_counts = dict(cursor.fetchall())

    conn.close()

    actual_total = total if total > 1 or len(logs) > 0 else 0

    stats = {
        'total': actual_total,
        'retest_only': retest_only,
        'retest_pct': round((retest_only / actual_total) * 100, 1) if actual_total > 0 else 0,
        'repaired': repaired,
        'repaired_pct': round((repaired / actual_total) * 100, 1) if actual_total > 0 else 0,
        'scrapped': scrapped,
        'scrapped_pct': round((scrapped / actual_total) * 100, 1) if actual_total > 0 else 0,
    }

    return render_template_string(HTML_TEMPLATE, logs=logs, stats=stats, err_counts=err_counts)

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
    
    # Use %s for PostgreSQL (Render/Neon) and ? for local SQLite
    p = "%s" if DB_URL else "?"
    
    query = f'''
        INSERT INTO debug_logs (project_number, serial_number, timestamp, error_code, action_type, replaced_components, action_taken, debug_technician, final_status)
        VALUES ({p}, {p}, {p}, {p}, {p}, {p}, {p}, {p}, {p})
    '''
    
    cursor.execute(query, (project_number, sn, timestamp, error_code, action_type, replaced_components, action_taken, tech, final_status))
    conn.commit()
    conn.close()

    return redirect(url_for('index'))

@app.route('/export')
def export():
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM debug_logs", conn)
    conn.close()

    export_path = "Debug_Traceability_Analytics.xlsx"

    with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
        # Sheet 1: Detailed Logs
        df.to_excel(writer, sheet_name='All Units Log', index=False)

        # Sheet 2: Resolution Breakdown (Initial Failure vs How Solved)
        if not df.empty and 'error_code' in df.columns and 'action_type' in df.columns:
            pivot_table = pd.crosstab(df['error_code'], df['action_type'], margins=True, margins_name='Total')
            pivot_table.to_excel(writer, sheet_name='Failure vs Resolution Summary')

        # Sheet 3: Percentage Summary
        if not df.empty and 'project_number' in df.columns and 'final_status' in df.columns:
            stats_summary = df.groupby(['project_number', 'final_status']).size().unstack(fill_value=0)
            stats_summary.to_excel(writer, sheet_name='Project Yield Summary')

    return send_file(export_path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
