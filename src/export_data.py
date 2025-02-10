"""
   Exp
"""

import sqlite3
import pandas as pd
def export_report(database_path):
    conn = sqlite3.connect(database_path)
    cur = conn.cursor()
    # Test the connection
    print("Test SQLite connection...")
    result = cur.execute('SELECT SQLITE_VERSION()')
    version = result.fetchone()
    print("SQLite version:", version)
    # obtain the bus name information
    cur.execute(r"SELECT IDFrom FROM LFR WHERE TYPE!=0 AND kV != 0;")
    bus_ID = cur.fetchall()
    # obtain the bus voltage magnitude information
    cur.execute(r"select kV from LFR WHERE TYPE!=0 AND kV != 0;")
    base_kV = cur.fetchall()
    cur.execute(r"select VoltMag from LFR WHERE TYPE!=0 AND kV != 0;")
    volt_mag = cur.fetchall()
    # obtain the bus voltage angle information
    cur.execute(r"select VoltAng from LFR WHERE TYPE!=0 AND kV != 0;")
    volt_ang = cur.fetchall()
    cur.close()
    # combine the voltage information
    result_bus = {"bus_ID": [item[0] for item in bus_ID],
                  "volt_mag": [item[0] / 100 for item in volt_mag],
                  "volt_ang": [item[0] for item in volt_ang]}
    result_bus = pd.DataFrame(result_bus)
    return result_bus