# Generate_daily_stats.py

import sys
import os
import configparser
import smtplib
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Optional, List, Tuple, Dict
from datetime import datetime

# --- Static Configuration ---
try:
    script_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
except NameError:
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

CONFIG_FILE_NAME = "config.ini"
CONFIG_FILE_PATH = os.path.join(script_dir, CONFIG_FILE_NAME)

try:
    import pyodbc
except ImportError:
    print("FATAL: The 'pyodbc' library is not installed. Please install it with 'pip install pyodbc'", file=sys.stderr)
    sys.exit(1)


def get_db_connection(db_cfg: Dict[str, str]) -> Optional[pyodbc.Connection]:
    """Establishes and returns a database connection using a config dictionary."""
    try:
        conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={db_cfg['server']};DATABASE={db_cfg['database']};UID={db_cfg['user']};PWD={db_cfg['password']};"
        print(f"Connecting to {db_cfg['database']} on {db_cfg['server']}...")
        return pyodbc.connect(conn_str)
    except Exception as e:
        print(f"ERROR: Could not connect to the database.\n{e}")
        return None

def fetch_agent_stats(conn: pyodbc.Connection) -> Optional[Tuple[List[pyodbc.Row], Optional[datetime]]]:
    """Runs the aggregation query for the most recent date and returns the results and the date."""
    sql_query = """
    WITH LatestDate AS (
        SELECT MAX(CAST(ProcessingDateTime AS DATE)) AS MaxDate
        FROM dbo.IndividualCallAnalyses
    )
    SELECT
        A.AgentName,
        SUM(CASE WHEN IEI.Finding = 'Positive' THEN 1 ELSE 0 END) AS PositiveFindings,
        SUM(CASE WHEN IEI.Finding = 'Negative' THEN 1 ELSE 0 END) AS NegativeFindings,
        SUM(CASE WHEN IEI.Finding = 'Neutral' THEN 1 ELSE 0 END) AS NeutralFindings,
        COUNT(IEI.Finding) AS TotalFindings,
        (
            (SUM(CASE WHEN IEI.Finding = 'Positive' THEN 1.0 ELSE 0 END) + (SUM(CASE WHEN IEI.Finding = 'Neutral' THEN 1.0 ELSE 0 END) / 2.0))
            / NULLIF(COUNT(IEI.Finding), 0)
        ) * 100 AS ScorePercentage,
        (SELECT MaxDate FROM LatestDate) as ReportDate
    FROM
        dbo.IndividualEvaluationItems AS IEI
    JOIN
        dbo.IndividualCallAnalyses AS ICA ON IEI.AnalysisID = ICA.AnalysisID
    JOIN
        dbo.Agents AS A ON ICA.AgentID = A.AgentID
    WHERE
        CAST(ICA.ProcessingDateTime AS DATE) = (SELECT MaxDate FROM LatestDate)
    GROUP BY
        A.AgentName
    HAVING COUNT(IEI.Finding) > 0
    ORDER BY
        A.AgentName;
    """
    try:
        print("Fetching statistics for the most recent analysis date...")
        cursor = conn.cursor()
        cursor.execute(sql_query)
        rows = cursor.fetchall()
        
        if not rows:
            print("No data found for the most recent analysis date."); return None, None

        report_date = rows[0].ReportDate
        print(f"Found statistics for {len(rows)} agents on {report_date.strftime('%Y-%m-%d')}.")
        return rows, report_date
    except pyodbc.ProgrammingError as e:
        if "Invalid column name 'ProcessingDateTime'" in str(e):
             print("\nCRITICAL DATABASE ERROR: The 'ProcessingDateTime' column is missing from 'IndividualCallAnalyses'.")
        else:
            print(f"ERROR: A database programming error occurred: {e}")
        return None, None
    except Exception as e:
        print(f"ERROR: An unexpected error occurred while running the query: {e}"); return None, None

def create_html_report(rows: List[pyodbc.Row], report_date: datetime) -> str:
    """Creates a formatted HTML string from the query results."""
    print("Generating HTML report...")
    html = f"""
    <html><head><style>
        body {{ font-family: 'Segoe UI', sans-serif; }} table {{ border-collapse: collapse; width: 95%; margin: 20px auto; }}
        th, td {{ border: 1px solid #ddd; text-align: left; padding: 12px; }}
        th {{ background-color: #004a99; color: white; text-align: center; }}
        tr:nth-child(even) {{ background-color: #f8f8f8; }} h2, p {{ text-align: center; color: #004a99; }}
        .score {{ font-weight: bold; }} .positive {{ color: green; }} .negative {{ color: red; }}
    </style></head><body>
        <h2>Agent QA Score Report</h2><p>Displaying Findings for: {report_date.strftime('%A, %B %d, %Y')}</p>
        <table><thead><tr><th>Agent Name</th><th>Score</th><th>Positive</th><th>Negative</th><th>Neutral</th><th>Total Reviewed</th></tr></thead><tbody>
    """
    for row in rows:
        score_str = f"{row.ScorePercentage:.1f}%" if row.ScorePercentage is not None else "N/A"
        html += f"""
        <tr><td>{row.AgentName}</td><td style="text-align: center;" class="score">{score_str}</td>
            <td style="text-align: center;" class="positive">{row.PositiveFindings}</td>
            <td style="text-align: center;" class="negative">{row.NegativeFindings}</td>
            <td style="text-align: center;">{row.NeutralFindings}</td>
            <td style="text-align: center;">{row.TotalFindings}</td></tr>"""
    html += "</tbody></table></body></html>"
    print("HTML report generated successfully.")
    return html

def send_email(smtp_cfg: Dict[str, str], email_cfg: Dict[str, str], html_report: str, report_date: datetime):
    """Constructs and sends an email with the HTML report."""
    try:
        from_addr = email_cfg['from']
        to_addrs = [addr.strip() for addr in email_cfg.get('to', '').split(';') if addr.strip()]
        if not to_addrs:
            print("ERROR: No 'TO' recipients found in [Management Report Emails] section of config.ini."); return

        recipients = to_addrs + [addr.strip() for addr in email_cfg.get('cc', '').split(';') if addr.strip()]
        
        msg = MIMEMultipart('alternative')
        msg['Subject'] = f"Agent QA Score Report for {report_date.strftime('%Y-%m-%d')}"
        msg['From'] = from_addr
        msg['To'] = ', '.join(to_addrs)
        if cc_list := [addr.strip() for addr in email_cfg.get('cc', '').split(';') if addr.strip()]:
            msg['Cc'] = ', '.join(cc_list)
        msg.attach(MIMEText(html_report, 'html'))
        
        pwd = base64.b64decode(smtp_cfg['password_b64']).decode('utf-8')

        print(f"Connecting to SMTP server: {smtp_cfg['server']}:{smtp_cfg['port']}...")
        with smtplib.SMTP(smtp_cfg['server'], int(smtp_cfg['port'])) as server:
            if smtp_cfg.get('usestarttls', 'true').lower() == 'true':
                server.starttls()
            server.login(smtp_cfg['uid'], pwd)
            print(f"Sending email to: {recipients}...")
            server.sendmail(from_addr, recipients, msg.as_string())
            print("Email sent successfully!")
    except Exception as e:
        print(f"ERROR: Failed to send email: {e}")

def main():
    """Main execution function."""
    if not os.path.exists(CONFIG_FILE_PATH):
        print(f"CRITICAL: Configuration file not found at '{CONFIG_FILE_PATH}'."); return

    config = configparser.ConfigParser()
    config.read(CONFIG_FILE_PATH)
    
    try:
        db_settings = dict(config.items('Database'))
        smtp_settings = dict(config.items('SMTP'))
        email_settings = dict(config.items('Management Report Emails'))
    except (configparser.Error, KeyError) as e:
        print(f"CRITICAL: Missing or invalid setting in '{CONFIG_FILE_NAME}': {e}"); return
    
    conn = get_db_connection(db_settings)
    if not conn: return
        
    try:
        report_data, report_date = fetch_agent_stats(conn)
        if report_data and report_date:
            html_content = create_html_report(report_data, report_date)
            send_email(smtp_settings, email_settings, html_content, report_date)
        else:
            print("No data found to generate a report.")
    finally:
        if conn:
            conn.close()
            print("Database connection closed.")

if __name__ == "__main__":
    main()