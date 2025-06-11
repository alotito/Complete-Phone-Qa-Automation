# JSON_DB_Importer.py

import sys
import os
import datetime
import shutil
import re
import logging
import argparse
import traceback
import configparser
import json
from typing import List, Tuple, Optional, Dict, Any, Set

try:
    import pyodbc
except ImportError:
    print("FATAL: The 'pyodbc' library is not installed. Please install it using 'pip install pyodbc'", file=sys.stderr, flush=True)
    sys.exit(1)


# --- Static Configuration ---
APP_NAME = "PhoneQA_DB_Importer"
CONFIG_FILE_NAME = "config.ini"
EXT_LIST_FILE_NAME = "ExtList.data"
PROCESSED_PREFIX = "Stored-"
FAILED_PREFIX = "BadData-"
COMBINED_REPORT_FILENAME = "Combined_Analysis_Report.json"


# --- Determine Script Directory ---
try:
    script_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
except NameError:
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

# --- Global Logger Instance ---
logger = logging.getLogger(APP_NAME)


def setup_logger(log_dir: str):
    """Configures the global logger instance."""
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f"{APP_NAME}_{datetime.datetime.now():%Y%m%d_%H%M%S}.log")
    logger.setLevel(logging.DEBUG)
    if logger.hasHandlers():
        logger.handlers.clear()
    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s')
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    ch.setFormatter(ch_formatter)
    logger.addHandler(ch)
    logger.info(f"Logger initialized. Log file at: {log_file}")

def find_latest_week_folder(root_dir: str) -> Optional[str]:
    """Finds the most recent 'Week ofYYYY-MM-DD' directory."""
    week_folders = []
    date_pattern = re.compile(r'Week of (\d{4}-\d{2}-\d{2})')
    try:
        for item in os.listdir(root_dir):
            full_path = os.path.join(root_dir, item)
            if os.path.isdir(full_path) and (match := date_pattern.search(item)):
                try:
                    folder_date = datetime.datetime.strptime(match.group(1), '%Y-%m-%d').date()
                    week_folders.append((folder_date, full_path))
                except ValueError:
                    logger.warning(f"Found folder '{item}' with matching pattern but invalid date.")
        if not week_folders:
            logger.warning(f"No directories matching 'Week ofYYYY-MM-DD' found in '{root_dir}'.")
            return None
        latest_folder = sorted(week_folders, key=lambda x: x[0], reverse=True)[0]
        logger.info(f"Found latest week folder: {latest_folder[1]}")
        return latest_folder[1]
    except FileNotFoundError:
        logger.error(f"Root directory '{root_dir}' not found.")
        return None
    except Exception as e:
        logger.error(f"Error scanning for latest week folder in '{root_dir}': {e}", exc_info=True)
        return None

def get_db_connection(config: configparser.ConfigParser) -> pyodbc.Connection:
    """Establishes and returns a pyodbc connection to the SQL Server database."""
    try:
        db_config = config['Database']
        conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={db_config['Server']};DATABASE={db_config['Database']};UID={db_config['User']};PWD={db_config['Password']};"
        conn = pyodbc.connect(conn_str)
        logger.info(f"Successfully connected to database '{db_config['Database']}' on server '{db_config['Server']}'.")
        return conn
    except Exception as e:
        logger.critical(f"Failed to establish database connection: {e}", exc_info=True)
        raise

def parse_extlist_data(extlist_path: str) -> Dict[str, Dict[str, str]]:
    """
    Parses the ExtList.data file into a dictionary keyed by EXTENSION.
    Safely handles lines with more than 3 columns.
    """
    members_by_ext = {}
    logger.info(f"Parsing agent data from: {extlist_path}")
    if not os.path.exists(extlist_path):
        logger.error(f"Agent data file not found: '{extlist_path}'. Cannot map agents."); return {}
    try:
        with open(extlist_path, 'r', encoding='utf-8') as f:
            for line in f:
                if not (s_line := line.strip()) or s_line.startswith('#'): continue
                parts = s_line.split('\t')
                # MODIFIED: Check for at least 3 parts, ignores any beyond the 3rd.
                if len(parts) >= 3 and (ext := parts[0].strip()):
                    members_by_ext[ext] = {"full_name": parts[1].strip(), "email": parts[2].strip(), "extension": ext}
        logger.info(f"Successfully parsed {len(members_by_ext)} members."); return members_by_ext
    except Exception as e:
        logger.error(f"Error parsing agent data file '{extlist_path}': {e}", exc_info=True); return {}

def extract_extension_from_path(file_path: str) -> Optional[str]:
    """Extracts the 4-digit extension from the file's directory path."""
    if match := re.search(r"Week of \d{4}-\d{2}-\d{2}[\\/](\d{4})", file_path):
        return match.group(1)
    logger.warning(f"Could not extract extension from path: '{file_path}'."); return None

def get_or_create_agent(cursor: pyodbc.Cursor, agent_details: Dict[str, str]) -> Optional[int]:
    """Gets the AgentID for a given agent, creating them if they don't exist."""
    agent_name, extension = agent_details.get("full_name"), agent_details.get("extension")
    if not agent_name or not extension:
        logger.error(f"Agent details missing name or extension: {agent_details}."); return None
    try:
        cursor.execute("SELECT AgentID FROM Agents WHERE Extension = ?", extension)
        if row := cursor.fetchone(): return row.AgentID
        logger.info(f"Agent '{agent_name}' with ext '{extension}' not found. Creating new record.")
        insert_sql = "INSERT INTO Agents (AgentName, EmailAddress, Extension) OUTPUT INSERTED.AgentID VALUES (?, ?, ?);"
        return cursor.execute(insert_sql, agent_name, agent_details.get('email'), extension).fetchval()
    except Exception as e:
        logger.error(f"DB error getting/creating agent '{agent_name}': {e}", exc_info=True); raise

def get_or_create_quality_points(cursor: pyodbc.Cursor, qp_texts: Set[str]) -> Dict[str, int]:
    """Efficiently gets IDs for existing quality points and creates non-existent ones."""
    if not (sanitized_qp_texts := {text.strip() for text in qp_texts if text}): return {}
    qp_map = {}
    try:
        placeholders = ', '.join(['?'] * len(sanitized_qp_texts))
        sql_select = f"SELECT QualityPointText, QualityPointID FROM QualityPointsMaster WHERE QualityPointText IN ({placeholders})"
        cursor.execute(sql_select, *list(sanitized_qp_texts))
        for row in cursor.fetchall(): qp_map[row.QualityPointText] = row.QualityPointID
        
        if new_qps := [text for text in sanitized_qp_texts if text not in qp_map]:
            logger.info(f"Found {len(new_qps)} new quality points to insert.")
            params = [(text, 1 if "[BONUS]" in text.upper() else 0) for text in new_qps]
            cursor.fast_executemany = True
            cursor.executemany("INSERT INTO QualityPointsMaster (QualityPointText, IsBonus) VALUES (?, ?)", params)
            cursor.execute(sql_select, *list(sanitized_qp_texts))
            for row in cursor.fetchall(): qp_map[row.QualityPointText] = row.QualityPointID
        
        return {orig_text: qp_map.get(orig_text.strip()) for orig_text in qp_texts if orig_text.strip() in qp_map}
    except Exception as e:
        logger.error(f"DB error getting/creating quality points: {e}", exc_info=True); raise

def process_individual_json(cursor: pyodbc.Cursor, json_data: Dict, file_path: str, agent_id: int, qp_map: Dict, ts: datetime.datetime):
    """Processes a single individual analysis JSON and inserts data into the database."""
    summary, remarks = json_data.get('call_summary', {}), json_data.get('concluding_remarks', {})
    sql = """INSERT INTO IndividualCallAnalyses (AgentID, TechDispatcherNameRaw, OriginalAudioFileName, CallDuration, ClientName,
               ClientFacilityCompany, TicketNumber, ClientCallbackNumber, TicketStatusType, CallSubjectSummary, 
               ConcludingRemarks_Positive, ConcludingRemarks_Negative, ConcludingRemarks_Coaching, ProcessingDateTime)
               OUTPUT INSERTED.AnalysisID VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);"""
    params = (agent_id, summary.get('tech_dispatcher_name'), os.path.basename(file_path).replace('_analysis.json', '.wav'),
              summary.get('call_duration'), summary.get('client_name'), summary.get('client_facility_company'),
              summary.get('ticket_number'), summary.get('client_callback_number'), summary.get('ticket_status_type'),
              summary.get('call_subject_summary'), remarks.get('summary_positive_findings'),
              remarks.get('summary_negative_findings'), remarks.get('coaching_plan_for_growth'), ts)
    analysis_id = cursor.execute(sql, params).fetchval()
    
    eval_params = [(analysis_id, qp_map.get(item.get('quality_point', '').strip()), item.get('finding'), item.get('explanation_snippets'))
                   for item in json_data.get('detailed_evaluation', []) if qp_map.get(item.get('quality_point', '').strip())]
    if eval_params:
        cursor.fast_executemany = True
        cursor.executemany("INSERT INTO IndividualEvaluationItems (AnalysisID, QualityPointID, Finding, ExplanationSnippets) VALUES (?, ?, ?, ?)", eval_params)
        logger.debug(f"Inserted {len(eval_params)} evaluation items for Individual AnalysisID {analysis_id}.")

def process_combined_json(cursor: pyodbc.Cursor, json_data: Dict, agent_id: int, qp_map: Dict, ts: datetime.datetime):
    """Processes the combined analysis JSON and inserts data into the database."""
    header, snapshot = json_data.get('report_header', {}), json_data.get('overall_performance_snapshot', {})
    counts = snapshot.get('aggregate_findings_counts', {})
    sql = """INSERT INTO CombinedAnalyses (AgentID, AnalysisPeriodNote, NumberOfReportsProvided, NumberOfReportsSuccessfullyAnalyzed,
               Snapshot_TotalCallsContributing, Snapshot_PositiveCount, Snapshot_NegativeCount, Snapshot_NeutralCount, ProcessingDateTime)
               OUTPUT INSERTED.CombinedAnalysisID VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);"""
    params = (agent_id, header.get('analysis_period_note'), header.get('number_of_reports_provided'), header.get('number_of_reports_successfully_analyzed'),
              snapshot.get('total_calls_contributing_to_aggregates'), counts.get('positive_count'), counts.get('negative_count'), counts.get('neutral_count'), ts)
    combined_id = cursor.execute(sql, params).fetchval()

    qual_summary = json_data.get('qualitative_summary_and_coaching_plan', {})
    cursor.fast_executemany = True
    if strengths := qual_summary.get('overall_strengths_observed', []):
        cursor.executemany("INSERT INTO CombinedAnalysisStrengths (CombinedAnalysisID, StrengthText) VALUES (?, ?)", [(combined_id, s) for s in strengths])
    if dev_areas := qual_summary.get('overall_areas_for_development', []):
        cursor.executemany("INSERT INTO CombinedAnalysisDevelopmentAreas (CombinedAnalysisID, DevelopmentAreaText) VALUES (?, ?)", [(combined_id, d) for d in dev_areas])
    for focus_item in qual_summary.get('consolidated_coaching_focus', []):
        if area_text := focus_item.get('area'):
            focus_id = cursor.execute("INSERT INTO CombinedAnalysisCoachingFocus (CombinedAnalysisID, AreaText) OUTPUT INSERTED.CoachingFocusID VALUES (?, ?);", combined_id, area_text).fetchval()
            if actions := focus_item.get('specific_actions', []):
                cursor.executemany("INSERT INTO CombinedAnalysisCoachingActions (CoachingFocusID, ActionText) VALUES (?, ?)", [(focus_id, a) for a in actions])
    
    qp_params = [(combined_id, qp_map.get(item.get('quality_point', '').strip()), item.get('findings_summary', {}).get('positive_count'), 
                  item.get('findings_summary', {}).get('negative_count'), item.get('findings_summary', {}).get('neutral_count'), item.get('trend_observation'))
                 for item in json_data.get('detailed_quality_point_analysis', []) if qp_map.get(item.get('quality_point', '').strip())]
    if qp_params:
        cursor.executemany("INSERT INTO CombinedAnalysisQualityPointDetails (CombinedAnalysisID, QualityPointID, FindingsSummary_Positive, FindingsSummary_Negative, FindingsSummary_Neutral, TrendObservation) VALUES (?, ?, ?, ?, ?, ?)", qp_params)

def process_folder(target_folder: str, config: configparser.ConfigParser):
    """Orchestrates the processing of all valid JSON files within a given folder."""
    logger.info(f"Starting processing for folder: {target_folder}")
    conn = None
    try:
        ext_map = parse_extlist_data(os.path.join(script_dir, EXT_LIST_FILE_NAME))
        files_to_process = [os.path.join(r, f) for r, _, f_list in os.walk(target_folder) for f in f_list if (f.endswith('_analysis.json') or COMBINED_REPORT_FILENAME in f) and not f.startswith((PROCESSED_PREFIX, FAILED_PREFIX))]

        if not files_to_process:
            logger.info("No new, valid JSON report files to process in this folder."); return
        logger.info(f"Found {len(files_to_process)} new JSON reports to process.")
        
        conn = get_db_connection(config)
        cursor = conn.cursor()
        ts = datetime.datetime.now()
        logger.info(f"Using processing timestamp for this batch: {ts.strftime('%Y-%m-%d %H:%M:%S.%f')}")

        for file_path in files_to_process:
            base_name = os.path.basename(file_path)
            logger.info(f"--- Processing file: {base_name} ---")
            try:
                extension = extract_extension_from_path(file_path)
                if not extension:
                    logger.error(f"Could not determine extension for '{base_name}'. Skipping file."); continue
                
                agent_details = ext_map.get(extension, {"full_name": f"Un-rostered Agent - {extension}", "email": None, "extension": extension})
                
                with open(file_path, 'r', encoding='utf-8') as f: json_data = json.load(f)
                
                conn.autocommit = False
                agent_id = get_or_create_agent(cursor, agent_details)
                if not agent_id: raise ValueError(f"Could not get/create ID for agent: {agent_details}")

                all_qps = {item.get('quality_point') for item in json_data.get("detailed_evaluation", []) if item.get('quality_point')}
                all_qps.update({item.get('quality_point') for item in json_data.get("detailed_quality_point_analysis", []) if item.get('quality_point')})
                qp_map = get_or_create_quality_points(cursor, all_qps)

                if COMBINED_REPORT_FILENAME in base_name:
                    process_combined_json(cursor, json_data, agent_id, qp_map, ts)
                else:
                    process_individual_json(cursor, json_data, file_path, agent_id, qp_map, ts)
                
                conn.commit()
                shutil.move(file_path, os.path.join(os.path.dirname(file_path), f"{PROCESSED_PREFIX}{base_name}"))
                logger.info(f"Successfully processed and renamed '{base_name}'.")
            except Exception as e:
                logger.error(f"Failed to process file '{file_path}': {e}", exc_info=True)
                if conn and not conn.autocommit: conn.rollback()
                try:
                    shutil.move(file_path, os.path.join(os.path.dirname(file_path), f"{FAILED_PREFIX}{base_name}"))
                except Exception as e_mv: logger.error(f"CRITICAL: Failed to rename error file '{file_path}': {e_mv}")
    
    except Exception as e:
        logger.critical(f"A major error occurred during folder processing: {e}", exc_info=True)
    finally:
        if conn: conn.close(); logger.info("Database connection closed.")

def main():
    """Main execution function."""
    log_dir = os.path.join(script_dir, "logs", APP_NAME)
    setup_logger(log_dir)
    
    try:
        config_path = os.path.join(script_dir, CONFIG_FILE_NAME)
        if not os.path.exists(config_path):
            logger.critical(f"Configuration file not found at '{config_path}'."); return
        
        config = configparser.ConfigParser(); config.read(config_path)
        source_root = config.get('Paths', 'ImporterSourceRoot')
        
        target_folder = find_latest_week_folder(source_root)
        if not target_folder:
            logger.info("No folder to process. Exiting."); return
        
        process_folder(target_folder, config)
    except (configparser.Error, KeyError) as e:
        logger.critical(f"Missing or invalid setting in '{CONFIG_FILE_NAME}': {e}")
    except Exception as e:
        logger.critical(f"An unhandled exception occurred in main: {e}", exc_info=True)
        
    logger.info(f"{APP_NAME} finished.")

if __name__ == '__main__':
    main()
