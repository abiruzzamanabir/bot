from flask import Flask, request, render_template, jsonify
import os
import shutil
import pandas as pd
import re
import time
import logging
import psutil

app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.INFO)

# Progress tracking
progress = {"percent": 0}
current_process = {"message": "No files processed yet."}

def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', ' ', filename)

def close_excel():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'EXCEL.EXE':
            proc.kill()

def process_files(excel_file, video_folder, output_base_folder):
    global progress, current_process
    progress["percent"] = 0  # Reset progress at the start
    current_process["message"] = "Starting file processing..."

    if not excel_file.endswith('.xlsx'):
        return [], [{"Error": "Invalid Excel file type."}], 0

    if not os.path.isdir(video_folder):
        return [], [{"Error": "Video folder does not exist."}], 0

    start_time = time.time()
    logging.info("Starting file processing...")
    
    try:
        df = pd.read_excel(excel_file).fillna('')
    except Exception as e:
        logging.error(f"Error reading Excel file: {str(e)}")
        return [], [{"Error": str(e)}], 0

    not_found_files = []
    failed_copies = []
    total_files = len(df)

    for index, row in df.iterrows():
        folder_name = sanitize_filename(row['Category'])
        subfolder_name = sanitize_filename(row['Winner'])
        original_file_name = f"{str(row['Original File Name'])}.mp4"
        new_file_name = f"{sanitize_filename(str(row['Campaign Name']))}.mp4"

        subfolder_path = os.path.join(output_base_folder, folder_name, subfolder_name)
        os.makedirs(subfolder_path, exist_ok=True)

        original_file_path = os.path.join(video_folder, original_file_name)
        destination_file_path = os.path.join(subfolder_path, new_file_name)

        if os.path.isfile(original_file_path):
            try:
                shutil.copy2(original_file_path, destination_file_path)
                logging.info(f"Copied {original_file_name} to {destination_file_path}")
                current_process["message"] = f"Copied: '{original_file_name}' to '{new_file_name}'"
            except Exception as e:
                failed_copies.append({
                    'Original File': original_file_name,
                    'Destination': destination_file_path,
                    'Error': str(e)
                })
        else:
            not_found_files.append(original_file_name)
            current_process["message"] = f"File not found: '{original_file_name}'"

        # Update progress
        progress["percent"] = int((index + 1) / total_files * 100)
        logging.info(f"Processed {index + 1}/{total_files}")

    execution_time = time.time() - start_time
    return not_found_files, failed_copies, execution_time

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    excel_url = request.form['excel_url']
    video_folder_url = request.form['video_folder_url']
    final_folder = request.form['final_folder']
    
    # Start file processing
    not_found_files, failed_copies, execution_time = process_files(excel_url, video_folder_url, final_folder)

    return render_template('results.html', 
                           not_found_files=not_found_files, 
                           failed_copies=failed_copies, 
                           execution_time=execution_time)

@app.route('/progress')
def get_progress():
    return jsonify(progress)

@app.route('/current_process')
def get_current_process():
    return jsonify(current_process)

if __name__ == '__main__':
    host = os.getenv('FLASK_HOST', '127.0.0.1')  # Default to localhost if not set
    port = int(os.getenv('FLASK_PORT', 8717))    # Default to port 8717 if not set
    app.run(host=host, port=port, debug=True)
