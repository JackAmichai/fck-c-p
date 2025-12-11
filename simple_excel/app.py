import os
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

# Helper: Convert Excel Column Letter (A, B, AA) to Index (0, 1, 26)
def col2num(col_str):
    num = 0
    for c in col_str:
        if c in "0123456789": return int(c) - 1 # Handle if user types numbers
        num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num - 1

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_files():
    try:
        # 1. Get Files
        if 'source_file' not in request.files or 'target_file' not in request.files:
            return jsonify({"error": "Missing files"}), 400
        
        source_file = request.files['source_file']
        target_file = request.files['target_file']
        
        # 2. Get Rules (The form data comes as a JSON string or list)
        # We expect a list of operations. 
        # Example form fields: rule[0][src_col], rule[0][match_mode]...
        # For simplicity in this demo, we parse the JSON payload sent by JS.
        import json
        rules = json.loads(request.form.get('rules_json'))

        # 3. Load Dataframes
        # header=None means we treat it as a grid of data, not assuming row 1 is headers immediately
        df_source = pd.read_excel(source_file, header=None) 
        df_target = pd.read_excel(target_file, header=None)

        # 4. Process Rules
        for rule in rules:
            # Parse inputs
            src_col_idx = col2num(rule['src_col'])
            tgt_col_idx = col2num(rule['tgt_col'])
            
            # Python uses 0-index, User uses 1-index for rows
            src_start_row = int(rule['src_row_start']) - 1
            tgt_start_row = int(rule['tgt_row_start']) - 1
            
            src_end_row = rule.get('src_row_end')
            if src_end_row and src_end_row.strip() != "":
                src_end_row = int(src_end_row)
            else:
                src_end_row = len(df_source) # Go to end if empty

            # Extract Data
            data_to_copy = df_source.iloc[src_start_row:src_end_row, src_col_idx]

            # --- LOGIC BRANCH ---
            if rule.get('is_advanced'):
                if not (rule.get('match_src_col') and rule.get('match_tgt_col')):
                    return jsonify({"error": "Advanced rule selected but match columns missing for one of the rules."}), 400
                
                # COMPLEX MODE (VLOOKUP REPLACEMENT)
                match_src_idx = col2num(rule['match_src_col'])
                match_tgt_idx = col2num(rule['match_tgt_col'])
                
                # Create a temporary mapping dictionary: {Key: Value}
                # We zip the Key Column and the Value Column from source
                source_map = dict(zip(
                    df_source.iloc[src_start_row:src_end_row, match_src_idx],
                    df_source.iloc[src_start_row:src_end_row, src_col_idx]
                ))

                # Apply map to target
                # We iterate target rows starting from tgt_start_row
                for i in range(tgt_start_row, len(df_target)):
                    key_val = df_target.iloc[i, match_tgt_idx]
                    if key_val in source_map:
                        df_target.iloc[i, tgt_col_idx] = source_map[key_val]

            else:
                # SIMPLE MODE (Direct Copy Paste)
                # We paste directly into target coordinates
                # Ensure we don't overflow target bounds
                rows_to_paste = list(data_to_copy)
                for i, val in enumerate(rows_to_paste):
                    current_tgt_row = tgt_start_row + i
                    # Expand dataframe if target is shorter than source paste
                    if current_tgt_row >= len(df_target):
                         # In a real app, we'd append rows. For MVP, we stop or fill.
                         break 
                    df_target.iloc[current_tgt_row, tgt_col_idx] = val

        # 5. Save and Return
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_target.to_excel(writer, index=False, header=False)
        output.seek(0)
        
        response = send_file(output, download_name="processed_excel.xlsx", as_attachment=True)
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        return response

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
