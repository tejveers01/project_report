import streamlit as st
import pandas as pd
import requests
import json
import time
import re
WATSONX_API_URL = "https://us-south.ml.cloud.ibm.com/ml/v1/text/generation?version=2023-05-29"
MODEL_ID = "meta-llama/llama-3-3-70b-instruct"
PROJECT_ID = "4152f31e-6a49-40aa-9b62-0ecf629aae42"
API_KEY = "KS5iR_XHOYc4N_xoId6YcXFjZR2ikINRdAyc2w2o18Oo"

def GetAccesstoken():
    auth_url = "https://iam.cloud.ibm.com/identity/token"
    
    headers = { 
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json"
    }
    
    data = {
        "grant_type": "urn:ibm:params:oauth:grant-type:apikey",
        "apikey": API_KEY
    }
    response = requests.post(auth_url, headers=headers, data=data)
    
    if response.status_code != 200:
        st.write(f"Failed to get access token: {response.text}")
        return None
    else:
        token_info = response.json()
        return token_info['access_token']
    


def generatePrompt(datas, tower):
    body = {
        "input": f"""
    
      Read all data from this table carefully:
         
        {datas}.
        
        get pecentage of {tower} from this table and give as a json
        Convert the decimal value into a percentage string (e.g., 0.04 → "4%")
        need only json not explanation or any other
        json fromat
        {{
         'tower_name': 'tower_name',
         'percentage': 'percentage_value'
        }}

        Note: Return the result strictly as a JSON object—no code, no explanations, and dont add any notes, and steps please only the JSON that contains towername and values.

        """, 
        "parameters": {
            "decoding_method": "greedy",
            "max_new_tokens": 8100,
            "min_new_tokens": 0,
            "stop_sequences": [";"],
            "repetition_penalty": 1.05,
            "temperature": 0.5
        },
        "model_id": MODEL_ID,
        "project_id": PROJECT_ID
    }
    
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {GetAccesstoken()}"
    }
    
    if not headers["Authorization"]:
        return "Error: No valid access token."
    
    response = requests.post(WATSONX_API_URL, headers=headers, json=body)
    
    if response.status_code != 200:
        st.write(f"Failed to generate prompt: {response.text}")
        return "Error generating prompt"
    # st.write(json_datas)
    return response.json()['results'][0]['generated_text'].strip()


def get_percentages(exceldatas):
    eden = []

    towers = ["Tower 4", "Tower 5", "Tower 6", "Tower 7"]

    workbook = pd.ExcelFile(exceldatas)

    def _normalize(text):
        return re.sub(r"[^a-z0-9]+", "", str(text).lower())

    def _find_tower_sheet(tower_name):
        tower_num = re.search(r"\d+", tower_name)
        if not tower_num:
            return None
        key = f"tower{tower_num.group(0)}"

        # Exact normalized match first, then contains match.
        for sheet in workbook.sheet_names:
            if _normalize(sheet) == key:
                return sheet
        for sheet in workbook.sheet_names:
            if key in _normalize(sheet):
                return sheet
        return None

    def _parse_numeric_percent(value):
        if pd.isna(value):
            return None

        if isinstance(value, str):
            cleaned = value.strip().replace(",", "")
            if not cleaned:
                return None
            if cleaned.endswith("%"):
                cleaned = cleaned[:-1].strip()
                try:
                    return float(cleaned)
                except ValueError:
                    return None
            try:
                value = float(cleaned)
            except ValueError:
                return None

        try:
            num = float(value)
        except (TypeError, ValueError):
            return None

        if 0 <= num <= 1:
            return num * 100
        if 0 <= num <= 100:
            return num
        return None

    def _find_complete_column(df):
        candidates = []
        for column in df.columns:
            normalized = _normalize(column)
            if any(keyword in normalized for keyword in ("complete", "completion", "progress")):
                candidates.append(column)

        if not candidates:
            return None

        for column in candidates:
            normalized = _normalize(column)
            if "msp" in normalized and "complete" in normalized:
                return column
        for column in candidates:
            normalized = _normalize(column)
            if normalized in {"complete", "percentcomplete", "completepercent", "msppercentcomplete"}:
                return column
        return candidates[0]

    def _read_sheet_with_completion(sheet_name):
        for header_row in range(6):
            df = pd.read_excel(workbook, sheet_name=sheet_name, header=header_row)
            completion_col = _find_complete_column(df)
            if completion_col is not None:
                return df, completion_col
        return None, None

    def _extract_completion_values(df, completion_col):
        parsed_values = df[completion_col].map(_parse_numeric_percent).dropna()
        valid_values = parsed_values[(parsed_values >= 0) & (parsed_values <= 100)]
        return valid_values

    def _fallback_extract_completion(sheet_name):
        raw_df = pd.read_excel(workbook, sheet_name=sheet_name, header=None)
        search_rows = min(len(raw_df.index), 12)

        for row_idx in range(search_rows):
            row = raw_df.iloc[row_idx]
            for col_idx, value in enumerate(row):
                normalized = _normalize(value)
                if any(keyword in normalized for keyword in ("complete", "completion", "progress")):
                    values = raw_df.iloc[row_idx + 1 :, col_idx].map(_parse_numeric_percent).dropna()
                    values = values[(values >= 0) & (values <= 100)]
                    if not values.empty:
                        return values
        return pd.Series(dtype=float)

    for i in towers:
        try:
            sheet_name = _find_tower_sheet(i)
            if not sheet_name:
                raise ValueError(f"Sheet not found for {i}. Available sheets: {workbook.sheet_names}")

            datas, completion_col = _read_sheet_with_completion(sheet_name)
            if datas is not None and completion_col is not None:
                numeric_series = _extract_completion_values(datas, completion_col)
            else:
                numeric_series = _fallback_extract_completion(sheet_name)

            if numeric_series.empty:
                raise ValueError(f"No numeric completion values found in {sheet_name}")

            # Use the average completion across the tower sheet for a stable structure percentage.
            structure_pct = int(round(numeric_series.mean()))
            eden.append({
                "Project":"Eden",
                "Tower Name":i,
                "Structure":f"{structure_pct}%",
                "Finishing":"0%"
            })
        except Exception as e:
            eden.append({
                "Project":"Eden",
                "Tower Name":i,
                "Structure":"Error While Process",
                "Finishing":"0%"
            })
            st.write(f"Error while processing {i} tower {e}")

    return eden
    #  for i in towers:
    #      datas = pd.read_excel(exceldatas, sheet_name=i)
