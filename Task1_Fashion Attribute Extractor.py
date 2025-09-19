import pandas as pd
import requests
import time
import base64
import google.generativeai as genai
import json
import re
from openpyxl import load_workbook
from google.api_core.exceptions import ResourceExhausted, PermissionDenied
from rapidfuzz import fuzz, process

excel_file_path = "Best_Seller_Tags (2).xlsx" 
df = pd.read_excel(excel_file_path, sheet_name="Tagging",header=3)
print("Available columns:", df.columns.tolist())
# clean and standardize attributes
def clean_attribute(value, allowed_list,threshold=75):
    if not value:
        return None
    value = value.strip().title()
    # Fuzzy match to closest option in allowed list
    match, score, _ = process.extractOne(value, allowed_list, scorer=fuzz.token_sort_ratio)
    
    if score >= threshold:
        return match
    else:
        return None

# Ensure clean new columns
for col in ["Length", "Silhoutte", "Sleeve Type", "Neckline", "Waistline"]:
    if col not in df.columns:
        df[col] = None

# Loop over URLs

for i, url in enumerate(df["Image URL"]):   # check this is the correct URL column
    try:
        if pd.isna(url) or not str(url).startswith("http"):
            continue

        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            genai.configure(api_key="AIzaSyA4jdCjpwC3N-3Bhkp8bv9KXuStF94hAP8")
            model = genai.GenerativeModel("gemini-1.5-flash")

            # Allowed categories
            length_options = ["Mini", "Midi", "Maxi", "Floor Length"]
            Silhoutte_options = ["A-Line", "Column", "Mermaid", "Ball Gowns", "Two Piece Set", "Jumpsuit"]
            sleeve_options = ["Sleeveless", "Short Sleeve", "Long Sleeve", "Strapless", "Spaghetti Straps", "Cap Sleeve", "Puff Sleeves"]
            neckline_options = ["V Neck", "Sweetheart", "Straight", "Square Neck", "Scoop", "One Shoulder", "Off The Shoulder", "Cowl", "Halter", "High Neck"]
            waistline_options = ["Natural Waist", "Dropped Waist", "Empire Waist", "Basque Waist", "Illusion Waist"]

            prompt = f"""
            You are a fashion attribute extractor. 
            Analyze this dress image and return ONLY one value for each attribute 
            using the predefined options below (do not invent new terms):

            Length options: {length_options}
            Silhoutte options: {Silhoutte_options}
            Sleeve Type options: {sleeve_options}
            Neckline options: {neckline_options}
            Waistline options: {waistline_options}

            Extract dress attributes in strict JSON format with fields:
            length, Silhoutte, sleeve_type, neckline, waistline
            """

            # Convert image → base64
            image_b64 = base64.b64encode(response.content).decode("utf-8")

            # Call Gemini
            try:
                response = model.generate_content([{"role": "user", "parts": [
                    {"text": prompt},
                    {"inline_data": {"mime_type": "image/jpeg", "data": image_b64}}
                ]}])
            except (ResourceExhausted, PermissionDenied) as quota_error:
                print(f"Quota exceeded or access denied: {quota_error}")
                print("Stopping execution due to Gemini quota limits.")
                break  # Stop processing if quota is exceeded

            except Exception as api_error:
                print(f"Gemini API error for {url}: {api_error}")
                continue

            response_text = response.text.strip()
            cleaned = re.sub(r"```(json)?", "", response_text).strip()
            print("Raw Gemini Output:", cleaned)

            try:
                ai_data = json.loads(cleaned)
            except:
                print(f"Could not parse response for {url}: {response.text}")
                continue

            # Save cleaned values to SAME row
            
            df.at[i,"Length"] = clean_attribute(ai_data.get("length"), length_options)
            df.at[i,"Silhoutte"] = clean_attribute(ai_data.get("Silhoutte"), Silhoutte_options)
            df.at[i,"Sleeve Type"] = clean_attribute(ai_data.get("sleeve_type"), sleeve_options)
            df.at[i,"Neckline"] = clean_attribute(ai_data.get("neckline"), neckline_options)
            df.at[i,"Waistline"] = clean_attribute(ai_data.get("waistline"), waistline_options)

            #print(f"Processed row {i}: {df.loc[i, ['Length','Silhoutte','Sleeve Type','Neckline','Waistline']].to_dict()}")
            with pd.ExcelWriter(excel_file_path, engine="openpyxl",mode="a", if_sheet_exists="overlay") as writer:
                df.to_excel(writer, sheet_name="Tagging", index=False,startrow=3)

                print(f" Structured output saved correctly into {excel_file_path}")
            time.sleep(2)

        else:
            print(f"Failed to fetch {url} - Status: {response.status_code}")
            time.sleep(3)

    except Exception as e:
        print(f"Error with {url}: {e}")


