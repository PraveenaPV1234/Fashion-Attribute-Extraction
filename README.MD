# Dress Attribute Extractor with Gemini AI

This Python script extracts fashion attributes from dress images in an Excel file using **Google Gemini AI** and saves the structured results back into the Excel sheet.  

It analyzes dress images and identifies attributes like length, silhouette, sleeve type, neckline, and waistline based on predefined options.


## Features

- Reads image URLs from an Excel sheet (`Best_Seller_Tags (2).xlsx`).
- Uses **Google Gemini AI** to analyze dress images.
- Extracts key attributes:
  - **Length** – Mini, Midi, Maxi, Floor Length
  - **Silhouette** – A-Line, Column, Mermaid, Ball Gown, Two Piece Set, Jumpsuit
  - **Sleeve Type** – Sleeveless, Short Sleeve, Long Sleeve, Strapless, Spaghetti Straps, Cap Sleeve, Puff Sleeves
  - **Neckline** – V Neck, Sweetheart, Straight, Square Neck, Scoop, One Shoulder, Off The Shoulder, Cowl, Halter, High Neck
  - **Waistline** – Natural Waist, Dropped Waist, Empire Waist, Basque Waist, Illusion Waist
- Cleans and standardizes attributes using **fuzzy matching**.
- Updates the original Excel sheet with structured, validated data.
- Handles API errors and Gemini quota limits gracefully.


## Requirements

- Python 3.8+
- Python libraries:
  - `pandas`
  - `requests`
  - `openpyxl`
  - `google-generativeai`
  - `rapidfuzz`
  
Install dependencies using:

pip install pandas requests openpyxl google-generativeai rapidfuzz
