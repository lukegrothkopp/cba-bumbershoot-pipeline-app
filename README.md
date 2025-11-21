# Bumbershoot & Cannonball – Revenue Pipeline Dashboard 🎟️

This repository contains a Streamlit app that visualizes the **revenue pipeline** for:

- **Sponsorships (Corporate Partnerships)** – e.g., Verizon, Expedia, BECU  
- **Public Investment** – e.g., Visit Seattle, Chamber of Commerce, Port of Seattle  

The app is built around a single Excel workbook that contains both **prospect** and **contact activity** data. It provides a clear view of pipeline stages, expected revenue, activity volume, and top deals across both categories.

---

## Features

The dashboard currently includes:

1. **Pipeline by Stage**  
   Columns for each stage:
   - Lead  
   - Under 50%  
   - 50–75%  
   - Over 75%  
   - Contracted  

   Each stage lists:
   - Prospect (Account Name)  
   - Owner  
   - Projected Annual Revenue  
   - Contracted Annual Revenue  
   - Expected Value  

   You can view:
   - All deals combined  
   - Only Sponsorship deals  
   - Only Public Investment deals  

2. **Top Deals by Expected Value**  
   Shows the **Top 3 Sponsorship** deals and **Top 3 Public Investment** deals ranked by **Expected Value ($)**.

3. **Activity Heat Map (Last 3 Weeks)**  
   A heat map of **contact activity** broken out by:
   - Partner Type: Sponsorship vs Public Investment  
   - Time Buckets: This Week, Last Week, Two Weeks Ago  

   This helps you quickly see if outreach is skewed toward one category or if certain weeks have been “quiet.”

4. **Pipeline Value by Stage**  
   Aggregated **Expected Value ($)** at each stage, broken out by partner type, plus:
   - A summary table of total expected value and deal counts by stage  
   - A bar chart showing the breakdown of **Sponsorship** vs **Public Investment**

   > Note: Deals marked as **Dead** are excluded from pipeline roll-ups.

5. **Snapshot KPIs**  
   At the top of the app, you’ll see:
   - Total Expected Value  
   - Total Projected Annual Revenue  
   - Total Contracted Annual Revenue  
   - Number of Active Prospects  

6. **Recent Activity Feed**  
   A table of the **last 10 contact events** with:
   - Prospect name  
   - Category (Sponsorship/Public)  
   - Contact date  
   - Contact type  
   - Contact owner  
   - Outcome & next step  

7. **Data Dictionary (Optional Helper)**  
   The `Data_Dictionary` sheet is surfaced in an expandable section so anyone using the app can quickly remind themselves what each field means.

---

## Data Model

The app assumes an Excel workbook with these sheets:

- `Sponsorships`  
- `Public Investment`  
- `Contact Detail`  
- `Data_Dictionary`  

A sample file (the one you shared) is referenced by default:

- `data/Grothko-BumbershootCBA-Prospecting-Testing-New-Version.xlsx`  
- Example path in this environment:  
  `[Grothko-BumbershootCBA-Prospecting-Testing-New-Version.xlsx](sandbox:/mnt/data/Grothko-BumbershootCBA-Prospecting-Testing-New-Version.xlsx)`

### Prospects (Sponsorships & Public Investment sheets)

Key columns the app looks for:

- `Prospect ID`  
- `Prospect (Account Name)`  
- `Owner`  
- `Projected Annual Revenue ($)`  
- `Contracted Annual Revenue ($)`  
- `Probability (%)`  – stored as fractions (e.g. 0.25, 0.5, 0.75) and converted to 25%, 50%, 75% in the app  
- `Expected Value ($)`  
- Stage flag columns:
  - `Lead`  
  - `Prospect`  
  - `Under 50%`  
  - `50-75%`  
  - `Over 75%`  
  - `Contracted`  
  - `Dead`  

The app combines the two prospect sheets into a single dataframe and tags each row using:

- `Partner Type` = `"Sponsorship"` or `"Public Investment"`

A derived `Stage Bucket` is calculated for each prospect, using:

1. `Dead` / `Contracted` flags and contracted revenue  
2. `Probability (%)` (converted to 0–100 scale if needed)  
3. Stage flag columns as a fallback  

### Contact Detail sheet

Key columns:

- `Prospect (Account Name)`  
- `Prospect (Sponsorship/Public)` – free-text used as a backup for partner type  
- `Contact Date`  
- `Follow-up Date` (optional but supported)  
- `Contact Type (email/phone/zoom/in-person)`  
- `Contact Owner`  
- `Contact Name`  
- `Outcome (left VM/spoke/meeting set/sent deck/etc.)`  
- `Next Step`  

The app uses contact dates to bucket touches into:
- This Week  
- Last Week  
- Two Weeks Ago  

and counts them by **Partner Type** to build the heat map.
