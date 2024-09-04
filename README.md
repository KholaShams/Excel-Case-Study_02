# Excel-Case-Study_02
---

## **Final Report: Analyzing a Ramadan Digital Marketing Campaign in Excel**
---
![image](https://github.com/user-attachments/assets/7e835369-d9c7-4fb8-9850-cf92dcff7c28)

--- 

### Table of Contents

1. [Introduction](#introduction)
2. [Data Cleaning, Preparation, and Analysis Steps](#data-cleaning-preparation-and-analysis-steps)
   - [TikTok Data Cleaning Process](#tiktok-data-cleaning-process)
   - [Meta Data Cleaning Process](#meta-data-cleaning-process)
   - [Snapchat Data Cleaning Process](#snapchat-data-cleaning-process)
3. [Pivot Tables and Key Insights](#pivot-tables-and-key-insights)
   - [Platform Analysis](#platform-analysis)
   - [Comprehensive Performance by Platform](#comprehensive-performance-by-platform)
   - [Campaign-Level Performance](#campaign-level-performance)
   - [Video Completion Rate (VTR) by Market and Platform](#video-completion-rate-vtr-by-market-and-platform)
   - [Engagement, CPC, CTR, and VTR Overview](#engagement-cpc-ctr-and-vtr-overview)
   - [Audience Segmentation Analysis](#audience-segmentation-analysis)
4. [Visualizations](#Visualizations)
5. [Macros Implementation](#macros-implementation)
   - [Data Refresh Macro](#data-refresh-macro)
   - [Performance Filter Macro](#performance-filter-macro)
   - [Conditional Formatting Macro](#conditional-formatting-macro)
   - [Formatting Summary Sheet Macro](#formatting-summary-sheet-macro)
6. [Conclusion and Recommendations](#conclusion-and-recommendations)


---

#### **Introduction**

This report delves into the intricate analysis of a Ramadan digital marketing campaign that was conducted across three major platforms: TikTok, Meta, and Snapchat. The primary goal of the analysis was to assess the performance of the campaign across these platforms, uncover key insights, and offer data-driven recommendations for future campaigns. This process involved extensive data cleaning, preparation, creating complex Pivot Tables, and implementing advanced Macros to automate and streamline the analysis.

---

#### **Data Cleaning, Preparation, and Analysis Steps**

The data cleaning process was cautiously structured to ensure the data was accurate, consistent, and ready for deep analysis. Below is a detailed account of the cleaning procedures for each dataset.

---

### **1. TikTok Data Cleaning Process:**

**Initial Data Structure:**
- **Columns:**`Sourcesheet`, `Market`, `Channel`, `Compaign Name`, `Compaign attributes`, `compaign id`, `Audience`, `Duration`, `Language` , `Format`, `Creative Variation`, `Amount Spent`, `Clicks`, `CPC`, `Paid Reach`, `Total Impressions`, `CTR`, `CPM`,`2 Second Video Views`, `Video Completions`, `VTR (2 Sec)`, `VTR (Complete)`, `Total Engagement`, `Engagement Rate`

- **Issues Identified:**
  - Redundant columns with overlapping information.
  - Inconsistent data formats, especially in text fields.
  - Potential for hidden characters or anomalies in the `Campaign Name` column due to concatenated information.

**Step-by-Step Cleaning:**

1. **Campaign Name Decomposition:**
   - **Objective:** The `Campaign Name` field contained multiple pieces of information separated by underscores (_) and tilde (~). This needed to be split to isolate the different attributes.
   - **Action:**
     - Used the **Text-to-Columns** feature in Excel.
     - Specified underscore (_) and tilde (~) as delimiters to break down the `Campaign Name` into distinct columns: `Channel Name`, `Channel`, `Objective`, `Audience`, and `Market`.
     - Post-split, renamed the new columns accordingly for clarity.
   - **Outcome:** The `Campaign Name` column was effectively decomposed into four distinct and meaningful attributes.

2. **Redundant Column Removal:**
   - **Objective:** To eliminate unnecessary columns that either provided duplicate information or were irrelevant for analysis.
   - **Action:**
     - Removed the `Ad Group Name` column as the key attributes had been extracted elsewhere. The column was deemed unnecessary for the final analysis.
     - Deleted the `Market` column since the market data was already embedded in the decomposed `Market` attribute from the `Campaign Name`.
     - Deleted the Ad Name after extracting valuable data from it by using `Text to Column` feature of Excel.
   - **Outcome:** The dataset was streamlined, reducing clutter and focusing on essential data points.

3. **Text Cleaning:**
   - **Objective:** To ensure consistency and remove any potential hidden characters or irregularities in text fields.
   - **Action:**
     - Applied the **TRIM** function across text columns to eliminate leading and trailing spaces.
     - Used **CLEAN** to remove any non-printable characters that might have been introduced during data import.
   - **Outcome:** All text fields were standardized, ensuring they were free from extraneous spaces and non-printable characters.

4. **Metric Calculation:**
   - **Objective:** To clarify column meanings and introduce new metrics for a comprehensive analysis.
   - **Action:**
     - Calculated `CTR` (Click-Through Rate) as `(Clicks / Impressions) * 100` and added it as a new column.
     - Calculated `CPC` (Cost Per Click) as `Spend / Clicks` and added it as a new column.
   - **Outcome:** The dataset was enriched with new metrics, and columns were renamed for clarity.

5. **Format Standardization:**
   - **Objective:** To standardize numeric and date formats for consistency across the dataset.
   - **Action:**
     - Reformatted all date columns to a consistent `DD/MM/YYYY` format.
     - Applied number formatting to `Impressions`, `Clicks`, `Spend`, and other metric columns to ensure uniformity (e.g., comma separators for thousands).
   - **Outcome:** The dataset’s numeric and date fields were consistent, facilitating smoother analysis.

6. **Validation:**
   - **Objective:** To verify the integrity of the cleaned data before analysis.
   - **Action:**
     - Cross-checked key metrics (e.g., `Clicks` vs. `CTR` and `CPC`) to ensure calculated values aligned with raw data.
     - Used conditional formatting to identify any anomalies, such as unusually high or low values that might indicate errors.
   - **Outcome:** The dataset was validated and ready for in-depth analysis.

---

### **2. Meta Data Cleaning Process:**

**Initial Data Structure:**
- **Columns:** `Source sheet`, `Market`, `Compaign attributes`, `Compaign Name`, `Compaign ID`, `Duration`, `Audience`, `Language`, `Format`, `Creative variations`, `Reach`, `Impressions`, `Amount spent (USD)`, `Link clicks`, `CPC`, `3-second video plays`, `Video plays at 100%`, `CTR(all)`, `CTR Evaluation`, `all ctr evaluation`, `VTR`, `Age Group`, `Post engagement`, `total engagement`, `Engagement Rate 2`

- **Issues Identified:**
  - The numeric attributes weren’t properly formatted. Attributes having Percentage data were also not properly Formatted. And some text based data wasn’t formatted either. 

**Step-by-Step Cleaning:**

1. **Text Normalization:**
   - **Objective:** To ensure consistency in text data.
   - **Action:**
     - Converted all text data to uppercase to avoid discrepancies caused by case sensitivity.
     - Removed special characters and extra spaces using a combination of **SUBSTITUTE** and **TRIM** functions.
   - **Outcome:** The text data was standardized, ensuring consistency across all records.

2. **Data Type Validation:**
   - **Objective:** To confirm that each column contained data of the appropriate type.
   - **Action:**
     - Checked that numeric columns (e.g., `Impressions`, `Clicks`, `Spend`) were properly formatted as numbers.
     - Verified that the date columns followed the `DD/MM/YYYY` format.
   - **Outcome:** All data types were correctly assigned, reducing the risk of errors during analysis.

3. **Column Consistency Check:**
   - **Objective:** To ensure that all relevant columns had consistent and complete data.
   - **Action:**
     - Used data validation techniques to flag any missing or anomalous values.
     - Applied **COUNTIF** and **ISBLANK** functions to identify empty cells or inconsistencies.
   - **Outcome:** All columns were confirmed to have consistent and complete data, ready for further analysis.

4. **Metric Calculation:**
   - **Objective:** To clarify column meanings and introduce new metrics for a comprehensive analysis.
   - **Action:**
     - Calculated `CTR` (Click-Through Rate) as `(Clicks / Impressions) * 100` and added it as a new column.
     - Calculated `CPC` (Cost Per Click) as `Spend / Clicks` and added it as a new column.
   - **Outcome:** The dataset was enriched with new metrics, and columns were renamed for clarity.

---

### **3. Snapchat Data Cleaning Process:**

**Initial Data Structure:**
- **Columns:** `Sourcesheet`, `Market`, `Channel`, `Compaign ID`, `Compaign Name`, `Campaign Strategy`, `Audience`, `Duration`, `Language`, `Format`, `Creative Variation`, `Amount Spent`, `Engagement Rate`, `Engagement`, `Clicks`, `CPC`, `CTR`, `Clicks Rate`, `Paid Reach`, `Total Impressions`, `Paid Frequency`, `Paid eCPM`, `2 Second Video Views`, `Video Completions`, `VTR%`
- **Issues Identified:**
  - Ambiguity in column names (`Swipe Ups` vs. `Clicks`), leading to potential confusion.
  - Missing columns for key metrics like CTR and CPC.

**Step-by-Step Cleaning:**

1. **Column Renaming and Metric Calculation:**
   - **Objective:** To clarify column meanings and introduce new metrics for a comprehensive analysis.
   - **Action:**
     - Renamed `Swipe Ups` to `Clicks` and `Swipe Up Rate` to `Click Rate` to align with standard industry terminology.
     - Calculated `CTR` (Click-Through Rate) as `(Clicks / Impressions) * 100` and added it as a new column.
     - Calculated `CPC` (Cost Per Click) as `Spend / Clicks` and added it as a new column.
   - **Outcome:** The dataset was enriched with new metrics, and columns were renamed for clarity.

2. **Data Formatting:**
   - **Objective:** To ensure consistent formatting of numeric values.
   - **Action:**
     - Applied number formatting to key metric columns, ensuring values like `Clicks`, `Impressions`, `Spend`, and calculated metrics were displayed with appropriate decimal places and thousand separators.
   - **Outcome:** The numeric data was consistently formatted, enhancing readability and accuracy.

3. **Anomaly Detection:**
   - **Objective:** To identify and correct any potential outliers or errors in the data.
   - **Action:**
     - Used conditional formatting to highlight any values that were significantly higher or lower than expected, based on historical trends.
     - Investigated and corrected identified anomalies, ensuring the integrity of the data.
   - **Outcome:** Anomalies were identified and addressed, ensuring the dataset was robust and reliable.

4. **Column Validation and Consistency:**
   - **Objective:** To ensure that all calculated metrics were accurate and consistent across the dataset.
   - **Action:**
     - Cross-referenced calculated columns (`CTR`, `CPC`) with raw data to verify accuracy.
     - Ensured that all columns had consistent data types and no missing values.
   - **Outcome:** All calculated metrics were validated, and the dataset was consistent across all columns.

---

### Pivot Tables and Key Insights

#### Platform Analysis
The analysis begins by examining the performance across different platforms: MetaData cleaned, Snapchat cleaned data, and TikTok cleaned Data. The following metrics were evaluated:

- **Cost Per Click (CPC)**:  
  - MetaData cleaned: **120.78**
  - Snapchat cleaned data: **25.03**
  - TikTok cleaned Data: **35.23**

This indicates that the CPC for MetaData cleaned is significantly higher compared to the other two platforms, suggesting a higher cost efficiency for Snapchat and TikTok in terms of clicks.

![Platform Analysis](https://github.com/user-attachments/assets/3eb6ed68-8a40-49d8-9473-c7f00200621d)

#### Comprehensive Performance by Platform
A deeper analysis was conducted by aggregating the key metrics (Clicks, CPC, and Amount Spent) across the three platforms:

- **MetaData cleaned**:  
  - Clicks: **419,081**
  - CPC: **120.78**
  - Amount Spent: **53,113.53 USD**

- **Snapchat cleaned data**:  
  - Clicks: **54,794**
  - CPC: **25.03**
  - Amount Spent: **23,049.20 USD**

- **TikTok cleaned Data**:  
  - Clicks: **79,388**
  - CPC: **35.23**
  - Amount Spent: **52,192.57 USD**

MetaData cleaned had the highest number of clicks and total spending, yet its CPC remains considerably higher. Snapchat, although having the lowest number of clicks, presents a cost-effective CPC.

![Comprehensive Performance by Platform](https://github.com/user-attachments/assets/c6ce3bd4-07b7-4047-85e9-12824321f5e3)

#### Campaign-Level Performance
The performance at the campaign level was scrutinized by evaluating Total Impressions, Clicks, and Click-Through Rate (CTR):

- **Top Campaigns by Impressions**:
  - **CN~MCDRamadan_CH~FBIG_MK~RIY_TG**:  
    - Impressions: **16,873,762**
    - Clicks: **31,766**
    - CTR: **9.41%**
  - **CN~MCDRamadan_CH~Tiktok_MK~JED_TG**:  
    - Impressions: **15,307,011**
    - Clicks: **17,744**
    - CTR: **8.82%**

- **Top Campaign by CTR**:
  - **CN~MCDRamadan_CH~Tiktok_MK~AE_TG**:  
    - CTR: **24.85%**  
    This campaign in the AE market has the highest CTR, indicating effective engagement with the target audience.

![Campaign Level Performance](https://github.com/user-attachments/assets/7faf0d06-4525-4851-adb2-e455afd64f07)

#### Video Completion Rate (VTR) by Market and Platform
The analysis of Video Completion Rate (VTR) across different markets and platforms yielded the following:

- **MetaData cleaned**:
  - Highest VTR in **AE (12.79%)** and **JED (14.13%)** markets.
- **Snapchat cleaned data**:
  - High VTR in **BH (2.94%)** and **RIY (2.18%)** markets.
- **TikTok cleaned Data**:
  - Noticeable VTR in **AE (0.06%)** and **KWT (0.06%)** markets.

MetaData cleaned demonstrates superior VTR across various markets, with Snapchat performing well in certain regions like BH.

![Video Completion Rate (VTR) by Market and Platform](https://github.com/user-attachments/assets/cbebc256-a07b-48ed-abbd-ecff5ed79cde)

#### Engagement, CPC, CTR, and VTR Overview
A consolidated analysis across MetaData cleaned, Snapchat cleaned data, and TikTok cleaned Data was conducted to evaluate Engagement Rate, CPC, CTR, and VTR:

- **MetaData cleaned**:
  - Engagement Rate: **67.97%**
  - CPC: **120.78**
  - CTR: **1.56%**
  - VTR: **61.25%**

- **Snapchat cleaned data**:
  - Engagement Rate: **0.19%**
  - CPC: **25.03**
  - CTR: **0.19%**
  - VTR: **8.34%**

- **TikTok cleaned Data**:
  - Engagement Rate: **0.31%**
  - CPC: **35.23**
  - CTR: **0.07%**
  - VTR: **0.24%**

MetaData cleaned stands out with the highest engagement rate and VTR, although with a higher CPC. Snapchat's CPC remains low but shows relatively lower engagement and VTR.

![Engagement, CPC, CTR, and VTR Overview](https://github.com/user-attachments/assets/89816923-c29a-436e-8030-38984cff9b2e)


#### Audience Segmentation Analysis
The final part of the analysis focuses on audience segmentation, specifically comparing Boomers and Millennials in terms of Link Clicks, Impressions, Amount Spent, and Conversion Rate:

- **Boomers**:
  - Link Clicks: **87,622**
  - Impressions: **25,098,281**
  - Amount Spent: **10,950.78 USD**
  - Conversion Rate: **0.35%**

- **Millennials**:
  - Link Clicks: **331,459**
  - Impressions: **81,852,438**
  - Amount Spent: **42,162.75 USD**
  - Conversion Rate: **0.40%**

Millennials demonstrate a higher conversion rate compared to Boomers, with a significantly higher volume of impressions and clicks, indicating more effective engagement with this demographic.

![Audience Segmentation Analysis](https://github.com/user-attachments/assets/0203b577-bd44-4214-9509-52c065620045)


 ### **Visualizations:**
Following is the final Dashboard Created that includes all the visualizations created through out to better visualize, understand, and find interesting insights from the data:

![Dashboard](https://github.com/user-attachments/assets/52ff439b-543b-4226-bd7e-6bb2c8e442d5)

---

### **Macros Implementation:**

To streamline the analysis and ensure the process could be easily replicated, I recorded and implemented several macros:

1. **Data Refresh Macro:**
   - **Function:** Automatically refresh all Pivot Tables with the latest data.
   - **Steps:**
     - Created a macro that refreshes all Pivot Tables across the workbook with a single click.
   - **Outcome:** Streamlined the data update process, ensuring the analysis was always based on the most current data. 

![Data Refresh Macro](https://github.com/user-attachments/assets/dc5f97d8-41da-4d76-ad67-faf29b6bc1a7)


2. **Performance Filter Macro:**
   - **Function:** Filter the summary sheets to highlight campaigns with a `Good` performance status as green.
   - **Steps:**
     - Recorded a macro to apply filters across the summary sheets, highlighting only the top-performing campaigns.
   - **Outcome:** Enabled quick access to high-performing campaigns, aiding in fast decision-making.

![Performance Filter Macro](https://github.com/user-attachments/assets/009e88b5-c1dc-4e65-873c-b182bcfa7e67)


3. **Conditional Formatting Macro:**
   - **Function:** Apply conditional formatting to highlight exceptional performance metrics.
   - **Steps:**
     - Created a macro that applies conditional formatting to engagement rate columns, coloring cells based on performance thresholds.
   - **Outcome:** Automated the process of visualizing high and low performers, making insights more accessible.
     
![Performance Filter Macro](https://github.com/user-attachments/assets/d74a898b-2c40-4c1b-8979-75c0009a3791)


4. **Formatting Summary Sheet Macro:**
- **Function:** Enhance the "Overall Performance Summary" sheet with consistent formatting.
- **Steps:**
  - **Header Styles:** Applies bold, white text on a dark blue background, centers text, and sets the font size to 12.
  - **Column Widths:** Adjusts widths for columns A through G to fit data.
  - **Number Formatting:** Formats columns for Engagement Rate, CPC, CTR, VTR, and Performance with appropriate number formats (percentage or currency).
  - **Autofit Rows:** Adjusts row heights to fit content.
  - **Borders:** Adds thin, continuous borders around the data range and inside cells for improved readability.
- **Outcome:** Creates a visually appealing and organized summary sheet, enhancing data presentation and readability.

![Formatting Summary Sheet Macro1](https://github.com/user-attachments/assets/9daadaeb-5154-4ec4-939a-b4658114b275)
![Formatting Summary Sheet Macro2](https://github.com/user-attachments/assets/79042187-19a7-48e2-9c47-e695d231fe64)


### **Conclusion and Recommendations:**

The detailed analysis of the Ramadan digital marketing campaign data across TikTok, Meta, and Snapchat revealed several key insights:

- **Platform Performance:**
  - **TikTok** showed the highest engagement rates with 0.3095, indicating strong user interaction. It also had notable conversion rates and visual representation metrics, making it a strong platform for engagement-focused campaigns.
  - **Meta** demonstrated cost efficiency with a CPC of $120.78 and substantial click volumes, suggesting it provides value for money in terms of cost-per-click.
  - **Snapchat** had the lowest CPC but also relatively low engagement and conversion rates, indicating it may not be as effective for high-impact campaigns compared to TikTok and Meta.

- **Market-Specific Insights:**
  - **AE (United Arab Emirates)** performed exceptionally well across TikTok and Meta, showing high engagement and significant total impressions. It should be a focal point for future campaigns.
  - **JED (Jeddah)** and **KW (Kuwait)** also exhibited strong performance metrics, particularly in engagement and impressions on TikTok, warranting increased focus.

- **Demographic Trends:**
  - **Millennials** (ages 25-34) had higher click volumes and conversion rates compared to **Boomers**, suggesting that targeting this demographic could yield better results.

**Recommendations:**
- **Focus on High-Performing Markets:** Prioritize markets like AE and JED where high engagement and significant impressions were observed. Allocate more resources to these areas for future campaigns.
- **Target Engaged Demographics:** Emphasize campaigns aimed at Millennials who showed higher engagement and conversion rates. Explore strategies to better reach Boomers and other age groups.
- **Utilize Platform Strengths:** Leverage TikTok for engagement-driven content and Meta for cost-efficient ad placements. Consider Snapchat for supplementary, targeted efforts if budget allows.

By applying these recommendations, future campaigns can be optimized for better performance and ROI.

---
