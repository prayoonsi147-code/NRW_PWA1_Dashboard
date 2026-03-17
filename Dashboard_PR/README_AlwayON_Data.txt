================================================================================
AlwayON DATA EXTRACTION - Complete Package
================================================================================

PROJECT: Dashboard_PR - AlwayON Data for กปภ.ข.1 (District 1)
EXTRACTION DATE: 2026-03-17
SOURCE: 11.ข้อมูล PWA Always-on สะสม เดือนมกราคม 2569 โดย กลส..xls

================================================================================
OUTPUT FILES GENERATED
================================================================================

1. AlwayON_Data_District1.js (13 KB)
   - JavaScript-ready format
   - Complete data structure for all 22 branches
   - Includes helper functions for data retrieval
   - Month codes: ต.ค.68, พ.ย.68, ธ.ค.68, ม.ค.69
   - Always-On values available ONLY for January 2569 (ม.ค.69)
   - Can be imported as module (Node.js compatible)
   - Usage: Import and call getAlwaysOnValue() or getJanuaryAlwaysOnRanking()

2. AlwayON_Data_District1.csv (2.2 KB)
   - CSV format with header row
   - Sorted by Always-On ranking (highest to lowest for January)
   - Includes branch names (Thai & English), Always-On %, rank, water bill, payment
   - Easy import to Excel, Google Sheets, databases
   - All 22 branches listed with complete January 2569 data

3. AlwayON_Data_Summary.txt (This file)
   - Comprehensive summary of extraction
   - Branch list with rankings
   - Data statistics and observations
   - Notes on data availability
   - Key findings

4. Excel_Structure_Reference.txt
   - Detailed explanation of Excel file structure
   - Column descriptions for each sheet
   - Sample data row showing all columns
   - Notes on data conversion
   - Information on unused sheets

5. README_AlwayON_Data.txt (This file)
   - Quick reference guide
   - File descriptions
   - Usage instructions
   - Contact information

================================================================================
QUICK FACTS
================================================================================

BRANCHES COVERED: 22 (all of กปภ.ข.1 - Eastern Water Authority District 1)
MONTHS INCLUDED: October 2568 - January 2569
ALWAYS-ON DATA AVAILABLE: January 2569 ONLY

TOP 5 PERFORMERS (January 2569):
  1. สาขาบ้านฉาง (Ban Chang) - 88.30%
  2. สาขาพัทยา (Pattaya) - 87.65%
  3. สาขาชลบุรี (Chonburi) - 86.85%
  4. สาขาระยอง (Rayong) - 86.41%
  5. สาขาบ้านบึง (Ban Bueng) - 86.24%

BOTTOM 3 PERFORMERS (January 2569):
  20. สาขาคลองใหญ่ (Khlong Yai) - 67.22%
  21. สาขาวัฒนานคร (Watthananakorn) - 63.13%
  22. สาขาอรัญประเทศ (Aranyaprathet) - 53.72%

AVERAGE: 78.78%
MEDIAN: 82.62%

================================================================================
HOW TO USE EACH FILE
================================================================================

FOR JAVASCRIPT/WEB DEVELOPMENT:
  1. Use AlwayON_Data_District1.js
  2. Example: 
     const value = alwayONData["ม.ค.69"].branches["สาขาชลบุรี"].always_on;
     console.log(value * 100); // Outputs: 86.85

FOR SPREADSHEETS/DATABASES:
  1. Use AlwayON_Data_District1.csv
  2. Import into Excel, Google Sheets, or database
  3. Columns: Branch Name, Always-On %, Rank, Water Bill, Payment

FOR REPORTING/REFERENCE:
  1. Use AlwayON_Data_Summary.txt for rankings and statistics
  2. Use Excel_Structure_Reference.txt to understand data source structure
  3. Use this README for quick overview

================================================================================
DATA NOTES
================================================================================

ALWAYS-ON VALUES:
- Format: Decimal (0-1 scale)
- Conversion: Multiply by 100 for percentage
- Example: 0.8685 = 86.85%
- Availability: January 2569 only

ACCOMPANYING DATA (All months):
- Water Bill: Total water bill amount (THB) for the month
- Payment Total: Total payment received for the month
- District: Always "1" for เขต 1 branches

MISSING DATA:
- October, November, December: No Always-On values
- January: Complete Always-On data available

================================================================================
BRANCH LIST (22 Total)
================================================================================

ตั้งจากยังไม่มีลำดับตัวเลข - All 22 branches in English:

1. Chonburi
2. Ban Bueng
3. Phanat Nikhom
4. Si Racha
5. Laem Chabang
6. Pattaya
7. Chachoengsao
8. Bang Pakong
9. Bang Khla
10. Phanom Sarakham
11. Rayong
12. Ban Chang
13. Pak Nam Pra Sae
14. Chanthaburi
15. Khlung
16. Trat
17. Khlong Yai
18. Sa Kaew
19. Watthananakorn
20. Aranyaprathet
21. Prachuap Khiri Khan
22. Kabin Buri

================================================================================
TECHNICAL DETAILS
================================================================================

SOURCE FILE DETAILS:
- File: 11.ข้อมูล PWA Always-on สะสม เดือนมกราคม 2569 โดย กลส..xls
- Location: /sessions/serene-quirky-wozniak/mnt/Claude Test Cowork/Dashboard_PR/ข้อมูลดิบ/AlwayON/
- Sheets: 7 (ต.ค.68, พ.ย.68, ธ.ค.68, ม.ค.69, Sheet2, รวมไฟล์, สะสม 4 เดือน)
- Used Sheets: ม.ค.69 (January 2569)
- Always-On Column: Column 16 (0-indexed)
- Data Rows: 5-26 (22 branches)

EXTRACTION METHOD:
- Python xlrd library
- Automated column detection
- Decimal to percentage conversion
- Rankings calculated by Always-On values (highest to lowest)

================================================================================
CONTACT / QUESTIONS
================================================================================

For questions about this data extraction:
- Check Excel_Structure_Reference.txt for file structure details
- Review AlwayON_Data_Summary.txt for data rankings and statistics
- Source file has additional tabs (Sheet2, รวมไฟล์) for aggregate data

All files generated 2026-03-17
================================================================================
