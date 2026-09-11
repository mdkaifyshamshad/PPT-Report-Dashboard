# 📊 SOP Report Dashboard Presentation

## Overview
This project contains a Python script to generate a professional PowerPoint presentation for SOP (Standard Operating Procedures) Report Dashboard.

## 📋 Presentation Details

### 6 Professional Slides with Detailed Explanations:

**Slide 1: Dashboard Overview & Filtering Capabilities**
- Year, Quarter, Month, and Workgroup filters
- Dynamic data fetching from SOP Library
- Apply Filter functionality

**Slide 2: SOP Statistics Overview**
- Total SOPs: 834
- Published SOPs: 633 (75.8%)
- In-Development SOPs: 201 (24.1%)

**Slide 3: SOPs Status Distribution (Pie Chart)**
- Approved: 633 (75.9%)
- Pending: 194 (23.3%)
- Draft: 7 (0.8%)

**Slide 4: Average Aging by Workgroup (Table)**
- Workgroup-wise aging data
- In-Development, Pending, Approved columns
- Pagination info included

**Slide 5: Published SOPs by Workgroup (Bar Chart)**
- AMER COE: 54 SOPs
- UK RTR Medical: 24 SOPs
- Others: 15-20 SOPs each

**Slide 6: In-Development SOPs by Workgroup (Bar Chart)**
- EMER UK ER: 21 SOPs
- Global Center: 16 SOPs
- Others: 5-14 SOPs each

## 🚀 How to Generate the PPT

### Requirements
```bash
pip install python-pptx
```

### Run the Script
```bash
python create_presentation.py
```

### Output
The script will generate: **SOP_Report_Dashboard.pptx**

## 📊 File Structure
```
PPT-Report-Dashboard/
├── create_presentation.py    (Main script)
├── README.md                 (This file)
└── SOP_Report_Dashboard.pptx (Generated file)
```

## 🎨 Design Features
- ✅ Professional color scheme (Dark Blue, Light Blue, Green, Cyan)
- ✅ Responsive layouts
- ✅ Detailed explanations on each slide
- ✅ Charts and data visualizations
- ✅ Business insights and key metrics
- ✅ Dynamic filtering logic explained

## 📌 Key Information

### Data Filtering Logic:
When you select **Year/Quarter/Month/Workgroup**:
1. Filter criteria is sent to SOP Library Database
2. Data is fetched based on selected filters
3. All visualizations update dynamically
4. Statistics recalculate automatically

### Chart Explanations:
- **Pie Chart (Slide 3)**: Shows approval status breakdown
- **Bar Charts (Slides 5-6)**: Shows workgroup-wise SOP counts
- **Table (Slide 4)**: Shows aging metrics by workgroup

## �� Business Use Cases
- Monitor SOP approval pipeline
- Track workgroup performance
- Identify bottlenecks in approval process
- Plan future SOP rollouts
- Measure SOP aging metrics

## 👤 Author
MD KAIFY

## 📅 Last Updated
September 11, 2026

---

**Ready to use! Generate your presentation now.** 🎯
