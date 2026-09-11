from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

# Create presentation
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Define color scheme
DARK_BLUE = RGBColor(25, 55, 109)
LIGHT_BLUE = RGBColor(217, 237, 247)
GREEN = RGBColor(0, 176, 80)
CYAN = RGBColor(0, 176, 240)
YELLOW = RGBColor(255, 192, 0)
WHITE = RGBColor(255, 255, 255)
GRAY = RGBColor(89, 89, 89)
LIGHT_GRAY = RGBColor(240, 240, 240)

def add_title_box(slide, title_text):
    """Add a title box to slide"""
    title_box = slide.shapes.add_shape(1, Inches(0), Inches(0.2), Inches(10), Inches(0.7))
    title_box.fill.solid()
    title_box.fill.fore_color.rgb = DARK_BLUE
    title_box.line.color.rgb = DARK_BLUE
    title_frame = title_box.text_frame
    title_frame.text = title_text
    title_frame.paragraphs[0].font.size = Pt(36)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = WHITE
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

def add_explanation_box(slide, title, content_list, left=0.5, top=1.1, width=9, height=5.8):
    """Add explanation box with bullet points"""
    # Box background
    exp_box = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    exp_box.fill.solid()
    exp_box.fill.fore_color.rgb = LIGHT_BLUE
    exp_box.line.color.rgb = DARK_BLUE
    exp_box.line.width = Pt(2)
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(left + 0.2), Inches(top + 0.15), Inches(width - 0.4), Inches(0.4))
    tf = title_box.text_frame
    tf.text = title
    tf.paragraphs[0].font.size = Pt(18)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = DARK_BLUE
    
    # Content
    content_box = slide.shapes.add_textbox(Inches(left + 0.3), Inches(top + 0.65), Inches(width - 0.6), Inches(height - 0.85))
    tf = content_box.text_frame
    tf.word_wrap = True
    
    for i, content in enumerate(content_list):
        if i == 0:
            tf.text = content
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
            p.text = content
            p.level = 0
        
        p.font.size = Pt(13)
        p.font.color.rgb = GRAY
        p.space_before = Pt(4)
        p.space_after = Pt(4)

# ==================== SLIDE 1: Title & Dashboard Overview ====================
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
background = slide1.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = WHITE

# Main title
title_box = slide1.shapes.add_shape(1, Inches(0), Inches(0.5), Inches(10), Inches(1.2))
title_box.fill.solid()
title_box.fill.fore_color.rgb = DARK_BLUE
title_box.line.color.rgb = DARK_BLUE
title_frame = title_box.text_frame
title_frame.text = "SOP Report Dashboard"
title_frame.paragraphs[0].font.size = Pt(60)
title_frame.paragraphs[0].font.bold = True
title_frame.paragraphs[0].font.color.rgb = WHITE
title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Subtitle
subtitle_box = slide1.shapes.add_textbox(Inches(1), Inches(1.9), Inches(8), Inches(0.5))
tf = subtitle_box.text_frame
tf.text = "Advanced SOP Library Management System with Dynamic Filtering"
tf.paragraphs[0].font.size = Pt(20)
tf.paragraphs[0].font.color.rgb = GRAY
tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# Filter section
filter_box = slide1.shapes.add_shape(1, Inches(0.5), Inches(2.7), Inches(9), Inches(4))
filter_box.fill.solid()
filter_box.fill.fore_color.rgb = LIGHT_BLUE
filter_box.line.color.rgb = DARK_BLUE
filter_box.line.width = Pt(2)

# Filter title
filter_title = slide1.shapes.add_textbox(Inches(0.7), Inches(2.9), Inches(8.6), Inches(0.4))
tf = filter_title.text_frame
tf.text = "📊 Advanced Filtering Capabilities"
tf.paragraphs[0].font.size = Pt(22)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = DARK_BLUE

# Filter options
filter_options = [
    "✓ Year Filter - Select specific year or view all years | Real-time data filtering",
    "✓ Quarter Filter - Q1, Q2, Q3, Q4 or All Quarters | Quarterly performance tracking",
    "✓ Month Filter - Choose specific month for detailed analysis | Month-on-month comparison",
    "✓ Workgroup Filter - Filter by specific workgroup/department | Department-wise SOP management"
]

y_pos = 3.5
for option in filter_options:
    text_box = slide1.shapes.add_textbox(Inches(1.2), Inches(y_pos), Inches(8), Inches(0.35))
    tf = text_box.text_frame
    tf.text = option
    tf.word_wrap = True
    tf.paragraphs[0].font.size = Pt(12)
    tf.paragraphs[0].font.color.rgb = GRAY
    y_pos += 0.4

# Action info
action_box = slide1.shapes.add_textbox(Inches(1), Inches(6.2), Inches(8), Inches(0.5))
tf = action_box.text_frame
tf.text = "🔄 Apply Filter Button → Fetches Updated Data from SOP Library Database → Display Changes Dynamically"
tf.paragraphs[0].font.size = Pt(13)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = GREEN
tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# ==================== SLIDE 2: SOP Statistics Overview ====================
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
background = slide2.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = WHITE

add_title_box(slide2, "Slide 2: SOP Statistics Overview")

# Explanation
explanations = [
    "📋 DATA SOURCE: SOP Library Database - Total records maintained",
    "🔢 TOTAL SOPs = 834: Sum of all SOPs across all workgroups (Published + In-Development)",
    "✅ TOTAL PUBLISHED SOPs = 633: Ready-to-use SOPs (75.8% of total) - actively implemented",
    "🔧 TOTAL IN-DEVELOPMENT SOPs = 201: Under development or review (24.1% of total) - future rollout",
    "📊 FILTERING: When you select Year/Quarter/Month/Workgroup → These numbers update automatically",
    "💡 KEY INSIGHT: 3 out of 4 SOPs are Published | Strong pipeline for continuous improvement"
]

add_explanation_box(slide2, "📌 What This Data Represents:", explanations)

# Statistics boxes
stats = [
    {"title": "Total SOPs", "value": "834", "left": 0.8, "color": RGBColor(0, 102, 204)},
    {"title": "Published SOPs", "value": "633", "left": 3.7, "color": GREEN},
    {"title": "In-Development", "value": "201", "left": 6.6, "color": CYAN}
]

for stat in stats:
    stat_box = slide2.shapes.add_shape(1, Inches(stat["left"]), Inches(6.2), Inches(2.8), Inches(1))
    stat_box.fill.solid()
    stat_box.fill.fore_color.rgb = RGBColor(217, 237, 247)
    stat_box.line.color.rgb = stat["color"]
    stat_box.line.width = Pt(2)
    
    title_textbox = slide2.shapes.add_textbox(Inches(stat["left"] + 0.1), Inches(6.25), Inches(2.6), Inches(0.35))
    tf = title_textbox.text_frame
    tf.text = stat["title"]
    tf.paragraphs[0].font.size = Pt(11)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = DARK_BLUE
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    value_textbox = slide2.shapes.add_textbox(Inches(stat["left"] + 0.1), Inches(6.55), Inches(2.6), Inches(0.6))
    tf = value_textbox.text_frame
    tf.text = stat["value"]
    tf.paragraphs[0].font.size = Pt(42)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = stat["color"]
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# ==================== SLIDE 3: SOPs Status Distribution (Pie Chart) ====================
slide3 = prs.slides.add_slide(prs.slide_layouts[6])
background = slide3.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = WHITE

add_title_box(slide3, "Slide 3: SOPs Status Distribution")

# Explanation for pie chart
pie_explanations = [
    "📊 PIE CHART DATA: Shows breakdown of all SOPs by their approval status",
    "✅ APPROVED (Green - 633 SOPs / 75.9%): Fully approved & ready for use across organization",
    "⏳ PENDING (Cyan - 194 SOPs / 23.3%): Awaiting final approval from stakeholders/management",
    "📝 DRAFT (Yellow - 7 SOPs / 0.8%): Still under development, not yet submitted for approval",
    "🔄 DYNAMIC UPDATE: When filter applied → Status distribution recalculates based on selected criteria",
    "💡 ANALYSIS: Majority (75.9%) are Approved | Small pending queue (23.3%) | Minimal drafts (0.8%)"
]

add_explanation_box(slide3, "📌 What This Pie Chart Shows:", pie_explanations, top=1.1, height=2.7)

# Add pie chart
x, y, cx, cy = Inches(0.7), Inches(4), Inches(4.5), Inches(3)
chart_data = CategoryChartData()
chart_data.chart_type = XL_CHART_TYPE.PIE

chart_data.add_category("Status")
chart_data.add_series("Count", (633, 194, 7))

chart = slide3.shapes.add_chart(XL_CHART_TYPE.PIE, x, y, cx, cy).chart
chart.has_legend = True
chart.legend.position = 2
chart.legend.include_in_layout = False

# Status details
details_box = slide3.shapes.add_textbox(Inches(5.5), Inches(4), Inches(4), Inches(3))
tf = details_box.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "● Approved"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = GREEN

p = tf.add_paragraph()
p.text = "633 (75.9%)"
p.font.size = Pt(12)
p.font.color.rgb = GRAY

p = tf.add_paragraph()
p.text = "Ready for Implementation"
p.font.size = Pt(10)
p.font.italic = True
p.font.color.rgb = RGBColor(150, 150, 150)
p.space_after = Pt(8)

p = tf.add_paragraph()
p.text = "● Pending"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = CYAN

p = tf.add_paragraph()
p.text = "194 (23.3%)"
p.font.size = Pt(12)
p.font.color.rgb = GRAY

p = tf.add_paragraph()
p.text = "Awaiting Approval"
p.font.size = Pt(10)
p.font.italic = True
p.font.color.rgb = RGBColor(150, 150, 150)
p.space_after = Pt(8)

p = tf.add_paragraph()
p.text = "● Draft"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = YELLOW

p = tf.add_paragraph()
p.text = "7 (0.8%)"
p.font.size = Pt(12)
p.font.color.rgb = GRAY

p = tf.add_paragraph()
p.text = "Under Development"
p.font.size = Pt(10)
p.font.italic = True
p.font.color.rgb = RGBColor(150, 150, 150)

# ==================== SLIDE 4: Average Aging by Workgroup (Table) ====================
slide4 = prs.slides.add_slide(prs.slide_layouts[6])
background = slide4.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = WHITE

add_title_box(slide4, "Slide 4: Average Aging by Workgroup (Days)")

# Explanation for table
table_explanations = [
    "📊 TABLE CONTENT: Shows average days SOPs have been in each status, grouped by workgroup",
    "🏢 WORKGROUP COLUMN: Different departments/teams managing SOPs (Denmark RTR, China PTP, etc.)",
    "🔧 IN-DEVELOPMENT: Average days SOPs spend in development phase before submission",
    "⏳ PENDING FOR APPROVAL: Average days SOPs wait for final approval from management",
    "✅ APPROVED: Average days since SOPs were approved & became active",
    "📈 DATA FILTERING: Select Year/Quarter/Month/Workgroup → Table recalculates averages for that selection",
    "💡 USAGE: Identify bottlenecks where SOPs are stuck | Monitor approval cycle efficiency"
]

add_explanation_box(slide4, "📌 What This Table Shows:", table_explanations, top=1.1, height=2.5)

# Table
left = Inches(0.5)
top = Inches(4)
width = Inches(9)
height = Inches(2.8)
rows = 6
cols = 4

table_shape = slide4.shapes.add_table(rows, cols, left, top, width, height).table

table_shape.columns[0].width = Inches(2.5)
table_shape.columns[1].width = Inches(2)
table_shape.columns[2].width = Inches(2)
table_shape.columns[3].width = Inches(2.5)

# Header
headers = ["Workgroup", "In-Development", "Pending for Approval", "Approved"]
for col, header in enumerate(headers):
    cell = table_shape.cell(0, col)
    cell.text = header
    cell.fill.solid()
    cell.fill.fore_color.rgb = DARK_BLUE
    tf = cell.text_frame
    tf.paragraphs[0].font.size = Pt(11)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = WHITE
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# Data
workgroup_data = [
    ("Denmark RTR", "-", "-", "87"),
    ("China PTP", "-", "-", "583"),
    ("China RTR", "-", "-", "411"),
    ("China OTC", "-", "-", "383"),
    ("China B RTR", "-", "-", "383")
]

for row, (workgroup, indiv, pending, approved) in enumerate(workgroup_data, 1):
    data = [workgroup, indiv, pending, approved]
    for col, value in enumerate(data):
        cell = table_shape.cell(row, col)
        cell.text = value
        if row % 2 == 0:
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT_GRAY
        tf = cell.text_frame
        tf.paragraphs[0].font.size = Pt(10)
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# Pagination
paging_text = slide4.shapes.add_textbox(Inches(0.5), Inches(7), Inches(9), Inches(0.4))
tf = paging_text.text_frame
tf.text = "📄 Page 1 of 25 | Showing 1-5 of 121 Workgroups | Export To Excel Available"
tf.paragraphs[0].font.size = Pt(11)
tf.paragraphs[0].font.color.rgb = GRAY
tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# ==================== SLIDE 5: Published SOPs by Workgroup ====================
slide5 = prs.slides.add_slide(prs.slide_layouts[6])
background = slide5.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = WHITE

add_title_box(slide5, "Slide 5: Published SOPs Count by Workgroup")

# Explanation
published_explanations = [
    "📊 BAR CHART: Shows count of APPROVED & PUBLISHED SOPs for each workgroup",
    "🎯 DATA INCLUDES: Only SOPs with status = 'Approved' (Not pending or draft)",
    "📈 BAR HEIGHTS: Taller bar = More published SOPs | AMER COE leads with 54 SOPs",
    "🏢 WORKGROUPS: AMER COE (54) → UK RTR Med (24) → Denmark RTR (20) → and so on",
    "🔄 FILTER IMPACT: Select Year/Quarter/Month/Workgroup → Chart updates to show counts for that period",
    "💡 INSIGHTS: AMER COE dominates (54 SOPs) | Others have 16-20 SOPs | Regional distribution visible"
]

add_explanation_box(slide5, "📌 What This Bar Chart Shows:", published_explanations, top=1.1, height=2.5)

# Chart
x, y, cx, cy = Inches(0.5), Inches(4), Inches(9), Inches(2.8)
chart_data = CategoryChartData()
chart_data.chart_type = XL_CHART_TYPE.COLUMN_CLUSTERED

workgroups = ["AMER COE", "UK RTR Med", "Denmark RTR", "US-RTR CCo", "US-RTR CCo", "US RTR", "Norway RTR", "Finland RTR", "Others"]
values = [54, 24, 20, 20, 19, 18, 17, 16, 15]

chart_data.categories = workgroups
chart_data.add_series("Published SOPs", tuple(values))

chart = slide5.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy).chart
chart.has_legend = False

# Chart value labels
plot = chart.plots[0]
plot.has_data_labels = True
data_labels = plot.data_labels
data_labels.number_format = '0'
data_labels.position = 4  # Above bar

# Insights
insights_text = slide5.shapes.add_textbox(Inches(1), Inches(7), Inches(8), Inches(0.4))
tf = insights_text.text_frame
tf.text = "🎯 KEY: AMER COE leads with 54 published SOPs | Strong workgroup distribution | Hover to see exact values"
tf.paragraphs[0].font.size = Pt(12)
tf.paragraphs[0].font.color.rgb = GREEN
tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# ==================== SLIDE 6: In-Development SOPs by Workgroup ====================
slide6 = prs.slides.add_slide(prs.slide_layouts[6])
background = slide6.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = WHITE

add_title_box(slide6, "Slide 6: In-Development SOPs Count by Workgroup")

# Explanation
indev_explanations = [
    "📊 BAR CHART: Shows count of SOPs currently IN-DEVELOPMENT for each workgroup",
    "🔧 DATA INCLUDES: Only SOPs with status = 'Draft' or 'In-Development' (Not approved yet)",
    "📈 BAR HEIGHTS: EMER UK ER leads with 21 SOPs under development | Global Center with 16",
    "🏢 WORKGROUPS: EMER UK ER (21) → Global Center (16) → US PTP HQ (14) → and continuing",
    "🔄 FILTER IMPACT: Select Year/Quarter/Month/Workgroup → Chart updates with filtered development counts",
    "💡 INSIGHTS: 201 total in-development SOPs | Active pipeline ensures future SOP rollouts | Monitor development progress"
]

add_explanation_box(slide6, "📌 What This Bar Chart Shows:", indev_explanations, top=1.1, height=2.5)

# Chart
x, y, cx, cy = Inches(0.5), Inches(4), Inches(9), Inches(2.8)
chart_data = CategoryChartData()
chart_data.chart_type = XL_CHART_TYPE.COLUMN_CLUSTERED

workgroups_dev = ["EMER UK ER", "Global Cen", "US PTP HQ", "APAC CCOE", "US-RTR CXM", "Global Fun", "US-RTR CCo", "Central AS", "India"]
values_dev = [21, 16, 14, 11, 9, 8, 8, 7, 5]

chart_data.categories = workgroups_dev
chart_data.add_series("In-Development SOPs", tuple(values_dev))

chart = slide6.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy).chart
chart.has_legend = False

# Chart value labels
plot = chart.plots[0]
plot.has_data_labels = True
data_labels = plot.data_labels
data_labels.number_format = '0'
data_labels.position = 4

# Insights
insights_text = slide6.shapes.add_textbox(Inches(1), Inches(7), Inches(8), Inches(0.4))
tf = insights_text.text_frame
tf.text = "🔄 KEY: EMER UK ER leading with 21 in-development SOPs | Strong future pipeline | Ensure timely completion"
tf.paragraphs[0].font.size = Pt(12)
tf.paragraphs[0].font.color.rgb = CYAN
tf.paragraphs[0].alignment = PP_ALIGN.CENTER

# Save presentation
prs.save("SOP_Report_Dashboard.pptx")
print("✅ PowerPoint presentation created successfully!")
print("\n" + "="*60)
print("📊 SOP REPORT DASHBOARD - 6 SLIDES CREATED")
print("="*60)
print("\n📑 SLIDE BREAKDOWN:\n")
print("Slide 1: Dashboard Overview & Advanced Filtering Capabilities")
print("   └─ Filters: Year, Quarter, Month, Workgroup")
print("   └─ Action: Apply Filter → Data Updates Dynamically\n")
print("Slide 2: SOP Statistics Overview")
print("   └─ Total SOPs: 834")
print("   └─ Published: 633 (75.8%)")
print("   └─ In-Development: 201 (24.1%)\n")
print("Slide 3: SOPs Status Distribution (Pie Chart)")
print("   └─ Approved: 633 (75.9%)")
print("   └─ Pending: 194 (23.3%)")
print("   └─ Draft: 7 (0.8%)\n")
print("Slide 4: Average Aging by Workgroup (Table)")
print("   └─ Shows aging days by: In-Development, Pending, Approved")
print("   └─ Data grouped by 5 major workgroups\n")
print("Slide 5: Published SOPs by Workgroup (Bar Chart)")
print("   └─ AMER COE: 54 SOPs (Leading)")
print("   └─ UK RTR Med: 24 SOPs")
print("   └─ Other workgroups: 15-20 SOPs each\n")
print("Slide 6: In-Development SOPs by Workgroup (Bar Chart)")
print("   └─ EMER UK ER: 21 SOPs (Leading)")
print("   └─ Global Center: 16 SOPs")
print("   └─ Other workgroups: 5-14 SOPs each\n")
print("="*60)
print("File saved: SOP_Report_Dashboard.pptx")
print("="*60)
