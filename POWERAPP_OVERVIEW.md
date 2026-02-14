# PowerApps Absence Viewer - Overview & Architecture

## Table of Contents
1. [Architecture Overview](#architecture-overview)
2. [Data Flow](#data-flow)
3. [Key Differences from HTML/JS Version](#key-differences)
4. [Prerequisites](#prerequisites)
5. [File Structure](#file-structure)

---

## Architecture Overview

### High-Level Structure

```
PowerApp Canvas App
├── Data Sources
│   ├── SharePoint Document Library (CSV files)
│   ├── Collections (in-memory data processing)
│   └── Context Variables (state management)
├── Screens
│   ├── Home Screen (file upload & settings)
│   ├── Weekly Report Screen
│   └── Daily Report Screen
├── Components
│   ├── Hierarchical Gallery (Market → Dept → Person)
│   ├── Date Range Picker
│   ├── Export Button (triggers Power Automate)
│   └── Status Legend
└── Power Automate Flows
    ├── Process CSV Files
    └── Export to Excel
```

### Key Components

#### 1. **Data Layer**
- **SharePoint Document Library**: Store the 3 CSV files
  - Absence_Data.csv
  - Organization_Data.csv
  - Holiday_Calendar.csv
- **Collections**: In-memory data structures for processing
  - `colAbsenceData` - Parsed absence records
  - `colOrgData` - Department to Market mapping
  - `colHolidayData` - Holiday calendar
  - `colProcessedData` - Calculated weekly/daily absences
  - `colWeeklyReport` - Formatted weekly report data
  - `colDailyReport` - Formatted daily report data

#### 2. **Business Logic Layer**
- Week number calculation (Friday-ending weeks)
- Working day distribution (Mon-Fri only)
- Hour/day conversion based on normal working hours
- Hierarchical grouping and aggregation
- Filtering (exclude Withdrawn/Denied statuses)

#### 3. **Presentation Layer**
- Nested galleries for hierarchy display
- Conditional formatting for weekends/holidays
- Expand/collapse state management
- Status indicators (colored icons)
- Dynamic column generation

---

## Data Flow

### 1. **File Upload & Parsing**
```
SharePoint Library → Power Automate → Parse CSV → Collections
```

**Process:**
1. User uploads CSV files to SharePoint document library
2. Power Automate flow triggers on file upload
3. Flow reads CSV content and parses it
4. Flow calls PowerApps to populate collections

**Alternative (Simpler):** Use SharePoint lists instead of CSV files
- Create 3 SharePoint lists with matching column structure
- Users maintain data directly in SharePoint
- PowerApps connects directly to lists (no parsing needed)

### 2. **Data Processing**
```
Raw Collections → Calculate Weeks → Distribute Hours → Group by Hierarchy → Report Collections
```

**Steps:**
1. Load data into collections on app start
2. Filter out invalid statuses (Withdrawn, Denied)
3. For each absence record:
   - Calculate date range (start to end)
   - Identify working days (Mon-Fri only)
   - Distribute hours evenly across working days
   - Calculate week numbers for each day
   - Map department to market sub-sector
4. Group by Market → Department → Person
5. Aggregate hours by week/day
6. Store in report collections

### 3. **Report Generation**
```
Report Collections → Nested Galleries → User Interaction → Display
```

**Weekly Report:**
- Generate 52 weeks from start date
- Display hierarchical galleries
- Apply holiday highlighting
- Enable expand/collapse

**Daily Report:**
- Generate up to 14 days from date range
- Display hierarchical galleries
- Apply weekend/holiday highlighting
- Enable expand/collapse

### 4. **Export to Excel**
```
Report Collections → Power Automate → Excel File → Download
```

**Process:**
1. User clicks Export button
2. PowerApps sends report data to Power Automate flow
3. Flow creates Excel file with formatting
4. Flow returns file download link to user

---

## Key Differences from HTML/JS Version

| Feature | HTML/JS Version | PowerApps Version |
|---------|----------------|-------------------|
| **File Upload** | Direct browser file upload | SharePoint library or Power Automate |
| **Data Storage** | localStorage | Collections (session) or SharePoint (persistent) |
| **Data Processing** | JavaScript loops/functions | Power Fx formulas (AddColumns, GroupBy, etc.) |
| **Hierarchical Display** | Nested HTML tables | Nested galleries with visibility toggles |
| **Tooltips** | CSS hover tooltips | Info icons + pop-up labels or screens |
| **Expand/Collapse** | DOM manipulation | Context variables + gallery filters |
| **Excel Export** | ExcelJS library (client-side) | Power Automate flow (server-side) |
| **Week Calculation** | Custom JavaScript function | Power Fx formula |
| **Conditional Formatting** | CSS classes | Gallery item conditionals |

---

## Prerequisites

### Required Services
1. **Microsoft 365 Account** with:
   - PowerApps license (included in many M365 plans)
   - SharePoint Online
   - Power Automate (included with PowerApps)

2. **SharePoint Site** for:
   - Document library (to store CSV files) OR
   - Custom lists (to store data directly)

3. **Development Environment**:
   - PowerApps Studio (web or desktop)
   - Access to Power Automate flow designer

### Optional but Recommended
- **Power BI** (for advanced reporting/dashboards)
- **Dataverse** (for enterprise-scale data storage)

---

## File Structure

This documentation is organized into the following files:

### Documentation Files
1. **POWERAPP_OVERVIEW.md** (this file) - Architecture and overview
2. **POWERAPP_DATA_SETUP.md** - SharePoint setup and data structure
3. **POWERAPP_FORMULAS.md** - All Power Fx formulas for calculations
4. **POWERAPP_UI_COMPONENTS.md** - Screen and control designs
5. **POWERAPP_STEP_BY_STEP.md** - Complete build instructions
6. **POWERAPP_POWER_AUTOMATE.md** - Flow designs for CSV parsing and Excel export
7. **POWERAPP_SAMPLE_DATA.md** - Sample CSV structures for testing

### Sample Files (to be created)
- Sample_Absence_Data.csv
- Sample_Organization_Data.csv
- Sample_Holiday_Calendar.csv

---

## Important PowerApps Limitations to Consider

### 1. **Data Delegation**
- PowerApps has a 2000 record delegation limit for some operations
- With ~2000 employees, you may hit this limit
- **Solution**: Use collections for processing (loads all data into memory)

### 2. **Gallery Performance**
- Nested galleries with many items can be slow
- **Solution**: Implement expand/collapse to show fewer items at once

### 3. **Formula Complexity**
- Complex formulas can be slow and hard to debug
- **Solution**: Break processing into multiple steps using collections

### 4. **Excel Export**
- No built-in Excel export with formatting
- **Solution**: Use Power Automate flow with Excel connector

### 5. **File Upload**
- Canvas apps don't have great CSV parsing
- **Solution**: Use SharePoint lists OR Power Automate to parse CSVs

---

## Recommended Implementation Approach

### Option A: SharePoint Lists (Recommended for Simplicity)
**Pros:**
- No CSV parsing needed
- Direct data connection
- Easy data maintenance
- Better performance

**Cons:**
- Users must enter data in SharePoint (not CSV)
- May need migration from existing CSV files

### Option B: CSV Files + Power Automate
**Pros:**
- Keeps existing CSV workflow
- Users can upload files

**Cons:**
- More complex setup
- Requires Power Automate flow for parsing
- Potential performance issues with large files

**Recommendation:** Start with Option A (SharePoint lists) for faster development, then add CSV import via Power Automate later if needed.

---

## Next Steps

1. **Read POWERAPP_DATA_SETUP.md** to set up your SharePoint environment
2. **Review POWERAPP_FORMULAS.md** to understand the calculation logic
3. **Study POWERAPP_UI_COMPONENTS.md** to see the screen designs
4. **Follow POWERAPP_STEP_BY_STEP.md** to build the app
5. **Configure POWERAPP_POWER_AUTOMATE.md** flows for CSV import and Excel export

---

## Questions & Decisions Needed

Before building, please decide:

1. **Data Source Preference:**
   - [ ] SharePoint Lists (simpler, recommended)
   - [ ] CSV files via Power Automate (matches current workflow)
   - [ ] Hybrid (SharePoint lists + CSV import option)

2. **Deployment Model:**
   - [ ] Single app for all users
   - [ ] Multiple apps per market/department
   - [ ] Embedded in SharePoint page

3. **Export Requirements:**
   - [ ] Excel export is critical (requires Power Automate setup)
   - [ ] Can live without Excel export initially
   - [ ] Alternative export format acceptable (CSV, PDF)

4. **User Permissions:**
   - Who can upload/edit absence data?
   - Who can view reports?
   - Who can export data?

---

## Support & Resources

- [PowerApps Documentation](https://docs.microsoft.com/en-us/powerapps/)
- [Power Fx Formula Reference](https://docs.microsoft.com/en-us/powerapps/maker/canvas-apps/formula-reference)
- [Power Automate Documentation](https://docs.microsoft.com/en-us/power-automate/)
- [SharePoint Lists Documentation](https://support.microsoft.com/en-us/office/introduction-to-lists-0a1c3ace-def0-44af-b225-cfa8d92c52d7)
