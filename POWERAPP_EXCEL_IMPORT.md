# PowerApps Absence Viewer - Weekly Excel Import Guide

## Overview

This guide covers importing a **weekly Excel file with 10,000+ rows** into SharePoint for use in PowerApps.

---

## Architecture for Large Datasets

### Data Flow

```
Weekly Excel File (10k+ rows)
    ↓
SharePoint Document Library
    ↓
Power Automate Flow (Triggered on upload)
    ↓
Parse Excel → Filter → Validate
    ↓
SharePoint List (AbsenceData)
    ↓
PowerApps (Date-filtered queries only)
```

---

## Step 1: SharePoint Document Library Setup

### Create Document Library

1. Navigate to your SharePoint site
2. Click "New" → "Document Library"
3. Name: **"Absence Data Files"**
4. Description: "Weekly absence data Excel files"

### Create Folder Structure

```
Absence Data Files/
├── Current/          (Current week's file)
├── Archive/          (Previous weeks)
└── Error/            (Failed imports)
```

### File Naming Convention

**Required format:** `Absence_Data_YYYY-MM-DD.xlsx`

Example: `Absence_Data_2026-02-14.xlsx`

**Why this matters:**
- Consistent naming enables automation
- Date in filename helps with version tracking
- Flow can extract upload date from filename

---

## Step 2: Excel File Requirements

### Table Structure

Your Excel file MUST have data in a **named table** (not just a range).

**To create a table in Excel:**
1. Select your data range (including headers)
2. Insert → Table (or Ctrl+T)
3. Check "My table has headers"
4. Click OK
5. Table Design tab → Table Name: **"AbsenceTable"**
6. Save file

### Column Names (Exact Match Required)

| Column Name | Type | Format |
|------------|------|--------|
| LAST NAME | Text | - |
| FIRST NAME | Text | - |
| DEPARTMENT | Text | XXX-ENV-### |
| LINE MANAGER | Text | - |
| ABSENCE START DATE | Date | mm/dd/yyyy or yyyy-mm-dd |
| ABSENCE END DATE | Date | mm/dd/yyyy or yyyy-mm-dd |
| ABSENCE DURATION | Number | Decimal |
| UOM | Text | H or D |
| ABSENCE TYPE | Text | - |
| ABSENCE STATUS | Text | - |
| ABSENCE COMMENT | Text | - |
| NORMAL WORKING HOURS | Number | Decimal |

---

## Step 3: Power Automate Import Flow

### Flow Design for 10k+ Rows

**Key Challenge:** Power Automate has execution time limits
- Standard flows: 5 minutes
- Premium flows: 30 minutes

**With 10k rows:**
- Parsing: ~2 minutes
- Validation: ~1 minute
- Import: ~10-15 minutes (at ~15 rows/second)
- **Total: 13-18 minutes**

**Solution: Batch Processing**

### Flow: Weekly Excel Import (Optimized for Large Files)

#### Trigger

**When a file is created (SharePoint)**
- Site Address: Your site
- Library Name: Absence Data Files
- Folder: /Current
- Include subfolders: No

#### Initialize Variables

```
varBatchSize = 500          (records per batch)
varTotalRows = 0            (counter)
varSuccessCount = 0         (successful imports)
varErrorCount = 0           (failed imports)
varStartTime = utcNow()     (for logging)
```

#### Get File Properties

**Get file metadata (SharePoint)**
- Site Address: Same
- Library Name: Absence Data Files
- Id: (from trigger)

Store:
- File Name
- File Size
- Modified date

#### List Rows in Excel Table

**List rows present in a table (Excel Online)**
- Location: SharePoint Site
- Document Library: Absence Data Files
- File: File Identifier (from trigger)
- Table: AbsenceTable

**Critical Setting:**
- Top Count: 15000 (set higher than your max expected rows)

#### Batch Processing Loop

Since we can't use actual batching easily in Power Automate, use **concurrency control**:

**Apply to each** (Excel rows)
- Settings → Concurrency Control: **ON**
- Degree of Parallelism: **25** (process 25 rows at once)

Inside loop:

```
┌─────────────────────────────────────────┐
│ Condition: Filter bad records            │
│ - Status not Withdrawn                   │
│ - Status not Denied                      │
│ - Department not blank                   │
│ - Start Date <= End Date                 │
└────┬───────────────────────┬────────────┘
     │ Yes                   │ No
     ▼                       ▼
┌──────────────────┐  ┌───────────────────┐
│ Create/Update    │  │ Log error         │
│ SharePoint item  │  │ Increment error   │
│                  │  │ count             │
└────┬─────────────┘  └───────────────────┘
     │
     ▼
┌──────────────────────────────────────────┐
│ Increment success count                  │
└──────────────────────────────────────────┘
```

#### Create or Update Item (Upsert Pattern)

**Challenge:** How to avoid duplicates when re-importing?

**Solution 1: Clear and Reload**
- Before loop: Delete all items from AbsenceData list
- Then import all rows fresh
- **Pros:** Simple, no duplicates
- **Cons:** Breaks any PowerApps user currently viewing data

**Solution 2: Upsert by Composite Key** (Recommended)
- Create unique ID: `Person + StartDate + EndDate`
- Check if item exists before creating

**Flow logic:**

```powerFx
// Compose unique key
concat(item()?['LAST NAME'], '_', item()?['FIRST NAME'], '_', formatDateTime(item()?['ABSENCE START DATE'], 'yyyy-MM-dd'), '_', formatDateTime(item()?['ABSENCE END DATE'], 'yyyy-MM-dd'))

// Get items with this key
Filter items: UniqueKey = (composed value above)

// Condition: Count > 0?
If Yes: Update existing item
If No: Create new item
```

**Specific actions:**

**Get items (SharePoint)**
- Site: Your site
- List: AbsenceData
- Filter Query:
  ```odata
  UniqueKey eq '@{outputs('Compose_UniqueKey')}'
  ```
- Top Count: 1

**Condition:** length(outputs('Get_items')?['body/value']) greater than 0

**If Yes branch:**
- Update item (SharePoint)
- Item ID: `first(outputs('Get_items')?['body/value'])?['ID']`
- Map all fields from Excel row

**If No branch:**
- Create item (SharePoint)
- Map all fields from Excel row
- UniqueKey: From Compose step

#### Post-Processing

After "Apply to each" completes:

**1. Move File to Archive**
```
Move file (SharePoint)
From: /Current/{filename}
To: /Archive/{filename}
```

**2. Send Summary Email**
```
Send email (V2)
To: admin@company.com
Subject: "Weekly Absence Import Complete"
Body:
  File: {filename}
  Total rows: {varTotalRows}
  Successful: {varSuccessCount}
  Errors: {varErrorCount}
  Duration: {difference between utcNow() and varStartTime}

  {if errors > 0, attach error log}
```

**3. Update SharePoint Metadata**

Create a "Import Log" list to track each import:
- ImportDate
- FileName
- RowsProcessed
- SuccessCount
- ErrorCount
- Duration
- Status (Success/Partial/Failed)

---

## Step 4: Error Handling & Monitoring

### Common Errors with Large Files

#### Error: "Flow timeout after 5 minutes"

**Cause:** Too many rows to process in time limit
**Solution:**
- Enable concurrency (25 parallel)
- Or upgrade to Premium connector (30 min timeout)
- Or split import into multiple smaller flows

#### Error: "SharePoint throttling - too many requests"

**Cause:** Creating 10k items too fast triggers throttle
**Solution:**
- Add delay: 100ms between creates (in Apply to each)
- Use batch operations (Create multiple items at once)
- Reduce concurrency from 25 to 10

#### Error: "Null reference in Excel column"

**Cause:** Excel cell is empty but SharePoint field is required
**Solution:**
- Add null checks in flow:
  ```
  if(empty(item()?['ABSENCE COMMENT']), '', item()?['ABSENCE COMMENT'])
  ```

### Retry Logic

For each SharePoint action:
1. Click "..." → Settings
2. Retry Policy: Exponential
3. Count: 3
4. Interval: PT10S (10 seconds)

### Error Notifications

**Add "Configure run after" on Send Email:**
- Run after: "Apply to each" has failed OR has timed out
- Send error notification with run ID

---

## Step 5: PowerApps Integration with Large Dataset

### Modified Data Loading Strategy

**DON'T do this (too slow):**
```powerFx
// ❌ Loads all 10k+ rows into memory
ClearCollect(colAbsenceData, AbsenceData);
```

**DO this instead:**
```powerFx
// ✅ Load only date-filtered records
ClearCollect(
    colAbsenceData,
    Filter(
        AbsenceData,
        AbsenceStartDate >= varReportStartDate &&
        AbsenceEndDate <= varReportEndDate &&
        AbsenceStatus <> "Withdrawn" &&
        AbsenceStatus <> "Denied"
    )
);

// Typically returns 500-2000 records instead of 10k+
```

### Performance Tips

**1. Use Views in SharePoint**

Create a view "Current Year Active":
- Filter: Year([AbsenceStartDate]) = 2026
- Filter: [AbsenceStatus] ≠ "Withdrawn"
- Filter: [AbsenceStatus] ≠ "Denied"

Then in PowerApps:
```powerFx
// Connect to the VIEW, not the full list
ClearCollect(colAbsenceData, 'AbsenceData Current Year Active');
```

**2. Lazy Loading**

Show Home Screen immediately, load data only when needed:

```powerFx
// App.OnStart - Load ONLY reference tables
ClearCollect(colOrgData, OrganizationData);
ClearCollect(colHolidayData, HolidayCalendar);
Set(varDataLoaded, true);

// btnGenerateReport.OnSelect - Load absence data on demand
ClearCollect(colAbsenceFiltered, Filter(AbsenceData, ...));
```

**3. Progress Indicators**

With large datasets, show progress:

```powerFx
// Show spinner
Set(varLoading, true);

// Update label
Set(varLoadingMessage, "Loading absence data...");

// Load data
ClearCollect(...);

// Update label
Set(varLoadingMessage, "Processing " & CountRows(colAbsenceData) & " records...");

// Process
// ...

// Hide spinner
Set(varLoading, false);
```

---

## Step 6: Testing with Large Data

### Test Scenarios

**1. Small File Test (100 rows)**
- Verify flow completes successfully
- Check data accuracy
- Validate all fields populated

**2. Medium File Test (1,000 rows)**
- Verify performance acceptable (<2 min)
- Check memory usage
- Validate no throttling

**3. Large File Test (10,000+ rows)**
- Verify flow completes (<20 min)
- Check for timeout issues
- Validate error handling
- Monitor SharePoint throttling

**4. Duplicate Test**
- Import same file twice
- Verify no duplicate records created
- Check upsert logic working

**5. Error Test**
- Import file with bad data (invalid dates, missing fields)
- Verify errors logged correctly
- Check partial import handling

---

## Step 7: Weekly Operations

### User Workflow

1. **Monday morning:** HR exports absence data from HR system
2. **Save as:** `Absence_Data_2026-02-14.xlsx`
3. **Upload to:** SharePoint → Absence Data Files → Current folder
4. **Flow triggers automatically** (5-20 minutes)
5. **Email confirmation** sent when complete
6. **PowerApps users** can now generate reports with latest data

### Monitoring

**Weekly checklist:**
- [ ] Import email received?
- [ ] Success count matches expected?
- [ ] Any errors reported?
- [ ] File moved to Archive?
- [ ] PowerApps showing latest data?

**Monthly review:**
- Archive old files (>6 months)
- Check SharePoint list size
- Review error logs
- Optimize flow if needed

---

## Performance Benchmarks

Expected timing for various operations:

| Operation | Rows | Time |
|-----------|------|------|
| Excel parse | 10,000 | ~2 min |
| SharePoint import (sequential) | 10,000 | ~30 min |
| SharePoint import (parallel 25) | 10,000 | ~12 min |
| PowerApps filter load | 10,000 | 10-15 sec |
| PowerApps filter load (indexed) | 10,000 | 2-3 sec |
| PowerApps date-filtered load | 1,000 | 1-2 sec |
| Weekly report generation | 1,000 | 3-5 sec |

---

## Troubleshooting

### Issue: Import takes too long

**Diagnosis:** Check flow run history → See which step is slow
**Solutions:**
- Increase concurrency (Settings → Concurrency Control)
- Reduce validation logic
- Consider premium connector for longer timeout

### Issue: SharePoint throttling

**Diagnosis:** Error message "Request exceeded limit"
**Solutions:**
- Add 100ms delay between creates
- Reduce concurrency from 25 to 10
- Use batch create operations

### Issue: Memory errors in PowerApps

**Diagnosis:** App crashes or becomes unresponsive
**Solutions:**
- Never load full 10k list into collection
- Always filter by date range first
- Use SharePoint views for pre-filtering
- Consider Dataverse for large datasets

---

## Advanced: Incremental Updates

If you only want to import **new or changed** records:

### Option 1: Delta File
HR system exports only changed records since last export

### Option 2: Last Modified Check
```powerFx
// In Power Automate
Filter Excel rows where:
  LastModifiedDate > (date of last import)
```

### Option 3: Weekly Partitions
Create separate lists per month/quarter:
- AbsenceData_2026_Q1
- AbsenceData_2026_Q2
- Keeps each list smaller
- Archive old quarters

---

## Next Steps

1. ✅ Set up SharePoint Document Library
2. ✅ Convert your Excel file to use named table
3. ✅ Create Power Automate import flow
4. ✅ Test with small dataset first
5. ✅ Test with full 10k+ dataset
6. ✅ Update PowerApps to use filtered queries
7. ✅ Document weekly upload process for HR team

---

## Summary

**Key Points for 10k+ Rows:**
- ✅ Use SharePoint Document Library for Excel file uploads
- ✅ Power Automate for automated import (not manual)
- ✅ Enable concurrency for faster processing
- ✅ Index SharePoint columns for performance
- ✅ PowerApps should NEVER load all 10k rows
- ✅ Always filter by date range first
- ✅ Monitor and optimize regularly

**Expected Performance:**
- Import: 10-20 minutes (automated)
- Report generation: 3-5 seconds (filtered data)
- User experience: Fast and responsive
