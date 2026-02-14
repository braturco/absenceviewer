# PowerApps Absence Viewer - Data Setup Guide

## Table of Contents
1. [SharePoint Lists Setup (Recommended)](#sharepoint-lists-setup)
2. [SharePoint Document Library Setup (Alternative)](#sharepoint-document-library-setup)
3. [Data Connections in PowerApps](#data-connections-in-powerapps)
4. [Initial Data Load](#initial-data-load)

---

## SharePoint Lists Setup (Recommended)

### Overview
Instead of CSV files, create 3 SharePoint lists to store your data. This provides:
- Direct data connection to PowerApps
- Better performance (no parsing needed)
- Easy data maintenance
- Better security and permissions

### List 1: Absence Data

**List Name:** `AbsenceData`

**Columns:**

| Column Name | Type | Settings | Example |
|------------|------|----------|---------|
| Title | Single line of text | (default, will use as ID) | Auto-number |
| LastName | Single line of text | Required | Smith |
| FirstName | Single line of text | Required | John |
| Department | Single line of text | Required | ABC-ENV-123 |
| LineManager | Single line of text | Required | Jane Doe |
| AbsenceStartDate | Date and time | Date only, Required | 1/6/2026 |
| AbsenceEndDate | Date and time | Date only, Required | 1/10/2026 |
| AbsenceDuration | Number | Decimal, Required | 40 |
| UOM | Choice | H, D (default H) | H |
| AbsenceType | Choice | See choices below | Annual Leave |
| AbsenceStatus | Choice | See choices below | Scheduled |
| AbsenceComment | Multiple lines of text | | Family vacation |
| NormalWorkingHours | Number | Decimal, Default 40 | 40 |

**AbsenceType Choices:**
- Annual Leave
- Sick Leave
- Personal Leave
- Bereavement
- Maternity/Paternity
- Training
- Other

**AbsenceStatus Choices:**
- Scheduled
- Awaiting approval
- Awaiting withdrawl approval
- Completed
- In progress
- Saved
- Withdrawn
- Denied

**SharePoint PowerShell Commands to Create List:**
```powershell
# Connect to SharePoint site
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/yoursite" -Interactive

# Create list
New-PnPList -Title "AbsenceData" -Template GenericList

# Add columns
Add-PnPField -List "AbsenceData" -DisplayName "LastName" -InternalName "LastName" -Type Text -Required
Add-PnPField -List "AbsenceData" -DisplayName "FirstName" -InternalName "FirstName" -Type Text -Required
Add-PnPField -List "AbsenceData" -DisplayName "Department" -InternalName "Department" -Type Text -Required
Add-PnPField -List "AbsenceData" -DisplayName "LineManager" -InternalName "LineManager" -Type Text -Required
Add-PnPField -List "AbsenceData" -DisplayName "AbsenceStartDate" -InternalName "AbsenceStartDate" -Type DateTime -DisplayFormat DateOnly -Required
Add-PnPField -List "AbsenceData" -DisplayName "AbsenceEndDate" -InternalName "AbsenceEndDate" -Type DateTime -DisplayFormat DateOnly -Required
Add-PnPField -List "AbsenceData" -DisplayName "AbsenceDuration" -InternalName "AbsenceDuration" -Type Number -Required
Add-PnPField -List "AbsenceData" -DisplayName "NormalWorkingHours" -InternalName "NormalWorkingHours" -Type Number

# Add choice fields
Add-PnPFieldFromXml -List "AbsenceData" -FieldXml '<Field Type="Choice" DisplayName="UOM" Name="UOM" Required="TRUE"><Default>H</Default><CHOICES><CHOICE>H</CHOICE><CHOICE>D</CHOICE></CHOICES></Field>'

Add-PnPFieldFromXml -List "AbsenceData" -FieldXml '<Field Type="Choice" DisplayName="AbsenceType" Name="AbsenceType"><CHOICES><CHOICE>Annual Leave</CHOICE><CHOICE>Sick Leave</CHOICE><CHOICE>Personal Leave</CHOICE><CHOICE>Bereavement</CHOICE><CHOICE>Maternity/Paternity</CHOICE><CHOICE>Training</CHOICE><CHOICE>Other</CHOICE></CHOICES></Field>'

Add-PnPFieldFromXml -List "AbsenceData" -FieldXml '<Field Type="Choice" DisplayName="AbsenceStatus" Name="AbsenceStatus"><CHOICES><CHOICE>Scheduled</CHOICE><CHOICE>Awaiting approval</CHOICE><CHOICE>Awaiting withdrawl approval</CHOICE><CHOICE>Completed</CHOICE><CHOICE>In progress</CHOICE><CHOICE>Saved</CHOICE><CHOICE>Withdrawn</CHOICE><CHOICE>Denied</CHOICE></CHOICES></Field>'

Add-PnPField -List "AbsenceData" -DisplayName "AbsenceComment" -InternalName "AbsenceComment" -Type Note
```

### List 2: Organization Data

**List Name:** `OrganizationData`

**Columns:**

| Column Name | Type | Settings | Example |
|------------|------|----------|---------|
| Title | Single line of text | (default, use as OfficeLocationID) | ABC-ENV |
| OfficeLocationID | Single line of text | Required, Indexed | ABC-ENV |
| MarketSubSectorName | Single line of text | Required | North America - Technology |

**SharePoint PowerShell Commands:**
```powershell
New-PnPList -Title "OrganizationData" -Template GenericList

Add-PnPField -List "OrganizationData" -DisplayName "OfficeLocationID" -InternalName "OfficeLocationID" -Type Text -Required
Set-PnPField -List "OrganizationData" -Identity "OfficeLocationID" -Values @{Indexed=$true}

Add-PnPField -List "OrganizationData" -DisplayName "MarketSubSectorName" -InternalName "MarketSubSectorName" -Type Text -Required
```

**Manual Setup:**
1. Navigate to your SharePoint site
2. Click "New" → "List"
3. Choose "Blank list", name it "OrganizationData"
4. Add columns as shown in table above
5. Make OfficeLocationID indexed (List Settings → Indexed Columns)

### List 3: Holiday Calendar

**List Name:** `HolidayCalendar`

**Columns:**

| Column Name | Type | Settings | Example |
|------------|------|----------|---------|
| Title | Single line of text | (default, use as holiday name) | New Year's Day |
| HolidayDate | Date and time | Date only, Required | 1/1/2026 |
| Holiday | Single line of text | Required | New Year's Day |
| AppliesTo | Single line of text | Required | All Employees |

**SharePoint PowerShell Commands:**
```powershell
New-PnPList -Title "HolidayCalendar" -Template GenericList

Add-PnPField -List "HolidayCalendar" -DisplayName "HolidayDate" -InternalName "HolidayDate" -Type DateTime -DisplayFormat DateOnly -Required
Add-PnPField -List "HolidayCalendar" -DisplayName "Holiday" -InternalName "Holiday" -Type Text -Required
Add-PnPField -List "HolidayCalendar" -DisplayName "AppliesTo" -InternalName "AppliesTo" -Type Text -Required
```

---

## SharePoint Document Library Setup (Alternative)

If you prefer to keep using CSV files:

### Create Document Library

1. Navigate to your SharePoint site
2. Click "New" → "Document Library"
3. Name it "AbsenceDataFiles"
4. Create 3 folders:
   - AbsenceData
   - OrganizationData
   - HolidayData

### File Naming Convention

Use consistent file names that PowerApps can reference:
- `Absence_Data_Latest.csv`
- `Organization_Data_Latest.csv`
- `Holiday_Calendar_Latest.csv`

### Metadata Columns (Optional)

Add columns to track file versions:
- UploadDate (Date/Time)
- UploadedBy (Person)
- RecordCount (Number)
- Status (Choice: Active, Archived)

---

## Data Connections in PowerApps

### Connecting to SharePoint Lists

**Step 1: Add Data Source**
1. In PowerApps Studio, click "Data" in left panel
2. Click "Add data"
3. Search for "SharePoint"
4. Select "SharePoint" connector

**Step 2: Connect to Lists**
1. Enter your SharePoint site URL: `https://yourtenant.sharepoint.com/sites/yoursite`
2. Click "Connect"
3. Select all three lists:
   - AbsenceData
   - OrganizationData
   - HolidayCalendar
4. Click "Connect"

**Result:** Lists appear in your Data sources panel and can be referenced in formulas

### Connecting to Document Library (for CSV files)

**Step 1: Add SharePoint Connection**
1. Click "Data" → "Add data"
2. Select "SharePoint"
3. Enter site URL
4. Select "AbsenceDataFiles" library
5. Click "Connect"

**Step 2: File Access**
- Use Power Automate flow to read CSV files (see POWERAPP_POWER_AUTOMATE.md)
- PowerApps cannot directly parse CSV files
- Flow will parse and return data to PowerApps

---

## Initial Data Load

### OnStart Formula (For SharePoint Lists)

Place this in your App's **OnStart** property:

```powerFx
// Clear existing collections
Clear(colAbsenceData);
Clear(colOrgData);
Clear(colHolidayData);

// Load Absence Data (filter out Withdrawn and Denied)
ClearCollect(
    colAbsenceData,
    Filter(
        AbsenceData,
        AbsenceStatus <> "Withdrawn" && AbsenceStatus <> "Denied"
    )
);

// Load Organization Data
ClearCollect(
    colOrgData,
    OrganizationData
);

// Load Holiday Calendar
ClearCollect(
    colHolidayData,
    HolidayCalendar
);

// Set global variables
Set(varDataLoaded, true);
Set(varRecordCount, CountRows(colAbsenceData));
Set(varStartDate, Date(2026, 1, 2)); // Default: Week 1 ends Jan 2, 2026
Set(varEndDate, Date(2026, 12, 31)); // Default: End of 2026
```

### Testing Data Load

**Add a Label to your screen:**
- Text: `"Loaded " & varRecordCount & " absence records"`

**Add buttons to test:**
- "Refresh Data" button:
  ```powerFx
  OnSelect =
  Refresh(AbsenceData);
  Refresh(OrganizationData);
  Refresh(HolidayCalendar);
  // Re-run OnStart
  Set(varDataLoaded, false);
  Set(varDataLoaded, true);
  ```

---

## Data Migration from CSV to SharePoint Lists

If you have existing CSV files and want to migrate to SharePoint lists:

### Option 1: Manual Import via SharePoint

1. Open your SharePoint list
2. Click "Integrate" → "Excel"
3. Click "Import from CSV"
4. Select your CSV file
5. Map columns
6. Click "Import"

### Option 2: Power Automate Flow

Create a flow that:
1. Triggers when CSV file is uploaded to document library
2. Parses CSV content
3. Creates items in SharePoint list

**Flow Template:**
```
Trigger: When a file is created (SharePoint)
├── Get file content (SharePoint)
├── Parse CSV (Data Operations - Parse JSON or Excel)
├── Apply to each row
│   └── Create item (SharePoint)
└── Send email notification (optional)
```

### Option 3: PowerShell Script

```powershell
# Connect to SharePoint
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/yoursite" -Interactive

# Read CSV
$csvData = Import-Csv -Path "C:\path\to\Absence_Data.csv"

# Import to list
foreach ($row in $csvData) {
    Add-PnPListItem -List "AbsenceData" -Values @{
        "LastName" = $row.'LAST NAME'
        "FirstName" = $row.'FIRST NAME'
        "Department" = $row.'DEPARTMENT'
        "LineManager" = $row.'LINE MANAGER'
        "AbsenceStartDate" = [datetime]$row.'ABSENCE START DATE'
        "AbsenceEndDate" = [datetime]$row.'ABSENCE END DATE'
        "AbsenceDuration" = [decimal]$row.'ABSENCE DURATION'
        "UOM" = $row.'UOM'
        "AbsenceType" = $row.'ABSENCE TYPE'
        "AbsenceStatus" = $row.'ABSENCE STATUS'
        "AbsenceComment" = $row.'ABSENCE COMMENT'
        "NormalWorkingHours" = [decimal]$row.'NORMAL WORKING HOURS'
    }
}
```

---

## Data Validation & Business Rules

### SharePoint Column Validation

**AbsenceData - End Date Must Be After Start Date:**
```
=[AbsenceEndDate]>=[AbsenceStartDate]
```
Validation Message: "End date must be on or after start date"

**AbsenceData - Duration Must Be Positive:**
```
=[AbsenceDuration]>0
```
Validation Message: "Absence duration must be greater than 0"

**OrganizationData - Office Location ID Format:**
```
=LEN([OfficeLocationID])=8
```
Validation Message: "Office Location ID must be exactly 8 characters (e.g., ABC-ENV)"

### PowerApps Validation

Add validation in PowerApps when users edit data:

```powerFx
// In your Edit Form or Input fields
If(
    DateValue(AbsenceEndDate_Input.Text) < DateValue(AbsenceStartDate_Input.Text),
    Notify("End date must be after start date", NotificationType.Error),
    // Proceed with save
    SubmitForm(AbsenceForm)
)
```

---

## Performance Considerations

### Indexing

**CRITICAL for 10,000+ records** - Add indexes to frequently filtered columns:

**AbsenceData:**
- AbsenceStatus (indexed) - Required
- Department (indexed) - Required
- AbsenceStartDate (indexed) - **Required for date filtering**
- AbsenceEndDate (indexed) - **Required for date filtering**

**Why This Matters:**
- Without indexes, queries on 10k+ rows will timeout
- Indexes enable SharePoint to pre-sort data for fast retrieval
- PowerApps delegation works better with indexed columns

**How to Add Indexes:**
1. Go to List Settings
2. Click "Indexed columns"
3. Click "Create a new index"
4. Select column
5. Save

**Create Compound Indexes (Advanced):**
1. Create index on AbsenceStartDate
2. Then add secondary column: AbsenceStatus
3. This optimizes the common filter: StartDate + Status

### Delegation Limits

PowerApps has a 2000 record delegation limit for some operations. To work around this:

**Use ClearCollect instead of Filter in formulas:**
```powerFx
// This loads ALL data into memory, bypassing delegation limit
ClearCollect(colAbsenceData, AbsenceData);

// Then filter the collection (no delegation limit)
Filter(colAbsenceData, AbsenceStatus = "Scheduled")
```

**Break large operations into chunks:**
```powerFx
// Load data in batches if you have >2000 records
ClearCollect(colBatch1, FirstN(AbsenceData, 2000));
Collect(colBatch2, FirstN(Filter(AbsenceData, ID > 2000), 2000));
// Combine batches
ClearCollect(colAllData, colBatch1);
Collect(colAllData, colBatch2);
```

---

## Security & Permissions

### SharePoint List Permissions

Set up appropriate permissions for your lists:

**Recommended Permission Levels:**

| User Role | AbsenceData | OrganizationData | HolidayCalendar |
|-----------|-------------|------------------|-----------------|
| HR Admin | Full Control | Full Control | Full Control |
| Manager | Edit (own items) | Read | Read |
| Employee | Read (own items) | Read | Read |
| Viewer | Read | Read | Read |

**How to Set:**
1. Go to List Settings
2. Click "Permissions for this list"
3. Click "Stop Inheriting Permissions"
4. Remove existing groups as needed
5. Click "Grant Permissions"
6. Add users/groups with appropriate permissions

### PowerApps Security

**Row-level security in PowerApps:**

Filter data based on current user:
```powerFx
Filter(
    colAbsenceData,
    Department = LookUp(
        colEmployeeDept,
        Email = User().Email,
        Department
    )
)
```

---

## Next Steps

1. ✅ Choose SharePoint Lists or Document Library approach
2. ✅ Create SharePoint lists/library using instructions above
3. ✅ Import sample data for testing
4. ✅ Connect PowerApps to SharePoint
5. ✅ Test data load with OnStart formula
6. ➡️ Continue to **POWERAPP_FORMULAS.md** for calculation logic

---

## Troubleshooting

### Common Issues

**Issue: "Delegation warning" in PowerApps**
- Solution: Use ClearCollect to load data into collections first

**Issue: "Cannot find data source"**
- Solution: Re-add SharePoint connection in Data panel

**Issue: "Access denied" when connecting**
- Solution: Ensure you have Read permissions on SharePoint lists

**Issue: CSV import fails**
- Solution: Check CSV file encoding (UTF-8), remove special characters

**Issue: Dates importing incorrectly**
- Solution: Use ISO format (YYYY-MM-DD) in CSV files
