# PowerApps Absence Viewer - Sample Data Guide

## Table of Contents
1. [CSV File Formats](#csv-file-formats)
2. [Sample Absence Data](#sample-absence-data)
3. [Sample Organization Data](#sample-organization-data)
4. [Sample Holiday Calendar](#sample-holiday-calendar)
5. [SharePoint List Import](#sharepoint-list-import)
6. [Data Validation Rules](#data-validation-rules)

---

## CSV File Formats

### General CSV Guidelines

- **Encoding:** UTF-8 (without BOM)
- **Delimiter:** Comma (,)
- **Quote character:** Double quote (")
- **Line endings:** CRLF (Windows) or LF (Unix)
- **Header row:** Required (first row must contain column names)
- **Date format:** YYYY-MM-DD (ISO 8601)
- **Decimal separator:** Period (.)

### Special Characters

If a field contains commas or quotes, wrap it in double quotes:
```csv
"Smith, Jr., John","Comment with, comma","Normal field"
```

If a field contains a double quote, escape it with another double quote:
```csv
"She said ""Hello""","Normal field"
```

---

## Sample Absence Data

### CSV Structure: Absence_Data.csv

**Columns (in this exact order):**

| Column Name | Data Type | Format | Required | Example |
|------------|-----------|--------|----------|---------|
| LAST NAME | Text | - | Yes | Smith |
| FIRST NAME | Text | - | Yes | John |
| DEPARTMENT | Text | XXX-ENV-### | Yes | ABC-ENV-123 |
| LINE MANAGER | Text | - | Yes | Doe, Jane |
| ABSENCE START DATE | Date | YYYY-MM-DD | Yes | 2026-01-06 |
| ABSENCE END DATE | Date | YYYY-MM-DD | Yes | 2026-01-10 |
| ABSENCE DURATION | Number | Decimal | Yes | 40 |
| UOM | Text | H or D | Yes | H |
| ABSENCE TYPE | Text | - | Yes | Annual Leave |
| ABSENCE STATUS | Text | - | Yes | Scheduled |
| ABSENCE COMMENT | Text | - | No | Family vacation |
| NORMAL WORKING HOURS | Number | Decimal | No | 40 |

### Sample File Content

```csv
LAST NAME,FIRST NAME,DEPARTMENT,LINE MANAGER,ABSENCE START DATE,ABSENCE END DATE,ABSENCE DURATION,UOM,ABSENCE TYPE,ABSENCE STATUS,ABSENCE COMMENT,NORMAL WORKING HOURS
Smith,John,ABC-ENV-123,Doe, Jane,2026-01-06,2026-01-10,40,H,Annual Leave,Scheduled,Family vacation,40
Johnson,Mary,ABC-ENV-123,Doe, Jane,2026-01-08,2026-01-09,16,H,Sick Leave,Scheduled,Flu symptoms,40
Williams,Robert,DEF-ENV-456,Brown, Tom,2026-01-13,2026-01-17,5,D,Annual Leave,Scheduled,Long weekend trip,40
Jones,Emily,DEF-ENV-456,Brown, Tom,2026-01-15,2026-01-15,8,H,Personal Leave,Awaiting approval,Doctor appointment,40
Davis,Michael,GHI-ENV-789,Wilson, Sarah,2026-02-03,2026-02-07,40,H,Annual Leave,Scheduled,Ski trip,40
Miller,Jennifer,GHI-ENV-789,Wilson, Sarah,2026-02-10,2026-02-14,32,H,Sick Leave,Completed,Back pain,32
Wilson,David,JKL-ENV-012,Taylor, Mark,2026-02-17,2026-02-21,40,H,Training,Scheduled,Leadership training,40
Moore,Jessica,JKL-ENV-012,Taylor, Mark,2026-02-18,2026-02-19,1.5,D,Personal Leave,Saved,Parent-teacher meeting,40
Taylor,James,MNO-ENV-345,Anderson, Lisa,2026-03-03,2026-03-07,40,H,Annual Leave,Scheduled,Spring break,40
Anderson,Sarah,MNO-ENV-345,Anderson, Lisa,2026-03-10,2026-03-11,2,D,Sick Leave,In progress,Migraine,37.5
Thomas,Daniel,PQR-ENV-678,Garcia, Carlos,2026-03-17,2026-03-21,40,H,Bereavement,Scheduled,Family funeral,40
Jackson,Lisa,PQR-ENV-678,Garcia, Carlos,2026-03-20,2026-03-20,8,H,Personal Leave,Scheduled,Car repair,40
White,Matthew,STU-ENV-901,Martinez, Rosa,2026-04-01,2026-04-02,16,H,Sick Leave,Completed,Cold,40
Harris,Ashley,STU-ENV-901,Martinez, Rosa,2026-04-07,2026-04-11,5,D,Annual Leave,Scheduled,Easter holiday,40
Martin,Christopher,VWX-ENV-234,Lee, Kevin,2026-04-14,2026-04-18,40,H,Annual Leave,Awaiting approval,Beach vacation,40
Thompson,Amanda,VWX-ENV-234,Lee, Kevin,2026-04-21,2026-04-22,12,H,Personal Leave,Scheduled,Home renovation,40
Garcia,Joshua,YZA-ENV-567,Rodriguez, Maria,2026-05-05,2026-05-09,40,H,Annual Leave,Scheduled,Visit family,40
Martinez,Melissa,YZA-ENV-567,Rodriguez, Maria,2026-05-12,2026-05-13,1.5,D,Sick Leave,Scheduled,Dental surgery,40
Robinson,Andrew,ABC-ENV-123,Doe, Jane,2026-05-19,2026-05-23,32,H,Training,Scheduled,Technical certification,40
Clark,Stephanie,ABC-ENV-123,Doe, Jane,2026-05-26,2026-05-30,40,H,Annual Leave,Scheduled,Memorial Day weekend,40
Rodriguez,Brian,DEF-ENV-456,Brown, Tom,2026-06-02,2026-06-06,5,D,Annual Leave,Scheduled,Camping trip,40
Lewis,Nicole,DEF-ENV-456,Brown, Tom,2026-06-09,2026-06-10,16,H,Personal Leave,Scheduled,Wedding,40
Lee,Ryan,GHI-ENV-789,Wilson, Sarah,2026-06-16,2026-06-20,40,H,Annual Leave,Withdrawn,Plans changed,40
Walker,Brittany,GHI-ENV-789,Wilson, Sarah,2026-06-23,2026-06-24,12,H,Sick Leave,Scheduled,Stomach bug,40
Hall,Kevin,JKL-ENV-012,Taylor, Mark,2026-07-01,2026-07-03,24,H,Annual Leave,Scheduled,Independence Day,40
Allen,Rachel,JKL-ENV-012,Taylor, Mark,2026-07-07,2026-07-11,40,H,Annual Leave,Denied,Coverage needed,40
Young,Justin,MNO-ENV-345,Anderson, Lisa,2026-07-14,2026-07-18,5,D,Annual Leave,Scheduled,Summer vacation,40
Hernandez,Lauren,MNO-ENV-345,Anderson, Lisa,2026-07-21,2026-07-25,40,H,Maternity/Paternity,Scheduled,Maternity leave,40
King,Tyler,PQR-ENV-678,Garcia, Carlos,2026-07-28,2026-08-01,40,H,Annual Leave,Scheduled,Music festival,40
Wright,Megan,PQR-ENV-678,Garcia, Carlos,2026-08-04,2026-08-05,1,D,Personal Leave,Scheduled,Moving day,40
```

### Validation Rules

- **LAST NAME, FIRST NAME:** Not blank
- **DEPARTMENT:** Format XXX-ENV-### where XXX is 3 letters, ENV is literal, ### is 3 digits
- **ABSENCE START DATE <= ABSENCE END DATE**
- **ABSENCE DURATION > 0**
- **UOM:** Must be "H" or "D"
- **ABSENCE STATUS:** One of: Scheduled, Awaiting approval, Awaiting withdrawl approval, Completed, In progress, Saved, Withdrawn, Denied
- **NORMAL WORKING HOURS:** If blank, defaults to 40

---

## Sample Organization Data

### CSV Structure: Organization_Data.csv

**Columns:**

| Column Name | Data Type | Format | Required | Example |
|------------|-----------|--------|----------|---------|
| OFFICELOCATION_ID | Text | First 8 chars of dept | Yes | ABC-ENV |
| MARKETSUBSECTOR_NAME | Text | - | Yes | North America - Technology |

### Sample File Content

```csv
OFFICELOCATION_ID,MARKETSUBSECTOR_NAME
ABC-ENV,North America - Technology
DEF-ENV,North America - Manufacturing
GHI-ENV,Europe - Technology
JKL-ENV,Europe - Manufacturing
MNO-ENV,Asia Pacific - Technology
PQR-ENV,Asia Pacific - Manufacturing
STU-ENV,Latin America - Technology
VWX-ENV,Latin America - Manufacturing
YZA-ENV,Middle East - Technology
```

### Notes

- **OFFICELOCATION_ID** must match the first 8 characters of departments in Absence Data
- Each department prefix should map to exactly one market sub-sector
- If a department is not in this mapping, it will show as "Unknown" in reports

---

## Sample Holiday Calendar

### CSV Structure: Holiday_Calendar.csv

**Columns:**

| Column Name | Data Type | Format | Required | Example |
|------------|-----------|--------|----------|---------|
| DATE | Date | YYYY-MM-DD | Yes | 2026-01-01 |
| HOLIDAY | Text | - | Yes | New Year's Day |
| APPLIES-TO | Text | - | Yes | All Employees |

### Sample File Content

```csv
DATE,HOLIDAY,APPLIES-TO
2026-01-01,New Year's Day,All Employees
2026-01-19,Martin Luther King Jr. Day,US Employees
2026-02-16,Presidents Day,US Employees
2026-04-03,Good Friday,All Employees
2026-04-06,Easter Monday,European Employees
2026-05-25,Memorial Day,US Employees
2026-07-03,Independence Day (Observed),US Employees
2026-07-04,Independence Day,US Employees
2026-09-07,Labor Day,US Employees
2026-10-12,Columbus Day,US Employees
2026-11-11,Veterans Day,US Employees
2026-11-26,Thanksgiving Day,US Employees
2026-11-27,Day after Thanksgiving,US Employees
2026-12-24,Christmas Eve,All Employees
2026-12-25,Christmas Day,All Employees
2026-12-31,New Year's Eve,All Employees
2026-04-01,April Fool's Day,No One (Example - not a real holiday)
2026-05-04,Star Wars Day,Nerds Only (Example)
2026-12-26,Boxing Day,UK/Canadian Employees
```

### Notes

- Holidays can apply to different employee groups (use APPLIES-TO to filter)
- Multiple holidays can occur in the same week
- Holidays falling on weekends may have "observed" dates on Friday/Monday

---

## SharePoint List Import

### Importing CSV to SharePoint Lists

#### Method 1: SharePoint UI

1. Navigate to your SharePoint list
2. Click "Integrate" → "From CSV"
3. Select your CSV file
4. Map columns (should auto-detect if headers match)
5. Click "Import"

#### Method 2: PowerShell

```powershell
# Install PnP PowerShell if not already installed
Install-Module -Name PnP.PowerShell

# Connect to site
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/AbsenceManagement" -Interactive

# Import Absence Data
$absenceData = Import-Csv -Path "C:\path\to\Absence_Data.csv"

foreach ($row in $absenceData) {
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
        "NormalWorkingHours" = if ($row.'NORMAL WORKING HOURS') { [decimal]$row.'NORMAL WORKING HOURS' } else { 40 }
    }
}

# Import Organization Data
$orgData = Import-Csv -Path "C:\path\to\Organization_Data.csv"

foreach ($row in $orgData) {
    Add-PnPListItem -List "OrganizationData" -Values @{
        "OfficeLocationID" = $row.'OFFICELOCATION_ID'
        "MarketSubSectorName" = $row.'MARKETSUBSECTOR_NAME'
    }
}

# Import Holiday Calendar
$holidays = Import-Csv -Path "C:\path\to\Holiday_Calendar.csv"

foreach ($row in $holidays) {
    Add-PnPListItem -List "HolidayCalendar" -Values @{
        "Title" = $row.'HOLIDAY'
        "HolidayDate" = [datetime]$row.'DATE'
        "Holiday" = $row.'HOLIDAY'
        "AppliesTo" = $row.'APPLIES-TO'
    }
}

Write-Host "Import complete!"
```

---

## Data Validation Rules

### Absence Data Validation

**Check 1: Date Range Validity**
```powerFx
// In PowerApps or SharePoint validation
ABSENCE END DATE >= ABSENCE START DATE
```

**Check 2: Department Format**
```regex
^[A-Z]{3}-ENV-[0-9]{3}$
```

**Check 3: Duration Positive**
```powerFx
ABSENCE DURATION > 0
```

**Check 4: Valid UOM**
```powerFx
UOM = "H" || UOM = "D"
```

**Check 5: Valid Status**
```powerFx
ABSENCE STATUS in [
    "Scheduled",
    "Awaiting approval",
    "Awaiting withdrawl approval",
    "Completed",
    "In progress",
    "Saved",
    "Withdrawn",
    "Denied"
]
```

### Organization Data Validation

**Check 1: Office Location ID Length**
```powerFx
LEN(OFFICELOCATION_ID) = 7 or 8
```

**Check 2: Unique Office Locations**
```sql
-- No duplicates
SELECT OFFICELOCATION_ID, COUNT(*)
FROM OrganizationData
GROUP BY OFFICELOCATION_ID
HAVING COUNT(*) > 1
```

### Holiday Calendar Validation

**Check 1: Valid Date**
```powerFx
YEAR(DATE) >= 2026
```

**Check 2: Unique Holiday per Date**
```sql
-- Same date can have different holidays for different groups
SELECT DATE, HOLIDAY, COUNT(*)
FROM HolidayCalendar
GROUP BY DATE, HOLIDAY
HAVING COUNT(*) > 1
```

---

## Generating Test Data

### PowerShell Script to Generate Random Test Data

```powershell
# Generate 100 random absence records

$firstNames = @("John", "Mary", "Robert", "Emily", "Michael", "Jennifer", "David", "Jessica", "James", "Sarah", "Daniel", "Lisa", "Matthew", "Ashley", "Christopher", "Amanda", "Joshua", "Melissa", "Andrew", "Stephanie")
$lastNames = @("Smith", "Johnson", "Williams", "Jones", "Davis", "Miller", "Wilson", "Moore", "Taylor", "Anderson", "Thomas", "Jackson", "White", "Harris", "Martin", "Thompson", "Garcia", "Martinez", "Robinson", "Clark")
$departments = @("ABC-ENV-123", "DEF-ENV-456", "GHI-ENV-789", "JKL-ENV-012", "MNO-ENV-345", "PQR-ENV-678", "STU-ENV-901", "VWX-ENV-234", "YZA-ENV-567")
$managers = @("Doe, Jane", "Brown, Tom", "Wilson, Sarah", "Taylor, Mark", "Anderson, Lisa", "Garcia, Carlos", "Martinez, Rosa", "Lee, Kevin", "Rodriguez, Maria")
$absenceTypes = @("Annual Leave", "Sick Leave", "Personal Leave", "Training", "Bereavement", "Maternity/Paternity")
$statuses = @("Scheduled", "Awaiting approval", "Completed", "In progress", "Saved")

$csv = "LAST NAME,FIRST NAME,DEPARTMENT,LINE MANAGER,ABSENCE START DATE,ABSENCE END DATE,ABSENCE DURATION,UOM,ABSENCE TYPE,ABSENCE STATUS,ABSENCE COMMENT,NORMAL WORKING HOURS`n"

for ($i = 0; $i -lt 100; $i++) {
    $lastName = $lastNames | Get-Random
    $firstName = $firstNames | Get-Random
    $dept = $departments | Get-Random
    $manager = $managers | Get-Random
    $startDate = (Get-Date "2026-01-01").AddDays((Get-Random -Minimum 0 -Maximum 365))
    $duration = Get-Random -Minimum 1 -Maximum 6  # 1-5 days
    $endDate = $startDate.AddDays($duration - 1)
    $uom = if ((Get-Random -Minimum 0 -Maximum 2) -eq 0) { "H" } else { "D" }
    $absenceDuration = if ($uom -eq "H") { $duration * 8 } else { $duration }
    $absenceType = $absenceTypes | Get-Random
    $status = $statuses | Get-Random
    $comment = "Generated test data"
    $normalHours = 40

    $csv += "$lastName,$firstName,$dept,$manager," + $startDate.ToString("yyyy-MM-dd") + "," + $endDate.ToString("yyyy-MM-dd") + ",$absenceDuration,$uom,$absenceType,$status,$comment,$normalHours`n"
}

$csv | Out-File -FilePath "Generated_Absence_Data.csv" -Encoding UTF8

Write-Host "Generated 100 test absence records in Generated_Absence_Data.csv"
```

---

## Data Quality Checks

Before importing, run these checks:

### Check 1: No Missing Required Fields
```powershell
Import-Csv "Absence_Data.csv" | Where-Object {
    [string]::IsNullOrWhiteSpace($_.'LAST NAME') -or
    [string]::IsNullOrWhiteSpace($_.'FIRST NAME') -or
    [string]::IsNullOrWhiteSpace($_.'DEPARTMENT')
} | Format-Table
```

### Check 2: Valid Date Ranges
```powershell
Import-Csv "Absence_Data.csv" | Where-Object {
    [datetime]$_.'ABSENCE END DATE' -lt [datetime]$_.'ABSENCE START DATE'
} | Format-Table
```

### Check 3: Department Mappings Exist
```powershell
$absences = Import-Csv "Absence_Data.csv"
$orgs = Import-Csv "Organization_Data.csv"

$absences | ForEach-Object {
    $deptPrefix = $_.DEPARTMENT.Substring(0, 7)
    if ($orgs.'OFFICELOCATION_ID' -notcontains $deptPrefix) {
        Write-Warning "Missing org mapping for: $deptPrefix"
    }
}
```

### Check 4: Duplicate Records
```powershell
Import-Csv "Absence_Data.csv" | Group-Object 'LAST NAME', 'FIRST NAME', 'ABSENCE START DATE' | Where-Object { $_.Count -gt 1 } | ForEach-Object {
    Write-Warning "Duplicate found: $($_.Name)"
}
```

---

## Appendix: Complete Sample Dataset

### Download Links

Create these files in your repository or SharePoint:

1. **Absence_Data_Sample.csv** - 30 records covering all scenarios
2. **Organization_Data_Sample.csv** - 9 market mappings
3. **Holiday_Calendar_Sample.csv** - 16 holidays for 2026

### Quick Import Instructions

1. Download all 3 sample CSV files
2. Navigate to your SharePoint site
3. For each list (AbsenceData, OrganizationData, HolidayCalendar):
   - Click "Integrate" → "From CSV"
   - Select corresponding CSV file
   - Verify column mappings
   - Click "Import"
4. Refresh your PowerApp
5. Click "Generate Weekly Report" to test

---

## Troubleshooting Data Import

### Issue: "Date format not recognized"

**Cause:** CSV using different date format
**Solution:** Ensure dates are in YYYY-MM-DD format

```powershell
# Convert dates in CSV
$data = Import-Csv "Absence_Data.csv"
$data | ForEach-Object {
    $_.'ABSENCE START DATE' = ([datetime]::ParseExact($_.'ABSENCE START DATE', 'M/d/yyyy', $null)).ToString('yyyy-MM-dd')
    $_.'ABSENCE END DATE' = ([datetime]::ParseExact($_.'ABSENCE END DATE', 'M/d/yyyy', $null)).ToString('yyyy-MM-dd')
}
$data | Export-Csv "Absence_Data_Fixed.csv" -NoTypeInformation
```

### Issue: "Special characters corrupted"

**Cause:** Incorrect file encoding
**Solution:** Save CSV as UTF-8

```powershell
# Re-save with UTF-8 encoding
Get-Content "Absence_Data.csv" | Out-File "Absence_Data_UTF8.csv" -Encoding UTF8
```

### Issue: "Decimal numbers importing as text"

**Cause:** Excel formatting or locale settings
**Solution:** Ensure decimal separator is period (.)

```powershell
# Fix decimal separators
(Get-Content "Absence_Data.csv") -replace ',(\d+),', '.$1,' | Set-Content "Absence_Data_Fixed.csv"
```

---

## Next Steps

1. ✅ Review sample data formats
2. ✅ Create your own test CSV files or use samples
3. ✅ Import data to SharePoint lists
4. ✅ Verify data in SharePoint
5. ➡️ Continue to **POWERAPP_STEP_BY_STEP.md** to build the PowerApp
