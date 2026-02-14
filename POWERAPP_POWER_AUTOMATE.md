# PowerApps Absence Viewer - Power Automate Flows

## Table of Contents
1. [Overview](#overview)
2. [Flow 1: Parse CSV Files](#flow-1-parse-csv-files)
3. [Flow 2: Export to Excel (Weekly Report)](#flow-2-export-to-excel-weekly-report)
4. [Flow 3: Export to Excel (Daily Report)](#flow-3-export-to-excel-daily-report)
5. [Testing Flows](#testing-flows)
6. [Error Handling](#error-handling)

---

## Overview

### Power Automate Flows for This Solution

1. **CSV Parser Flow** (optional) - Parses uploaded CSV files and populates SharePoint lists
2. **Weekly Excel Export Flow** - Creates formatted Excel workbook from weekly report data
3. **Daily Excel Export Flow** - Creates formatted Excel workbook from daily report data

### Prerequisites

- Power Automate license (included with PowerApps)
- Access to SharePoint site
- OneDrive for Business or SharePoint document library for output files

---

## Flow 1: Parse CSV Files

### Purpose

Automatically parse CSV files uploaded to SharePoint and populate the three data lists.

### When to Use

- If you want to keep CSV workflow instead of direct SharePoint list entry
- For bulk imports from existing systems
- For scheduled data updates

### Flow Design

#### Trigger

**When a file is created or modified in a folder**
- Site Address: `https://yourtenant.sharepoint.com/sites/AbsenceManagement`
- Folder: `/AbsenceDataFiles`

#### Actions

```
┌─────────────────────────────────────────┐
│ Trigger: File Created/Modified          │
└────────────────┬────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│ Condition: Check file extension         │
│ File Extension = .csv                    │
└────────────────┬────────────────────────┘
                 │ Yes
                 ▼
┌─────────────────────────────────────────┐
│ Get file content                         │
│ File Identifier: File.Id                 │
└────────────────┬────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│ Condition: Check filename                │
│ Contains "Absence_Data"?                 │
└────┬───────────────────────┬────────────┘
     │ Yes                   │ No
     ▼                       ▼
┌──────────────────┐  ┌───────────────────┐
│ Parse Absence    │  │ Check if Org or   │
│ CSV              │  │ Holiday CSV       │
└────┬─────────────┘  └────┬──────────────┘
     │                     │
     ▼                     ▼
┌──────────────────────────────────────────┐
│ Parse CSV (Data Operations)              │
│ Content: File Content                    │
│ Schema: {see below}                      │
└────────────────┬─────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│ Apply to each row                        │
├─────────────────────────────────────────┤
│   Condition: Status not Withdrawn/Denied │
│   ├─ Yes:                                │
│   │   └─ Create item in AbsenceData list│
│   └─ No: Skip                            │
└─────────────────────────────────────────┘
```

### Step-by-Step Setup

#### Step 1: Create Flow

1. Go to [make.powerautomate.com](https://make.powerautomate.com)
2. Click "Create" → "Automated cloud flow"
3. Name: "Parse Absence CSV Files"
4. Trigger: "When a file is created or modified (SharePoint)"
5. Click "Create"

#### Step 2: Configure Trigger

- Site Address: `https://yourtenant.sharepoint.com/sites/AbsenceManagement`
- Library Name: `AbsenceDataFiles`
- Folder: `/`

#### Step 3: Add Condition - Check File Extension

1. Add "Condition" action
2. Choose value: `File name with extension` (from trigger)
3. Condition: `ends with`
4. Value: `.csv`

#### Step 4: Get File Content

In "If yes" branch:

1. Add "Get file content" (SharePoint)
2. Site Address: `Same as trigger`
3. File Identifier: `Identifier` (from trigger)

#### Step 5: Parse CSV

1. Add "Parse CSV" action (Data Operations → Parse CSV) or use "Compose" with expression
2. Content: `File content` (from Get file content)
3. Since there's no built-in CSV parser, use "Create HTML table" then parse:

**Alternative: Use Excel Connector**

1. Add "Create table" (Excel Online)
   - Location: `OneDrive`
   - Document Library: `OneDrive`
   - File: `(create temp Excel file)`
   - Table: Create from CSV

2. Add "List rows present in a table" (Excel Online)
   - File: `Same as above`
   - Table: `Table1`

**Better Alternative: Use HTTP + Parse**

```powerFx
// Expression to split CSV into array
split(body('Get_file_content'), decodeUriComponent('%0A'))
```

#### Step 6: Parse Each Row

1. Add "Apply to each"
2. Select output from previous step
3. Inside loop:
   - Add "Compose" to split row by comma:
     ```
     split(item(), ',')
     ```
   - Add variables for each column:
     ```
     LastName: outputs('Compose_split_row')[0]
     FirstName: outputs('Compose_split_row')[1]
     Department: outputs('Compose_split_row')[2]
     ...
     ```

#### Step 7: Filter and Create Items

1. Add "Condition" to check status:
   - AbsenceStatus not equals "Withdrawn"
   - AND AbsenceStatus not equals "Denied"

2. In "If yes" branch, add "Create item" (SharePoint):
   - Site Address: `Your site`
   - List Name: `AbsenceData`
   - Map fields:
     ```
     Title: auto-generated
     LastName: variables('LastName')
     FirstName: variables('FirstName')
     Department: variables('Department')
     LineManager: variables('LineManager')
     AbsenceStartDate: variables('AbsenceStartDate')
     AbsenceEndDate: variables('AbsenceEndDate')
     AbsenceDuration: float(variables('AbsenceDuration'))
     UOM: variables('UOM')
     AbsenceType: variables('AbsenceType')
     AbsenceStatus: variables('AbsenceStatus')
     AbsenceComment: variables('AbsenceComment')
     NormalWorkingHours: float(variables('NormalWorkingHours'))
     ```

### Simplified Approach: Manual Trigger + Parameter

Instead of file trigger, create a manual flow that accepts CSV content as parameter:

1. Trigger: "PowerApps V2"
2. Add input: "Text" → "CSVContent"
3. Parse and process as above
4. Call from PowerApps:
   ```powerFx
   'ParseCSV'.Run(
       Concat(
           AbsenceDataUpload.Attachments,
           Content,
           "|||"
       )
   )
   ```

---

## Flow 2: Export to Excel (Weekly Report)

### Purpose

Create a formatted Excel workbook from weekly report data with proper styling, grouped by market sub-sector.

### Flow Design

```
┌─────────────────────────────────────────┐
│ Trigger: PowerApps (Manual)              │
│ Inputs:                                  │
│   - ReportData (JSON string)             │
│   - ReportName (string)                  │
└────────────────┬────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│ Parse JSON                               │
│ Content: ReportData                      │
│ Schema: {see below}                      │
└────────────────┬────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│ Create Excel File (OneDrive)            │
│ - Create blank workbook                  │
│ - Add worksheets per market              │
└────────────────┬────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│ Apply to each Market                     │
├─────────────────────────────────────────┤
│   Add worksheet                          │
│   Name: MarketSubSectorName              │
│   Add header rows (month + week)         │
│   Apply to each Department               │
│     Add department row                   │
│     Apply to each Person                 │
│       Add person row with week data      │
│   Add total row                          │
│   Format cells (colors, borders)         │
└────────────────┬────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│ Get file sharing link                    │
│ Link Type: View and edit                │
│ Scope: Organization                      │
└────────────────┬────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────┐
│ Respond to PowerApps                     │
│ Output: FileUrl (string)                 │
└─────────────────────────────────────────┘
```

### Step-by-Step Setup

#### Step 1: Create Flow

1. Create → "Instant cloud flow"
2. Name: "Export Weekly Absence Report to Excel"
3. Trigger: "PowerApps V2"
4. Click "Create"

#### Step 2: Add Input Parameters

1. Click "+ Add an input"
2. Add "Text" input:
   - Name: `ReportData`
   - Description: `JSON string of weekly report data`
3. Add another "Text" input:
   - Name: `ReportName`
   - Description: `Name for the report file`

#### Step 3: Parse JSON Input

1. Add "Parse JSON" (Data Operations)
2. Content: `ReportData` (from PowerApps)
3. Schema: Click "Use sample payload to generate schema"

**Sample JSON structure:**
```json
{
  "markets": [
    {
      "marketName": "North America - Technology",
      "departments": [
        {
          "departmentId": "ABC-ENV-123",
          "persons": [
            {
              "personName": "Smith, John",
              "lineManager": "Doe, Jane",
              "weeks": [
                {"weekNumber": 1, "totalHours": 8, "statuses": "Scheduled"},
                {"weekNumber": 2, "totalHours": 16, "statuses": "Scheduled,Completed"}
              ]
            }
          ]
        }
      ]
    }
  ],
  "weeks": [
    {"weekNumber": 1, "endDate": "2026-01-02", "month": "Jan 2026"},
    {"weekNumber": 2, "endDate": "2026-01-09", "month": "Jan 2026"}
  ]
}
```

#### Step 4: Create Excel File

1. Add "Create file" (OneDrive for Business)
2. Folder Path: `/Absence Reports`
3. File Name: `@{concat(triggerBody()['text'], ' - ', formatDateTime(utcNow(), 'yyyy-MM-dd'), '.xlsx')}`
4. File Content: `(leave blank - will create empty Excel file)`

**Alternative: Use Excel Template**

1. Create template file manually with formatting
2. Use "Copy file" action to duplicate template
3. Then populate with data

#### Step 5: Create Table Structure

Since Excel connector has limitations, use this approach:

1. Add "Create table" (Excel Online)
   - Location: `OneDrive for Business`
   - Document Library: `OneDrive`
   - File: `File Identifier` (from Create file)
   - Table range: `A1:AZ1000` (large enough range)
   - Table has headers: Yes
   - Headers: Build dynamically with expression:
     ```
     concat('Department,Line Manager,Person,', join(body('Parse_JSON')?['weeks']?['weekNumber'], ','), ',Total')
     ```

#### Step 6: Add Rows to Table

This is complex in Power Automate. **Recommended approach:**

**Use "Run script" (Excel Online - Script)**

1. Create Office Script in Excel:

```javascript
function main(workbook: ExcelScript.Workbook, reportData: string) {
  const data = JSON.parse(reportData);

  // Get or create sheet
  let sheet = workbook.getWorksheet("Weekly Report");
  if (!sheet) {
    sheet = workbook.addWorksheet("Weekly Report");
  }

  // Build headers
  let row = 1;
  sheet.getRange("A1").setValue("Department");
  sheet.getRange("B1").setValue("Line Manager");
  sheet.getRange("C1").setValue("Person");

  let col = 4;
  data.weeks.forEach(week => {
    sheet.getRange(1, col).setValue(`Week ${week.weekNumber}`);
    col++;
  });
  sheet.getRange(1, col).setValue("Total");

  // Add data rows
  row = 2;
  data.markets.forEach(market => {
    // Market header row
    sheet.getRange(row, 1).setValue(market.marketName);
    sheet.getRange(row, 1, 1, col).getFormat().getFill().setColor("#a0a0a0");
    sheet.getRange(row, 1, 1, col).getFormat().getFont().setBold(true);
    row++;

    market.departments.forEach(dept => {
      // Department header row
      sheet.getRange(row, 1).setValue(dept.departmentId);
      sheet.getRange(row, 1, 1, col).getFormat().getFill().setColor("#d0d0d0");
      sheet.getRange(row, 1, 1, col).getFormat().getFont().setBold(true);
      row++;

      dept.persons.forEach(person => {
        // Person row
        sheet.getRange(row, 1).setValue(dept.departmentId);
        sheet.getRange(row, 2).setValue(person.lineManager);
        sheet.getRange(row, 3).setValue(person.personName);

        let personTotal = 0;
        let col = 4;
        person.weeks.forEach(week => {
          sheet.getRange(row, col).setValue(week.totalHours);
          personTotal += week.totalHours;
          col++;
        });
        sheet.getRange(row, col).setValue(personTotal);

        // Alternate row colors
        if (row % 2 === 0) {
          sheet.getRange(row, 1, 1, col).getFormat().getFill().setColor("#f0f0f0");
        }

        row++;
      });
    });
  });

  // Auto-fit columns
  sheet.getUsedRange().getFormat().autofitColumns();

  return "Success";
}
```

2. In Power Automate, add "Run script" action:
   - Location: `OneDrive for Business`
   - Document Library: `OneDrive`
   - File: `File Identifier` (from Create file)
   - Script: `(select your script)`
   - reportData: `ReportData` (from PowerApps)

#### Step 7: Get Sharing Link

1. Add "Create sharing link for a file or folder" (OneDrive)
2. File: `File Identifier` (from Create file)
3. Link Type: `View and edit`
4. Link Scope: `Organization`

#### Step 8: Return URL to PowerApps

1. Add "Respond to PowerApps" action
2. Add output:
   - Name: `FileUrl`
   - Type: `Text`
   - Value: `Link` (from Create sharing link)

### Simplified Alternative: Use HTML Table + Email

If Office Scripts is not available:

1. Convert data to HTML table with formatting
2. Send email with HTML table embedded
3. Or save HTML as file

```powerFx
// In PowerApps, create HTML from data
Set(
    varHTMLTable,
    "<table border='1'>" &
    Concat(
        colWeeklyReport,
        "<tr><td>" & PersonName & "</td>" &
        Concat(
            WeekData,
            "<td>" & Text(TotalHours, "0.0") & "</td>",
            ""
        ) &
        "</tr>",
        ""
    ) &
    "</table>"
);

// Send to flow
'EmailReport'.Run(varHTMLTable)
```

---

## Flow 3: Export to Excel (Daily Report)

### Design

Same as Weekly Report flow, but with these differences:

1. **Input data structure:**
   - Days instead of weeks
   - Date format: "Mon 1/6" instead of "Week 1 (01/02)"

2. **Column headers:**
   - Use day names and dates

3. **Cell formatting:**
   - Weekend columns: grey background (#e0e0e0)
   - Holiday columns: yellow background (#FFF9C4)
   - Regular columns: white

4. **Office Script modifications:**

```javascript
// In Office Script for daily report
data.days.forEach(day => {
  const header = `${day.dayName} ${day.date}`;
  sheet.getRange(1, col).setValue(header);

  // Apply weekend formatting to header
  if (day.isWeekend) {
    sheet.getRange(1, col).getFormat().getFill().setColor("#e0e0e0");
  }

  // Apply holiday formatting to header
  if (day.isHoliday) {
    sheet.getRange(1, col).getFormat().getFill().setColor("#FFF9C4");
  }

  col++;
});
```

Follow same steps as Weekly Report flow, just adapt for daily structure.

---

## Testing Flows

### Test Weekly Export Flow

1. In Power Automate, open flow
2. Click "Test" → "Manually" → "Test"
3. Enter sample JSON:
```json
{
  "ReportData": "{\"markets\":[{\"marketName\":\"Test Market\",\"departments\":[{\"departmentId\":\"TEST-001\",\"persons\":[{\"personName\":\"Test User\",\"lineManager\":\"Test Manager\",\"weeks\":[{\"weekNumber\":1,\"totalHours\":8}]}]}]}],\"weeks\":[{\"weekNumber\":1,\"endDate\":\"2026-01-02\",\"month\":\"Jan 2026\"}]}",
  "ReportName": "Test Report"
}
```
4. Click "Run flow"
5. Verify Excel file created in OneDrive
6. Open file and check formatting

### Test from PowerApps

1. Add test button to PowerApps:
```powerFx
OnSelect =
Set(
    varTestExport,
    'ExportWeeklyAbsenceReporttoExcel'.Run(
        JSON(
            {
                markets: [...],
                weeks: [...]
            },
            JSONFormat.IncludeBinaryData
        ),
        "Test Export"
    )
);
Launch(varTestExport.FileUrl)
```

2. Click button
3. Verify file downloads and opens

---

## Error Handling

### Common Errors

#### Error: "File not found"

**Cause:** Path in OneDrive doesn't exist
**Solution:**
1. Add "Create folder" action before "Create file"
2. Ensure folder path is correct

#### Error: "Invalid JSON"

**Cause:** Data from PowerApps not properly formatted
**Solution:**
1. In PowerApps, use `JSON()` function with correct format
2. Add error handling in flow:
```
// In Parse JSON step, add "Configure run after"
// → "has failed" → Add error notification
```

#### Error: "Script execution timeout"

**Cause:** Too much data for Office Script (>2000 rows)
**Solution:**
1. Split data into multiple sheets
2. Or use chunked processing:
```javascript
// Process in batches of 500 rows
for (let i = 0; i < persons.length; i += 500) {
  const batch = persons.slice(i, i + 500);
  processBatch(batch);
}
```

#### Error: "Access denied"

**Cause:** Flow doesn't have permission to OneDrive/SharePoint
**Solution:**
1. Re-authenticate connections
2. Ensure flow runs under correct user context

### Add Error Notifications

1. After each major step, add "Configure run after"
2. Set to run on "has failed" or "has timed out"
3. Add "Send an email (V2)" action:
```
To: your-email@company.com
Subject: Flow Failed: Export Absence Report
Body:
Flow Name: @{workflow().name}
Run ID: @{workflow().run.id}
Error: @{outputs('Previous_Action')?['error']?['message']}
```

### Add Retry Logic

For unreliable steps (like HTTP requests):

1. Click "..." on action
2. Click "Settings"
3. Under "Retry Policy":
   - Type: `Exponential`
   - Count: `3`
   - Interval: `PT10S` (10 seconds)

---

## Performance Optimization

### 1. Minimize Loops

Instead of nested "Apply to each", use "Select" to transform data first:

```
// Select action to flatten structure
@{body('Parse_JSON')?['markets']?[0]?['departments']?[0]?['persons']}
```

### 2. Use Concurrent Processing

In "Apply to each" settings:
- Enable "Concurrency Control"
- Degree of Parallelism: `20` (for most operations)

### 3. Batch Operations

Instead of creating items one by one, use "Batch" operations:

```
// HTTP action to SharePoint REST API
POST _api/web/lists/getbytitle('AbsenceData')/$batch
```

### 4. Cache Lookups

If you need to lookup the same data multiple times, cache it in a variable:

```
// Initialize variable
{
  "orgMappings": @{body('List_rows_from_OrganizationData')}
}

// Then use filter on variable instead of repeated queries
```

---

## Advanced: CSV Import Flow with UI

Create a better CSV import experience:

1. **PowerApps UI:**
   - Add "Attachment" control
   - User uploads CSV file
   - Preview first 10 rows in gallery
   - "Import" button

2. **Flow triggered from PowerApps:**
   - Receives file content from attachment
   - Parses CSV
   - Validates data
   - Imports to SharePoint list
   - Returns success/error message with count

3. **PowerApps displays result:**
   - "Successfully imported 234 records"
   - Or "Failed: 5 rows had errors" (show details)

---

## Next Steps

1. ✅ Review flow designs
2. ✅ Create Office Scripts for Excel formatting
3. ✅ Test flows with sample data
4. ✅ Add error handling
5. ✅ Deploy to production
6. ➡️ Return to **POWERAPP_STEP_BY_STEP.md** to integrate flows with PowerApp
