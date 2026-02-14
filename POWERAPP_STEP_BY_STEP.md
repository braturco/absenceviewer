# PowerApps Absence Viewer - Step-by-Step Build Guide

## Table of Contents
1. [Prerequisites Check](#prerequisites-check)
2. [Phase 1: SharePoint Setup](#phase-1-sharepoint-setup)
3. [Phase 2: Create PowerApp](#phase-2-create-powerapp)
4. [Phase 3: Data Connections](#phase-3-data-connections)
5. [Phase 4: Build Home Screen](#phase-4-build-home-screen)
6. [Phase 5: Data Processing Logic](#phase-5-data-processing-logic)
7. [Phase 6: Build Weekly Report Screen](#phase-6-build-weekly-report-screen)
8. [Phase 7: Build Daily Report Screen](#phase-7-build-daily-report-screen)
9. [Phase 8: Power Automate Flows](#phase-8-power-automate-flows)
10. [Phase 9: Testing](#phase-9-testing)
11. [Phase 10: Deployment](#phase-10-deployment)

---

## Prerequisites Check

Before starting, ensure you have:

- [ ] Microsoft 365 account with PowerApps license
- [ ] SharePoint site with appropriate permissions
- [ ] Access to Power Automate
- [ ] Sample CSV data for testing (see POWERAPP_SAMPLE_DATA.md)
- [ ] PowerApps Studio installed or web access

**Time Estimate:** 6-8 hours for initial build, 2-3 hours for testing and refinement

---

## Phase 1: SharePoint Setup

### Step 1.1: Create SharePoint Site (if needed)

1. Navigate to SharePoint Home
2. Click "Create site"
3. Choose "Team site"
4. Name: "Absence Management"
5. Description: "Employee absence tracking and reporting"
6. Click "Finish"

### Step 1.2: Create SharePoint Lists

Follow instructions in **POWERAPP_DATA_SETUP.md** to create these three lists:

1. **AbsenceData**
   - Create list manually or via PowerShell
   - Add all required columns
   - Set up column validation

2. **OrganizationData**
   - Create list
   - Add columns
   - Create index on OfficeLocationID

3. **HolidayCalendar**
   - Create list
   - Add columns

**✓ Verification:**
- All 3 lists created
- Columns match specifications
- Sample data added to each list

### Step 1.3: Load Sample Data

**Option A: Manual Entry**
1. Click "New" in each list
2. Enter sample records (at least 10-20 for testing)

**Option B: Import from CSV**
1. Click "Integrate" → "Excel"
2. Click "Import from CSV"
3. Select your CSV file
4. Map columns
5. Click "Import"

**Option C: PowerShell Script**
Use the PowerShell script in **POWERAPP_DATA_SETUP.md** section "Data Migration"

**Sample Data:**
- At least 20 absence records across multiple employees
- At least 5 department-to-market mappings
- At least 10 holidays for 2026

---

## Phase 2: Create PowerApp

### Step 2.1: Create Blank Canvas App

1. Navigate to [make.powerapps.com](https://make.powerapps.com)
2. Click "Create" → "Blank app" → "Blank canvas app"
3. App name: "Absence Viewer"
4. Format: Tablet (16:9)
5. Click "Create"

**Result:** PowerApps Studio opens with a blank screen

### Step 2.2: Configure App Settings

1. Click "File" → "Settings" → "General"
   - Set orientation: Landscape
   - Enable "Enable modern controls" (optional)

2. Click "Display"
   - Scale to fit: On
   - Lock aspect ratio: On
   - Lock orientation: Landscape

3. Click "Advanced settings"
   - Data row limit for non-delegable queries: 2000
   - Enable formula bar: On

4. Click "Save" (top right)

### Step 2.3: Rename Initial Screen

1. In Tree View (left panel), select "Screen1"
2. Right-click → Rename → "HomeScreen"

---

## Phase 3: Data Connections

### Step 3.1: Add SharePoint Connections

1. Click "Data" icon in left panel
2. Click "Add data"
3. Search "SharePoint"
4. Click "SharePoint" connector
5. Enter your SharePoint site URL:
   ```
   https://yourtenant.sharepoint.com/sites/AbsenceManagement
   ```
6. Click "Connect"
7. Select all three lists:
   - ☑ AbsenceData
   - ☑ OrganizationData
   - ☑ HolidayCalendar
8. Click "Connect"

**✓ Verification:**
- All 3 lists appear in Data sources panel
- No error messages
- Lists show column names when expanded

### Step 3.2: Test Data Connection

1. Insert a Label on HomeScreen
2. Set Text property:
   ```powerFx
   CountRows(AbsenceData)
   ```
3. Should display a number (count of records)
4. If you see "0", check SharePoint permissions
5. Delete the test label

---

## Phase 4: Build Home Screen

### Step 4.1: Set Up App.OnStart

1. In Tree View, select "App"
2. In property dropdown (top left), select "OnStart"
3. Paste this formula (see **POWERAPP_FORMULAS.md** for complete version):

```powerFx
// ===== GLOBAL CONSTANTS =====
Set(gblWeek1EndDate, Date(2026, 1, 2));
Set(gblCurrentYear, 2026);

// ===== COLOR PALETTE =====
Set(gblColors, {
    MarketHeader: RGBA(160, 160, 160, 1),
    DeptHeader: RGBA(208, 208, 208, 1),
    PersonEven: RGBA(240, 240, 240, 1),
    PersonOdd: RGBA(255, 255, 255, 1),
    OverallTotal: RGBA(136, 136, 136, 1),
    HolidayWeek: RGBA(255, 224, 130, 1),
    HolidayCell: RGBA(255, 249, 196, 1),
    Weekend: RGBA(224, 224, 224, 1),
    ZeroValue: RGBA(204, 204, 204, 1)
});

Set(gblStatusColors, {
    Scheduled: RGBA(0, 128, 0, 1),
    AwaitingApproval: RGBA(255, 255, 0, 1),
    AwaitingWithdrawl: RGBA(255, 165, 0, 1),
    Completed: RGBA(0, 0, 255, 1),
    InProgress: RGBA(128, 0, 128, 1),
    Saved: RGBA(128, 128, 128, 1)
});

// ===== DEFAULT DATE RANGES =====
Set(varStartDate, Date(2026, 1, 2));
Set(varEndDate, Date(2026, 12, 31));
Set(varDailyStartDate, Today());
Set(varDailyEndDate, DateAdd(Today(), 6, TimeUnit.Days));

// ===== LOAD DATA =====
ClearCollect(
    colAbsenceRaw,
    Filter(
        AbsenceData,
        AbsenceStatus <> "Withdrawn" &&
        AbsenceStatus <> "Denied" &&
        !IsBlank(Department)
    )
);

ClearCollect(colOrgData, OrganizationData);
ClearCollect(colHolidayData, HolidayCalendar);

// ===== STATUS FLAGS =====
Set(varDataLoaded, true);
Set(varRecordCount, CountRows(colAbsenceRaw));
```

4. Click "Run OnStart" (three dots menu next to App)

**✓ Verification:**
- No error messages in formula bar
- Variables panel shows varDataLoaded = true
- varRecordCount shows number of records

### Step 4.2: Add Title and Status Panel

1. Insert → Label
   - Name: `lblTitle`
   - Text: `"Company Absence Viewer"`
   - X: `40`
   - Y: `40`
   - Width: `Parent.Width - 80`
   - Height: `60`
   - Font: Segoe UI
   - Size: 24
   - Font weight: Bold
   - Color: `RGBA(51, 51, 51, 1)`

2. Insert → Container (modern control) or Rectangle (classic)
   - Name: `conDataStatus`
   - X: `40`
   - Y: `120`
   - Width: `Parent.Width - 80`
   - Height: `120`
   - Fill: `RGBA(240, 240, 240, 1)`

3. Inside conDataStatus, Insert → Label × 3

   **Label 1:**
   - Name: `lblAbsenceStatus`
   - Text:
     ```powerFx
     If(
         varRecordCount > 0,
         "✓ Loaded " & varRecordCount & " absence records",
         "⚠ No absence data loaded"
     )
     ```
   - X: `20`
   - Y: `10`
   - Width: `Parent.Width - 40`
   - Color:
     ```powerFx
     If(varRecordCount > 0, RGBA(0, 128, 0, 1), RGBA(255, 0, 0, 1))
     ```

   **Label 2:**
   - Name: `lblOrgStatus`
   - Text:
     ```powerFx
     If(
         CountRows(colOrgData) > 0,
         "✓ Organization data loaded (" & CountRows(colOrgData) & " mappings)",
         "⚠ No organization data loaded"
     )
     ```
   - X: `20`
   - Y: `40`
   - Width: `Parent.Width - 40`
   - Color:
     ```powerFx
     If(CountRows(colOrgData) > 0, RGBA(0, 128, 0, 1), RGBA(255, 0, 0, 1))
     ```

   **Label 3:**
   - Name: `lblHolidayStatus`
   - Text:
     ```powerFx
     If(
         CountRows(colHolidayData) > 0,
         "✓ Holiday calendar loaded (" & CountRows(colHolidayData) & " holidays)",
         "⚠ No holiday calendar loaded"
     )
     ```
   - X: `20`
   - Y: `70`
   - Width: `Parent.Width - 40`
   - Color:
     ```powerFx
     If(CountRows(colHolidayData) > 0, RGBA(0, 128, 0, 1), RGBA(255, 0, 0, 1))
     ```

### Step 4.3: Add Date Pickers

1. Insert → Label
   - Text: `"Weekly Report Start Date:"`
   - X: `40`
   - Y: `260`

2. Insert → Date picker
   - Name: `dpStartDate`
   - DefaultDate: `varStartDate`
   - X: `250`
   - Y: `260`
   - OnChange:
     ```powerFx
     Set(varStartDate, dpStartDate.SelectedDate)
     ```

3. Repeat for end date and daily report dates (see **POWERAPP_UI_COMPONENTS.md**)

### Step 4.4: Add Navigation Buttons

1. Insert → Button
   - Name: `btnGenerateWeekly`
   - Text: `"Generate Weekly Report"`
   - X: `40`
   - Y: `370`
   - Width: `300`
   - Height: `50`
   - Fill: `RGBA(0, 120, 212, 1)`
   - Color: `White`
   - OnSelect: (leave blank for now, we'll add in Phase 6)

2. Insert → Button
   - Name: `btnGenerateDaily`
   - Text: `"Generate Daily Report"`
   - X: `40`
   - Y: `500`
   - Width: `300`
   - Height: `50`
   - Fill: `RGBA(0, 120, 212, 1)`
   - Color: `White`
   - OnSelect: (leave blank for now, we'll add in Phase 7)

3. Insert → Button
   - Name: `btnRefresh`
   - Text: `"Refresh Data"`
   - X: `40`
   - Y: `630`
   - Width: `300`
   - Height: `50`
   - Fill: `RGBA(102, 102, 102, 1)`
   - Color: `White`
   - OnSelect:
     ```powerFx
     Refresh(AbsenceData);
     Refresh(OrganizationData);
     Refresh(HolidayCalendar);
     // Re-run OnStart (copy formula from App.OnStart)
     ```

**✓ Verification:**
- All controls visible and aligned
- Status labels show green checkmarks
- Date pickers show default dates
- Buttons display correctly

---

## Phase 5: Data Processing Logic

### Step 5.1: Create Processing Function (Button)

Since PowerApps doesn't support custom functions, we'll create a "Process Data" button that runs all processing steps.

1. On HomeScreen, Insert → Button
   - Name: `btnProcessData`
   - Text: `"Process Data (Debug)"`
   - X: `40`
   - Y: `700`
   - Visible: `false` (we'll hide this in production)

2. Set OnSelect property (this is the core processing logic):

```powerFx
// ===== STEP 1: EXPAND ABSENCES TO DAILY RECORDS =====
Set(varProcessing, true);
Notify("Processing absence data...", NotificationType.Information);

ClearCollect(
    colDailyAbsences,
    ForAll(
        colAbsenceRaw As _absence,
        With(
            {
                _totalDays: DateDiff(
                    _absence.AbsenceStartDate,
                    _absence.AbsenceEndDate,
                    TimeUnit.Days
                ) + 1,
                _normalHours: If(
                    IsBlank(_absence.NormalWorkingHours),
                    40,
                    _absence.NormalWorkingHours
                ),
                _totalHours: If(
                    _absence.UOM = "D",
                    _absence.AbsenceDuration * (_normalHours / 5),
                    _absence.AbsenceDuration
                )
            },
            With(
                {
                    _workingDays: Filter(
                        AddColumns(
                            Sequence(_totalDays, 0),
                            "AbsDate",
                            DateAdd(_absence.AbsenceStartDate, Value, TimeUnit.Days),
                            "DayOfWeek",
                            Weekday(
                                DateAdd(_absence.AbsenceStartDate, Value, TimeUnit.Days),
                                StartOfWeek.Monday
                            )
                        ),
                        DayOfWeek >= 1 && DayOfWeek <= 5
                    ),
                    _workingDayCount: CountRows(_workingDays)
                },
                If(
                    _workingDayCount > 0,
                    AddColumns(
                        _workingDays,
                        "LastName", _absence.LastName,
                        "FirstName", _absence.FirstName,
                        "PersonName", _absence.LastName & ", " & _absence.FirstName,
                        "Department", _absence.Department,
                        "LineManager", _absence.LineManager,
                        "AbsenceType", _absence.AbsenceType,
                        "AbsenceStatus", _absence.AbsenceStatus,
                        "AbsenceComment", _absence.AbsenceComment,
                        "HoursThisDay", _totalHours / _workingDayCount,
                        "DeptID", Left(_absence.Department, 8),
                        "MarketSubSector", Coalesce(
                            LookUp(
                                colOrgData,
                                OfficeLocationID = Left(_absence.Department, 8),
                                MarketSubSectorName
                            ),
                            "Unknown"
                        ),
                        "WeekNumber",
                        With(
                            {
                                _friday: DateAdd(
                                    AbsDate,
                                    Mod(5 - DayOfWeek + 7, 7),
                                    TimeUnit.Days
                                )
                            },
                            1 + Round(
                                DateDiff(gblWeek1EndDate, _friday, TimeUnit.Days) / 7,
                                0
                            )
                        )
                    ),
                    Blank()
                )
            )
        )
    )
);

Notify("Daily absences expanded: " & CountRows(colDailyAbsences) & " records", NotificationType.Success);

// ===== STEP 2: CREATE WEEKLY SUMMARY =====
ClearCollect(
    colWeeklySummary,
    AddColumns(
        GroupBy(
            colDailyAbsences,
            "MarketSubSector",
            "Department",
            "PersonName",
            "LineManager",
            "WeekNumber",
            "GroupedRecords"
        ),
        "TotalHours", Sum(GroupedRecords, HoursThisDay),
        "Statuses", Concat(Distinct(GroupedRecords, AbsenceStatus), Value, ", ")
    )
);

Notify("Weekly summary created: " & CountRows(colWeeklySummary) & " records", NotificationType.Success);

Set(varProcessing, false);
Notify("Data processing complete!", NotificationType.Success);
```

3. Click the button to test processing

**✓ Verification:**
- colDailyAbsences created with records
- colWeeklySummary created with aggregated data
- Check Variables panel to inspect data
- No error messages

### Step 5.2: Test Data Processing

1. Insert → Gallery (Blank vertical)
   - Name: `galTestDaily`
   - Items: `FirstN(colDailyAbsences, 10)`
   - Visible: `false` (debug only)

2. Inside gallery, Insert → Label
   - Text:
     ```powerFx
     ThisItem.PersonName & " - " & Text(ThisItem.AbsDate, "mm/dd") &
     " - Week " & ThisItem.WeekNumber & " - " &
     Text(ThisItem.HoursThisDay, "0.0") & "h"
     ```

3. Preview the gallery to verify data looks correct

4. Delete debug gallery when satisfied

---

## Phase 6: Build Weekly Report Screen

### Step 6.1: Create New Screen

1. Insert → New screen → Blank
2. Rename to `WeeklyReportScreen`

### Step 6.2: Add Header Controls

Follow **POWERAPP_UI_COMPONENTS.md** to add:
- Back button (navigate to HomeScreen)
- Title label
- Export button (placeholder for now)
- Expand All / Collapse All buttons

### Step 6.3: Generate 52 Weeks Collection

Add a button on HomeScreen (or in btnGenerateWeekly OnSelect):

```powerFx
// Generate 52 weeks starting from varStartDate
ClearCollect(
    colWeeks,
    AddColumns(
        Sequence(52, 0),
        "WeekOffset", Value,
        "WeekEndDate",
        DateAdd(
            DateAdd(
                varStartDate,
                Mod((5 - Weekday(varStartDate, StartOfWeek.Monday) + 7), 7),
                TimeUnit.Days
            ),
            Value * 7,
            TimeUnit.Days
        )
    )
);

// Add week numbers and formatting
ClearCollect(
    colWeeks,
    AddColumns(
        colWeeks,
        "WeekNumber",
        With(
            {_friday: WeekEndDate},
            1 + Round(DateDiff(gblWeek1EndDate, _friday, TimeUnit.Days) / 7, 0)
        ),
        "MonthName", Text(WeekEndDate, "mmm yyyy"),
        "HeaderText",
        "Week " & WeekNumber & "\n" & Text(WeekEndDate, "mm/dd"),
        "HasHoliday",
        CountRows(
            Filter(
                colHolidayData,
                HolidayDate >= DateAdd(WeekEndDate, -6, TimeUnit.Days) &&
                HolidayDate <= WeekEndDate
            )
        ) > 0
    )
);
```

### Step 6.4: Build Hierarchical Data Structure

**Simplified Approach:**

Create 3 separate collections for Market, Department, and Person levels.

See **POWERAPP_FORMULAS.md** section "Build Weekly Report Data Structure" for complete formulas.

For testing, start with a simpler flat structure:

```powerFx
ClearCollect(
    colWeeklyReport,
    AddColumns(
        colWeeklySummary,
        "RowType", "Person",
        "Level", 3,
        "IsVisible", true
    )
);
```

### Step 6.5: Add Week Header Gallery

On WeeklyReportScreen:

1. Insert → Horizontal gallery
   - Name: `galWeekHeaders`
   - Items: `colWeeks`
   - X: `450` (after fixed columns)
   - Y: `100`
   - Width: `Parent.Width - 470`
   - Height: `80`
   - TemplateSize: `60`
   - TemplatePadding: `2`

2. Inside template, add Label:
   - Text: `ThisItem.HeaderText`
   - Fill:
     ```powerFx
     If(ThisItem.HasHoliday, gblColors.HolidayWeek, RGBA(242, 242, 242, 1))
     ```
   - Align: Center

### Step 6.6: Add Fixed Column Headers

Add 3 labels for Market/Department/Person columns (see **POWERAPP_UI_COMPONENTS.md**)

### Step 6.7: Add Main Data Gallery

1. Insert → Vertical gallery (Blank)
   - Name: `galRows`
   - Items: `colWeeklyReport` (simplified for now)
   - X: `0`
   - Y: `180`
   - Width: `450`
   - Height: `Parent.Height - 200`
   - TemplateSize: `40`

2. Inside template:
   - Add Label for PersonName
   - Add Horizontal gallery for week data
   - Follow **POWERAPP_UI_COMPONENTS.md** for complete layout

### Step 6.8: Add Week Data Gallery Inside Each Row

Inside galRows template:

1. Insert → Horizontal gallery
   - Name: `galWeekData`
   - Items:
     ```powerFx
     // Get week data for this person
     AddColumns(
         colWeeks,
         "TotalHours",
         Sum(
             Filter(
                 colWeeklySummary,
                 PersonName = galRows.Selected.PersonName &&
                 Department = galRows.Selected.Department &&
                 WeekNumber = ThisRecord.WeekNumber
             ),
             TotalHours
         )
     )
     ```
   - X: `450`
   - Y: `0`
   - Width: `Parent.Width - 450`
   - Height: `40`
   - TemplateSize: `60`

2. Inside galWeekData template, add Label:
   - Text:
     ```powerFx
     If(
         ThisItem.TotalHours = 0,
         "0",
         Text(Round(ThisItem.TotalHours * 2, 0) / 2, "0.0")
     )
     ```

### Step 6.9: Wire Up Generate Button

Back on HomeScreen, set btnGenerateWeekly OnSelect:

```powerFx
// Run all processing
btnProcessData.OnSelect;

// Generate weeks
/* paste week generation formula */;

// Navigate to report
Navigate(WeeklyReportScreen, ScreenTransition.Fade)
```

**✓ Verification:**
- Click "Generate Weekly Report"
- WeeklyReportScreen displays
- Week headers show correctly
- Person names and hours display
- Scroll horizontally to see all 52 weeks

---

## Phase 7: Build Daily Report Screen

### Step 7.1: Create New Screen

1. Insert → New screen → Blank
2. Rename to `DailyReportScreen`

### Step 7.2: Generate Days Collection

Add to btnGenerateDaily OnSelect:

```powerFx
// Validate date range
If(
    DateDiff(varDailyStartDate, varDailyEndDate, TimeUnit.Days) > 13,
    Notify("Daily report limited to 14 days", NotificationType.Warning),

    // Process data
    btnProcessData.OnSelect;

    // Generate days
    With(
        {_dayCount: DateDiff(varDailyStartDate, varDailyEndDate, TimeUnit.Days) + 1},
        ClearCollect(
            colDays,
            AddColumns(
                Sequence(_dayCount, 0),
                "DayOffset", Value,
                "AbsDate", DateAdd(varDailyStartDate, Value, TimeUnit.Days),
                "DayOfWeek", Weekday(DateAdd(varDailyStartDate, Value, TimeUnit.Days), StartOfWeek.Monday),
                "HeaderText",
                Text(DateAdd(varDailyStartDate, Value, TimeUnit.Days), "ddd m/d"),
                "IsWeekend", Weekday(DateAdd(varDailyStartDate, Value, TimeUnit.Days), StartOfWeek.Monday) >= 6,
                "IsHoliday",
                CountRows(
                    Filter(
                        colHolidayData,
                        HolidayDate = DateAdd(varDailyStartDate, Value, TimeUnit.Days)
                    )
                ) > 0
            )
        )
    );

    // Create daily summary
    ClearCollect(
        colDailySummary,
        AddColumns(
            GroupBy(
                Filter(
                    colDailyAbsences,
                    AbsDate >= varDailyStartDate && AbsDate <= varDailyEndDate
                ),
                "MarketSubSector",
                "Department",
                "PersonName",
                "LineManager",
                "AbsDate",
                "GroupedRecords"
            ),
            "TotalHours", Sum(GroupedRecords, HoursThisDay),
            "Statuses", Concat(Distinct(GroupedRecords, AbsenceStatus), Value, ", ")
        )
    );

    Navigate(DailyReportScreen, ScreenTransition.Fade)
)
```

### Step 7.3: Build Daily Report UI

Follow same pattern as WeeklyReportScreen, but:
- Use `colDays` instead of `colWeeks`
- Column width: 80px instead of 60px
- Apply weekend/holiday highlighting
- See **POWERAPP_UI_COMPONENTS.md** for complete layout

---

## Phase 8: Power Automate Flows

### Step 8.1: Create "Export to Excel" Flow

See **POWERAPP_POWER_AUTOMATE.md** for complete flow design.

**Quick setup:**

1. Navigate to [make.powerautomate.com](https://make.powerautomate.com)
2. Create → Instant cloud flow
3. Name: "Export Absence Report to Excel"
4. Trigger: PowerApps
5. Add actions:
   - Parse JSON (parse report data from PowerApps)
   - Create Excel file
   - Format Excel (apply styles)
   - Store in OneDrive or SharePoint
   - Return file URL to PowerApps
6. Save flow

### Step 8.2: Add Flow to PowerApp

1. In PowerApps Studio, click "Power Automate" icon
2. Click "Add flow"
3. Select "Export Absence Report to Excel"
4. Flow appears in your app

### Step 8.3: Wire Up Export Button

On WeeklyReportScreen, set btnExport OnSelect:

```powerFx
Set(varExporting, true);

// Call flow
Set(
    varExportResult,
    'ExportAbsenceReporttoExcel'.Run(
        JSON(colWeeklyReport, JSONFormat.IncludeBinaryData),
        "Weekly Report " & Text(Now(), "yyyy-mm-dd")
    )
);

// Download file
If(
    !IsBlank(varExportResult.fileurl),
    Launch(varExportResult.fileurl);
    Notify("Export complete!", NotificationType.Success),
    Notify("Export failed", NotificationType.Error)
);

Set(varExporting, false);
```

---

## Phase 9: Testing

### Step 9.1: Unit Testing

Test each component individually:

**Data Loading:**
- [ ] AbsenceData loads correctly
- [ ] OrganizationData loads correctly
- [ ] HolidayCalendar loads correctly
- [ ] Counts match expected

**Data Processing:**
- [ ] Absences expand to daily records
- [ ] Working days calculated correctly (Mon-Fri only)
- [ ] Hours distributed evenly
- [ ] Week numbers calculated correctly
- [ ] Market sub-sectors mapped correctly

**Weekly Report:**
- [ ] 52 weeks generated
- [ ] Week headers display correctly
- [ ] Holiday weeks highlighted
- [ ] Person rows show correct hours
- [ ] Zero hours displayed as "0" in grey
- [ ] Totals calculate correctly

**Daily Report:**
- [ ] Up to 14 days generated
- [ ] Weekend columns highlighted
- [ ] Holiday dates highlighted
- [ ] Daily hours correct
- [ ] Date range validation works

### Step 9.2: Integration Testing

Test complete workflows:

1. **Full Weekly Report:**
   - Load data
   - Select date range
   - Generate report
   - Verify accuracy
   - Export to Excel
   - Open Excel file
   - Verify formatting

2. **Full Daily Report:**
   - Load data
   - Select 7-day range
   - Generate report
   - Verify accuracy
   - Export to Excel

3. **Edge Cases:**
   - Employee with absence spanning weekends
   - Employee with absence in "Days" UOM
   - Employee with multiple absences in same week
   - Week with multiple holidays
   - Employee in unknown market sub-sector

### Step 9.3: Performance Testing

Test with production-scale data:

- [ ] Load 2000+ absence records
- [ ] Generate weekly report
- [ ] Measure load time (should be < 10 seconds)
- [ ] Test scrolling performance
- [ ] Test expand/collapse responsiveness

### Step 9.4: User Acceptance Testing

Have actual users test:

- [ ] Can they navigate easily?
- [ ] Are reports easy to read?
- [ ] Is data accurate?
- [ ] Any missing features?
- [ ] Any confusing UI elements?

---

## Phase 10: Deployment

### Step 10.1: Publish App

1. Click "File" → "Save"
2. Click "Publish"
3. Click "Publish this version"
4. Add release notes
5. Click "Publish"

### Step 10.2: Share App

1. Click "File" → "Share"
2. Add users or groups:
   - HR Team (Can Edit)
   - Managers (Can Use)
   - Employees (Can View)
3. Send email notification
4. Click "Share"

### Step 10.3: Create App Documentation

Create user guide covering:
- How to access the app
- How to generate reports
- How to interpret status colors
- How to export to Excel
- Who to contact for support

### Step 10.4: Schedule Data Refresh

If using SharePoint lists:
- Set up automated data import from existing systems
- Create Power Automate flow to import CSV files weekly
- Schedule holiday calendar updates annually

### Step 10.5: Monitor Usage

1. In PowerApps admin center, monitor:
   - Number of active users
   - Error logs
   - Performance metrics

2. Set up alerts for:
   - Flow failures
   - App errors
   - Slow performance

---

## Troubleshooting

### Common Issues

**Issue: "Delegation warning" on Filter**
- Solution: Use ClearCollect to load data into memory first

**Issue: Slow performance with 2000+ records**
- Solution: Implement pagination or date range filters

**Issue: Week numbers incorrect**
- Solution: Verify gblWeek1EndDate is set correctly (Friday, Jan 2, 2026)

**Issue: Excel export fails**
- Solution: Check Power Automate flow run history for errors

**Issue: Data doesn't refresh**
- Solution: Add Refresh() calls before ClearCollect

**Issue: Expand/collapse doesn't work**
- Solution: Check IsVisible formula and context variables

---

## Next Steps After Build

1. **Enhancements:**
   - Add filtering by department/market
   - Add search functionality
   - Add email report scheduling
   - Add manager approval workflow

2. **Optimization:**
   - Implement caching for better performance
   - Add loading spinners
   - Optimize formulas for delegation

3. **Integration:**
   - Connect to HR systems via API
   - Sync with Outlook calendar
   - Integrate with Teams

4. **Reporting:**
   - Add summary dashboards
   - Create Power BI embedded reports
   - Add trend analysis

---

## Completion Checklist

- [ ] All SharePoint lists created and populated
- [ ] PowerApp created with all 3 screens
- [ ] Data connections working
- [ ] Data processing formulas implemented
- [ ] Weekly report functional
- [ ] Daily report functional
- [ ] Power Automate flows configured
- [ ] Export to Excel working
- [ ] Testing completed
- [ ] App published
- [ ] Users granted access
- [ ] Documentation created
- [ ] Training provided

**Congratulations on building your PowerApps Absence Viewer!**
