# PowerApps Absence Viewer - UI Components Guide

## Table of Contents
1. [App Structure Overview](#app-structure-overview)
2. [Home Screen](#home-screen)
3. [Weekly Report Screen](#weekly-report-screen)
4. [Daily Report Screen](#daily-report-screen)
5. [Hierarchical Gallery Pattern](#hierarchical-gallery-pattern)
6. [Status Legend Component](#status-legend-component)
7. [Export Button](#export-button)

---

## App Structure Overview

### Screen Navigation

```
App
├── HomeScreen
│   ├── Title
│   ├── Data Status Panel
│   ├── Date Pickers
│   └── Navigation Buttons
├── WeeklyReportScreen
│   ├── Header
│   ├── Expand/Collapse Controls
│   ├── Status Legend
│   ├── Horizontal Gallery (Week Headers)
│   ├── Vertical Gallery (Hierarchical Rows)
│   └── Export Button
└── DailyReportScreen
    ├── Header
    ├── Expand/Collapse Controls
    ├── Status Legend
    ├── Horizontal Gallery (Day Headers)
    ├── Vertical Gallery (Hierarchical Rows)
    └── Export Button
```

### Color Palette

```powerFx
// Define in App.OnStart for consistency
Set(gblColors, {
    MarketHeader: RGBA(160, 160, 160, 1),      // #a0a0a0
    DeptHeader: RGBA(208, 208, 208, 1),        // #d0d0d0
    PersonEven: RGBA(240, 240, 240, 1),        // #f0f0f0
    PersonOdd: RGBA(255, 255, 255, 1),         // white
    OverallTotal: RGBA(136, 136, 136, 1),      // #888888
    HolidayWeek: RGBA(255, 224, 130, 1),       // #FFE082
    HolidayCell: RGBA(255, 249, 196, 1),       // #FFF9C4
    Weekend: RGBA(224, 224, 224, 1),           // #e0e0e0
    ZeroValue: RGBA(204, 204, 204, 1)          // #ccc
});

Set(gblStatusColors, {
    Scheduled: RGBA(0, 128, 0, 1),             // green
    AwaitingApproval: RGBA(255, 255, 0, 1),    // yellow
    AwaitingWithdrawl: RGBA(255, 165, 0, 1),   // orange
    Completed: RGBA(0, 0, 255, 1),             // blue
    InProgress: RGBA(128, 0, 128, 1),          // purple
    Saved: RGBA(128, 128, 128, 1)              // gray
});
```

---

## Home Screen

### Layout

```
┌─────────────────────────────────────┐
│  Company Absence Viewer             │
│                                      │
│  [✓ Loaded 1,234 absence records]   │
│  [✓ Organization data loaded]       │
│  [✓ Holiday calendar loaded]        │
│                                      │
│  Start Date: [  2026-01-02  ]       │
│  End Date:   [  2026-12-31  ]       │
│                                      │
│  [ Generate Weekly Report ]         │
│                                      │
│  Daily Report Start: [ 2026-02-10 ] │
│  Daily Report End:   [ 2026-02-17 ] │
│                                      │
│  [ Generate Daily Report ]          │
│                                      │
│  [ Refresh Data ]                   │
└─────────────────────────────────────┘
```

### Controls

#### Title Label
```
Name: lblTitle
Text: "Company Absence Viewer"
Font: Segoe UI, 24, Bold
Color: RGBA(51, 51, 51, 1)
X: 40
Y: 40
Width: Parent.Width - 80
Height: 60
```

#### Data Status Panel

**Container with status labels:**

```powerFx
// Container
Name: conDataStatus
X: 40
Y: 120
Width: Parent.Width - 80
Height: 120
Fill: RGBA(240, 240, 240, 1)

// Label 1: Absence Data Status
Name: lblAbsenceStatus
Text: If(
    varRecordCount > 0,
    "✓ Loaded " & varRecordCount & " absence records",
    "⚠ No absence data loaded"
)
Color: If(varRecordCount > 0, RGBA(0, 128, 0, 1), RGBA(255, 0, 0, 1))
X: 20
Y: 10
Width: Parent.Width - 40

// Label 2: Organization Data Status
Name: lblOrgStatus
Text: If(
    CountRows(colOrgData) > 0,
    "✓ Organization data loaded (" & CountRows(colOrgData) & " mappings)",
    "⚠ No organization data loaded"
)
Color: If(CountRows(colOrgData) > 0, RGBA(0, 128, 0, 1), RGBA(255, 0, 0, 1))
X: 20
Y: 40
Width: Parent.Width - 40

// Label 3: Holiday Calendar Status
Name: lblHolidayStatus
Text: If(
    CountRows(colHolidayData) > 0,
    "✓ Holiday calendar loaded (" & CountRows(colHolidayData) & " holidays)",
    "⚠ No holiday calendar loaded"
)
Color: If(CountRows(colHolidayData) > 0, RGBA(0, 128, 0, 1), RGBA(255, 0, 0, 1))
X: 20
Y: 70
Width: Parent.Width - 40
```

#### Date Pickers

```powerFx
// Start Date Label
Name: lblStartDate
Text: "Start Date:"
X: 40
Y: 260

// Start Date Picker
Name: dpStartDate
DefaultDate: varStartDate
X: 150
Y: 260
OnChange: Set(varStartDate, dpStartDate.SelectedDate)

// End Date Label
Name: lblEndDate
Text: "End Date:"
X: 40
Y: 310

// End Date Picker
Name: dpEndDate
DefaultDate: varEndDate
X: 150
Y: 310
OnChange: Set(varEndDate, dpEndDate.SelectedDate)
```

#### Buttons

```powerFx
// Generate Weekly Report Button
Name: btnGenerateWeekly
Text: "Generate Weekly Report"
X: 40
Y: 370
Width: 300
Height: 50
Fill: RGBA(0, 120, 212, 1)
Color: White
OnSelect:
    // Process data and navigate
    // Call processing formulas (see POWERAPP_FORMULAS.md)
    Set(varProcessing, true);

    // Expand absences to daily records
    ClearCollect(...); // See formulas doc

    // Generate weekly summary
    ClearCollect(...); // See formulas doc

    Set(varProcessing, false);
    Navigate(WeeklyReportScreen, ScreenTransition.Fade)

// Generate Daily Report Button
Name: btnGenerateDaily
Text: "Generate Daily Report"
X: 40
Y: 500
Width: 300
Height: 50
Fill: RGBA(0, 120, 212, 1)
Color: White
OnSelect:
    // Validate date range
    If(
        DateDiff(varDailyStartDate, varDailyEndDate, TimeUnit.Days) > 13,
        Notify("Daily report limited to 14 days", NotificationType.Warning),
        Set(varProcessing, true);
        // Process daily data
        ClearCollect(...); // See formulas doc
        Set(varProcessing, false);
        Navigate(DailyReportScreen, ScreenTransition.Fade)
    )

// Refresh Data Button
Name: btnRefresh
Text: "Refresh Data"
X: 40
Y: 630
Width: 300
Height: 50
Fill: RGBA(102, 102, 102, 1)
Color: White
OnSelect:
    Refresh(AbsenceData);
    Refresh(OrganizationData);
    Refresh(HolidayCalendar);
    // Re-run OnStart collections
    // (paste OnStart formula here or create a separate "Initialize" function)
    Notify("Data refreshed", NotificationType.Success)
```

---

## Weekly Report Screen

### Layout Overview

```
┌────────────────────────────────────────────────┐
│ < Back          Weekly Absence Report   Export│
│                                                │
│ [ Expand All ]  [ Collapse All ]              │
│                                                │
│ Legend: ● Scheduled  ● Awaiting approval ...  │
│                                                │
│ ┌──────────────────────────────────────────┐  │
│ │ Market│Dept│Person│Jan 2026│Feb 2026│...│  │
│ │                   │W1│W2│W3│W4│W5│W6│  │  │
│ ├──────────────────────────────────────────┤  │
│ │ ▼ North America   │  │  │  │  │  │  │  │  │
│ │   ▼ ABC-ENV-123   │  │  │  │  │  │  │  │  │
│ │       Smith,John  │8 │16│0 │24│0 │8 │  │  │
│ │       Doe,Jane    │0 │8 │8 │0 │16│0 │  │  │
│ │   ▶ DEF-ENV-456   │  │  │  │  │  │  │  │  │
│ │ ▶ Europe          │  │  │  │  │  │  │  │  │
│ └──────────────────────────────────────────┘  │
│                                                │
│ Scroll horizontally for more weeks ───────>   │
└────────────────────────────────────────────────┘
```

### Header Controls

```powerFx
// Back Button (Icon)
Name: icoBack
Icon: Icon.ChevronLeft
X: 20
Y: 20
Width: 60
Height: 60
OnSelect: Navigate(HomeScreen, ScreenTransition.Fade)

// Title Label
Name: lblWeeklyTitle
Text: "Weekly Absence Report"
X: 100
Y: 30
Font: Segoe UI, 20, Bold

// Export Button
Name: btnExportWeekly
Text: "Export to Excel"
X: Parent.Width - 220
Y: 30
Width: 200
Height: 40
OnSelect: // See Export section below
```

#### Expand/Collapse Buttons

```powerFx
// Expand All Button
Name: btnExpandAll
Text: "Expand All"
X: 20
Y: 100
Width: 150
Height: 40
OnSelect:
    Set(varExpandAllMarkets, true);
    Set(varExpandAllDepts, true)

// Collapse All Button
Name: btnCollapseAll
Text: "Collapse All"
X: 180
Y: 100
Width: 150
Height: 40
OnSelect:
    Set(varExpandAllMarkets, false);
    Set(varExpandAllDepts, false)
```

### Status Legend

```powerFx
// Container for legend
Name: conLegend
X: 20
Y: 160
Width: Parent.Width - 40
Height: 40
Fill: RGBA(245, 245, 245, 1)

// Label
Name: lblLegendTitle
Text: "Status Legend:"
X: 10
Y: 10
Font: Bold

// Horizontal Gallery for status items
Name: galLegend
X: 120
Y: 0
Width: Parent.Width - 130
Height: 40
Items: Table(
    {Status: "Scheduled", Color: gblStatusColors.Scheduled},
    {Status: "Awaiting approval", Color: gblStatusColors.AwaitingApproval},
    {Status: "Awaiting withdrawl approval", Color: gblStatusColors.AwaitingWithdrawl},
    {Status: "Completed", Color: gblStatusColors.Completed},
    {Status: "In progress", Color: gblStatusColors.InProgress},
    {Status: "Saved", Color: gblStatusColors.Saved}
)
TemplateSize: 180
Layout: Horizontal

// Inside galLegend template:
  // Status Label
  Text: ThisItem.Status
  X: 25
  Y: 10

  // Status Circle
  Name: cirStatus
  X: 5
  Y: 12
  Width: 16
  Height: 16
  Fill: ThisItem.Color
  BorderRadius: 8
```

### Week Header Gallery (Horizontal)

This gallery shows the week columns horizontally.

```powerFx
Name: galWeekHeaders
X: 450  // After the 3 fixed columns
Y: 220
Width: Parent.Width - 470
Height: 80
Items: colWeeks  // Collection of 52 weeks
Layout: Horizontal
TemplateSize: 60  // Each week column is 60px wide
TemplatePadding: 2

// Inside template:
  // Container
  Name: conWeekHeader
  Width: 60
  Height: 80
  Fill: If(
      ThisItem.HasHoliday,
      gblColors.HolidayWeek,
      RGBA(242, 242, 242, 1)
  )

  // Week Label
  Name: lblWeek
  Text: "Week " & ThisItem.WeekNumber & "\n(" & Text(ThisItem.WeekEndDate, "mm/dd") & ")"
  X: 0
  Y: 0
  Width: Parent.Width
  Height: Parent.Height
  Align: Center
  Font: Segoe UI, 9
  Wrap: true

  // Tooltip Icon (if has holiday)
  Name: icoHoliday
  Icon: Icon.Info
  Visible: ThisItem.HasHoliday
  X: 40
  Y: 5
  Width: 16
  Height: 16
  OnSelect:
      Set(varTooltipText, ThisItem.HolidayNames);
      Set(varShowTooltip, true)
```

### Fixed Column Headers

```powerFx
// Market Sub Sector Header
Name: lblMarketHeader
Text: "Market Sub Sector"
X: 0
Y: 220
Width: 150
Height: 80
Fill: RGBA(242, 242, 242, 1)
Align: Center
Font: Bold

// Department Header
Name: lblDeptHeader
Text: "Department"
X: 150
Y: 220
Width: 150
Height: 80
Fill: RGBA(242, 242, 242, 1)
Align: Center
Font: Bold

// Person Header
Name: lblPersonHeader
Text: "Person"
X: 300
Y: 220
Width: 150
Height: 80
Fill: RGBA(242, 242, 242, 1)
Align: Center
Font: Bold
```

---

## Hierarchical Gallery Pattern

### Challenge

PowerApps doesn't support truly nested galleries well. We need a "flattened" approach where each row knows its level and parent.

### Solution: Flat Table with Expand/Collapse State

**Data Structure:**
```powerFx
colWeeklyReportRows: [
    {
        RowType: "Market",
        Level: 1,
        MarketID: "North_America",
        MarketName: "North America",
        DepartmentID: "",
        PersonName: "",
        IsExpanded: true,
        WeekData: [...]
    },
    {
        RowType: "Department",
        Level: 2,
        MarketID: "North_America",
        MarketName: "North America",
        DepartmentID: "ABC-ENV-123",
        PersonName: "",
        IsExpanded: true,
        WeekData: [...]
    },
    {
        RowType: "Person",
        Level: 3,
        MarketID: "North_America",
        MarketName: "North America",
        DepartmentID: "ABC-ENV-123",
        PersonName: "Smith, John",
        IsExpanded: false,
        WeekData: [...]
    },
    ...
]
```

### Main Vertical Gallery

```powerFx
Name: galRows
X: 0
Y: 300
Width: 450  // Fixed columns width
Height: Parent.Height - 320
Items: Filter(
    colWeeklyReportRows,
    // Show row if:
    // 1. It's a market (always show)
    // 2. It's a department and its market is expanded
    // 3. It's a person and its department is expanded
    RowType = "Market" ||
    (RowType = "Department" && LookUp(colWeeklyReportRows, MarketID = ThisRecord.MarketID && RowType = "Market").IsExpanded) ||
    (RowType = "Person" && LookUp(colWeeklyReportRows, DepartmentID = ThisRecord.DepartmentID && RowType = "Department").IsExpanded)
)
TemplateSize: 40
TemplatePadding: 1

// Inside galRows template:

  // Background - conditional based on row type
  Name: recRowBackground
  Fill: Switch(
      ThisItem.RowType,
      "Market", gblColors.MarketHeader,
      "Department", gblColors.DeptHeader,
      "Person", If(CountRows(Filter(galRows.AllItems, RowType = "Person" && MarketID = ThisItem.MarketID)) Mod 2 = 0, gblColors.PersonEven, gblColors.PersonOdd),
      White
  )

  // Expand/Collapse Icon (only for Market and Department)
  Name: icoExpand
  Icon: If(ThisItem.IsExpanded, Icon.ChevronDown, Icon.ChevronRight)
  Visible: ThisItem.RowType <> "Person"
  X: 10 + (ThisItem.Level - 1) * 20  // Indent based on level
  Y: 10
  Width: 24
  Height: 24
  OnSelect:
      // Toggle expansion
      If(
          ThisItem.RowType = "Market",
          Patch(
              colWeeklyReportRows,
              LookUp(colWeeklyReportRows, MarketID = ThisItem.MarketID && RowType = "Market"),
              {IsExpanded: !ThisItem.IsExpanded}
          ),
          ThisItem.RowType = "Department",
          Patch(
              colWeeklyReportRows,
              LookUp(colWeeklyReportRows, DepartmentID = ThisItem.DepartmentID && RowType = "Department"),
              {IsExpanded: !ThisItem.IsExpanded}
          )
      )

  // Market Label (Column 1)
  Name: lblMarket
  Text: If(ThisItem.RowType = "Market", ThisItem.MarketName, "")
  X: 40
  Y: 0
  Width: 110
  Height: 40
  Font: If(ThisItem.RowType = "Market", Bold, Normal)

  // Department Label (Column 2)
  Name: lblDepartment
  Text: If(ThisItem.RowType = "Department", ThisItem.DepartmentID, "")
  X: 150
  Y: 0
  Width: 150
  Height: 40
  Font: If(ThisItem.RowType = "Department", Bold, Normal)

  // Person Label (Column 3)
  Name: lblPerson
  Text: If(ThisItem.RowType = "Person", ThisItem.PersonName, "")
  X: 300
  Y: 0
  Width: 150
  Height: 40
```

### Week Data Gallery (Horizontal - inside each row)

```powerFx
Name: galWeekData
X: 450
Y: 0
Width: Parent.Width - 450
Height: 40
Items: ThisItem.WeekData  // Array of 52 weeks for this row
Layout: Horizontal
TemplateSize: 60
TemplatePadding: 2

// Inside galWeekData template:

  // Cell Container
  Name: conWeekCell
  Width: 60
  Height: 40
  Fill: If(
      // Check if this week is a holiday
      LookUp(colWeeks, WeekNumber = ThisItem.WeekNumber).HasHoliday,
      gblColors.HolidayCell,
      Transparent
  )

  // Hours Label
  Name: lblHours
  Text: If(
      ThisItem.TotalHours = 0,
      "0",
      Text(Round(ThisItem.TotalHours * 2, 0) / 2, "0.0")
  )
  X: 5
  Y: 5
  Width: 40
  Height: 30
  Color: If(ThisItem.TotalHours = 0, gblColors.ZeroValue, Black)
  Font: If(galRows.Selected.RowType <> "Person", Bold, Normal)
  Align: Right

  // Status Indicators (colored circles) - only for Person and Department rows
  Name: galStatuses
  X: 45
  Y: 5
  Width: 10
  Height: 30
  Visible: galRows.Selected.RowType <> "Market"
  Items: If(
      !IsBlank(ThisItem.Statuses),
      Split(ThisItem.Statuses, ","),
      []
  )
  Layout: Vertical
  TemplateSize: 10

    // Circle inside galStatuses
    Name: cirStatus
    Width: 8
    Height: 8
    Fill: Switch(
        Trim(ThisItem.Value),
        "Scheduled", gblStatusColors.Scheduled,
        "Awaiting approval", gblStatusColors.AwaitingApproval,
        "Awaiting withdrawl approval", gblStatusColors.AwaitingWithdrawl,
        "Completed", gblStatusColors.Completed,
        "In progress", gblStatusColors.InProgress,
        "Saved", gblStatusColors.Saved,
        Black
    )
    BorderRadius: 4

  // Tooltip Icon
  Name: icoTooltip
  Icon: Icon.Info
  Visible: ThisItem.TotalHours > 0 && galRows.Selected.RowType = "Person"
  X: 50
  Y: 25
  Width: 12
  Height: 12
  OnSelect:
      // Show tooltip with daily breakdown
      Set(varTooltipData, ThisItem.Details);
      Set(varShowTooltip, true)
```

---

## Daily Report Screen

### Differences from Weekly Report

1. **Column width**: 80px instead of 60px (more space for day names)
2. **Weekend highlighting**: Grey background for Sat/Sun columns
3. **Holiday highlighting**: Yellow background (different from weekend)
4. **Date range**: Up to 14 days instead of 52 weeks
5. **Header format**: "Mon 1/6" instead of "Week 1 (01/02)"

### Day Header Gallery

```powerFx
Name: galDayHeaders
X: 450
Y: 220
Width: Parent.Width - 470
Height: 80
Items: colDays  // Collection of up to 14 days
Layout: Horizontal
TemplateSize: 80
TemplatePadding: 2

// Inside template:
  Name: conDayHeader
  Fill: If(
      ThisItem.IsHoliday,
      gblColors.HolidayWeek,
      ThisItem.IsWeekend,
      gblColors.Weekend,
      RGBA(242, 242, 242, 1)
  )

  Text: ThisItem.HeaderText  // "Mon 1/6"
  Align: Center
  Font: Segoe UI, 10, If(ThisItem.IsWeekend, Italic, Normal)
```

### Day Data Gallery (Horizontal)

```powerFx
Name: galDayData
X: 450
Y: 0
Width: Parent.Width - 450
Height: 40
Items: ThisItem.DayData  // Array of daily hours for this row
Layout: Horizontal
TemplateSize: 80
TemplatePadding: 2

// Inside template:
  Fill: If(
      // Check if this day is a holiday
      LookUp(colDays, AbsDate = ThisItem.AbsDate).IsHoliday,
      gblColors.HolidayCell,
      LookUp(colDays, AbsDate = ThisItem.AbsDate).IsWeekend,
      gblColors.Weekend,
      Transparent
  )

  // Hours display (same as weekly)
  Text: If(ThisItem.TotalHours = 0, "0", Text(Round(ThisItem.TotalHours * 2, 0) / 2, "0.0"))
```

---

## Export Button

### Using Power Automate Flow

```powerFx
Name: btnExport
Text: "Export to Excel"
OnSelect:
    // Prepare data for export
    Set(varExportData, JSON(colWeeklyReportRows, JSONFormat.IncludeBinaryData));

    // Call Power Automate flow
    Set(varExportResult, 'ExportToExcel'.Run(varExportData, "Weekly Report"));

    // Show download link
    If(
        !IsBlank(varExportResult.FileUrl),
        Notify("Export complete! Downloading...", NotificationType.Success);
        Launch(varExportResult.FileUrl),
        Notify("Export failed: " & varExportResult.Error, NotificationType.Error)
    )
```

**See POWERAPP_POWER_AUTOMATE.md for the complete flow design.**

---

## Tooltip Implementation

### Challenge
PowerApps doesn't have native HTML tooltips. Need to create custom tooltip screen/panel.

### Solution: Pop-up Container

```powerFx
// Tooltip Container (on each screen)
Name: conTooltip
Visible: varShowTooltip
X: Min(galRows.Selected.X + 100, Parent.Width - 300)  // Position near click
Y: Min(galRows.Selected.Y, Parent.Height - 200)
Width: 280
Height: Min(CountRows(varTooltipData) * 30 + 40, 200)
Fill: White
BorderColor: RGBA(200, 200, 200, 1)
BorderThickness: 1
ZIndex: 1000

// Close Icon
Name: icoCloseTooltip
Icon: Icon.Cancel
X: Parent.Width - 30
Y: 5
Width: 24
Height: 24
OnSelect: Set(varShowTooltip, false)

// Tooltip Content Gallery
Name: galTooltipContent
X: 10
Y: 30
Width: Parent.Width - 20
Height: Parent.Height - 40
Items: varTooltipData  // Details array
TemplateSize: 30

  // Tooltip text label
  Text: ThisItem.AbsenceType & ": " & ThisItem.AbsenceComment &
        " (" & Text(Round(ThisItem.Hours * 2, 0) / 2, "0.0") & "h)"
  X: 5
  Y: 0
  Width: Parent.Width - 10
  Size: 10
  Wrap: true
```

---

## Responsive Design Tips

### Use Parent.Width and Parent.Height

```powerFx
// Gallery that fills remaining space
Width: Parent.Width - 40
Height: Parent.Height - galHeaders.Height - 100
```

### Breakpoints for Mobile

```powerFx
// Hide detailed columns on small screens
Visible: Parent.Width > 600
```

### Horizontal Scroll for Week/Day Columns

```powerFx
// Enable horizontal scrolling for many columns
galWeekData.Overflow: Horizontal
```

---

## Performance Optimization

### Virtualization

Use galleries instead of repeating controls - PowerApps virtualizes gallery items automatically.

### Limit Initial Load

```powerFx
// Start with markets collapsed
Set(varExpandAllMarkets, false);
Set(varExpandAllDepts, false);
```

### Use Delegation-Safe Filters

```powerFx
// Pre-filter data in collections instead of gallery Items property
ClearCollect(
    colFiltered,
    Filter(colWeeklyReportRows, ...)
);

// Then use simple gallery
galRows.Items: colFiltered
```

---

## Next Steps

1. ✅ Understand UI component structure
2. ➡️ Continue to **POWERAPP_STEP_BY_STEP.md** for complete build instructions
3. Review **POWERAPP_POWER_AUTOMATE.md** for Excel export flow
