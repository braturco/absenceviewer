# PowerApps Absence Viewer - Formula Reference

## Table of Contents
1. [Week Number Calculation](#week-number-calculation)
2. [Working Days Distribution](#working-days-distribution)
3. [Data Processing Pipeline](#data-processing-pipeline)
4. [Grouping and Aggregation](#grouping-and-aggregation)
5. [Weekly Report Formulas](#weekly-report-formulas)
6. [Daily Report Formulas](#daily-report-formulas)
7. [Helper Functions](#helper-functions)

---

## Week Number Calculation

### Friday-Ending Week Calculation

In your system, weeks end on Friday, and week 1 of 2026 ends on January 2, 2026 (Friday).

**Custom Function Approach:**

Since PowerApps doesn't support custom functions like JavaScript, we'll create a reusable formula using `With()`:

```powerFx
// Calculate week number for a given date
// Usage: Replace _dateInput with your date variable/field
With(
    {
        _week1End: Date(2026, 1, 2),  // Jan 2, 2026 (Friday)
        _inputDate: _dateInput,
        _inputFriday: DateAdd(
            _inputDate,
            (5 - Weekday(_inputDate, StartOfWeek.Monday)) Mod 7,
            TimeUnit.Days
        )
    },
    With(
        {
            _diffDays: DateDiff(_week1End, _inputFriday, TimeUnit.Days),
            _diffWeeks: Round(_diffDays / 7, 0)
        },
        If(
            1 + _diffWeeks <= 0,
            1 + _diffWeeks + 52,
            1 + _diffWeeks
        )
    )
)
```

**Simplified Version (Global Variable):**

Set this in App.OnStart:
```powerFx
Set(gblWeek1EndDate, Date(2026, 1, 2)); // Friday, Jan 2, 2026
```

Then use this formula wherever you need week number:
```powerFx
// For a given date (_thisDate)
With(
    {
        _friday: DateAdd(
            _thisDate,
            (5 - Weekday(_thisDate, StartOfWeek.Monday)) Mod 7,
            TimeUnit.Days
        )
    },
    1 + Round(DateDiff(gblWeek1EndDate, _friday, TimeUnit.Days) / 7, 0)
)
```

**Explanation:**
1. Find the Friday of the week containing `_thisDate`
2. Calculate days between Week 1's Friday (Jan 2, 2026) and this Friday
3. Divide by 7 to get week offset
4. Add 1 to get week number (Week 1 = 1, not 0)

### Get Friday for a Given Date

```powerFx
// Returns the Friday of the week containing _thisDate
DateAdd(
    _thisDate,
    (5 - Weekday(_thisDate, StartOfWeek.Monday)) Mod 7,
    TimeUnit.Days
)
```

**Explanation:**
- `Weekday(_thisDate, StartOfWeek.Monday)` returns 1-7 (Mon=1, Fri=5, Sun=7)
- `(5 - Weekday(...))` gives days until Friday
- `Mod 7` ensures non-negative result
- `DateAdd` adds those days to get Friday

---

## Working Days Distribution

### Calculate Working Days Between Two Dates

PowerApps doesn't have a built-in "working days" function, so we need to calculate it:

```powerFx
// For a given start and end date, count working days (Mon-Fri)
// Returns a table of working days

With(
    {
        _start: _startDate,
        _end: _endDate,
        _totalDays: DateDiff(_startDate, _endDate, TimeUnit.Days) + 1
    },
    Filter(
        Sequence(_totalDays, 0),
        With(
            {
                _currentDay: DateAdd(_start, Value, TimeUnit.Days),
                _dayOfWeek: Weekday(_currentDay, StartOfWeek.Monday)
            },
            _dayOfWeek >= 1 && _dayOfWeek <= 5  // Mon-Fri only
        )
    )
)
```

**Result:** A table with column `Value` containing offsets from start date (0, 1, 2, ...) for working days only

### Convert Days to Hours Based on Normal Working Hours

```powerFx
// Convert absence duration from Days to Hours
With(
    {
        _normalHours: 40,  // or from field: _record.NormalWorkingHours
        _hoursPerDay: _normalHours / 5,
        _duration: _record.AbsenceDuration,
        _uom: _record.UOM
    },
    If(
        _uom = "D",
        _duration * _hoursPerDay,  // Days to Hours
        _duration  // Already in Hours
    )
)
```

### Round Hours to Nearest 0.5

```powerFx
Round(_hours * 2, 0) / 2
```

---

## Data Processing Pipeline

### Step 1: Load and Filter Raw Data

```powerFx
// App.OnStart - Load absence data
ClearCollect(
    colAbsenceRaw,
    Filter(
        AbsenceData,
        AbsenceStatus <> "Withdrawn" &&
        AbsenceStatus <> "Denied" &&
        !IsBlank(Department) &&
        Department contains "-ENV-"  // Filter departments
    )
);
```

### Step 2: Expand Absences to Daily Records

This is the most complex step - we need to "explode" each absence into individual days.

```powerFx
// Create a collection with one row per working day per absence
ClearCollect(
    colDailyAbsences,
    // For each absence record
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
            // Generate sequence of days
            With(
                {
                    _workingDays: Filter(
                        AddColumns(
                            Sequence(_totalDays, 0),
                            "AbsDate",
                            DateAdd(_absence.AbsenceStartDate, Value, TimeUnit.Days),
                            "DayOfWeek",
                            Weekday(DateAdd(_absence.AbsenceStartDate, Value, TimeUnit.Days), StartOfWeek.Monday)
                        ),
                        DayOfWeek >= 1 && DayOfWeek <= 5  // Mon-Fri only
                    ),
                    _workingDayCount: CountRows(_workingDays)
                },
                // For each working day, create a record
                AddColumns(
                    _workingDays,
                    "LastName", _absence.LastName,
                    "FirstName", _absence.FirstName,
                    "Department", _absence.Department,
                    "LineManager", _absence.LineManager,
                    "AbsenceType", _absence.AbsenceType,
                    "AbsenceStatus", _absence.AbsenceStatus,
                    "AbsenceComment", _absence.AbsenceComment,
                    "HoursThisDay", _totalHours / _workingDayCount,  // Evenly distribute
                    "DeptID", Left(_absence.Department, 8),
                    "MarketSubSector", LookUp(
                        colOrgData,
                        OfficeLocationID = Left(_absence.Department, 8),
                        MarketSubSectorName
                    ),
                    "WeekNumber", // Calculate week number for this day
                    With(
                        {
                            _friday: DateAdd(
                                AbsDate,
                                (5 - DayOfWeek) Mod 7,
                                TimeUnit.Days
                            )
                        },
                        1 + Round(DateDiff(gblWeek1EndDate, _friday, TimeUnit.Days) / 7, 0)
                    )
                )
            )
        )
    )
);
```

**What this does:**
1. For each absence record in `colAbsenceRaw`
2. Calculate total days and total hours (converting days to hours if needed)
3. Generate a sequence of all calendar days in the absence period
4. Filter to working days only (Mon-Fri)
5. Count working days
6. For each working day:
   - Create a record with all absence details
   - Assign hours evenly distributed: `totalHours / workingDayCount`
   - Calculate week number for that day
   - Look up Market Sub Sector from organization data

### Step 3: Add Week Numbers to Daily Data

If you didn't add week numbers in Step 2, add them separately:

```powerFx
ClearCollect(
    colDailyAbsencesWithWeeks,
    AddColumns(
        colDailyAbsences,
        "WeekNumber",
        With(
            {
                _friday: DateAdd(
                    AbsDate,
                    (5 - Weekday(AbsDate, StartOfWeek.Monday)) Mod 7,
                    TimeUnit.Days
                )
            },
            1 + Round(DateDiff(gblWeek1EndDate, _friday, TimeUnit.Days) / 7, 0)
        )
    )
);
```

---

## Grouping and Aggregation

### Group by Market → Department → Person → Week

```powerFx
// Create hierarchical structure
ClearCollect(
    colGroupedData,
    GroupBy(
        colDailyAbsencesWithWeeks,
        "MarketSubSector",
        "Markets",
        GroupBy(
            Markets,
            "Department",
            "Departments",
            GroupBy(
                Departments,
                "LastName", "FirstName",
                "Persons",
                GroupBy(
                    Persons,
                    "WeekNumber",
                    "Weeks",
                    "TotalHours", Sum(Weeks, HoursThisDay)
                )
            )
        )
    )
);
```

**Note:** PowerApps doesn't support deeply nested GroupBy like this. Instead, use a different approach:

### Alternative: Flatten and Summarize

```powerFx
// Create a flat summary table
ClearCollect(
    colWeeklySummary,
    AddColumns(
        GroupBy(
            colDailyAbsencesWithWeeks,
            "MarketSubSector",
            "Department",
            "LastName",
            "FirstName",
            "LineManager",
            "WeekNumber",
            "GroupedRecords"
        ),
        "TotalHours", Sum(GroupedRecords, HoursThisDay),
        "Statuses", Concat(Distinct(GroupedRecords, AbsenceStatus), Value, ", "),
        "PersonName", LastName & ", " & FirstName
    )
);
```

**Result:** A table with columns:
- MarketSubSector
- Department
- PersonName (Last, First)
- LineManager
- WeekNumber
- TotalHours
- Statuses (comma-separated)

---

## Weekly Report Formulas

### Generate 52 Weeks from Start Date

```powerFx
// Generate table of 52 weeks starting from varStartDate
ClearCollect(
    colWeeks,
    AddColumns(
        Sequence(52, 0),  // 0 to 51
        "WeekOffset", Value,
        "WeekEndDate", DateAdd(
            // Find Friday on or after varStartDate
            DateAdd(
                varStartDate,
                (5 - Weekday(varStartDate, StartOfWeek.Monday) + 7) Mod 7,
                TimeUnit.Days
            ),
            Value * 7,
            TimeUnit.Days
        )
    )
);

// Add week numbers
ClearCollect(
    colWeeks,
    AddColumns(
        colWeeks,
        "WeekNumber",
        With(
            {
                _friday: WeekEndDate
            },
            1 + Round(DateDiff(gblWeek1EndDate, _friday, TimeUnit.Days) / 7, 0)
        ),
        "MonthName", Text(WeekEndDate, "mmm yyyy"),
        "HeaderText", "Week " & WeekNumber & "\n" & Text(WeekEndDate, "mm/dd")
    )
);
```

### Check if Week Has Holiday

```powerFx
// Add holiday indicator to weeks
ClearCollect(
    colWeeks,
    AddColumns(
        colWeeks,
        "HasHoliday", CountRows(
            Filter(
                colHolidayData,
                // Check if holiday date falls in this week
                HolidayDate >= DateAdd(WeekEndDate, -6, TimeUnit.Days) &&
                HolidayDate <= WeekEndDate
            )
        ) > 0,
        "HolidayNames", Concat(
            Filter(
                colHolidayData,
                HolidayDate >= DateAdd(WeekEndDate, -6, TimeUnit.Days) &&
                HolidayDate <= WeekEndDate
            ),
            Holiday,
            ", "
        )
    )
);
```

### Build Weekly Report Data Structure

```powerFx
// Create collection for weekly report
// This creates a denormalized table for easy gallery binding

// Step 1: Get unique markets
ClearCollect(
    colMarkets,
    Distinct(colWeeklySummary, MarketSubSector)
);

// Step 2: For each market, get departments
ClearCollect(
    colWeeklyReportRows,
    ForAll(
        colMarkets As _market,
        With(
            {
                _depts: Distinct(
                    Filter(colWeeklySummary, MarketSubSector = _market.Value),
                    Department
                )
            },
            // Market header row
            {
                RowType: "Market",
                MarketSubSector: _market.Value,
                Department: "",
                PersonName: "",
                LineManager: "",
                IsExpanded: true,
                Level: 1,
                // Calculate totals for all 52 weeks
                WeekData: ForAll(
                    colWeeks As _week,
                    {
                        WeekNumber: _week.WeekNumber,
                        TotalHours: Sum(
                            Filter(
                                colWeeklySummary,
                                MarketSubSector = _market.Value &&
                                WeekNumber = _week.WeekNumber
                            ),
                            TotalHours
                        )
                    }
                )
            },
            // Department rows for this market
            ForAll(
                _depts As _dept,
                With(
                    {
                        _persons: Distinct(
                            Filter(
                                colWeeklySummary,
                                MarketSubSector = _market.Value &&
                                Department = _dept.Value
                            ),
                            PersonName
                        )
                    },
                    // Department header row
                    {
                        RowType: "Department",
                        MarketSubSector: _market.Value,
                        Department: _dept.Value,
                        PersonName: "",
                        LineManager: "",
                        IsExpanded: true,
                        Level: 2,
                        WeekData: ForAll(
                            colWeeks As _week,
                            {
                                WeekNumber: _week.WeekNumber,
                                TotalHours: Sum(
                                    Filter(
                                        colWeeklySummary,
                                        MarketSubSector = _market.Value &&
                                        Department = _dept.Value &&
                                        WeekNumber = _week.WeekNumber
                                    ),
                                    TotalHours
                                )
                            }
                        )
                    },
                    // Person rows for this department
                    ForAll(
                        _persons As _person,
                        {
                            RowType: "Person",
                            MarketSubSector: _market.Value,
                            Department: _dept.Value,
                            PersonName: _person.Value,
                            LineManager: First(
                                Filter(
                                    colWeeklySummary,
                                    PersonName = _person.Value &&
                                    Department = _dept.Value
                                )
                            ).LineManager,
                            IsExpanded: false,
                            Level: 3,
                            WeekData: ForAll(
                                colWeeks As _week,
                                {
                                    WeekNumber: _week.WeekNumber,
                                    TotalHours: Sum(
                                        Filter(
                                            colWeeklySummary,
                                            PersonName = _person.Value &&
                                            Department = _dept.Value &&
                                            WeekNumber = _week.WeekNumber
                                        ),
                                        TotalHours
                                    ),
                                    Statuses: First(
                                        Filter(
                                            colWeeklySummary,
                                            PersonName = _person.Value &&
                                            Department = _dept.Value &&
                                            WeekNumber = _week.WeekNumber
                                        )
                                    ).Statuses,
                                    Details: Filter(
                                        colDailyAbsencesWithWeeks,
                                        PersonName = _person.Value &&
                                        Department = _dept.Value &&
                                        WeekNumber = _week.WeekNumber
                                    )
                                }
                            )
                        }
                    )
                )
            )
        )
    )
);
```

**Warning:** The above formula creates VERY deeply nested structures. PowerApps may struggle with this.

**Simpler Approach: Three Separate Collections**

```powerFx
// Collection 1: Market totals by week
ClearCollect(
    colMarketWeekly,
    AddColumns(
        GroupBy(
            colWeeklySummary,
            "MarketSubSector",
            "WeekNumber",
            "GroupedRecords"
        ),
        "TotalHours", Sum(GroupedRecords, TotalHours)
    )
);

// Collection 2: Department totals by week
ClearCollect(
    colDepartmentWeekly,
    AddColumns(
        GroupBy(
            colWeeklySummary,
            "MarketSubSector",
            "Department",
            "WeekNumber",
            "GroupedRecords"
        ),
        "TotalHours", Sum(GroupedRecords, TotalHours),
        "Statuses", Concat(Distinct(GroupedRecords, Statuses), Value, ", ")
    )
);

// Collection 3: Person totals by week (already have in colWeeklySummary)
```

---

## Daily Report Formulas

### Generate Days Between Start and End Date

```powerFx
// Generate table of days (up to 14)
With(
    {
        _dayCount: DateDiff(varDailyStartDate, varDailyEndDate, TimeUnit.Days) + 1
    },
    If(
        _dayCount > 14,
        Notify("Daily report limited to 14 days", NotificationType.Warning);
        ClearCollect(colDays, Blank()),
        ClearCollect(
            colDays,
            AddColumns(
                Sequence(_dayCount, 0),
                "DayOffset", Value,
                "AbsDate", DateAdd(varDailyStartDate, Value, TimeUnit.Days),
                "DayOfWeek", Weekday(DateAdd(varDailyStartDate, Value, TimeUnit.Days), StartOfWeek.Monday),
                "DayName", Text(DateAdd(varDailyStartDate, Value, TimeUnit.Days), "ddd"),
                "DateText", Text(DateAdd(varDailyStartDate, Value, TimeUnit.Days), "m/d"),
                "HeaderText", Text(DateAdd(varDailyStartDate, Value, TimeUnit.Days), "ddd m/d"),
                "IsWeekend", Weekday(DateAdd(varDailyStartDate, Value, TimeUnit.Days), StartOfWeek.Monday) >= 6,
                "IsHoliday", CountRows(
                    Filter(
                        colHolidayData,
                        HolidayDate = DateAdd(varDailyStartDate, Value, TimeUnit.Days)
                    )
                ) > 0
            )
        )
    )
);
```

### Build Daily Report Data Structure

```powerFx
// Create daily summary (similar to weekly, but by date instead of week)
ClearCollect(
    colDailySummary,
    AddColumns(
        GroupBy(
            Filter(
                colDailyAbsencesWithWeeks,
                AbsDate >= varDailyStartDate && AbsDate <= varDailyEndDate
            ),
            "MarketSubSector",
            "Department",
            "LastName",
            "FirstName",
            "LineManager",
            "AbsDate",
            "GroupedRecords"
        ),
        "TotalHours", Sum(GroupedRecords, HoursThisDay),
        "Statuses", Concat(Distinct(GroupedRecords, AbsenceStatus), Value, ", "),
        "PersonName", LastName & ", " & FirstName
    )
);
```

---

## Helper Functions

### Format Hours Display

```powerFx
// Round to nearest 0.5 and format
Text(Round(_hours * 2, 0) / 2, "0.0")
```

### Get Status Color

```powerFx
// Return color based on status
Switch(
    _status,
    "Scheduled", RGBA(0, 128, 0, 1),  // green
    "Awaiting approval", RGBA(255, 255, 0, 1),  // yellow
    "Awaiting withdrawl approval", RGBA(255, 165, 0, 1),  // orange
    "Completed", RGBA(0, 0, 255, 1),  // blue
    "In progress", RGBA(128, 0, 128, 1),  // purple
    "Saved", RGBA(128, 128, 128, 1),  // gray
    RGBA(0, 0, 0, 1)  // default black
)
```

### Check if Date is Working Day

```powerFx
// Returns true if Mon-Fri
With(
    {_dow: Weekday(_date, StartOfWeek.Monday)},
    _dow >= 1 && _dow <= 5
)
```

---

## Complete App.OnStart Formula

Putting it all together:

```powerFx
// ===== APP INITIALIZATION =====

// Set global constants
Set(gblWeek1EndDate, Date(2026, 1, 2)); // Week 1 ends Friday, Jan 2, 2026
Set(gblCurrentYear, 2026);

// Set default date ranges
Set(varStartDate, Date(2026, 1, 2));
Set(varEndDate, Date(2026, 12, 31));
Set(varDailyStartDate, Today());
Set(varDailyEndDate, DateAdd(Today(), 6, TimeUnit.Days));

// ===== LOAD DATA FROM SHAREPOINT =====

// Load absence data (filtered)
ClearCollect(
    colAbsenceRaw,
    Filter(
        AbsenceData,
        AbsenceStatus <> "Withdrawn" &&
        AbsenceStatus <> "Denied" &&
        !IsBlank(Department) &&
        Department contains "-ENV-"
    )
);

// Load organization data
ClearCollect(colOrgData, OrganizationData);

// Load holiday data
ClearCollect(colHolidayData, HolidayCalendar);

// Set status
Set(varDataLoaded, true);
Set(varRecordCount, CountRows(colAbsenceRaw));

// (Do NOT process reports yet - wait for user to click "Generate Report")
```

---

## Performance Optimization Tips

1. **Use Concurrent() for independent operations:**
```powerFx
Concurrent(
    ClearCollect(colAbsenceRaw, Filter(...)),
    ClearCollect(colOrgData, OrganizationData),
    ClearCollect(colHolidayData, HolidayCalendar)
);
```

2. **Cache expensive calculations:**
```powerFx
// Calculate once, reuse many times
Set(varTotalHours, Sum(colWeeklySummary, TotalHours));
```

3. **Use With() to avoid recalculating:**
```powerFx
With(
    {_hours: Sum(Filter(...), HoursThisDay)},
    Text(_hours, "0.0") & " hours (" & Text(_hours/8, "0.0") & " days)"
)
```

---

## Next Steps

Continue to **POWERAPP_UI_COMPONENTS.md** to see how these formulas are used in screen controls and galleries.
