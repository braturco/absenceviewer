# PowerApps Absence Viewer - Complete Documentation

## 📋 Overview

This documentation set provides everything you need to recreate your HTML/JavaScript absence viewer as a Microsoft PowerApps Canvas App. The PowerApps version will integrate seamlessly with your Microsoft 365 environment while maintaining all the functionality of your current application.

---

## 🎯 What This App Does

### Current Features (HTML/JS Version)
✅ Upload and parse 3 CSV files (Absence, Organization, Holiday)
✅ Calculate week numbers with Friday-ending weeks
✅ Distribute absence hours evenly across working days (Mon-Fri)
✅ Handle both "Hours" and "Days" units of measure
✅ Filter out "Withdrawn" and "Denied" statuses
✅ Generate weekly reports (52 weeks) with hierarchical grouping
✅ Generate daily reports (up to 14 days) with weekend/holiday highlighting
✅ Expand/collapse functionality for markets and departments
✅ Status indicators (colored circles)
✅ Tooltips with daily breakdowns
✅ Export to Excel with formatting

### PowerApps Version Will Provide
✅ All features from HTML/JS version
✅ Better integration with Microsoft 365
✅ Direct connection to SharePoint lists (no CSV parsing needed)
✅ Role-based access control
✅ Mobile-friendly responsive design
✅ Automated data refresh
✅ Enhanced security and compliance

---

## 📚 Documentation Files

### 1. **POWERAPP_OVERVIEW.md** - Start Here!
**Read this first to understand the overall architecture.**

**Contents:**
- High-level architecture diagram
- Data flow explanation
- Key differences from HTML/JS version
- Prerequisites and requirements
- Implementation approach options

**Time to read:** 15 minutes

---

### 2. **POWERAPP_DATA_SETUP.md** - SharePoint Configuration
**Follow this to set up your data sources.**

**Contents:**
- SharePoint lists creation (3 lists)
- Column definitions and settings
- Data connections in PowerApps
- Initial data load procedures
- Migration from CSV to SharePoint

**Time to complete:** 1-2 hours

---

### 3. **POWERAPP_FORMULAS.md** - Calculation Logic
**Reference this while building your app for all Power Fx formulas.**

**Contents:**
- Week number calculation (Friday-ending weeks)
- Working days distribution (Mon-Fri only)
- Hour/day conversion formulas
- Data processing pipeline
- Grouping and aggregation logic
- Complete App.OnStart formula

**Complexity:** Advanced - contains complex nested formulas

---

### 4. **POWERAPP_UI_COMPONENTS.md** - Screen Designs
**Use this as your UI blueprint.**

**Contents:**
- Screen layouts (Home, Weekly Report, Daily Report)
- Control specifications (position, size, properties)
- Color palette definitions
- Hierarchical gallery pattern
- Status legend component
- Tooltip implementation

**Time to build:** 3-4 hours

---

### 5. **POWERAPP_STEP_BY_STEP.md** - Build Instructions
**Follow this guide from start to finish to build the complete app.**

**Contents:**
- 10 phases with detailed steps
- Verification checkpoints
- Testing procedures
- Deployment instructions
- Troubleshooting guide
- Completion checklist

**Time to complete:** 6-8 hours (initial build)

---

### 6. **POWERAPP_POWER_AUTOMATE.md** - Integration Flows
**Set up these flows for CSV import and Excel export.**

**Contents:**
- CSV parser flow (optional)
- Excel export flow (weekly report)
- Excel export flow (daily report)
- Office Scripts for Excel formatting
- Error handling and retry logic

**Time to complete:** 2-3 hours

---

### 7. **POWERAPP_SAMPLE_DATA.md** - Test Data
**Use this to generate or import test data.**

**Contents:**
- CSV file format specifications
- 30+ sample absence records
- Organization mapping samples
- Holiday calendar for 2026
- PowerShell scripts for data generation
- Import instructions

**Time to complete:** 30 minutes

---

## 🚀 Quick Start Guide

### Option A: Fastest Path (SharePoint Lists)

**Recommended for:** Quick deployment, easier maintenance

1. **Setup (1-2 hours)**
   - Read: POWERAPP_OVERVIEW.md
   - Create SharePoint lists: POWERAPP_DATA_SETUP.md
   - Import sample data: POWERAPP_SAMPLE_DATA.md

2. **Build (4-6 hours)**
   - Follow: POWERAPP_STEP_BY_STEP.md (Phases 1-7)
   - Reference: POWERAPP_FORMULAS.md and POWERAPP_UI_COMPONENTS.md as needed

3. **Test & Deploy (1-2 hours)**
   - Follow: POWERAPP_STEP_BY_STEP.md (Phases 9-10)

**Total Time: 6-10 hours**

---

### Option B: CSV Workflow (Keep Existing Process)

**Recommended for:** Maintaining current CSV-based workflow

1. **Setup (2-3 hours)**
   - Read: POWERAPP_OVERVIEW.md
   - Create SharePoint lists: POWERAPP_DATA_SETUP.md
   - Create CSV parser flow: POWERAPP_POWER_AUTOMATE.md

2. **Build (4-6 hours)**
   - Follow: POWERAPP_STEP_BY_STEP.md (Phases 1-7)
   - Reference: POWERAPP_FORMULAS.md and POWERAPP_UI_COMPONENTS.md

3. **Integration (2-3 hours)**
   - Set up CSV upload workflow: POWERAPP_POWER_AUTOMATE.md
   - Test CSV import functionality

4. **Test & Deploy (1-2 hours)**
   - Follow: POWERAPP_STEP_BY_STEP.md (Phases 9-10)

**Total Time: 9-14 hours**

---

## 📊 Key Features Breakdown

### 1. Week Number Calculation

**Challenge:** Weeks end on Friday, week 1 of 2026 ends Jan 2, 2026

**Solution:**
```powerFx
// Set in App.OnStart
Set(gblWeek1EndDate, Date(2026, 1, 2));

// Calculate week number for any date
With(
    {_friday: DateAdd(_date, Mod(5 - Weekday(_date, StartOfWeek.Monday) + 7, 7), TimeUnit.Days)},
    1 + Round(DateDiff(gblWeek1EndDate, _friday, TimeUnit.Days) / 7, 0)
)
```

📖 **See:** POWERAPP_FORMULAS.md → "Week Number Calculation"

### 2. Distributing Hours Across Working Days

**Challenge:** Evenly distribute absence hours across Mon-Fri only

**Solution:**
```powerFx
// For each absence, generate working days only
Filter(
    Sequence(DateDiff(StartDate, EndDate, Days) + 1, 0),
    Weekday(DateAdd(StartDate, Value, Days), StartOfWeek.Monday) <= 5
)

// Then divide total hours by working day count
HoursPerDay = TotalHours / CountRows(WorkingDays)
```

📖 **See:** POWERAPP_FORMULAS.md → "Working Days Distribution"

### 3. Hierarchical Display

**Challenge:** PowerApps doesn't support nested galleries natively

**Solution:** Flat table with expand/collapse state
```powerFx
// Each row knows its level and visibility state
{
    RowType: "Market" | "Department" | "Person",
    Level: 1 | 2 | 3,
    IsExpanded: true | false,
    ParentID: "...",
    Data: [...]
}

// Gallery filters based on parent expansion state
Filter(
    colRows,
    RowType = "Market" ||
    (RowType = "Department" && Parent.IsExpanded) ||
    (RowType = "Person" && Department.IsExpanded)
)
```

📖 **See:** POWERAPP_UI_COMPONENTS.md → "Hierarchical Gallery Pattern"

### 4. Excel Export with Formatting

**Challenge:** PowerApps can't directly create formatted Excel files

**Solution:** Power Automate + Office Scripts
```javascript
// Office Script runs in Excel Online
function main(workbook, reportData) {
    // Parse JSON data
    // Create sheets per market
    // Add headers with month groupings
    // Apply cell formatting (colors, borders)
    // Return success
}
```

📖 **See:** POWERAPP_POWER_AUTOMATE.md → "Flow 2: Export to Excel"

---

## ⚠️ Important Considerations

### Data Delegation Limit

**Issue:** PowerApps has a 2,000 record delegation limit
**Impact:** With ~2,000 employees, you may hit this limit
**Solution:** Use `ClearCollect` to load all data into memory first

```powerFx
// Instead of direct Filter on data source
Filter(AbsenceData, ...)  // ❌ May hit delegation limit

// Use ClearCollect first
ClearCollect(colAbsence, AbsenceData);  // ✅ Loads all into memory
Filter(colAbsence, ...)  // ✅ No delegation limit
```

📖 **See:** POWERAPP_DATA_SETUP.md → "Performance Considerations"

### Complex Formula Performance

**Issue:** Nested ForAll loops with 2,000+ records can be slow
**Impact:** Data processing may take 5-10 seconds
**Solution:**
- Show loading spinner during processing
- Process data only when user clicks "Generate Report"
- Cache processed results in collections

📖 **See:** POWERAPP_FORMULAS.md → "Performance Optimization Tips"

### Excel Export Limitations

**Issue:** Office Scripts have execution time limits (typically 30-60 seconds)
**Impact:** Very large reports may timeout
**Solution:**
- Export one market per worksheet
- Process in batches if needed
- Simplify formatting for large datasets

📖 **See:** POWERAPP_POWER_AUTOMATE.md → "Error Handling"

---

## 🛠️ Prerequisites Checklist

Before starting, ensure you have:

**Access & Licenses:**
- [ ] Microsoft 365 account with PowerApps license
- [ ] SharePoint Online access with site creation permissions
- [ ] Power Automate access (included with PowerApps)
- [ ] Permissions to create SharePoint lists

**Knowledge:**
- [ ] Basic familiarity with PowerApps Studio
- [ ] Understanding of your current absence data structure
- [ ] Access to existing CSV files for data migration

**Time:**
- [ ] 6-10 hours for initial build (Option A)
- [ ] 9-14 hours for CSV workflow (Option B)
- [ ] Additional 2-3 hours for testing and refinement

---

## 📈 Implementation Phases

### Phase 1: Planning (1 hour)
- [ ] Read POWERAPP_OVERVIEW.md
- [ ] Choose implementation approach (SharePoint Lists vs CSV)
- [ ] Review sample data in POWERAPP_SAMPLE_DATA.md
- [ ] Identify stakeholders and access requirements

### Phase 2: Data Setup (1-3 hours)
- [ ] Create SharePoint site (if needed)
- [ ] Create 3 SharePoint lists using POWERAPP_DATA_SETUP.md
- [ ] Import or create sample data
- [ ] Verify data structure and relationships

### Phase 3: App Foundation (2 hours)
- [ ] Create blank PowerApps canvas app
- [ ] Connect to SharePoint data sources
- [ ] Set up App.OnStart with formulas from POWERAPP_FORMULAS.md
- [ ] Build Home Screen using POWERAPP_UI_COMPONENTS.md

### Phase 4: Data Processing (2 hours)
- [ ] Implement absence expansion logic
- [ ] Create weekly/daily summary collections
- [ ] Test with sample data
- [ ] Verify calculations match expected results

### Phase 5: Weekly Report Screen (2 hours)
- [ ] Build screen layout
- [ ] Create week header gallery
- [ ] Build hierarchical data gallery
- [ ] Implement expand/collapse functionality
- [ ] Test with real data

### Phase 6: Daily Report Screen (1-2 hours)
- [ ] Build screen layout (similar to weekly)
- [ ] Create day header gallery
- [ ] Add weekend/holiday highlighting
- [ ] Test with various date ranges

### Phase 7: Power Automate Flows (2-3 hours)
- [ ] Create Excel export flow (weekly)
- [ ] Create Excel export flow (daily)
- [ ] Create Office Scripts for formatting
- [ ] Optional: Create CSV import flow
- [ ] Test all flows end-to-end

### Phase 8: Testing (1-2 hours)
- [ ] Unit test each component
- [ ] Integration test complete workflows
- [ ] Performance test with production data
- [ ] User acceptance testing

### Phase 9: Deployment (1 hour)
- [ ] Publish app
- [ ] Share with users
- [ ] Create user documentation
- [ ] Set up support process

### Phase 10: Monitoring & Iteration (Ongoing)
- [ ] Monitor usage and errors
- [ ] Collect user feedback
- [ ] Plan enhancements
- [ ] Schedule regular data refreshes

---

## 🎓 Learning Path

### New to PowerApps?

**Start with these resources:**
1. [PowerApps Documentation](https://docs.microsoft.com/powerapps) (Microsoft)
2. [Power Fx Formula Reference](https://docs.microsoft.com/powerapps/maker/canvas-apps/formula-reference)
3. Read POWERAPP_OVERVIEW.md for project context
4. Practice with POWERAPP_SAMPLE_DATA.md before using real data

**Key Concepts to Understand:**
- Collections (in-memory data tables)
- Galleries (repeating data display)
- Context variables vs Global variables
- Data delegation
- Power Fx formulas (similar to Excel)

### Experienced with PowerApps?

**Skip to:**
- POWERAPP_FORMULAS.md for calculation logic
- POWERAPP_DATA_SETUP.md for data structure
- POWERAPP_STEP_BY_STEP.md Phase 5+ for building

**Focus on:**
- Week calculation algorithm (unique to this app)
- Hierarchical gallery pattern (advanced technique)
- Performance optimization with 2,000+ records

---

## 🤝 Getting Help

### Documentation Order for Troubleshooting

1. **Error Messages:**
   - POWERAPP_STEP_BY_STEP.md → "Troubleshooting" section
   - POWERAPP_POWER_AUTOMATE.md → "Error Handling"

2. **Formula Issues:**
   - POWERAPP_FORMULAS.md → Find specific calculation
   - Check syntax in [Power Fx Reference](https://docs.microsoft.com/powerapps/maker/canvas-apps/formula-reference)

3. **UI Problems:**
   - POWERAPP_UI_COMPONENTS.md → Review control specifications
   - Verify data structure in collections

4. **Data Issues:**
   - POWERAPP_DATA_SETUP.md → Validate SharePoint list structure
   - POWERAPP_SAMPLE_DATA.md → Check data format

### Common Issues Quick Reference

| Issue | Solution | Reference |
|-------|----------|-----------|
| "Delegation warning" | Use ClearCollect first | POWERAPP_DATA_SETUP.md |
| Week numbers incorrect | Verify gblWeek1EndDate | POWERAPP_FORMULAS.md |
| Slow performance | Check formula nesting | POWERAPP_FORMULAS.md |
| Excel export fails | Check flow run history | POWERAPP_POWER_AUTOMATE.md |
| Data not loading | Refresh SharePoint connection | POWERAPP_STEP_BY_STEP.md |
| Expand/collapse broken | Check IsExpanded variable | POWERAPP_UI_COMPONENTS.md |

---

## 🔄 Migration from HTML/JS to PowerApps

### What Stays the Same
✅ Business logic (week calculation, hour distribution)
✅ Data structure (3 CSV files → 3 SharePoint lists)
✅ Report layouts (weekly 52 weeks, daily 14 days)
✅ Status colors and indicators
✅ Hierarchical grouping (Market → Dept → Person)

### What Changes
🔄 **File Upload** → SharePoint data connection
🔄 **localStorage** → SharePoint lists (persistent)
🔄 **JavaScript functions** → Power Fx formulas
🔄 **HTML tables** → PowerApps galleries
🔄 **CSS tooltips** → Custom pop-up components
🔄 **ExcelJS library** → Power Automate + Office Scripts

### Side-by-Side Comparison

| Feature | HTML/JS | PowerApps |
|---------|---------|-----------|
| **Data Input** | CSV file upload | SharePoint lists |
| **Storage** | Browser localStorage | Cloud (SharePoint) |
| **Processing** | Client-side JS | Power Fx formulas |
| **UI Framework** | HTML/CSS/DOM | PowerApps controls |
| **Styling** | CSS classes | Control properties |
| **Export** | ExcelJS (client) | Power Automate (server) |
| **Access** | Web browser | M365 ecosystem |
| **Authentication** | None (local files) | Azure AD (SSO) |
| **Mobile** | Not optimized | Responsive design |
| **Offline** | Works offline | Requires connection |

---

## 🎯 Success Criteria

Your PowerApps implementation is successful when:

**Functionality:**
- [ ] All 3 data sources loading correctly
- [ ] Week numbers calculating accurately (Friday-ending weeks)
- [ ] Hours distributed correctly across working days only
- [ ] Weekly report shows 52 weeks with proper grouping
- [ ] Daily report shows up to 14 days with weekend/holiday highlighting
- [ ] Expand/collapse works for markets and departments
- [ ] Status indicators display with correct colors
- [ ] Excel export creates properly formatted files

**Performance:**
- [ ] Report generation completes in under 10 seconds
- [ ] App loads in under 5 seconds
- [ ] Smooth scrolling through galleries
- [ ] Expand/collapse responds instantly

**User Experience:**
- [ ] Intuitive navigation
- [ ] Clear visual feedback (loading spinners, etc.)
- [ ] Helpful error messages
- [ ] Mobile-friendly (if needed)

**Maintenance:**
- [ ] Data easy to update (SharePoint lists)
- [ ] Code well-documented with comments
- [ ] Error handling for edge cases
- [ ] Admin documentation created

---

## 📦 Deliverables

When you complete this project, you will have:

1. **PowerApps Canvas App**
   - Home screen with data status
   - Weekly report screen (52 weeks)
   - Daily report screen (14 days)

2. **SharePoint Lists** (3 lists)
   - AbsenceData
   - OrganizationData
   - HolidayCalendar

3. **Power Automate Flows** (2-3 flows)
   - Excel export (weekly)
   - Excel export (daily)
   - Optional: CSV import

4. **Office Scripts** (for Excel formatting)

5. **Documentation**
   - User guide
   - Admin guide
   - Support procedures

---

## 🚦 Next Steps

### Ready to Start?

1. **First-Time Builders:**
   ```
   1. Read: POWERAPP_OVERVIEW.md (15 min)
   2. Setup: POWERAPP_DATA_SETUP.md (1-2 hours)
   3. Build: POWERAPP_STEP_BY_STEP.md (follow all phases)
   ```

2. **Experienced PowerApps Developers:**
   ```
   1. Skim: POWERAPP_OVERVIEW.md (5 min)
   2. Review: POWERAPP_FORMULAS.md (focus on week calculation)
   3. Build: POWERAPP_STEP_BY_STEP.md (skip basic steps)
   ```

### Questions to Answer Before Starting

1. **Data Source:**
   - [ ] Will you use SharePoint lists? (recommended)
   - [ ] Or keep CSV upload workflow? (more complex)

2. **Users:**
   - [ ] Who will use the app?
   - [ ] What permissions do they need?
   - [ ] How many concurrent users?

3. **Deployment:**
   - [ ] Single app for everyone?
   - [ ] Separate apps per department?
   - [ ] Embedded in SharePoint page?

4. **Timeline:**
   - [ ] When do you need this completed?
   - [ ] Time available per week?
   - [ ] Testing period needed?

---

## 📞 Support

For help with this documentation:
1. Review the specific documentation file for your issue
2. Check the "Troubleshooting" sections
3. Refer to Microsoft's official PowerApps documentation
4. Search PowerApps Community forums

---

## ✅ Documentation Checklist

Make sure you've reviewed:

- [ ] **POWERAPP_README.md** (this file) - Overview and navigation
- [ ] **POWERAPP_OVERVIEW.md** - Architecture and approach
- [ ] **POWERAPP_DATA_SETUP.md** - SharePoint configuration
- [ ] **POWERAPP_FORMULAS.md** - All calculation logic
- [ ] **POWERAPP_UI_COMPONENTS.md** - Screen designs
- [ ] **POWERAPP_STEP_BY_STEP.md** - Build instructions
- [ ] **POWERAPP_POWER_AUTOMATE.md** - Integration flows
- [ ] **POWERAPP_SAMPLE_DATA.md** - Test data and formats

**Good luck building your PowerApps Absence Viewer!** 🎉

---

*Last Updated: 2026-02-14*
*Documentation Version: 1.0*
*PowerApps Version: Canvas App*
