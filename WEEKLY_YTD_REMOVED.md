# YTD Section Removed from Weekly HTML Reports

## Change Summary

Removed Year-to-Date (YTD) sections from weekly HTML reports as they're not applicable for 7-day periods.

## What Was Changed

### HTML Report Generation
**Before**: Weekly reports included empty YTD sections with no data
**After**: Weekly reports show only the weekly period data - no YTD section at all

### Specific Changes

1. **YTD Data Processing** (lines 1549-1628)
   - Wrapped all YTD data aggregation in `if not self.weekly:`
   - Weekly reports skip YTD calculations entirely

2. **YTD Chart Generation** (lines 1642-1654)
   - YTD line charts only generated for monthly reports
   - Weekly reports set empty strings for YTD chart variables

3. **YTD HTML Section** (lines 1656-1732)
   - Generated as separate `ytd_section_html` variable
   - Conditionally included in main template
   - Weekly reports: empty string
   - Monthly reports: full YTD section with filters, charts, tables

4. **JavaScript Chart Init** (lines 1898-1947)
   - YTD chart initialization only for monthly reports
   - Weekly reports skip ytdLanguageChart and ytdUserChart creation
   - Conditional array building for image/chart IDs

### Result

**Weekly Reports Now Show**:
- ✅ Weekly period data (previous 7 days)
- ✅ Language breakdown charts and tables
- ✅ User productivity charts and tables  
- ✅ Interactive filters
- ❌ NO YTD sections (removed)

**Monthly Reports Still Show**:
- ✅ Monthly period data
- ✅ Year-to-Date sections with trends
- ✅ All original functionality intact

## File Size Impact

Weekly HTML reports are now ~30-40% smaller without the YTD sections.

## Testing

Run weekly report to verify:
```bash
python3 generate_report.py --weekly
```

Expected: HTML report shows only weekly data, no YTD heading or sections.

