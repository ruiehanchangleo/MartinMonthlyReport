# Weekly Report Recipient Configuration

## Change Summary

Weekly reports now go only to Leo Chang, while monthly reports continue going to all 3 recipients.

## Configuration

### xtm_config.json

```json
{
  "email_recipients": [
    "BreaB@FamilySearch.org",
    "leo.chang@familysearch.org",
    "MG@familysearch.org"
  ],
  "weekly_recipients": [
    "leo.chang@familysearch.org"
  ],
  "error_recipients": [
    "leo.chang@familysearch.org"
  ]
}
```

## Recipient Lists

| Report Type | Recipients | Field Name |
|------------|-----------|------------|
| **Monthly Reports** | BreaB, Leo, MG (3 people) | `email_recipients` |
| **Weekly Reports** | Leo only (1 person) | `weekly_recipients` |
| **Error Alerts** | Leo only (1 person) | `error_recipients` |

## How It Works

The `send_email_via_outlook()` method checks the report type:

```python
if self.weekly:
    recipients = self.config.get('weekly_recipients', self.config['email_recipients'])
else:
    recipients = self.config['email_recipients']
```

- **Weekly mode**: Uses `weekly_recipients` (falls back to `email_recipients` if not set)
- **Monthly mode**: Uses `email_recipients`

## Schedule

### Monday 9:00 AM (Weekly)
- Report: Previous 7 days
- Recipient: `leo.chang@familysearch.org`
- Format: HTML + Excel
- Auto-send: Yes

### 1st of Month 9:00 AM (Monthly)  
- Report: Previous month + YTD
- Recipients: `BreaB@FamilySearch.org`, `leo.chang@familysearch.org`, `MG@familysearch.org`
- Format: HTML + Excel
- Auto-send: Yes

## Changing Recipients

### To Change Weekly Recipients

Edit `xtm_config.json`:
```json
"weekly_recipients": [
  "new.person@familysearch.org",
  "another.person@familysearch.org"
]
```

### To Change Monthly Recipients

Edit `xtm_config.json`:
```json
"email_recipients": [
  "person1@familysearch.org",
  "person2@familysearch.org",
  "person3@familysearch.org"
]
```

### To Change Error Alert Recipients

Edit `xtm_config.json`:
```json
"error_recipients": [
  "admin@familysearch.org"
]
```

## Fallback Behavior

If `weekly_recipients` is not defined in the config, weekly reports fall back to using `email_recipients` (all 3 people). This ensures backward compatibility.

## Testing

```bash
# Test weekly report (should go to Leo only)
python3 generate_report.py --weekly

# Test monthly report (should go to all 3)
python3 generate_report.py

# Check the output will show recipient count
```

Expected output:
- Weekly: "Email draft created with 1 recipient"
- Monthly: "Email draft created with 3 recipients"
