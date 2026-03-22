# RowMailer

スプレッドシートの行データが条件を満たしたら、テンプレートメールを自動送信するGoogle Sheets アドオン。

A Google Sheets Add-on that automatically sends template emails when row data meets your conditions.

## Features
- Condition-based email trigger (=, !=, >, <, contains, is not empty, ...)
- `{{ColumnHeader}}` placeholder replacement in subject & body
- Send log (timestamp written to log column, LockService duplicate prevention)
- CC / BCC (fixed address or column reference)
- Timer triggers: hourly / every 6 hours / daily 9:00
- Test send (first matching row to your own address)
- Free: 50/day | Pro: 500/day + unlimited template library

## Plans
| Feature | Free | Pro |
|---------|------|-----|
| Daily limit | 50 | 500 |
| Templates | 1 | Unlimited |
| Triggers | Yes | Yes |
| CC/BCC | Yes | Yes |

**Pro license key format**: `ROWMAILER-PRO-XXXXXXXXXXXXXXXX`

## MIT License
