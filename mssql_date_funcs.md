| Function               | Description                                               | Return Type         | Simple Example                              |
|------------------------|-----------------------------------------------------------|---------------------|---------------------------------------------|
| `YEAR()`               | Extracts the year part of a date.                         | `INT`               | `SELECT YEAR('2024-09-16');`               |
| `MONTH()`              | Extracts the month part of a date.                        | `INT`               | `SELECT MONTH('2024-09-16');`              |
| `DAY()`                | Extracts the day part of a date.                          | `INT`               | `SELECT DAY('2024-09-16');`                |
| `DATEPART()`           | Returns a specific part of a date (e.g., year, month).    | `INT`               | `SELECT DATEPART(quarter, '2024-09-16');`  |
| `DATENAME()`           | Returns the name of a specific part of a date.            | `NVARCHAR`          | `SELECT DATENAME(month, '2024-09-16');`    |
| `GETDATE()`            | Returns the current date and time.                        | `DATETIME`          | `SELECT GETDATE();`                        |
| `SYSDATETIME()`        | Returns the current system date and time with precision.  | `DATETIME2`         | `SELECT SYSDATETIME();`                    |
| `DATEADD()`            | Adds an interval to a date (e.g., days, months).          | `DATETIME`/`DATE`   | `SELECT DATEADD(day, 10, '2024-09-16');`   |
| `DATEDIFF()`           | Returns the difference between two dates in units.        | `BIGINT`            | `SELECT DATEDIFF(day, '2024-01-01', GETDATE());` |
| `FORMAT()`             | Formats a date as a string.                               | `NVARCHAR`          | `SELECT FORMAT(GETDATE(), 'MM/dd/yyyy');`  |
| `EOMONTH()`            | Returns the last day of the month.                        | `DATE`              | `SELECT EOMONTH('2024-09-16');`            |
| `SWITCHOFFSET()`       | Adjusts a date with a time zone offset.                   | `DATETIMEOFFSET`    | `SELECT SWITCHOFFSET(SYSDATETIMEOFFSET(), '-05:00');` |
| `CONVERT()`            | Converts a date to a specific format.                     | `VARCHAR`/`NVARCHAR`| `SELECT CONVERT(VARCHAR(10), GETDATE(), 120);` |
| `ISDATE()`             | Checks if a value is a valid date.                        | `INT`               | `SELECT ISDATE('2024-09-16');`             |
| `GETUTCDATE()`         | Returns the current UTC date and time.                    | `DATETIME`          | `SELECT GETUTCDATE();`                     |
| `SYSDATETIMEOFFSET()`  | Returns the current date and time with time zone.         | `DATETIMEOFFSET`    | `SELECT SYSDATETIMEOFFSET();`              |
