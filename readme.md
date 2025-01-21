# Excel Generator WPF Application

A WPF application to generate an Excel report for a selected month and year. Supports Persian and Gregorian calendars, with highlighted holidays.

## Features
- **Persian & Gregorian Calendars**: Choose between the two calendar systems.
- **Holiday Highlighting**: Automatically colors holidays.
- **Work Schedule**: Generate a schedule with day, date, hours, and description.
- **Export to Excel**: Save the schedule as an Excel file.

## Requirements
- .NET Framework 4.7.2 or later
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) NuGet package

## Usage

1. **Select Month & Year**: Pick the desired month and year.
2. **Select Language**: Choose between Persian or English.
3. **Generate Excel**: Click "Generate Excel" to create and save the Excel file.

## Example Excel Layout

| Day of Week | Date       | Start Hour | End Hour | Difference | Description |
|-------------|------------|------------|----------|------------|-------------|
| Monday      | 2024-02-01 | 08:00      | 17:00    | 9:00       | Work day    |

## License
This project is licensed under the MIT License.

## Other versions
The outdated version of the application that supports .NET Framework 4.7.2 is available here:

[SarvenExcelGenerator v2](https://github.com/amin-norollah/sarvenExcelGenerator2)
