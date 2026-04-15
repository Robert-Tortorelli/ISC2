# ISC2

Projects supporting ISC2 chapter event coordination.

## Projects

### ISC2 Event Coordinator

A Windows console application that downloads attendee registration data from Constant Contact and generates:

- **Registration Spreadsheet** — CSV derived from the Constant Contact Attendee report
- **Registration List** — CSV of attendee names and email addresses
- **Attendance Sheet** — populated DOCX and PDF for sign-in at the event
- **Name Tags** — populated DOCX and PDF for printing attendee name tags

#### Requirements

- Windows 10 or later
- Microsoft Word (for PDF conversion)
- CMake 3.15 or later
- MSVC (Visual Studio 2019 or later)

#### Build

```
cmake -S "ISC2 Event Coordinator" -B "ISC2 Event Coordinator/build"
cmake --build "ISC2 Event Coordinator/build" --config Release
```

#### Usage

Run `EventCoordinator.exe` from the project directory. The program will prompt for Constant Contact credentials or allow selection of an existing Registration Spreadsheet CSV.
