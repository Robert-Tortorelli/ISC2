# ISC2

Projects supporting ISC2 chapter event coordination.

## Projects

### ISC2 Event Coordinator

A C++ Windows console application that uses Constant Contact attendee registration data and Microsoft Teams meeting attendance data to generate the following documents that support ISC2 Chapter Meetings:

- **Registration Spreadsheet** — CSV derived from the Constant Contact Attendee report
- **Registration List** — CSV of attendee names and email addresses required by building security
- **Attendance Sheet** — A PDF for printing
- **Name Tags** — A PDF for printing on name tag paper
- **ISC2 CPE Spreadsheet** — CSV of attendee information needed to submit CPE evidence to ISC2 for qualifying attendees

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

- **Running EventCoordinator’s OAuth2 Setup Wizard:** EventCoordinator --setup
- **Running EventCoordinator to generate documents:** EventCoordinator
