# ISC2

Projects supporting ISC2 Chapters.

## Projects

### ISC2 Event Coordinator V2.0
EventCoordinator is a Windows command line program that generates Microsoft 365 (Office) and PDF documents for ISC2 events:

**Registration Documents Needed Before an Event:** EventCoordinator --DocReg

•	Attendance Sheets used to sign-in attendees

•	Attendee Name Tag Labels

•	Registration Lists used by building security

These are derived from Constant Contact event registration data obtained directly using the Constant Contact API or from Constant Contact Attendee Reports that have been downloaded manually.

**Attendance Documents Needed After an Event:** EventCoordinator --DocAttend

•	CPE Eligible Attendees Report submitted to ISC2 on behalf of attendees in order to give them ISC2 CPE credits for attending events

This is derived from post-meeting attendance data obtained from:
Attendance Sheets (after signing by a meeting’s in-person attendees)

Microsoft Teams Attendance data obtained directly using the Microsoft Graph API or from Teams Attendance files downloaded manually using Teams.

#### Build

```
cmake -S "ISC2 Event Coordinator" -B "ISC2 Event Coordinator/build"
cmake --build "ISC2 Event Coordinator/build" --config Release
```
