# ClassScheduleToICS

A bookmarklet script to export your Queen’s class schedule as an ICS file.

The script parses the SOLUS class schedule using a known layout, warning the user if anything could not be
exported (e.g., classes with a *TBA* timeslot). It produces an iCalendar file which can be imported into
programs such as Google Calendar and Outlook. It can also combine classes with duplicate listings (e.g., labs
held in multiple rooms simultaneously) into one calendar event.
[FileSaver.js][1] is used to create the download link.

This script was originally adapted from [the Chrome extension by Keanu Lee][2] for Waterloo’s self-service
system (also a PeopleSoft product). It has since been rewritten to offer cross-browser support,
valid iCalendar ([RFC 5545][3]) output with correct timezone handling, and additional features.

[1]: https://github.com/eligrey/FileSaver.js/
[2]: https://github.com/keanulee/ClassScheduleToICS
[3]: http://tools.ietf.org/html/rfc5545

## Use
A bookmarklet is provided which loads the main script from a server, allowing for updates.
Since SOLUS uses a secure connection, the script must be served with an SSL certificate to work in all browsers.
Alternatively, the content of makeics.js can be pasted directly into the javascript console.

The bookmarklet and instructions are available [here](http://blog.whither.ca/export-solus-course-calendar/).
