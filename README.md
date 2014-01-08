# ClassScheduleToICS

A bookmarklet script to export your Queen’s class schedule as an ICS file.

The script generates an ICS file from the course list (currently using a
hardcoded element order), and uses [FileSaver.js][1] to create the download link.

This script is based on [one by Keanu Lee][2] for Waterloo’s self-service
system, but differences with SOLUS (although both use the same PeopleSoft software) 
required changes. I opted for a cross-browser bookmarklet over a Chrome
extension, and used a different method of producing ICS files to include
timezone information. As a result, it’s an almost complete rewrite. [Source][3] 

[1]: https://github.com/eligrey/FileSaver.js/
[2]: https://github.com/keanulee/ClassScheduleToICS
[3]: https://github.com/theoceanwalker/ClassScheduleToICS


## Use
The bookmarklet and instructions are available [here](http://blog.whither.ca/export-solus-course-calendar/).

