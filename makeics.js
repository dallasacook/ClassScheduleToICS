/**
 * Class schedule to .ics file bookmarklet
 * Leo Koppel
 * Based on the script by Keanu Lee
 * (https://github.com/keanulee/ClassScheduleToICS)
 *
 * Depends on FileSaver.js (github.com/eligrey/FileSaver.js) and, in some
 * browsers, Blob.js (github.com/eligrey/Blob.js).
 *
 * License: MIT (see LICENSE.md)
 */

var ver = '140719';
var frame = parent.TargetContent;
var allowed_weekdays = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];
var num_courses = 0, num_rows = 0, num_problem_rows = 0;
var link_text = 'Download .ics file';

// 11:30AM -> 41400
function time_to_seconds(time_str) {
    // time_str can be in the form "2:30PM" or "14:30" -- varies by browser for some reason.
    var m = time_str.match(/(\d*):(\d*)(\wM)?/);
    var hour = parseInt(m[1], 10);
    var min = parseInt(m[2], 10);
    if(m[3] == 'PM' && hour < 12) {
        hour += 12;
    }
    return (hour * 60 + min) * 60;
}

function pad(n) {
    if(n < 10) {
        return '0' + n;
    }
    return n.toString();
}

// JS Date -> 20130602T130000
function date_to_string(date) {
    return date.getFullYear() + pad(date.getMonth() + 1) + pad(date.getDate()) +
    'T' + pad(date.getHours()) + pad(date.getMinutes()) + pad(date.getSeconds());
}

function title_case(str) {
    return str.replace(/\w[^\s-]*/g, function(txt) {
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
}

// Escape special characters as in iCalendar spec on Text values
// (http://tools.ietf.org/html/rfc5545#section-3.3.11)
function escape_ics_text(text) {
    return text.replace(/[;,\\]/g, '\\$&').replace(/\r\n|\r|\n/gm, '\\n');
}


// Return iCalendar string given array of class event info
function create_ics(class_events) {

    var s = 'BEGIN:VCALENDAR\r\n' +
    'PRODID:-//Leo Koppel//Queen\'s Soulless Calendar Exporter v' + ver + '//EN\r\n' +
    'VERSION:2.0\r\n' +
    // timezone definition from http://erics-notes.blogspot.ca/2013/05/fixing-ics-time-zone.html
    'BEGIN:VTIMEZONE\r\n' +
    'TZID:America/New_York\r\n' +
    'X-LIC-LOCATION:America/New_York\r\n' +
    'BEGIN:DAYLIGHT\r\n' +
    'TZOFFSETFROM:-0500\r\n' +
    'TZOFFSETTO:-0400\r\n' +
    'TZNAME:EDT\r\n' +
    'DTSTART:19700308T020000\r\n' +
    'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU\r\n' +
    'END:DAYLIGHT\r\n' +
    'BEGIN:STANDARD\r\n' +
    'TZOFFSETFROM:-0400\r\n' +
    'TZOFFSETTO:-0500\r\n' +
    'TZNAME:EST\r\n' +
    'DTSTART:19701101T020000\r\n' +
    'RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU\r\n' +
    'END:STANDARD\r\n' +
    'END:VTIMEZONE\r\n';

    var i, c;
    for(i=0; i<class_events.length; i++) {
        c = class_events[i];
        s += ('\r\nBEGIN:VEVENT\r\n' +
        'DTSTART;TZID=America/New_York:' + date_to_string(c.start_date) + '\r\n' +
        'DTEND;TZID=America/New_York:' + date_to_string(c.end_date) + '\r\n' +
        'SUMMARY:' + escape_ics_text(c.course_code + ' ' + c.component) + '\r\n' +
        'LOCATION:' + escape_ics_text(title_case(c.room)) + '\r\n' +
        'DESCRIPTION:' + escape_ics_text(c.course_code + ' - ' + c.course_name + ' ' + c.component + '. ' + c.instructor) + '\r\n' +
        'RRULE:FREQ=WEEKLY;UNTIL=' + date_to_string(c.range_end_date) + 'Z' + '\r\n' +
        'END:VEVENT\r\n');
    }

    s += '\r\nEND:VCALENDAR\r\n';
    return s;
}

// Parse a single row (given as an array of table cell content) into an event
// object. If successful, add it to the output array and return true.
// If the row could not be understood, throw an exception.
function parse_row(course_code, course_name, cells, output_array) {
    // Sometimes solus lists extra rows with no date/time (?). Ignore them.
    if(cells[3].trim().length === 0) {
        throw ('Row is missing Days & Times field.');
    }

    // Ignore the first two columns:
    //class_nbr = cells[0]; //ignore
    //section = cells[1]; //ignore

    // Fields used in human-readable text properties
    // if component (lecture or tutorial or lab) is omitted, it is the same as previous row
    var component;
    if(cells[2].trim().length > 0) {
        component = cells[2].trim();
    } else {
        component = output_array[output_array.length-1].component;
    }

    var room = cells[4].trim();
    var instructor = cells[5].trim();
    var start_and_end = cells[6].split(' - ');

    // Fields used in date and calendar rule properties
    var days_and_times = cells[3].split(' ');
    var input_weekday = days_and_times[0].trim();
    // e.g. 'Mo'
    var input_start_time = days_and_times[1].trim();
    // e.g. '8:30AM'
    // days_and_times[2] is '-'
    var input_end_time = days_and_times[3].trim();
    // e.g. '9:30AM'
    var range_start_date = new Date(Date.parse(start_and_end[0]));

    // annoyingly, UNTIL must be given in UTC time
    // even then, there were odd issues with calendars ending recurring events a day earlier
    // this quick fix just adds a day.
    var range_end_date = new Date(Date.parse(start_and_end[1]));
    range_end_date.setTime(range_end_date.getTime() + 24 * (60 * 60 * 1000));

    // Now we need to get the actual date of the class.
    // This is not trivial as "start_day" could be before this - probably the monday of that week, but not for sure
    // We have, e.g. "Mo 11:30AM - 12:30PM" from which we can get day of week and we know it is after range_start_date.
    // JS days start at 0 for Sunday, and so does allowed_weekdays
    //
    // Also, it could be that SOLUS gives more than one day, e.g. "TuTh" for both Tues. and Thurs.
    // Thus we treat the field as a string of one or more two-character weekdays to begin with
    var start_days = [];

    var i;
    for(i = 0; i < input_weekday.length; i += 2) {
        // Check each two-character substring against valid weekdays
        var single_start_day = allowed_weekdays.indexOf(input_weekday.slice(i, i + 2));
        if(single_start_day == -1) {
            // It's no good
            throw ('Unexpected weekday format: ' + allowed_weekdays);
        } else {
            start_days.push(single_start_day);
        }
    }

    if(start_days.length > 0) {
        // Now add an event object for each weekday (though there is usually only one)
        for(i = 0; i < start_days.length; i++) {
            var range_start_day = range_start_date.getDay();
            var incr = (7 - range_start_day + start_days[i]) % 7;
            // number of days until the first occurrence of that weekday, after range_start_date

            // The real event start and end dates. Assume no class runs through midnight (I don't know how SOLUS would show this anyway).
            var start_date = new Date(range_start_date.getTime() + incr * (24 * 60 * 60 * 1000) + time_to_seconds(input_start_time) * 1000);
            var end_date = new Date(range_start_date.getTime() + incr * (24 * 60 * 60 * 1000) + time_to_seconds(input_end_time) * 1000);

            output_array.push({
                start_date : start_date,
                end_date : end_date,
                range_end_date : range_end_date,
                course_code : course_code,
                course_name : course_name,
                component : component,
                room : room,
                instructor : instructor
            });
        }
        return true;
    }

    return false;
}

function get_class_events() {
    var class_events = [];
    if(frame.$('.PSGROUPBOXWBO').length === 0) {
        throw "Course tables not found.";
    }

    // for each course
    frame.$('.PSGROUPBOXWBO:gt(0)').each(function() {
        var _course_title_parts = frame.$(this).find('td:eq(0)').text().split(' - ');
        var course_code = _course_title_parts[0].trim();
        var course_name = _course_title_parts[1].trim();

        // for each row
        frame.$(this).find("tr:gt(7)").each(function() {

            try {
                var cells = frame.$(this).find('td').map(function() {
                    return frame.$(this).text();
                });
                var valid_row = parse_row(course_code, course_name, cells, class_events);
                if(valid_row) {
                    // highlight it in green
                    frame.$(this).find('td').addClass('ics_c_g');
                }
            } catch(err) {
                // add the row to the 'could not parse' count and highlight it red
                num_problem_rows += 1;
                frame.$(this).find('td').addClass('ics_c_r');
            }

            num_rows += 1;
        });
        // end each row

        num_courses += 1;
    });
    // end each course

    return class_events;
}

// Calculate non-overlapping hours of class given events array
// Note this implementation only throws away overlapping classes that start
// at the same time (TODO)
function get_hours_of_class(class_events) {
    var i, c, k, total_hours = 0;
    var class_times = [];
    for(i=0; i<class_events.length; i++) {
        c = class_events[i];
        var match = false;
        for(k=0; k<class_times.length; k++) {
            if(class_times[k].start === c.start_date.getTime() && class_times[k].end == c.end_date.getTime()) {
                match = true;
                break;
            }
        }
        if(!match) {
            class_times.push({start: c.start_date.getTime(), end: c.end_date.getTime()});
            h = (c.end_date - c.start_date) / (60 * 60 * 1000);
            total_hours += h
        }
    }
    return total_hours;
}

// Find and combine overlapping "duplicate" classes in event array
// A "duplicate" is generally a lab with multiple events listed at the same
// time, only in different rooms
function combine_duplicate_classes(input_events) {
    var i, k, p, x, y, dup;
    var output_events = [];
    for(i=0; i<input_events.length; i++) {
        x = input_events[i];
        for(k=0; k<output_events.length; k++) {
            y = output_events[k];
            matches = [];
            dup = true;
            for(p in x) {
                if(x.start_date.getTime() != y.start_date.getTime()
                || x.end_date.getTime() != y.end_date.getTime()
                || x.start_date.getTime() != y.start_date.getTime()
                || x.range_end_date.getTime() != y.range_end_date.getTime()
                || x.component != y.component
                || x.course_code != y.course_code
                ) {
                    dup = false;
                    break;
                }
            }
            if(dup) {
                // Just add to the room field
                y.room += ', ' + x.room;
                break;
            }
        }
        if(!dup) {
            output_events.push(x);
        }
    }
    return output_events;
}

// Create the results infobox and show a spinner while additional scripts are
// loaded & run
function show_loading_box() {
    // remove infobox from any previous run
    var infobox = frame.document.getElementById('ics_box');
    if(infobox) {
        infobox.parentNode.removeChild(infobox);
    }

    frame.document.getElementById('ACE_DERIVED_REGFRM1_GROUP_BOX').insertAdjacentHTML('afterend',
    '<div id="ics_box" style="border:1px solid rgb(114, 175, 69); padding:0px 1em;' +
    'font-family:Verdana,sans-serif; font-size:0.9em; background-color:#ebffeb;">' +
    '<div id="ics_spinner" style="text-align:center; font-family:monospace;">' +
    '..</div></div>');

    var spinner = frame.document.getElementById('ics_spinner');
    var k = 0, s = ['.:', ':.'];
    setInterval(function() {
        k = +!k;
        spinner.innerHTML = s[k];
    }, 200);

}

function initBookmarklet() {
    if(parent.TargetContent.location.pathname.indexOf('SSR_SSENRL_LIST.GBL') == -1) {
        throw "List view not found.";
    }

    var scripts = ['https://googledrive.com/host/0B4PDwhAa-jNITkc4MTh5M1BoZG8/filesaver.js'];
    try {
        var checkBlobSupport = !!new frame.Blob;
    } catch (e) {
        scripts.push('https://googledrive.com/host/0B4PDwhAa-jNITkc4MTh5M1BoZG8/blob.js');
    }
    
    frame.$.when.apply(this, frame.$.map(scripts, frame.$.getScript)).done(runBookmarklet);
}

// Parse the page to create the iCalendar file
// Create the download link for a file containing ics_content
// and show info about the script results
function runBookmarklet() {

    var class_events = get_class_events();
    var num_events = class_events.length;

    // Remove duplicate labs
    class_events = combine_duplicate_classes(class_events);
    var num_removed = num_events - class_events.length;
    num_events = class_events.length;

    // Construct iCalendar file
    var ics_content = create_ics(class_events);

    // Construct message to user
    var msg = '';

    if(num_events > 0) {
        msg += '<b>Success!</b> A calendar file was created.';

        // Safari problems require manual workarounds for now
        if (navigator.userAgent.indexOf('Safari') != -1 && navigator.userAgent.indexOf('Chrome') == -1) {
            msg += '</p><p><b>Safari users:</b> If the file opens in a new window instead of downloading, ' +
            'use <b>Save As</b> (Ctrl-S or Cmd-S) and save it <b>with a .ics extension.</b> ' +
            'For example, as <i>classes.ics</i>. Then import into your calendar software.';
        }
    } else {
        msg += '<b>Failure.</b> A calendar file could not be created.';
    }

    // Create page elements
    var infobox = frame.$('#ics_box');
    frame.$('#ics_spinner').remove();

    infobox.append('<p id="ics_results_msg">' + msg + '</p>');
    var download_p = frame.$('<p style="text-align:center">').appendTo(infobox);

    var download_button = frame.$('<button type="button" id="ics_download_link">' + link_text + '</button>').appendTo(download_p);

    infobox.append('<p>' + 'I found <b>' + num_rows + '</b> row' + (num_rows == 1 ? '' : 's') +
                   ' under <b>' + num_courses + '</b> course' + (num_courses == 1 ? '' : 's') + '.</br>' +
                   'I could not understand <b>' + num_problem_rows + '</b> row' + (num_problem_rows == 1 ? '' : 's') +
                   ' (highlighted in <span class="ics_c_r">red</span>).</br>' +
                   ' <b>' + num_removed + '</b> overlapping event' + (num_removed == 1 ? '' : 's') + ' were merged into another.</br>' +
                   'The calendar file contains <b>' + num_events + '</b> event' + (num_events == 1 ? '' : 's') + '.</p>' +
                   '<p><i>Experimental:</i> You have <b>' + (+get_hours_of_class(class_events).toFixed(1)) + '</b> hours of class weekly.' +
                   '</p>');

    infobox.append('<p style="font-size:0.5em; text-align:right;">' +
                   '<a href="http://blog.whither.ca/export-solus-course-calendar/" target="_blank">Instructions</a>' +
                   ' <a href="https://github.com/leokoppel/ClassScheduleToICS/issues" target="_blank">Issues</a>' +
                   ' <i>v' + ver + '</i>' + '</p>');

    download_button.click(function() {
        var blob = new frame.Blob([ics_content], {
            type : "text/plain;charset=utf-8"
        });
        frame.saveAs(blob, "coursecalendar.ics");
        return false;
    });

    // Add styling for info box and previously highlighted (classed) rows
    frame.$('<style type="text/css"> ' +
      '.ics_c_r { background-color: #e37d7d; } ' +
      '.ics_c_g { background-color: #ebffeb; } ' +
      '</style>').appendTo("head");

    if(num_events > 0) {
        download_button.css('font-size', '1.5em');
    } else {
        download_button.hide();
        infobox.css('background-color', '#edc2c2');
    }

}

(function() {

    try {
        // Show spinner before loading any other scripts
        show_loading_box();

        // the minimum version of jQuery we want
        var jquery_ver = "1.10.0";

        if(frame === undefined) {
            throw "TargetContent frame not found.";
        }

        // check prior inclusion and version
        if(frame.jQuery === undefined || frame.jQuery.fn.jquery < jquery_ver) {
            var done = false;
            var script = frame.document.createElement("script");
            script.src = "https://ajax.googleapis.com/ajax/libs/jquery/" + jquery_ver + "/jquery.min.js";
            script.onload = script.onreadystatechange = function() {
                if (!done && (!this.readyState || this.readyState == "loaded" || this.readyState == "complete")) {
                    done = true;
                    initBookmarklet();
                }
            };
            frame.document.getElementsByTagName("head")[0].appendChild(script);
        } else {
            initBookmarklet();
        }
    } catch(err) {
        var msg = "Schedule exporter didn't work :(\n" +
                  "Make sure you are on the \"List View\" of \" My Class Schedule\".\n\n" +
                  "Otherwise, report this: \n" + 'v' + ver + '\n' + err + '\n';
        alert(msg);
    }
})();