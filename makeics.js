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

var ver = '140720';
var frame = parent.TargetContent;
var allowed_weekdays = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];
var num_courses = 0, num_rows = 0, num_problem_rows = 0;
var link_text = 'Download .ics file';

// "11:30AM" -> 41400
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

// Add (only one) leading zero if needed to make a two-digit number
function pad(n) {
    if(n < 10) {
        return '0' + n;
    }
    return n.toString();
}

// JS Date -> 20130602T130000 (in its own timezone)
function date_to_string(date) {
    return date.getFullYear() + pad(date.getMonth() + 1) + pad(date.getDate()) +
    'T' + pad(date.getHours()) + pad(date.getMinutes()) + pad(date.getSeconds());
}

// JS Date -> 20130602T210000Z (using UTC time)
function date_to_string_UTC(date) {
    return date.getUTCFullYear() + pad(date.getUTCMonth() + 1) + pad(date.getUTCDate()) +
    'T' + pad(date.getUTCHours()) + pad(date.getUTCMinutes()) + pad(date.getUTCSeconds()) + 'Z';
}

// Turn capitalized names into friendlier title-case names
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


// Return an iCalendar object (a string), given an array of class event info
// This should conform to the iCalendar spec (http://tools.ietf.org/html/rfc5545.html)
function create_ics(class_events) {


    // date the calendar object was created (right now)
    var dtstamp = date_to_string_UTC(new Date());

    var s = 'BEGIN:VCALENDAR\r\n' +
    'VERSION:2.0\r\n' +
    'PRODID:-//Leo Koppel//Queen\'s Soulless Calendar Exporter v' + ver + '//EN\r\n' +
    'BEGIN:VTIMEZONE\r\n' +
    'TZID:America/Toronto\r\n' +
    'BEGIN:STANDARD\r\n' +
    'DTSTART:16011104T020000\r\n' +
    'RRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=11\r\n' +
    'TZOFFSETFROM:-0400\r\n' +
    'TZOFFSETTO:-0500\r\n' +
    'END:STANDARD\r\n' +
    'BEGIN:DAYLIGHT\r\n' +
    'DTSTART:16010311T020000\r\n' +
    'RRULE:FREQ=YEARLY;BYDAY=2SU;BYMONTH=3\r\n' +
    'TZOFFSETFROM:-0500\r\n' +
    'TZOFFSETTO:-0400\r\n' +
    'END:DAYLIGHT\r\n' +
    'END:VTIMEZONE\r\n';

    var i, c;
    for(i=0; i<class_events.length; i++) {
        // Construct an event calendar component for each class event.
        // Note about dates:
        // DTSTART and DTEND and given as fixed dates in Eastern Time (since
        // SOLUS shows schedules in the Univesity's own time zone), ignoring
        // the Date object's own timezone offset (which varies by client).
        // DTSTAMP and UNTIL dates however, must be given in UTC time. Just
        // an idiosyncrasy of the standard.
        c = class_events[i];
        s += ('BEGIN:VEVENT\r\n' +
        'DTSTAMP:' + dtstamp + '\r\n' +
        'UID:' + dtstamp + i + '@' + window.location.hostname + '\r\n' +
        'DTSTART;TZID=America/Toronto:' + date_to_string(c.start_date) + '\r\n' +
        'DTEND;TZID=America/Toronto:' + date_to_string(c.end_date) + '\r\n' +
        'SUMMARY:' + escape_ics_text(c.course_code + ' ' + c.component) + '\r\n' +
        'LOCATION:' + escape_ics_text(title_case(c.room)) + '\r\n' +
        'DESCRIPTION:' + escape_ics_text(c.course_code + ' - ' + c.course_name + ' ' + c.component + '. ' + c.instructor) + '\r\n' +
        'RRULE:FREQ=WEEKLY;UNTIL=' + date_to_string_UTC(c.range_end_date) + '\r\n' +
        'END:VEVENT\r\n');
    }

    s += 'END:VCALENDAR\r\n';

    // Fold long lines
    // (RFC5445: "Lines of text SHOULD NOT be longer than 75 octets, excluding the line break")
    // Split using CRLF followed by space
    return s.replace(/(.{75})/g,"$1\r\n ");
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

    // Add a day to the end of the range as SOLUS treats it as inclusive, unlike iCalendar
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
            total_hours += (c.end_date - c.start_date) / (60 * 60 * 1000);
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
            output_events.push(frame.$.extend({}, x)); // clone the object
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


// Currently, either combine duplicate labs or leave as is
function apply_options(class_events, combine) {
    if(combine) {
        class_events = combine_duplicate_classes(class_events);
    }
    return class_events;
}

// Parse the page to create the iCalendar file
// Create the download link for a file containing ics_content
// and show info about the script results
function runBookmarklet() {

    var orig_class_events = get_class_events();

    // Remove duplicate labs
    var class_events = apply_options(orig_class_events, true);

    // Construct iCalendar file
    var ics_content = create_ics(class_events);

    // Construct message to user
    var msg = '';
    var success = (class_events.length > 0);

    if(success) {
        msg += '<h3>Success!</h3> A calendar file was created.';

        // Safari problems require manual workarounds for now
        if (navigator.userAgent.indexOf('Safari') != -1 && navigator.userAgent.indexOf('Chrome') == -1) {
            msg += '</p><p><b>Safari users:</b> If the file opens in a new window instead of downloading, ' +
            'use <b>Save As</b> (Ctrl-S or Cmd-S) and save it <b>with a .ics extension.</b> ' +
            'For example, as <i>classes.ics</i>. Then import into your calendar software.';
        }
    } else {
        msg += '<h3>Failure.</h3> A calendar file could not be created.';
    }

    // Create page elements
    var infobox = frame.$('#ics_box');
    frame.$('#ics_spinner').hide();

    infobox.prepend('<p id="ics_results_msg">' + msg + '</p>'); //prepend before spinner
    var download_p = frame.$('<div style="text-align:center">').appendTo(infobox);

    var download_button = frame.$('<button type="button" id="ics_download_link">' + link_text + '</button>').appendTo(download_p);

    if(success) {
        infobox.append('<h3>Options</h3>' +
                       '<input type="checkbox" id="ics_opt_combine" checked/><label for="ics_opt_combine">' +
                       'Combine duplicate events (only Room differs)</label>');
    }

    var statbox = frame.$('<div id="ics_stats">').appendTo(infobox);

    // Rewrite the stats div. Function so it can be called when options change.
    function update_stats() {
        msg = '<h3>Stats</h3>' +
              '<p>' + 'I found <b>' + num_rows + '</b> row' + (num_rows == 1 ? '' : 's') +
              ' under <b>' + num_courses + '</b> course' + (num_courses == 1 ? '' : 's') + '.</br>';
        if(num_problem_rows > 0) {
            msg += 'I could not export <b>' + num_problem_rows + '</b> row' + (num_problem_rows == 1 ? '' : 's') +
                   ' (highlighted in <span class="ics_c_r">red</span>).</br>';
        }
        msg += 'The calendar file contains <b>' + class_events.length + '</b> event' + (class_events.length == 1 ? '' : 's') + '.</p>';
        if(success) {
            msg += '<p><i>Experimental:</i> You have <b>' + (+get_hours_of_class(class_events).toFixed(1)) + '</b>' +
                   ' hours of class weekly.</p>';
        }

        statbox.html(msg);
    }

    update_stats();

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


    frame.$('#ics_opt_combine').change(function() {

        frame.$('#ics_spinner').height(download_p.height());
        download_p.hide();
        frame.$('#ics_spinner').show();

        setTimeout(function() {
            // Updating the info is virtually instantaneous
            // This timeout is just so the user notices the change
            frame.$('#ics_spinner').hide();
            download_p.show();
        }, 80);

        class_events = apply_options(orig_class_events, this.checked);

        // Reconstruct iCalendar file
        ics_content = create_ics(class_events);
        update_stats();
    });


    // Add styling for info box and previously highlighted (classed) rows
    frame.$('<style type="text/css"> ' +
      '.ics_c_r { background-color: #e37d7d; } ' +
      '.ics_c_g { background-color: #ebffeb; } ' +
      'h3 { font-size: 1em; }' +
      '</style>').appendTo("head");

    if(success) {
        download_button.css('font-size', '1.5em');
    } else {
        download_button.hide();
        infobox.css({'background-color': '#edc2c2',
                     'border-color': '#a37272'});
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