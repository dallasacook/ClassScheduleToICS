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

var ver = '140712';
var frame = parent.TargetContent;
var allowed_weekdays = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];
var num_problem_rows = 0;
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
    return (hour*60 +min)*60;
}

function pad(n) {
      if (n<10) {
          return '0'+n;
      }
      return n;
}

// JS Date -> 20130602T130000
function date_to_string(date) {
     return date.getFullYear()
        +pad( date.getMonth() + 1 )
        +pad( date.getDate() )
        +'T'
        +pad( date.getHours() )
        +pad( date.getMinutes() )
        +pad( date.getSeconds() );
}

function title_case(str)
{
    return str.replace(/\w\S*/g, function(txt){
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
        });
}

function create_ics_wrap(events) {
        return 'BEGIN:VCALENDAR\r\n'
        +'PRODID:-//Leo Koppel//Queen\'s Soulless Calendar Exporter v' + ver + '//EN\r\n'
        +'VERSION:2.0\r\n'
        
        // timezone definition from http://erics-notes.blogspot.ca/2013/05/fixing-ics-time-zone.html
        +'BEGIN:VTIMEZONE\r\n'
        +'TZID:America/New_York\r\n'
        +'X-LIC-LOCATION:America/New_York\r\n'
        +'BEGIN:DAYLIGHT\r\n'
        +'TZOFFSETFROM:-0500\r\n'
        +'TZOFFSETTO:-0400\r\n'
        +'TZNAME:EDT\r\n'
        +'DTSTART:19700308T020000\r\n'
        +'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU\r\n'
        +'END:DAYLIGHT\r\n'
        +'BEGIN:STANDARD\r\n'
        +'TZOFFSETFROM:-0400\r\n'
        +'TZOFFSETTO:-0500\r\n'
        +'TZNAME:EST\r\n'
        +'DTSTART:19701101T020000\r\n'
        +'RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU\r\n'
        +'END:STANDARD\r\n'
        +'END:VTIMEZONE\r\n'
        + events.join('\r\n')
        +'END:VCALENDAR\r\n';
}

// Parse a single row (given as an array of table cell content) and return the ICS string.
// If the row should be ignored return false
function row_to_ics(course_code, course_name, cells) {
    // Sometimes solus lists extra rows with no date/time (?). Ignore them.
          if(cells[3].trim().length == 0) {
              return false;
          }

          //class_nbr = cells[0]; //ignore
          //section = cells[1]; //ignore
          // if component (lecture or tutorial or lab) is omitted, it is the same as above
          if(cells[2].trim().length > 0) {
              component = cells[2].trim();
          } 
          var days_and_times = cells[3].split(' ');
          var room = cells[4].trim();
          var instructor = cells[5].trim();
          var start_and_end = cells[6].split(' - ');

          var input_weekday = days_and_times[0].trim(); // e.g. 'Mo'
          var input_start_time = days_and_times[1].trim(); // e.g. '8:30AM'
          // days_and_times[2] is '-'
          var input_end_time = days_and_times[3].trim(); // e.g. '9:30AM'     
          var range_start_date = new Date(Date.parse(start_and_end[0]));
          
          // annoyingly, UNTIL must be given in UTC time
          // even then,, there were odd issues with calendars ending recurring events a day earlier
          // this quick fix just to adds a day.
          var range_end_date = new Date(Date.parse(start_and_end[1]));
          range_end_date.setTime(range_end_date.getTime() + 24*(60*60*1000));
          
                    
          // Now we need to get the actual date of the class.
          // This is not trivial as "start_day" could be before this - probably the monday of that week, but not for sure
          // We have, e.g. "Mo 11:30AM - 12:30PM" from which we can get day of week and we know it is after range_start_date.
          // JS days start at 0 for Sunday, and so does allowed_weekdays
          var start_day = allowed_weekdays.indexOf(input_weekday);
          
          if(start_day == -1 && input_weekday.length > 2) {
              // It could be that SOLUS gives more than one day, e.g. "TuTh" for both Tues. and Thurs.
              // In this case, split it up and recurse.
              var valid_weekdays = true;
              var new_rows = [];
              for(var i=0; i<input_weekday.length; i+=2) {
                  var single_weekday = input_weekday.slice(i,i+2);
                  if(allowed_weekdays.indexOf(single_weekday) == -1) {
                      valid_weekdays = false;
                      break;
                  } else {
                      var new_cells = cells.slice(0);
                      new_cells[3] = single_weekday + ' ' + days_and_times.slice(1).join(' ');
                      new_rows.push(new_cells);
                  }
              }
              if(valid_weekdays) {
                  // now recurse for new, single weekday rows.
                  return new_rows.map(function(e) {return row_to_ics(course_code,course_name, e);}).join('\r\n');
              }
         
              throw ('Unexpected weekday format: ' + allowed_weekdays);
          }
          var range_start_day = range_start_date.getDay();
          var incr = (7-range_start_day+start_day)%7;
          
          // The real event start and end dates. Assume no class runs through midnight.
          var start_date = new Date(range_start_date.getTime() + incr*(24*60*60*1000) + time_to_seconds(input_start_time)*1000);
          var end_date = new Date(range_start_date.getTime() + incr*(24*60*60*1000) + time_to_seconds(input_end_time)*1000);
          
          return ('BEGIN:VEVENT\r\n'
          +'DTSTART;TZID=America/New_York:' + date_to_string(start_date) + '\r\n'
          +'DTEND;TZID=America/New_York:' + date_to_string(end_date) + '\r\n'
          +'SUMMARY:' + course_code + ' ' + component + '\r\n'
          +'LOCATION:' + title_case(room) + '\r\n'
          +'DESCRIPTION:' + course_code + ' - ' + course_name + ' ' + component + '. ' + instructor + '\r\n'
          +'RRULE:FREQ=WEEKLY;UNTIL=' + date_to_string(range_end_date) + 'Z' + '\r\n'
          +'END:VEVENT\r\n');

}

function create_ics() {
    var ics_events = [];
    if(frame.$('.PSGROUPBOXWBO').length == 0) {
        throw "Course tables not found.";
    }

    // for each course
    frame.$('.PSGROUPBOXWBO:gt(0)').each(function() { 
        _course_title_parts = frame.$(this).find('td:eq(0)').text().split(' - ');
        var course_code = _course_title_parts[0].trim();
        var course_name = _course_title_parts[1].trim();
       
       var component = '';
       
       // for each event
       frame.$(this).find("tr:gt(7)").each(function() {
          var cells = frame.$(this).find('td').map(function() { return frame.$(this).text(); });
          try {
            var event_string = row_to_ics(course_code, course_name, cells);
            if (event_string) {
                // now append to the ics string
                ics_events.push(event_string);
                frame.$(this).find('td').css('background', '#ebffeb'); // mark in light green
            }
          }
          catch(err) {
            // add the row to the 'could not parse' count and highlight it
            num_problem_rows += 1;
            frame.$(this).find('td').css('background', 'red');
          }

       }); // end each event
        
    }); // end each course

    if(ics_events.length == 0) {
        throw "No class entries found.";
    }
    return create_ics_wrap(ics_events);
}

        
function initBookmarklet() {

        if(parent.TargetContent.location.pathname.indexOf('SSR_SSENRL_LIST.GBL') == -1) {
            throw "List view not found.";
        }

       try {
           var checkBlobSupport = !!new frame.Blob;
           frame.$.getScript('https://googledrive.com/host/0B4PDwhAa-jNITkc4MTh5M1BoZG8/filesaver.js').done(runBookmarklet);
       } catch (e) {
           frame.$.when(frame.$.getScript('https://googledrive.com/host/0B4PDwhAa-jNITkc4MTh5M1BoZG8/filesaver.js'),
               frame.$.getScript('https://googledrive.com/host/0B4PDwhAa-jNITkc4MTh5M1BoZG8/blob.js')).done(runBookmarklet);
       }

}

function runBookmarklet() {

    var checkBlobSupport = !!new frame.Blob;

    ics_content = create_ics();

    frame.$('#ics_download').remove();

    frame.$('.PATRANSACTIONTITLE').append(' <span id="ics_download">('
    +'<a href="javascript:void(0)" id="ics_download_link">' + link_text + '</a>'
    +')</span>');

    frame.$('#ics_download_link').click( function() {
        var blob = new frame.Blob([ics_content], {type: "text/plain;charset=utf-8"});
        frame.saveAs(blob, "coursecalendar.ics");
        return false;
    });


    var msg = '';

    if(num_problem_rows > 0) {
        msg += 'Success! ICS file was created, but I could not understand '
        + num_problem_rows + ' '
        + (num_problem_rows > 1 ? 'rows. These are' : 'row. This is')
        + ' highlighted in red.';
    } else {
        msg += 'Success! All exported rows are highlighted in green.';
    }

    msg += '\n\nClick on the "Download .ics file" link to download.';

    /* Safari problems require manual workarounds for now */
    if (navigator.userAgent.indexOf('Safari') != -1 && navigator.userAgent.indexOf('Chrome') == -1) {
        msg += '\n\nSafari users: If the file opens in a new window instead of downloading, '
             + 'use Save As (Ctrl-S or Cmd-S) and save it with a .ics extension. '
             + 'Then import into your calendar software.';
    }

    alert(msg);
}

(function(){

    try {
    // the minimum version of jQuery we want
    var jquery_ver = "1.10.0";
   
        if (frame === undefined) {
            throw "TargetContent frame not found.";
        }

        // check prior inclusion and version
        if (frame.jQuery === undefined || frame.jQuery.fn.jquery < jquery_ver) {
            var done = false;
            var script = frame.document.createElement("script");
            script.src = "https://ajax.googleapis.com/ajax/libs/jquery/" + jquery_ver + "/jquery.min.js";
            script.onload = script.onreadystatechange = function(){
                if (!done && (!this.readyState || this.readyState == "loaded" || this.readyState == "complete")) {
                    done = true;
                    initBookmarklet();
                }
            };
            frame.document.getElementsByTagName("head")[0].appendChild(script);
        } else {
            initBookmarklet();
        }
     }    
    catch(err){
        var msg="Schedule exporter didn't work :(\n"
        +"Make sure you are on the \"List View\" of \" My Class Schedule\".\n\n"
        +"Otherwise, report this: \n"
        +'v' + ver + '\n'
        +err + '\n';

        alert(msg);
    }
})();
