/**
 * Class schedule to .ics file bookmarklet
 * Leo Koppel
 * Rewritten from the script by Keanu Lee (https://github.com/keanulee/ClassScheduleToICS)
 */
var ver='0.1'
var frame = parent.TargetContent;
var weekdays_input = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];

// 11:30AM -> 41400
function time_to_seconds(time_str) {
    m = time_str.match(/(\d*):(\d*)(\wM)/);
    hour = parseInt(m[1]);
    min = parseInt(m[2]);
    if(m[3] == 'PM') hour += 12;
    return (hour*60 +min)*60;
}

function pad(n) {
      if (n<10) return '0'+n;
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
        +pad( date.getSeconds() )
        +'Z';
}

function title_case(str)
{
    return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
}

function create_ics_wrap(events) {
        ics_content = 'BEGIN:VCALENDAR\r\n'
        +"PRODID:-//Leo Koppel//Queen's Soulless Calendar Exporter//EN\r\n"
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
        +'END:VTIMEZONE\r\n';

        ics_content += events.join('\r\n')
        ics_content += 'END:VCALENDAR\r\n';
        
        return ics_content;

}

function create_ics() {
     
    ics_events = [];
    if($('.PSGROUPBOXWBO').length == 0) {
        throw "Course tables not found.";
    }
    

    
    // for each course
    frame.$('.PSGROUPBOXWBO:gt(0)').each(function() { 
        _course_title_parts = $(this).find('td:eq(0)').text().split(' - ');
        course_code = _course_title_parts[0].trim();
        course_name = _course_title_parts[1].trim();
       
       var component = '';
       
       // for each event
       $(this).find("tr:gt(7)").each(function() {
          cells = $(this).find('td').map(function() { return $(this).text(); });
          
          // Sometimes solus lists extra rows with no date/time (?). Ignore them.
          if(cells[3].trim().length == 0) {
              return;
          }
          
          //class_nbr = cells[0]; //ignore
          //section = cells[1]; //ignore
          // if component (lecture or tutorial or lab) is omitted, it is the same as above
          if(cells[2].trim().length > 0) {
              component = cells[2].trim();
          } 
          days_and_times = cells[3].split(' ');
          room = cells[4].trim();
          instructor = cells[5].trim();
          start_and_end = cells[6].split(' - ');

          input_weekday = days_and_times[0].trim(); // e.g. 'Mo'
          input_start_time = days_and_times[1].trim(); // e.g. '8:30AM'
          // days_and_times[2] is '-'
          input_end_time = days_and_times[3].trim(); // e.g. '9:30AM'     
          range_start_date = new Date(Date.parse(start_and_end[0]));
          
          // annoyingly, UNTIL must be given in UTC time
          // even then,, there were odd issues with calendars ending recurring events a day earlier
          // this quick fix just to adds a day.
          range_end_date = new Date(Date.parse(start_and_end[1]));
          range_end_date.setTime(range_end_date.getTime() + 24*(60*60*1000));
          
                    
          // Now we need to get the actual date of the class.
          // This is not trivial as "start_day" could be before this - probably the monday of that week, but not for sure
          // We have, e.g. "Mo 11:30AM - 12:30PM" from which we can get day of week and we know it is after range_start_date.
          // JS days start at 0 for Sunday, and so does weekdays_input
          start_day = weekdays_input.indexOf(input_weekday);
          range_start_day = range_start_date.getDay();
          incr = (7-range_start_day+start_day)%7;
          
          // The real event start and end dates. Assume no class runs through midnight.
          start_date = new Date(range_start_date.getTime() + incr*(24*60*60*1000) + time_to_seconds(input_start_time)*1000);
          end_date = new Date(range_start_date.getTime() + incr*(24*60*60*1000) + time_to_seconds(input_end_time)*1000);
          
          // now append to the ics string
          ics_events.push('BEGIN:VEVENT\r\n'
          +'DTSTART;TZID=America/New_York:' + date_to_string(start_date) + '\r\n'
          +'DTEND;TZID=America/New_York:' + date_to_string(end_date) + '\r\n'
          +'SUMMARY:' + course_code + ' ' + component + '\r\n'
          +'LOCATION:' + title_case(room) + '\r\n'
          +'DESCRIPTION:' + course_code + ' - ' + course_name + ' ' + component + '. ' + instructor + '\r\n'
          +'RRULE:FREQ=WEEKLY;UNTIL=' + date_to_string(range_end_date) + '\r\n'
          +'END:VEVENT\r\n');

       }); // end each event
        
    }); // end each course
    
    if(ics_events.length == 0) {
        throw "No class entries found."
    }
    
    return create_ics_wrap(ics_events);
}

try {
(function(){
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
        
        function initBookmarklet() {
            (frame.CalendarBookmarklet = function(){
                ics_content = create_ics();
                                
                $('#ics_download').remove();
                
                frame.$('.PATRANSACTIONTITLE').append(' <span id="ics_download">('
                +'<a href="data:text/calendar;charset=utf8,'
                +encodeURIComponent(ics_content)
                +'" download="coursecalendar.ics">Download .ics file</a>'
                +')</span>');
            })();
        }
})();
}    
    catch(err){
        alert("Schedule exporter didn't work :(\n"
        +"Make sure you are on the \"List View\" of \" My Class Schedule\".\n\n"
        +"Otherwise, report this: \n"
        +'v' + ver + '\n'
        +err);
}
