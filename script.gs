const CONFIG = {
    // Which emails to scan. To limit to one sender, add e.g. "from:portal@school.org"
    SEARCH_QUERY: 'from:nypdsound@gmail.com has:attachment filename:ics newer_than:14d',
    //SEARCH_QUERY: 'has:attachment filename:ics newer_than:14d',
    CALENDAR_ID: 'primary',          // or a specific calendar's ID from its settings page
    PROCESSED_LABEL: 'Sound-Job-Added', // Gmail label applied once an email is handled
    CHECK_EVERY_MINUTES: 15,         // must be 1, 5, 10, 15 or 30
};

// Outlook often uses Windows time zone names instead of IANA ones.
const WINDOWS_TZ = {
    'Eastern Standard Time': 'America/New_York',
    'Central Standard Time': 'America/Chicago',
    'Mountain Standard Time': 'America/Denver',
    'US Mountain Standard Time': 'America/Phoenix',
    'Pacific Standard Time': 'America/Los_Angeles',
    'Alaskan Standard Time': 'America/Anchorage',
    'Hawaiian Standard Time': 'Pacific/Honolulu',
    'GMT Standard Time': 'Europe/London',
    'W. Europe Standard Time': 'Europe/Berlin',
    'UTC': 'UTC',
};

/** Run once: creates the timer trigger and does a first pass. */
function setup() {
    ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'importIcsFromGmail')
    .forEach(t => ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger('importIcsFromGmail')
    .timeBased()
    .everyMinutes(CONFIG.CHECK_EVERY_MINUTES)
    .create();
    importIcsFromGmail();
}

function importIcsFromGmail() {
    const label = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL)
    || GmailApp.createLabel(CONFIG.PROCESSED_LABEL);
    const calendar = CONFIG.CALENDAR_ID === 'primary'
    ? CalendarApp.getDefaultCalendar()
    : CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);

    const query = `${CONFIG.SEARCH_QUERY} -label:${CONFIG.PROCESSED_LABEL}`;
    const threads = GmailApp.search(query, 0, 50).reverse(); // oldest first so updates apply in order

    threads.forEach(thread => {
        thread.getMessages().forEach(msg => {
            msg.getAttachments().forEach(att => {
                const name = (att.getName() || '').toLowerCase();
                const type = (att.getContentType() || '').toLowerCase();
                if (!name.endsWith('.ics') && !type.includes('text/calendar')) return;
                try {
                    const ics = parseIcs(att.getDataAsString('UTF-8'));
                    const fallbackTz = resolveTz(ics.calTz) || calendar.getTimeZone();
                    ics.events.forEach(raw => upsertEvent(calendar, toEvent(raw, fallbackTz, ics.method)));
                } catch (e) {
                    console.error(`Failed on "${msg.getSubject()}": ${e}`);
                }
            });
        });
        thread.addLabel(label);
    });
}

// ---------- Calendar ----------

function upsertEvent(calendar, ev) {
    const props = PropertiesService.getScriptProperties();
    const key = 'uid:' + ev.uid;
    const trackedId = props.getProperty(key);
    let tracked = null;
    if (trackedId) {
        try { tracked = calendar.getEventById(trackedId); } catch (e) { tracked = null; }
    }

    // If Google already added this as a real invite, leave it alone.
    if (!tracked) {
        try { if (calendar.getEventById(ev.uid)) return; } catch (e) { /* not found */ }
    }

    if (ev.cancelled) {
        if (tracked) tracked.deleteEvent();
        props.deleteProperty(key);
        return;
    }

    // Update in place when possible
    if (tracked && tracked.isAllDayEvent() === ev.allDay) {
        tracked.setTitle(ev.title);
        tracked.setDescription(ev.description);
        tracked.setLocation(ev.location);
        if (ev.allDay) tracked.setAllDayDates(ev.start, ev.end);
        else tracked.setTime(ev.start, ev.end);
        return;
    }
    if (tracked) tracked.deleteEvent();

    const opts = { description: ev.description, location: ev.location };
    const created = ev.allDay
    ? calendar.createAllDayEvent(ev.title, ev.start, ev.end, opts)
    : calendar.createEvent(ev.title, ev.start, ev.end, opts);
    props.setProperty(key, created.getId());
}

// ---------- ICS parsing ----------

function parseIcs(text) {
    const lines = text.replace(/\r?\n[ \t]/g, '').split(/\r?\n/); // unfold wrapped lines
    const events = [];
    let calTz = null, method = null, cur = null, nested = 0;

    for (const line of lines) {
        if (!line) continue;
        if (line === 'BEGIN:VEVENT') { cur = {}; nested = 0; continue; }
        if (line === 'END:VEVENT') { if (cur) events.push(cur); cur = null; continue; }
        if (cur && line.startsWith('BEGIN:')) { nested++; continue; } // skip VALARM etc.
        if (cur && line.startsWith('END:')) { nested--; continue; }

        const prop = parseLine(line);
        if (!prop) continue;
        if (!cur && prop.name === 'X-WR-TIMEZONE') calTz = prop.value;
        if (!cur && prop.name === 'METHOD') method = prop.value.toUpperCase();
        if (cur && nested === 0 && !(prop.name in cur)) cur[prop.name] = prop;
    }
    return { events, calTz, method };
}

function parseLine(line) {
    let inQuotes = false, i = 0;
    for (; i < line.length; i++) {
        const c = line[i];
        if (c === '"') inQuotes = !inQuotes;
        else if (c === ':' && !inQuotes) break;
    }
    if (i >= line.length) return null;
    const [name, ...rawParams] = line.slice(0, i).split(';');
    const params = {};
    rawParams.forEach(p => {
        const eq = p.indexOf('=');
        if (eq > 0) params[p.slice(0, eq).toUpperCase()] = p.slice(eq + 1).replace(/^"|"$/g, '');
    });
        return { name: name.toUpperCase(), params, value: line.slice(i + 1) };
}

function toEvent(raw, fallbackTz, method) {
    if (!raw.DTSTART) throw new Error('Event has no start time');
    const text = k => (raw[k] ? unescapeIcs(raw[k].value) : '');
    const start = parseIcsDate(raw.DTSTART, fallbackTz);

    let end;
    if (raw.DTEND) end = parseIcsDate(raw.DTEND, fallbackTz).date;
    else if (raw.DURATION) end = new Date(start.date.getTime() + durationMs(raw.DURATION.value));
    else end = start.allDay ? addDays(start.date, 1) : new Date(start.date.getTime() + 36e5);

    return {
        uid: text('UID') || `${text('SUMMARY')}|${raw.DTSTART.value}`,
        title: text('SUMMARY') || '(No title)',
        description: text('DESCRIPTION'),
        location: text('LOCATION'),
        start: start.date,
        end,
        allDay: start.allDay,
        cancelled: method === 'CANCEL' || text('STATUS').toUpperCase() === 'CANCELLED',
    };
}

function parseIcsDate(prop, fallbackTz) {
    const v = prop.value.trim();
    if (prop.params.VALUE === 'DATE' || /^\d{8}$/.test(v)) {
        return { date: new Date(+v.slice(0, 4), +v.slice(4, 6) - 1, +v.slice(6, 8)), allDay: true };
    }
    const m = v.match(/^(\d{8})T(\d{6})(Z?)$/);
    if (!m) throw new Error('Unrecognized date: ' + v);
    const tz = m[3] === 'Z' ? 'UTC' : (resolveTz(prop.params.TZID) || fallbackTz);
    return { date: Utilities.parseDate(m[1] + m[2], tz, 'yyyyMMddHHmmss'), allDay: false };
}

function resolveTz(tzid) {
    if (!tzid) return null;
    if (WINDOWS_TZ[tzid]) return WINDOWS_TZ[tzid];
    // Pull an IANA name out of things like "/citadel.org/20070227_1/America/New_York"
    const m = tzid.match(/[A-Za-z]+\/[A-Za-z_]+(?:\/[A-Za-z_]+)?/);
    return m ? m[0] : null;
}

function durationMs(d) {
    const m = d.match(/^[+-]?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/);
    if (!m) return 36e5;
    const [, w = 0, days = 0, h = 0, min = 0, s = 0] = m;
    return ((((+w * 7 + +days) * 24 + +h) * 60 + +min) * 60 + +s) * 1000;
}

function addDays(date, n) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate() + n);
}

function unescapeIcs(v) {
    return v.replace(/\\n/gi, '\n').replace(/\\([,;\\])/g, '$1');
}
