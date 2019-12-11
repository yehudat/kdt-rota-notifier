/* 
 *******************
 * Useful articles *
 *******************
 * How to use a first row (headers), as hash keys with the values from the next rows. Pass this hash to functions like here: https://gitlab.com/barkerd427/conference-scripts/blob/cc2346257b69060264024fee0386ccbf057a9e46/confctl.gs#L22
 */

/*
 * Emailing error messages
 */
function errHandler(
    e,
    func_name
){
    var message  = e.message + '\n in file: ' + e.fileName + ' on line: ' + e.lineNumber;
    var sendto   = 'yehuda.tsimanis@gmail.com';
    var subject  = '[Rota_Reminders] Error occurred in ' + func_name;
    GmailApp.sendEmail(sendto, subject, message); 
}

/*
 * Create a data base of all the names and contact details from a spreadsheet
 */
function create_contacts_db(
    ss, 
    sheet_name
) {
    var sheet      = ss.getSheetByName(sheet_name);
    var data_range = sheet.getDataRange();
    var values     = data_range.getValues();

    var status   = 0;
    var contacts = {};

    if (!values) {
        Logger.log('No data found!');
    } else {
        for (var row = 0; row < values.length; row++) {
            if (values[row][1] === undefined) {
                Logger.log('WARNING: No email address found for %s !', values[row][0]);
                status = 1;
            } else {
                contacts[values[row][0]] = values[row][1];
            }
        }
    }
    if (status) {
        return null;
    } else {
        return contacts;
    }
}

/*
 * Adding Email & SMS reminders
 */
function set_notifications(
    event,
    min_before_a
) {
    for (var idx=0; idx<min_before_a.length; idx++) {   
        event.addEmailReminder(min_before_a[idx]);
        //event.addSmsReminder(min_before_a[idx]);
    }
}

/*
 * Create an event and set a value for an eventId
 */
function create_event(
    cal, 
    event_h
) {
    var new_event = cal.createAllDayEvent(event_h['parasha'], event_h['event_date'], event_h['event_opts']);

    //TODO: define parameters for the notification times
    var min_before_a = [20160/*2weeks*/, 10080/*1week*/, 1440/*24hours*/];
    set_notifications(new_event, min_before_a);

    var event_id  = new_event.getId();
    return event_id;
}

/*
 * Updates existing events according to new data in the spreadsheet. eventId stays untouched
 * https://stackoverflow.com/questions/40142760/google-apps-script-calendar-compare-if-2-events-are-in-the-same-day
 * Comparison needs to be made on event object versus event object basis
 */
function update_existing_event(
    cal,
    event_id,
    event_h
) {
    var existing_event = cal.getEventById(event_id);

    //If the title doesn't have "Kiddush Setup Rota", meaning we are refering to some random object 
    var re = new RegExp('[Kiddush Setup Rota]');
    //Is an existing event a kiddush setup event?
    try {
        if (re.exec(existing_event.getTitle())) {
            if (typeof existing_event === undefined) {
                Logger.log("WARNING: Event '%s' has a status '%s' (aka not replied)", existing_event.getTitle(), existing_event.status);
            } else {
                //event.setTitle(parasha);
                //event.setAllDayDate(event_date);
                //event.addGuest(email);
                Logger.log("Deleting eventId=%s", event_id);
                existing_event.deleteEvent();
            }
            event_id = create_event(
                    cal, 
                    event_h
                    );
            Logger.log("Created eventId=%s", event_id);
        } else {
            throw new Error("The existing event '%s' isn't a Kiddush Setup Event", existing_event.getTitle()); 
        }
    } catch(e) {
        e.message = "Event: \'" + existing_event.getTitle() + "\' : " + e.message;
        errHandler(e, "update_existing_event");
        event_id = undefined;
    }
    return event_id;
}

/*
 * Uses SpreadsheetApp to fetch shifts and CalendarApp to turn it to appointments.
 * The appointments will include notifications to emails coming from another spreadsheet.
 * Shifts DB: 
 */
function kiddush_setup_rota() {
    var calendar_id     = 'yehuda.tsimanis@gmail.com';
    var cal             = CalendarApp.getCalendarById(calendar_id);

    var spreadsheet_id  = '1-LMOTmoQVjpLCong8nayEmvSCE8BdSiaCQhrpp24e3g';
    var sheet_name      = "Rota";
    var ss              = SpreadsheetApp.openById(spreadsheet_id);
    var sheet           = ss.getSheetByName(sheet_name);
    var data_range      = sheet.getDataRange();
    var values          = data_range.getValues();

    var contacts_db     = create_contacts_db(
            SpreadsheetApp.openById(spreadsheet_id), 
            "Contacts"
            );

    if (contacts_db === null) {
        throw new Error("Some contacts are missing email addresses. Please, see warnings in the log.");
    }

    if (!values) {
        Logger.log('No data found!');
    } else {
        Logger.log("Scanning the ss with " + values.length + " potential events");
        for (var row = 0; row < values.length; row++) {
            // Skips blank rows
            if (values[row][0] == "") {
                break;
            }

            // Print columns A and E, which correspond to indices 0 and 4.
            var formatted_date = Utilities.formatDate(values[row][1], "GMT+2", "dd/MM/yyyy");

            var parasha      = "[Kiddush Setup Rota] Parashat " + values[row][0];
            var event_date   = values[row][1];
            var guests       = contacts_db[values[row][2]] + ',' + contacts_db[values[row][3]] + ',' + contacts_db[values[row][4]] + ',' + contacts_db[values[row][5]] + ',';
            var description  = "Captain of the week: " + values[row][2];

            //If an event is older then today, don't update or delete it
            var today = new Date;
            if (event_date < today) {
                Logger.log("WARNING: %s of %s is an old event. Skipping...", parasha, formatted_date);
                continue;
            }

            var event_opts   = {
                'description': description,
                'guests'     : guests,
                'sendInvites': 'True',
            }
            var event_h      = {
                'parasha'      : parasha,
                'event_date'   : event_date,
                'event_opts'   : event_opts
            }

            var event_id_column =   6 + 1;
            var event_id_row    = row + 1;

            var event_id        = sheet.getRange(event_id_row, event_id_column).getValue();

            if (event_id == "") {
                Logger.log("Creating a new event: %s dated: %s", parasha, formatted_date);
                event_id = create_event(
                        cal,
                        event_h
                        );
            } else {
                Logger.log("Updating an event: %s dated: %s", parasha, formatted_date);
                event_id = update_existing_event(
                        cal,
                        event_id,
                        event_h
                        );
            }
            if (event_id) {
                Logger.log("Saving eventId=" + event_id + " in ss[" + event_id_row + "][" + event_id_column + "]");
                //for further identifaction
                sheet.getRange(event_id_row, event_id_column).setValue(event_id); // row+1, as row begins with 1, while the iterator with 1
            }
        }
    }
}

//TODO:HIGH  : Don't update events that haven't been changed,as it notifies via email & people would ignore too many emails.
//TODO:HIGH  : check no event (parasha/festival name) appears twice in the sheet. Do a case-insensitive check.
//TODO:HIGH  : check that no date is assigned to 2 different events in the sheet.
//TODO:HIGH  : update_existing_event(): move the assertion part and getEventById() to the main function
//TODO:HIGH  : "No data found!" - add try-catch block
//TODO:MEDIUM: Twilio or addSmsReminder(): doesn't add SMS notification: https://www.smsclientreminders.com/google_calendar_text_reminders
//TODO:MEDIUM: update_existing_event(): modify an event rather then overriding it
//TODO:MEDIUM: add event.setColor(CalendarApp.EventColor.GREEN) colouring to different types of events
//TODO:LOW   : upgrade the logging by using https://developers.google.com/apps-script/reference/base/console
//TODO:LOW   : add some comment column to the 'description'
function main() {
    kiddush_setup_rota(); 
}
