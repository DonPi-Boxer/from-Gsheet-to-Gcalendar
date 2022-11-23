// important variables in sheet:
// var rangeCalendarName = "B1"
// var rangeCalendarID = "F1"

//// https://stackoverflow.com/questions/39319514/how-to-add-other-users-to-a-calendar

// https://developers.google.com/apps-script/reference/calendar Build-in API Google Appsscrpt for Calendar

///// Add participants to the calendar ?

//// Note add user werk met deze 'rigide data; --> hoe pakken we dit uit naar een array ?
//// moet nu iedereen de API opstarten ???



function probeer() {
var iDCalendar = SpreadsheetApp.getActiveSheet().getRange("D1").getValue();
var calendar = CalendarApp.getCalendarById(iDCalendar);

}

// All relevant general variables are under here:

var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

var dayIndex = 0;
var eventIndex = 1;
var startTimeIndex = 2;
var endTimeIndex = 3;
var locationIndex = 4;
var descriptionIndex = 5;
var teammemberIndex = 6;
var colourIndex =7;
var reminderIndex = 8;
var aanpassingIndex = 9;
var deletedornotIndex = 10;
var eventIDIndex = 11;

        
// set timezone:
  var timeZone = Session.getScriptTimeZone();
//Create varialbes for postions of: Calendar Name, link, Url description 
var cellLinkCal = "C1";
var cellIdCal = "E1";
var cellAddEditor = "J1";
var cellAddViewer = "J2";


var rangeDisplayCalendarName = "A1";
var rangeCalendarName = "B1";
var rangeUrl = "D1";
var rangeCalendarID = "F1";

//Dell from and to text
var cellDelFromText = "A2";
var cellDelToText = "C2";
//Dell from and to datet value
var cellDeleFromValue = "B2";
var cellDeleToValue = "D2";
// Set the names to put in abovestanding positions
var setCalName= "Name Google Calendar: ";
var setLinkCal = "Calendar link: ";
var setLinkID = "Calendar ID: ";
var setDelFrom = "Delete from: ";
var setDelTo = "Delete until: ";
var setAddEditor = "Add editors: " ;
var setAddViewers = "Add viewers: ";

var sheetname = SpreadsheetApp.getActiveSheet().getName();



/// variables for adding editors and subscribers
var firstCellOfMails = "L1";
var firstMailCellRange = SpreadsheetApp.getActiveSheet().getRange(firstCellOfMails);
var rowMailsIndex = 1;
var firstMailCollumnIndex = 11;

var rowsMailsIndexViewer = 2;

var rowMailsA1 = 1;
var firsCollumnMailsA1 = "L";


var startpointentries = "A5";
var endcolumncalendarentries = "L";

 
var calendarName = activesheet.getRange(rangeCalendarName).getValue();
var calendarId = activesheet.getRange(rangeCalendarID).getValue();


//Create variables for positions of: Calendar entrie headers
var cellDate = "A4";
var cellEventName = "B4";
var cellStart = "C4";
var cellEnd = "D4";

var cellLocation = "E4"
var cellDescription = "F4"
var teammembers = "G4"
var colourdescription = "H4"
var reminderdescription = "I4"
var cellAChange = "J4";
var cellDelete ="K4";
var cellEventid = "L4";

// Set the names to put in abovestanding positions. NOte: how to put this in bolt ???
var setDate = "Date: "
var setEventName = "Event name: "
var setCellStart = "Starttime: "
var setCellEnd = "Endtime: "
var setLocation = "Location: "
var setDescription = "Description" ; 
var setTeaammembers = "(Taken) betrokken teams/members"
var setcolordesciptiion = "Status/ colour: "
var setReminderdescription = "Reminder days:"
var setChange = "Change "
var setDelete = "Delete "
var setEventid = "Event ID: "



/// Data validation for status/color:

  /// Declare the datavalidation variabele for colour
   
  var option = new Array();
      option[0]= "Definitive" ;
      option[1]="Concept";
      option[2]="Plan"

  

// Set date DeleFrom and DeleteTo on the current date
function createcalender()
{
/// Check if you already own this Calendar
SpreadsheetApp.getActiveSpreadsheet().toast("Creating a new Google Calendar with the name: " + calendarName )

var newcal = CalendarApp.createCalendar(calendarName).setTimeZone(timeZone);
var newcalId = newcal.getId();
Logger.log(newcalId);
 
//  Logger.log("The name of the Calendar belonging to the given ID is " + calendarId);
  //set calendar link in cell D1
var urlpart1 = "https://calendar.google.com/calendar/embed?src="
var urltotal =  urlpart1 + newcalId;
//  Logger.log("The link to the Calendar belonging to the specified Calendar is " + urltotal);

  //Set calendar Id in cell F1

//  Logger.log("The ID the Calendar belonging to the given ID is " + calendarId);


/////PART OF THE SETLINKENIDINSHEET
activesheet.getRange(rangeUrl).setValue(urltotal);
activesheet.getRange(rangeCalendarID).setValue(newcalId);

SpreadsheetApp.getActiveSpreadsheet().toast("A new Google Calendar with the name: " + calendarName + " is created.")

}


function setformatofDelete() {
activesheet.getRange(cellDeleFromValue).setNumberFormat("dd-mm-yyyy").setFontWeight("normal");
activesheet.getRange(cellDeleToValue).setNumberFormat("dd-mm-yyyy").setFontWeight("normal");
}

function createsheet() {
//get the activesheet


setformatofDelete()

// Set all variables in positions of the sheet. Here we set all header entries. Not the values themselves !
activesheet.getRange(rangeDisplayCalendarName).setValue(setCalName).setFontWeight("bold")
activesheet.getRange(cellLinkCal).setValue(setLinkCal).setFontWeight("bold");
activesheet.getRange(cellIdCal).setValue(setLinkID).setFontWeight("bold");
activesheet.getRange(cellDelFromText).setValue(setDelFrom).setFontWeight("bold");
activesheet.getRange(cellDelToText).setValue(setDelTo).setFontWeight("bold");
activesheet.getRange(cellDate).setValue(setDate).setFontWeight("bold");
activesheet.getRange(cellStart).setValue(setCellStart).setFontWeight("bold");
activesheet.getRange(cellEnd).setValue(setCellEnd).setFontWeight("bold");
activesheet.getRange(cellEventName).setValue(setEventName).setFontWeight("bold");
activesheet.getRange(cellLocation).setValue(setLocation).setFontWeight("bold");
activesheet.getRange(cellDescription).setValue(setDescription).setFontWeight("bold");
activesheet.getRange(teammembers).setValue(setTeaammembers).setFontWeight("bold");
activesheet.getRange(colourdescription).setValue(setcolordesciptiion).setFontWeight("bold");
activesheet.getRange(reminderdescription).setValue(setReminderdescription).setFontWeight("bold");
activesheet.getRange(cellAChange).setValue(setChange).setFontWeight("bold");
activesheet.getRange(cellDelete).setValue(setDelete).setFontWeight("bold");
activesheet.getRange(cellEventid).setValue(setEventid).setFontWeight("bold");


activesheet.getRange(cellAddEditor).setValue(setAddEditor).setFontWeight("bold");
activesheet.getRange(cellAddViewer).setValue(setAddViewers).setFontWeight("bold");

  // Set date DeleFrom and DeleteTo on the current date, such that those cell are not empty, thus the onOpenTrigger works properly
  var deleDay = new Date();
 
  var fomatteddate  = Utilities.formatDate(deleDay, timeZone,"yyyy-MM-dd");
  var deledatestring = fomatteddate.substring(0,10);

  // set right date formats + input this day as delete from and delete to
  activesheet.getRange(cellDeleFromValue).setValue(deledatestring).setNumberFormat("dd-mm-yyyy");
  activesheet.getRange(cellDeleToValue).setValue(deledatestring).setNumberFormat("dd-mm-yyyy");

  activesheet.getRange("A5:A").setNumberFormat("dd-mm-yyyy");
  activesheet.getRange("C5:C").setNumberFormat("h:mm");
  activesheet.getRange("D5:D").setNumberFormat("h:mm");

  activesheet.getRange("A:Z").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  /// Conditional formatting for color:
  //var rule = new Array
  /// rule[0] = "Definitive"

activesheet.autoResizeColumn(1);
activesheet.autoResizeColumn(2);


// set dv vooor status agenda:

/// Declare DV:
    var dv = SpreadsheetApp.newDataValidation();
    dv.setAllowInvalid(false);
    dv.setHelpText("Choose an event status/color");
    dv.requireValueInList(option, true);

  activesheet.getRange(5, colourIndex+1).setDataValidation(dv);

}


function createTotalSheetAndCalendar () {
createcalender()
createsheet()
}

// Change Name of the Calendar if this get's changed in the Google Sheet
function alterCalendarName() {
  
  //declare variables: active sheet, cell of calendar ID, cell of Calendar name
  
 

  // call the ID of the altered calendar. Call the agenda
  var calendarId = activesheet.getRange(rangeCalendarID).getValue();
  var alteredCalendar = CalendarApp.getCalendarById(calendarId);
 
  //get new name of the calendar
  
  var calendararrayTest = CalendarApp.getOwnedCalendarsByName(calendarName);
Logger.log(calendararrayTest)
if (calendararrayTest == 0) {


  // set new name of the calendar in the Google Calendar
  alteredCalendar.setName(calendarName);

  //popup: name of the calendar is changed to..
  SpreadsheetApp.getActiveSpreadsheet().toast("The name of the Google Calendar is changed into: " + calendarName);
}
else {
var ui = SpreadsheetApp.getUi();
var response = ui.alert("'You already own a Google Calendar with the name '" + calendarName + "' please pick a different name and try again.", ui.ButtonSet.OK);
SpreadsheetApp.getActiveSpreadsheet().toast("Changing the name of Google Calendar '" + calendarName + "' is canceled.")

}
  

}


function addEditors() {

var addedUsers = [];

// Find all Filled in Cells in the row with Emails: we start at firstIndexOfMails and end with lastCollumn command:

var lastMailCollumnIndex = SpreadsheetApp.getActiveSheet().getLastColumn();
Logger.log(lastMailCollumnIndex);
Logger.log(firstMailCollumnIndex);
var lastMailCellRange = SpreadsheetApp.getActiveSheet().getRange(rowMailsIndex,lastMailCollumnIndex);
// Logger.log(lastMailCellRange);
// Logger.log(firstMailCellRange);


var totalAmmOfMails = (lastMailCollumnIndex - firstMailCollumnIndex +1);
//Logger.log(totalAmmOfMails );
var mailArray = SpreadsheetApp.getActiveSheet().getRange(rowMailsIndex,firstMailCollumnIndex,1, totalAmmOfMails).getValues();
// Logger.log(mailArray);
var mailVector = mailArray[0];
// Logger.log(mailVector);


for (i=0; i<mailVector.length; i++) {
  
  
  var userEmail = mailVector[i];
  var collumnsUserEmail = firstMailCollumnIndex + i;
 // Logger.log(userEmail);

  if (userEmail.includes("@"))  {


/// Stukje API voor toevoegen van de email
  var resource = {
      'scope': {
       'type': 'user',
       'value': userEmail
    },
     'role': 'owner'
    };
  //  Logger.log("dit includes at " + userEmail);
    Calendar.Acl.insert(resource, calendarId);
/// API eindigt hierboven; niet aanzitten !

/// Knip substring email zonder de mailinfo en zit hem in de gecallde Cell
var indexofat = userEmail.indexOf("@");
var mailAlleenNaam = userEmail.substring(0,indexofat);
activesheet.getRange(rowMailsIndex, collumnsUserEmail).setValue(mailAlleenNaam);

/// Create array met alle toegevoegde Emails: Display deze in de toast.

///AddedUser to put in the toast message:
addedUsers.push(userEmail);
  }
}


var addedUsersString = [];

for (j=0; j<addedUsers.length; j++) {

var addedUsersString = addedUsersString + " " + addedUsers[j];
}
var calendarName = SpreadsheetApp.getActiveSheet().getRange(rangeCalendarName).getValue();
SpreadsheetApp.getActiveSpreadsheet().toast("Editor invitation for '" + calendarName + "' is send to: "   +addedUsersString);
}

function addViewers() {

var addedUsers = [];


// zet dit boven functie !!!!

// Find all Filled in Cells in the row with Emails: we start at firstIndexOfMails and end with lastCollumn command:

var lastMailCollumnIndex = SpreadsheetApp.getActiveSheet().getLastColumn();
// Logger.log(lastMailCollumnIndex);
// Logger.log(firstMailCollumnIndex);
var lastMailCellRange = SpreadsheetApp.getActiveSheet().getRange(rowsMailsIndexViewer,lastMailCollumnIndex);
// Logger.log(lastMailCellRange);
// Logger.log(firstMailCellRange);


var totalAmmOfMails = (lastMailCollumnIndex - firstMailCollumnIndex +1);
// Logger.log(totalAmmOfMails );
var mailArray = SpreadsheetApp.getActiveSheet().getRange(rowsMailsIndexViewer,firstMailCollumnIndex,1, totalAmmOfMails).getValues();
// Logger.log(mailArray);
var mailVector = mailArray[0];
// Logger.log(mailVector);


for (i=0; i<mailVector.length; i++) {
  
  
  var userEmail = mailVector[i];
  var collumnsUserEmail = firstMailCollumnIndex + i;
 // Logger.log(userEmail);

  if (userEmail.includes("@"))  {


/// Stukje API voor toevoegen van de email
  var resource = {
      'scope': {
       'type': 'user',
       'value': userEmail
    },
     'role': 'reader'
    };
  //  Logger.log("dit includes at " + userEmail);
    Calendar.Acl.insert(resource, calendarId);
/// API eindigt hierboven; niet aanzitten !

/// Knip substring email zonder de mailinfo en zit hem in de gecallde Cell
var indexofat = userEmail.indexOf("@");
var mailAlleenNaam = userEmail.substring(0,indexofat);
activesheet.getRange(rowsMailsIndexViewer, collumnsUserEmail).setValue(mailAlleenNaam);

/// Create array met alle toegevoegde Emails: Display deze in de toast.

///AddedUser to put in the toast message:
addedUsers.push(userEmail);
  }
}


var addedUsersString = [];

for (j=0; j<addedUsers.length; j++) {

var addedUsersString = addedUsersString + " " + addedUsers[j];
}
var calendarName = SpreadsheetApp.getActiveSheet().getRange(rangeCalendarName).getValue();
SpreadsheetApp.getActiveSpreadsheet().toast("Reader invitaion for '" + calendarName + "' is send to: "   +addedUsersString);
}


// Function voor Omzetten van de Activiteiten sheet (sheet2) naar Kalender sheet
function zetActiviteitOpCalendar() {
    

  // Verkrijg time zone

  
  // get active sheet
  var spreadsheetevents = SpreadsheetApp.getActive().getActiveSheet();
  //get targeted calendar
  var calendarIdevents = spreadsheetevents.getRange(rangeCalendarID).getValue();
  Logger.log("calendar is " + calendarIdevents);
  var calOutEvents = CalendarApp.getCalendarById(calendarIdevents);
Logger.log("calendar is " + calOutEvents);
  

  
  var startpointentriesA1 = spreadsheetevents.getRange(startpointentries).getA1Notation();

  
  var endrowentriess = spreadsheetevents.getLastRow();
  var endpointentries = endcolumncalendarentries+endrowentriess;
  var endpointentriesA1 = spreadsheetevents.getRange(endpointentries).getA1Notation();

  Logger.log("A1end is " + endpointentriesA1);

  Logger.log("A1start is " + startpointentriesA1);
  
  var stringstart = Utilities.formatString(startpointentriesA1);
  var stringend = Utilities.formatString(endpointentriesA1);
  // Verkrijg: Sheet en kalender 
  var starttoend = stringstart +":"+ stringend;
  Logger.log("A1start is " + starttoend);
  

  // Put relevante data voor input in kalender (is gekozen range). Let op veranderen van deze range !!
  // Huidige matrix bestaat uit: [Start, Einde, Event]
  // rangeevent --> how many events are there actually ? --> put in search range voor callInevents
 
  var calInEvents = spreadsheetevents.getRange(starttoend).getValues();

  // Teller voor aantal added, aantal deleted en aantal gewijzigde evenementen.
  var numberaddedevents = 0;
  var numberalteredevents = 0;
  var numberdeletedevents = 0;

/// Declare DV:
    var dv = SpreadsheetApp.newDataValidation();
    dv.setAllowInvalid(false);
    dv.setHelpText("Choose an event status/color");
    dv.requireValueInList(option, true);


  // Zet data matrix om naar variabelen voor in calendar. (Eventueel mogelijk 2 forloops te gebruiken om shift[0] etc niet handmatig in te stellen)
  for (x=0; x<calInEvents.length; x++) {

  var  totalcount = x;

 // Logger.log(calInEvents[x]);
  // Kies een rij in de matrix   
      var shift = calInEvents[x];
  // Deze 2 variabelen bepalen of forloop wel of niet gestart wordt (aanpassingen voor aanpassingen, eventId voor nieuwe)  
      var day = shift[dayIndex];
      var startTime = shift[startTimeIndex];
      var endTime = shift[endTimeIndex];
      var event = shift[eventIndex];  
      var location = shift[locationIndex];
      var description = shift[descriptionIndex];
      var teammembers = shift[teammemberIndex];
      var colour = shift[colourIndex];
      var remindertime = 24 * 60* shift[reminderIndex];
      var aanpassing = shift[aanpassingIndex];
      var deletedornot = shift[deletedornotIndex];
      var eventID = shift[eventIDIndex];

  Logger.log("dag is" + day);
  Logger.log("startime is" + startTime);
  Logger.log("endTime is " +endTime);   
Logger.log("event is" + event);
Logger.log("location is " + location);
Logger.log("description is" + description);
Logger.log("colour is" + colour);
Logger.log("remindertime is" + remindertime);
Logger.log("aanpassing is" + aanpassing);
Logger.log("deletedornot is" + deletedornot);
Logger.log("event Id is" + eventID);

  if (event == "" )  {
  continue
  }

  if (day == "" )  {
  continue
  }

 
// Set data validation for colour in evy colour cell

activesheet.getRange((5 + x - numberdeletedevents), colourIndex+1).setDataValidation(dv);
  // Declare colour value          
    if (colour == "Definitive") {
    activesheet.getRange(5 + x - numberdeletedevents, colourIndex+1).setBackground("green");
    activesheet.getRange(5 + x - numberdeletedevents, colourIndex+1).setFontColor("white");
    var colourValue = 10;

  }

    else if (colour == "Concept") {
    activesheet.getRange(5 + x - numberdeletedevents, colourIndex+1).setBackground("yellow");
    activesheet.getRange(5 + x - numberdeletedevents, colourIndex+1).setFontColor("black");
    var colourValue = 5;

  }
    
      else if (colour == "Plan") {
    activesheet.getRange(5 + x - numberdeletedevents, colourIndex+1).setBackground("blue");
    activesheet.getRange(5 + x - numberdeletedevents, colourIndex+1).setFontColor("white");
    var colourValue = 9;

  }

    else if (colour == ""){
     activesheet.getRange(5 + x - numberdeletedevents, colourIndex+1).setBackground("white");
    }
   

  // If statement: voeg Id toe bij ieder toegevoet event. If no eventId en hele dag --> add new alldayevent. Alsook geen description --> zonder description
  if (eventID == "" && startTime == "" && event != ""){


/// Declaren reminder time:
  /// Declareren kleurindex:

          if (description == "" && teammembers == "") {
    // Create new event: set aanpassing op nee, set verwijder op nee, set new ID
        var newDagEvent = calOutEvents.createAllDayEvent(event, day, {location: location} );
        var newDagEventId = newDagEvent.getId();
        activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(newEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        //Logger + counter
  activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
  activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
  activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(newDagEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd"); 
  var datestring = fomatteddate.substring(0,10);
  SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is added. It is set for the whole day on " + datestring);
  var numberaddedevents = (numberaddedevents+1);
        
        }

       else if (description != "" && teammembers == "") {
    var newDagEvent = calOutEvents.createAllDayEvent(event, day, {description: description, location: location} );
    var newDagEventId = newDagEvent.getId();
  activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
  activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
  activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(newDagEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd"); 
  var datestring = fomatteddate.substring(0,10);
  SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is added. It is set for the whole day on " + datestring);
  var numberaddedevents = (numberaddedevents+1);
      
        }

        
      else  if (description == "" && teammembers != "") {
        var newDagEvent = calOutEvents.createAllDayEvent(event, day, {description: teammembers, location: location} );
        var newDagEventId = newDagEvent.getId();
     
        //Logger + counter
  activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
  activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
  activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(newDagEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd"); 
  var datestring = fomatteddate.substring(0,10);
  SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is added. It is set for the whole day on " + datestring);
  var numberaddedevents = (numberaddedevents+1);
        }
      else  if (description != "" && teammembers != "") {
        var totaldescription =  setTeaammembers + "\n" + teammembers + "\n" + "Description of event:" + "\n" + description;
        Logger.log("total description is" + totaldescription);
        var newDagEvent = calOutEvents.createAllDayEvent(event, day, {description: totaldescription, location: location} );
        var newDagEventId = newDagEvent.getId();
 
        //Logger + counter 
  activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
  activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
  activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(newDagEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd"); 
  var datestring = fomatteddate.substring(0,10);
  SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is added. It is set for the whole day on " + datestring);
  var numberaddedevents = (numberaddedevents+1);
      
        }
   


if (colour != ""){
newDagEvent.setColor(colourValue);


}

if (remindertime > 0) {
  newDagEvent.addEmailReminder(remindertime);
  newDagEvent.addPopupReminder(remindertime);
  }
  }

    else if (eventID == "" && startTime != "" && event != "") { // Check the event doesn't exist
  // ALs event ID bestaat niet en is niet hele dag --> create specific time ID

  
  /// Declareren kleurindex:
       
    //Correctly format date and time  

        /// format date
        var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd"); 
        var datestring = fomatteddate.substring(0,10);

        ///// format start
        var formattedstarttime = Utilities.formatDate(startTime, timeZone, "HH:mm:ss");
        var timestartstring = formattedstarttime.substring(0,7);
        var eventStartString = (datestring + " " + timestartstring);
        var eventStartDate = new Date(eventStartString);
        Logger.log("eventstart is " + eventStartDate);


       ///Format end
        var formattedendttime = Utilities.formatDate(endTime, timeZone, "HH:mm:ss");
        var timeendstring = formattedendttime.substring(0,7);
        var eventEndString = (datestring + " " + timeendstring);
        var eventEndDate = new Date(eventEndString);

         Logger.log("event end is " + eventEndDate);

        if ((description == "") &&  (teammembers == "")) {
    // Create new event: set aanpassing op nee, set verwijder op nee, set new ID
    var newEvent = calOutEvents.createEvent(event, eventStartDate, eventEndDate, {location: location} );
      var newEventId = newEvent.getId();
        activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(newEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        //Logger + counter
        SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is added. It takes place on " + datestring + " from " + timestartstring + " to " + timeendstring);
        var numberaddedevents = (numberaddedevents+1);
        
        }

       else if (description != "" && teammembers == "") {
        var newEvent = calOutEvents.createEvent(event, eventStartDate, eventEndDate, {description: description, location: location} );
        var newEventId = newEvent.getId();
        activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(newEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        //Logger + counter
        SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is added. It takes place on " + datestring + " from " + timestartstring + " to " + timeendstring);
        var numberaddedevents = (numberaddedevents+1);
      
        }

        
      else  if (description == "" && teammembers != "") {
        var newEvent = calOutEvents.createEvent(event, eventStartDate, eventEndDate, {description: teammembers, location: location} );
        var newEventId = newEvent.getId();
        activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(newEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        //Logger + counter
        SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is added. It takes place on " + datestring + " from " + timestartstring + " to " + timeendstring);
        var numberaddedevents = (numberaddedevents+1);
      
        }
      else  if (description != "" && teammembers != "") {
        var totaldescription =  setTeaammembers + "\n" + teammembers + "\n" + "Description of event:" + "\n" + description;
        var newEvent = calOutEvents.createEvent(event, eventStartDate, eventEndDate, {description: totaldescription, location: location} );
        var newEventId = newEvent.getId();
        activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(newEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        //Logger + counter
        SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is added. It takes place on " + datestring + " from " + timestartstring + " to " + timeendstring);
        var numberaddedevents = (numberaddedevents+1);
      
        }

            
      if (colour != ""){
      newEvent.setColor(colourValue);
 }

      if (remindertime > 0) {
        newEvent.addEmailReminder(remindertime);
        newEvent.addPopupReminder(remindertime);
        }
        

      }
      
    // If statement: als aanpassing ja en time1 is dag -> delete en recreate all day event
    else if (aanpassing == true && startTime == "" && event != "")  {
      var alteredevent = calOutEvents.getEventById(eventID);
      alteredevent.deleteEvent();
  activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).setValue("false");
  activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).setValue("false");

      
          if ((description == "") && (teammembers == "")) {
    // Create new event: set aanpassing op nee, set verwijder op nee, set new ID


      var alteredDagEvent = calOutEvents.createAllDayEvent(event, day, {location: location} );
     var alteredDagEventID = alteredDagEvent.getId();
 
      activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).setValue("false");
      activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).setValue("false");
      activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(alteredDagEventID).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd"); 
  var datestring = fomatteddate.substring(0,10);
  SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is altered. It is set for the whole day on " + datestring);
  var numberaddedevents = (numberaddedevents+1);
        
        }

       else if (description != "" && teammembers == "") {
  var teammembersTot = setTeaammembers + "\n" + teammembers;
  var alteredDagEvent = calOutEvents.createAllDayEvent(event, day, {description: teammembersTot, location: location} );
  var alteredDagEventID = alteredDagEvent.getId();
  activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).setValue("false");
  activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).setValue("false");
  activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(alteredDagEventID).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd"); 
  var datestring = fomatteddate.substring(0,10);
  SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is altered. It is set for the whole day on " + datestring);
  var numberaddedevents = (numberaddedevents+1);
      
        }

        
      else  if (description == "" && teammembers != "") {
     var alteredDagEvent = calOutEvents.createAllDayEvent(event, day, {description: teammembers, location: location} );
       var  alteredDagEventID = alteredEvent.getId();
     
        //Logger + counter
  activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).setValue("false");
  activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).setValue("false");
  activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(alteredDagEventID).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd"); 
  var datestring = fomatteddate.substring(0,10);
  SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is altered. It is set for the whole day on " + datestring);
  var numberaddedevents = (numberaddedevents+1);
        }

      else  if (description != "" && teammembers != "") {
        var totaldescription =  setTeaammembers + "\n" + teammembers + "\n" + "Description of event:" + "\n" + description;
       var alteredDagEvent = calOutEvents.createAllDayEvent(event, day, {description: totaldescription, location: location} );
      var alteredDagEventID = alteredDagEvent.getId();
 
        //Logger + counter 
  activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).setValue("false");
  activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).setValue("false");
  activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(alteredDagEventID).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd"); 
  var datestring = fomatteddate.substring(0,10);
  SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is altered. It is set for the whole day on " + datestring);
  var numberaddedevents = (numberaddedevents+1);
      
        }

    
      

      
      if (colour != ""){
      alteredDagEvent.setColor(colourValue);
 }

      if (remindertime > 0) {
        alteredDagEvent.addEmailReminder(remindertime);
        alteredDagEvent.addPopupReminder(remindertime);
        }

      
    }

    // IF statement: als aanpassing en time1 is NIET hele dag --> delete en recreate specific event
      else if (aanpassing == true && startTime != "" && event != ""){

  activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).setValue("false");
  activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).setValue("false");

    
    //Declare variables for in Calendar
      //Correctly format date and time 
        var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd");
        var datestring = fomatteddate.substring(0,10);
        var formattedstarttime = Utilities.formatDate(startTime, timeZone, "HH:mm:ss");
        var timestartstring = formattedstarttime.substring(0,7);
        var formattedendttime = Utilities.formatDate(endTime, timeZone, "HH:mm:ss");
        var timeendstring = formattedendttime.substring(0,7);
        var  eventStartString = (datestring + " " + timestartstring);
        var  eventEndString = (datestring + " " + timeendstring);
        var eventStartDate = new Date(eventStartString);
        var eventEndDate = new Date(eventEndString);
        var deletedEvent = calOutEvents.getEventById(eventID);
        deletedEvent.deleteEvent(); 


   if (description == "" && teammembers == "") {
    // Create new event: set aanpassing op nee, set verwijder op nee, set new ID
       var alteredEvent = calOutEvents.createEvent(event, eventStartDate, eventEndDate, {location: location} );
     var alteredEventId = alteredEvent.getId();
        activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(alteredEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        //Logger + counter
        SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is altered. It takes place on " + datestring + " from " + timestartstring + " to " + timeendstring);
        var numberaddedevents = (numberaddedevents+1);
        
        }

       else if (description != "" && teammembers == "") {
       var alteredEvent = calOutEvents.createEvent(event, eventStartDate, eventEndDate, {description: description, location: location} );
       var alteredEventId = alteredEvent.getId();
        activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(alteredEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        //Logger + counter
        SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is altered. It takes place on " + datestring + " from " + timestartstring + " to " + timeendstring);
        var numberaddedevents = (numberaddedevents+1);
      
        }

        
      else  if (description == "" && teammembers != "") {
        var alteredEvent = calOutEvents.createEvent(event, eventStartDate, eventEndDate, {description: teammembers, location: location} );
       var alteredEventId = alteredEvent.getId();
        activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(alteredEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        //Logger + counter
        SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is altered. It takes place on " + datestring + " from " + timestartstring + " to " + timeendstring);
        var numberaddedevents = (numberaddedevents+1);
      
        }
      else  if (description != "" && teammembers != "") {
        var totaldescription =  setTeaammembers + "\n" + teammembers + "\n" + "Description of event:" + "\n" + description;
      var alteredEvent = calOutEvents.createEvent(event, eventStartDate, eventEndDate, {description: totaldescription, location: location} );
     var alteredEventId = alteredEvent.getId();
        activesheet.getRange(5 + x - numberdeletedevents, aanpassingIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, deletedornotIndex+1).insertCheckboxes();
        activesheet.getRange(5 + x - numberdeletedevents, eventIDIndex+1).setValue(alteredEventId).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        //Logger + counter
        SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is altered. It takes place on " + datestring + " from " + timestartstring + " to " + timeendstring);
        var numberaddedevents = (numberaddedevents+1);
      
        }

        
        if (colour != ""){
alteredEvent.setColor(colourValue);

}

if (remindertime > 0) {
       alteredEvent.addEmailReminder(remindertime);
      alteredEvent.addPopupReminder(remindertime);
  }
        
        
var alteredEventId = alteredEvent.getId();
              
//Logger + counter
       SpreadsheetApp.getActiveSpreadsheet().toast("Event " + event + " is altered. It takes place on " + datestring + " from " + timestartstring + " to " +timeendstring);
        var numberalteredevents = (numberalteredevents + 1);

      if (colour != ""){
      alteredEvent.setColor(colourValue);
      }

      if (remindertime > 0) {
        alteredEvent.addEmailReminder(remindertime);
        alteredEvent.addPopupReminder(remindertime);
        }


    }

    // als deleted is true --> delete event + clear de hele row (en delete de row ook)

    // delete timed event
      else if  (deletedornot == true && startTime != "") {
      var deletedevent = calOutEvents.getEventById(eventID);
       // Formate date in time for input in logger
       var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd");
        var datestring = fomatteddate.substring(0,10);
        var formattedstarttime = Utilities.formatDate(startTime, timeZone, "HH:mm:ss");
        var timestartstring = formattedstarttime.substring(0,7);
        var formattedendttime = Utilities.formatDate(endTime, timeZone, "HH:mm:ss");
        var timeendstring = formattedendttime.substring(0,7);

        //delete
        deletedevent.deleteEvent();
        SpreadsheetApp.getActiveSpreadsheet().toast(event + " is deleted on date " + datestring + " from " + timestartstring + " to " +timeendstring);

      }
        
        /// delete alldayevent

        else if  (deletedornot == true && startTime == "") {
      var deletedevent = calOutEvents.getEventById(eventID);
       // Formate date in time for input in logger
       var fomatteddate  = Utilities.formatDate(day, timeZone,"yyyy-MM-dd");
       
        //delete
        deletedevent.deleteEvent();
        SpreadsheetApp.getActiveSpreadsheet().toast(event + " is deleted on date ") ; 
        //Logger
        
        activesheet.deleteRow(5+ x - numberdeletedevents);
        var numberdeletedevents = (numberdeletedevents + 1);
} 
  }


  /// Set data validation for future new event

    activesheet.getRange((5 + calInEvents.length - numberdeletedevents), colourIndex+1).setDataValidation(dv);

  if (numberaddedevents >0 ){
  SpreadsheetApp.getActiveSpreadsheet().toast("Total number of created events is " + numberaddedevents);
} 
  // if counter > 0 --> geef log van de counter:
  if (numberalteredevents > 0){
  SpreadsheetApp.getActiveSpreadsheet().toast("Total number of altered events is " + numberalteredevents);
}
if (numberdeletedevents >0 ){
  SpreadsheetApp.getActiveSpreadsheet().toast("Total number of deleted events is " + numberdeletedevents);
}

}
//// NOTE: cleart niet van tot en met maar van+1 tot: hoe fiksen we dit ??? 
//// Later zorg !
// Clear de Activiteiten calendar
function clearCalendarActiviteiten() {

    var fromDate = activesheet.getRange("B2").getValue();
  var toDate = activesheet.getRange("D2").getValue();

   var fomatteddateFrom  = Utilities.formatDate(fromDate, timeZone,"yyyy-MM-dd");
        var fromDatestring = fomatteddateFrom.substring(0,10);

var fomatteddateTo  = Utilities.formatDate(toDate, timeZone,"yyyy-MM-dd");
        var toDatestring = fomatteddateTo.substring(0,10);

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Are you sure you want to clear all activitities in Calendar '" + calendarName + "' from " + fromDatestring + " to " + toDatestring , ui.ButtonSet.YES_NO);

  // Welke calendar ?

  if (response == ui.Button.NO)
{
  return
}  
else {
  var spreadsheetevents = SpreadsheetApp.getActive().getActiveSheet();
  //get targeted calendar
  var calendarIdevents = spreadsheetevents.getRange("F1").getValue();
  var calOutEvents = CalendarApp.getCalendarById(calendarIdevents);
  // Welke dagen ? Note afhaneklijkheid van range als gegeven in Variabele spreadsheet !


  var numberofdeleted = 0;
  // Ga door list event en delete elk individueel event
  var events = calOutEvents.getEvents(fromDate, toDate);
  for(var i=0; i<events.length;i++){
    var ev = events[i];
        ev.deleteEvent();
        var numberofdeleted = (numberofdeleted+1);
      }
    SpreadsheetApp.getActiveSpreadsheet().toast("Deleted  " + numberofdeleted + " event. Between the date " + fromDatestring + " and " + toDatestring);
}
  }

/// Clear alle activiteiten uit de sheet van tot en met de gegeven datum. Note: nu bewegen we door van tm datum de IDs in de calendar te vinden en deze te zoeken in de sheet: als gevonden --> delete. Kan wellicht sneller door ook gewoon door de datums heen te loopen in de sheet zelf ?
function clearsheet()

{

  // Get spreadsheet. Get range of events. Get timezone
//var timeZone = Session.getScriptTimeZone();
  var spreadsheetevents = SpreadsheetApp.getActive().getActiveSheet();
  var rangeevents = spreadsheetevents.getLastRow();
  var sheetname = SpreadsheetApp.getActiveSheet().getName();

  // Count numbers of rows that have been deleted, so we van adjust for that in our sheet when we need to delete a 2nd row
  var alreadydeleted = 0;

  // get dates inbetween we want to clear. Set right format
  var fromDate = spreadsheetevents.getRange("B2").getValue();
  var toDate = spreadsheetevents.getRange("D2").getValue();

  var fomatteddateFrom  = Utilities.formatDate(fromDate, timeZone,"yyyy-MM-dd");
        var fromDatestring = fomatteddateFrom.substring(0,10);

var fomatteddateTo  = Utilities.formatDate(toDate, timeZone,"yyyy-MM-dd");
        var toDatestring = fomatteddateTo.substring(0,10);

 var ui = SpreadsheetApp.getUi();
   var response = ui.alert("Are you sure you want to clear all activitities in this sheet '" + sheetname + "' from " + fromDatestring + " to " + toDatestring , ui.ButtonSet.YES_NO);

     if (response == ui.Button.NO)
{
  return
} 


else {
  //Set right format
  var formatFromdate = new Date(fromDate).getTime();
  var formatTodate = new Date(toDate).getTime();

//  Logger.log("fromdate is " + fromDate);
//  Logger.log("fromdatestring is " + fromDatestring);
//  Logger.log("formatFromdate is " + formatFromdate);
//  Logger.log("toDate is " + toDate);
//  Logger.log("todatestring is " + toDatestring);
//  Logger.log("formatFromdate is " + formatTodate);

  /// Get the dates of ALL events in the spreadsheet
  var datesofAllEvents = spreadsheetevents.getRange("A5:A" + rangeevents).getValues();
 // Logger.log("datesofAllevents is " + datesofAllEvents);

  /// Check for each individual date whether it is in between from and todate --> if yes: delete from sheet
  for (i=0; i<datesofAllEvents.length; i++) {
    // Get date of this iteration
    var dateofi = datesofAllEvents[i];

    var formatDateofi = new Date(dateofi).getTime();
 //   Logger.log("todatestring is " + dateofistring);
//  Logger.log("formatdateofI is " + formatDateofi);
      
        if (formatDateofi <= formatTodate && formatDateofi >= formatFromdate) {
         
        Logger.log("event on this time gets deleted out sheet " + formatDateofi);
        
        // -already deleted makes sure that we keep the rows up to date to  their actual number after being deleted
        spreadsheetevents.deleteRow(5 + i - alreadydeleted);
        // Count the number of times something has been deleted: adjust the deleterow statement above for this such that rows get updated to their 
        // new position
        var alreadydeleted = (alreadydeleted + 1);
      }
  }
  Logger.log("Number of deleted events is " +  alreadydeleted);
}
}
  
 
  
  

function clearSheetenCalendar() {

clearCalendarActiviteiten()
clearsheet()

}

function deleteCalendar() {

  var ui = SpreadsheetApp.getUi();
   var response = ui.alert("Are you sure you want to delete Google Calendar:" + calendarName , ui.ButtonSet.YES_NO);

     if (response == ui.Button.NO)
{
  return
} 

else {
CalendarApp.getCalendarById(calendarId).deleteCalendar();
}
}

function deleteSheet ()  {

  
    var ui = SpreadsheetApp.getUi();
   var response = ui.alert("Are you sure you want to delete Google Sheet:" + sheetname  , ui.ButtonSet.YES_NO);

     if (response == ui.Button.NO)
{
  SpreadsheetApp.getActiveSpreadsheet().toast("Deletetion of Google Sheet '" + sheetname + "' is canceled.")
  return
} 
SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet();
}

function deleteSheetEnCalendar() {
deleteSheet()
deleteCalendar()
}

/// Note: add ui submenu with: change name of calendar

//Set de name for the called upon Calendar (called by ID in B1, which is set by hand) in D1


// set Gui for current sheet ! Dus zeg "clear activiteiten van SHEET (NAAM) uit Calendar (NAAM)"
// Function coor creeeren GUI

// NOTE: GUI GEEFT NU; YYYY-MM-DD, WIL DIT HET LIEFSST ANDERSOM; HOE FIKSEN WE DIT ?
// function setGUISync(nameCalendar) {
  function setGUISync() {

  var nameCalendar = SpreadsheetApp.getActive().getActiveSheet().getRange(rangeCalendarName).getValue();
  var namesheet = SpreadsheetApp.getActive().getActiveSheet().getName();
  var timeZone = Session.getScriptTimeZone();

  //set submenu 1) Sync
  var subMenusync = SpreadsheetApp.getUi().createMenu("Update Calendar")
    .addItem("Update activities to Google Calendar '" + calendarName + " ", 'zetActiviteitOpCalendar') 
    
  //set submenu 2) Clear
  /// dit is voor input van tot wanneer er gecleard wordt.
  var spreadsheetevents = SpreadsheetApp.getActive().getActiveSheet();
//  var spreadsheet = SpreadsheetApp.getActive().getSheetByName("Variabelen voor script");

  var fromDate = spreadsheetevents.getRange(cellDeleFromValue).getValue();
  var toDate = spreadsheetevents.getRange(cellDeleToValue).getValue();
  // format from en to data voor implementatie in gui

  var fromfomatteddate  = Utilities.formatDate(fromDate, timeZone,"dd-MM-yyyy");
  var fromdategui = fromfomatteddate.substring(0,10);
  var tofomatteddate  = Utilities.formatDate(toDate, timeZone,"dd-MM-yyy");
  var todategui = tofomatteddate.substring(0,10);
  //Declare variabelen voor in de gui
  var subMenuclear = SpreadsheetApp.getUi().createMenu("Clear Calendar")
    .addItem("Clear Google Sheet from " +fromdategui + ' to ' + todategui , 'clearsheet')
    .addItem("Clear Google Calendar from " + fromdategui + ' to ' + todategui , 'clearCalendarActiviteiten')
    .addItem("Clear Google Sheet and Google Calendar from " + fromdategui + " to " + todategui , 'clearSheetenCalendar')

  //creeer totaal menu en zet de submenu's er in 
  var ui = SpreadsheetApp.getUi();
    ui.createMenu('Sync Google Calendar')
    //Reference submenu
    .addSubMenu(subMenusync)
    .addSeparator()
    .addSubMenu(subMenuclear)
    .addToUi()
}

function setGuiManage() {

    /// Nieuwe ui hieronder:
    /// Note: naam van Calendar update nu niet automatisch --> add onedit ?
var nameCalendar = SpreadsheetApp.getActive().getActiveSheet().getRange(rangeCalendarName).getValue();
var namesheet = SpreadsheetApp.getActive().getActiveSheet().getName();
    
    var subMenuCreate = SpreadsheetApp.getUi().createMenu("Create calendar")
    .addItem("Create calendar with name: " + nameCalendar, 'createTotalSheetAndCalendar') 

    var subMenuChangeName = SpreadsheetApp.getUi().createMenu("Change name of the calendar")
    .addItem("Change the name of the calendar into: " + nameCalendar, 'alterCalendarName')

    var subMenuDelete = SpreadsheetApp.getUi().createMenu("Delete")
    .addItem("Delete calendar '" + nameCalendar + "'" , 'deleteCalendar')
    .addItem("Delete sheet '" + namesheet + "'", 'deleteSheet')
    .addItem("Delete calendar '" + nameCalendar + "' and sheet '" + namesheet + "'", 'deleteSheetEnCalendar')

  var submenuEditor = SpreadsheetApp.getUi().createMenu("Add users to calendar")
    .addItem("Add editors to calendar '" + nameCalendar + "'", "addEditors")
    .addItem("Add viewers to calendar '" + nameCalendar + "'", "addViewers")
  var ui2 = SpreadsheetApp.getUi();
    ui2.createMenu('Manage Google Calendar')
    .addSubMenu(subMenuCreate)
    .addSubMenu(subMenuChangeName)
    .addSubMenu(submenuEditor)
    .addSeparator()
    .addSubMenu(subMenuDelete)
    .addToUi()
}

// function setGUIs(nameCalendar) {
  function setGUIs() {
setGuiManage()
setGUISync()

}

//// Zet hierin ook die begindatumachtigtorrie aii danku ! EN add by the onselectchange!



//// Trigger follow here under:

 function onOpen(e) {

   // set veriables for Calendar function

    // Set date DeleFrom and DeleteTo on the current date, such that those cell are not empty, thus the onOpenTrigger works properly

   // spreadsheetevents.getRange(cellDeleFrom).setValue(deleDay);
  //  spreadsheetevents.getRange(cellDeleTo).setValue(deleDay);
 // setGUI zodra de spreadsheet geopend wordt. LET OP; DE E TUSSEN HAAKJES MOET BLIJVEN STAAN !!!!
setGUIs()
 // functioneel ivm evt verkeerd omtunen van de kalendar naam 
createsheet()
// Understanding things are for the changing tabs onselectionchange:
  var prop = PropertiesService.getScriptProperties();
  var sheetName = e.range.getSheet().getSheetName();
  prop.setProperty("previousSheet", sheetName);
   }


function setDV(rowDV , collumnDV){

    var option = new Array();
      option[0]= "Definitive" ;
      option[1]="Concept";
      option[2]="Plan"

    var dv = SpreadsheetApp.newDataValidation();
    dv.setAllowInvalid(false);
    dv.setHelpText("Choose an event status/color");
    dv.requireValueInList(option, true);

SpreadsheetApp.getActiveSheet().getRange(rowDV,collumnDV).setDataValidation(dv);

}





function onEdit(e) {


   var range = e.range;
   var rangeA1 = range.getA1Notation();
   var rangevalue = range.getValue();
   var rangeCollumn = e.range.getColumn();
   var rangeRow = e.range.getRow();
     
  // CHanging Delete from:
  if(rangeA1 == "B2") {

    var fomatteddateFrom  = Utilities.formatDate(rangevalue, timeZone,"dd-MM-yyyy");
    var fromDatestring = fomatteddateFrom.substring(0,10);

    setGUIs
    setformatofDelete()
    SpreadsheetApp.getActiveSpreadsheet().toast("Changed the first date to delete into: "  + fromDatestring); 
  
}

//CHanging Delete To:
if(rangeA1 == "D2") {
 
    var fomatteddateTo  = Utilities.formatDate(rangevalue, timeZone,"dd-MM-yyyy");
    var toDatestring = fomatteddateTo.substring(0,10);

    setGUIs()
    setformatofDelete()
    SpreadsheetApp.getActiveSpreadsheet().toast("Changed the last date to delete into: "  + toDatestring );   
  }

// Changing Calendar name
if (rangeA1 == "B1"){
  createsheet()
  setGUIs()
  
}

if (rangeCollumn == 8) {

var rowtaget = rangeRow+1;
var collumntaget = rangeCollumn; 
setDV(rowtaget, collumntaget)
}
    }

/// On selection change werkt !!!! Vraag niet hoe, maar het werkt. VUl toe te voegen dingen toe onder IF statement (niet in !!!)

function onSelectionChange(e) {
  const prop = PropertiesService.getScriptProperties();
  const previousSheet = prop.getProperty("previousSheet");
  const range = e.range;
  const sheet = e.range.getSheet();
  const nameCalendar = sheet.getRange(1,2).getValue();
  const a1Notation = range.getA1Notation();
  const sheetName = range.getSheet().getSheetName();
  if (sheetName == previousSheet) {
    return;
    // When the tab is changed, this script is run.  
  } 
  prop.setProperty("previousSheet", sheetName);
  //setGUIs(nameCalendar)
  setGUIs()  
}