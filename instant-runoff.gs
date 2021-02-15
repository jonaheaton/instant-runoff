/*
Multi-part Instant-runoff voting with Google Form and Google Apps Script

Author: Jonah Eaton
Date Updated: 2021-02-15

Source Author: Chris Cartland
Source Date created: 2012-04-29
Edited by Jonah Eaton to handle a ballot with multiple elections, multi-part ballots and easier to use google form interface


Read usage instructions online
https://github.com/jonaheaton/multipart-instant-runoff


This project may contain bugs. Use at your own risk.


Steps to run an election.
* Go to Google Drive. Create a new Google Form.
* Create questions according to instructions on GitHub -- https://github.com/jonaheaton/multipart-instant-runoff
* From the form spreadsheet go to "Tools" -> "Script Editor..."
* Copy the code from instant-runoff.gs into the editor.
* Configure settings in the editor and match the settings with the names of your sheets.
* From the form spreadsheet go to "Instant Runoff" -> "Setup".
    * If this is not an option, run the function setup_instant_runoff() directly from the Script Editor.
* Create keys in the sheet named "Keys".
* Send out the live form for voting. If you are using keys, don't forget to distribute unique secret keys to voters.
* From the form spreadsheet go to "Instant Runoff" -> "Run".
    * If this is not an option, run the function run_instant_runoff() directly from the Script Editor.

*/


/* Settings */ 
var PositionList = ["Chair","Director of Operations","Treasurer","Department Liaison"];
var RankingNames = ["1st Choice","2nd","3rd","4th","5th","6th","7th","8th"];
var ColLetters = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O"];
var MAX_NUM_VOTES = 25;

var BASE_ROW = 2;
//var BASE_COLUMN = 6;
var MAX_ROW = MAX_NUM_VOTES;

var USING_KEYS = false;
var VOTE_SHEET_KEYS_COLUMN = 2;
var KEYS_SHEET_NAME = "Keys";
var USED_KEYS_SHEET_NAME = "Used Keys";

var CHANGE_COLORS = true;
var RUN_RUNOFF = true;
/* End Settings */


/* Global variables */

var NUM_COLUMNS;

/* End global variables */

function setup_instant_runoff() {
  create_menu_items();
}

function create_menu_items() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [ {name: "Setup", functionName: "setup_instant_runoff"},
                        {name: "Run", functionName: "run_all_elections"} ];
    ss.addMenu("Instant Runoff", menuEntries);
}

/* Create menus */
function onOpen() {
    setup_instant_runoff();
}

/* Create menus when installed */
function onInstall() {
    onOpen();
}

/*  generate the election results message */
function get_election_message(VOTE_SHEET_NAME) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(VOTE_SHEET_NAME) == null) {
    var full_message = VOTE_SHEET_NAME + " not found"}
  else{
  var myout = run_instant_runoff(VOTE_SHEET_NAME,"");
var winner1 = myout.winner;
var outcome = myout.outcome;
    if (RUN_RUNOFF){
    if (outcome>0){
    var myout1 = run_instant_runoff(VOTE_SHEET_NAME,winner1);
    var runnerup1 = myout1.winner;}
  else{  var runnerup1 = "N/A"
      }
  var full_message = "Winner: " + winner1 + "\\n Runnerup: " + runnerup1;
    } else {
      var full_message = "Winner: " + winner1 ;
    }
  }
  return full_message
}

/* run the election */
function run_all_elections() {
   var full_message = get_election_message(PositionList[0]);
   for(let i = 1; i < PositionList.length; i++){

  full_message = full_message + "\\n" + get_election_message(PositionList[i]);
   }
  full_message = full_message + ".\\n Date and time: " +  Utilities.formatDate(new Date(), "PST", "yyyy-MM-dd HH:mm:ss");
Browser.msgBox(full_message);
}



/* format the google sheet for an election() */
function set_up_sheet() {
  for(let i = 0; i < PositionList.length; i++){
   new_sheet_for_position(PositionList[i]); 
   format_sheet_for_election(PositionList[i]);
   lock_sheet_for_position(PositionList[i])
  }
}

/* create a sheet for a specific Position */
function new_sheet_for_position(PositionName) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var itt = activeSpreadsheet.getSheetByName(PositionName);
  if (!itt){
var ssNew = activeSpreadsheet.insertSheet(PositionName);
ssNew.appendRow(['=FILTER(Responses!B1:AC103,IF(REGEXMATCH(Responses!B1:AC1, "'+ PositionName +'"), 1, 0))']);
}
}

/* count the number of running for a specific PositionName */
function how_many_candidates(PositionName) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = activeSpreadsheet.getSheetByName(PositionName);
  var numCol = Sheet.getLastColumn();
  var numCandidates = 0;
  var data = Sheet.getSheetValues(1,1,1,numCol);
  for(var i = 0; i<numCol;i++){
    if(data[0][i].includes(PositionName)){
      numCandidates = numCandidates+1;
    }
  }
return numCandidates
// Logger.log(numCandidates)
}

/* format the sheet for counting for a specific Position */
function format_sheet_for_election(PositionName) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var numCandidates = how_many_candidates(PositionName);
  var ss = activeSpreadsheet.getSheetByName(PositionName);
  // var numCandidates = ss.getLastColumn();
  var lastLet = ColLetters[numCandidates-1];
  for(let i = 0; i < numCandidates; i++){
    var cell = ss.getRange(1,i+numCandidates+1)
    cell.setValue(RankingNames[i]);
    var cell = ss.getRange(2,i+numCandidates+1)
    var myLet = ColLetters[i+numCandidates];
    // ss.getRange(2,i+numCandidates+1,20,1).setValues(outerArray);
    cell.setValue("=INDEX($A$1:$"+ lastLet + "$1,MIN(IF($A2:$"+ lastLet +"2="+ myLet +"$1,COLUMN($A:$"+ lastLet +"),9999)))")
    var destination = ss.getRange(2,i+numCandidates+1,MAX_NUM_VOTES,1);
    cell.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  }
}


/* lock the sheet associated with specific Position */
function lock_sheet_for_position(PositionName){
  // Protect the active sheet, then remove all other users from the list of editors.
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = activeSpreadsheet.getSheetByName(PositionName);
  var protection = ss.protect().setDescription('Sample protected sheet');

  // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
  // permission comes from a group, the script throws an exception upon removing the group.
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
    }
}


function run_instant_runoff(VOTE_SHEET_NAME,bad_candidate) {
  /* Determine number of voting columns */
  var active_spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var row1_range = active_spreadsheet.getSheetByName(VOTE_SHEET_NAME).getRange("A1:1");
  var BASE_COLUMN = how_many_candidates(VOTE_SHEET_NAME)+1;
  NUM_COLUMNS = get_num_columns_with_values(row1_range) - BASE_COLUMN + 1;


  /* Reset state */
  missing_keys_used_sheet_alert = false;
  
  /* Begin */
  clear_background_color(VOTE_SHEET_NAME);
  
  var results_range = get_range_with_values(VOTE_SHEET_NAME, BASE_ROW, BASE_COLUMN, NUM_COLUMNS);
  
  if (results_range == null) {
    Browser.msgBox("No votes. Looking for sheet: " + VOTE_SHEET_NAME);
    return;
  }
  // Keys are used to prevent voters from voting twice.
  // Keys also allow voters to change their vote.
  // If keys_range == null then we are not using keys. 
  var keys_range = null;
  
  // List of valid keys
  var valid_keys;
  
  if (USING_KEYS) {
    keys_range = get_range_with_values(VOTE_SHEET_NAME, BASE_ROW, VOTE_SHEET_KEYS_COLUMN, 1);
    if (keys_range == null) {
      Browser.msgBox("Using keys and could not find column with submitted keys. " + 
                     "Looking in column " + VOTE_SHEET_KEYS_COLUMN + 
                     " in sheet: " + VOTE_SHEET_NAME);
      return;
    }
    var valid_keys_range = get_range_with_values(KEYS_SHEET_NAME, BASE_ROW, 1, 1);
    if (valid_keys_range == null) {
      var results_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KEYS_SHEET_NAME);
      if (results_sheet == null) {
        Browser.msgBox("Looking for list of valid keys. Cannot find sheet: " + KEYS_SHEET_NAME);
      } else {
        Browser.msgBox("List of valid keys cannot be found in sheet: " + KEYS_SHEET_NAME);
      }
      return;
    }
    valid_keys = range_to_array(valid_keys_range);
  }
  
  /* candidates is a list of names (strings) */
  var candidates = get_all_candidates(results_range);
  //var badcandidate = candidates[5];
  remove_candidate(candidates,"#NUM!");
  remove_candidate(candidates,bad_candidate);
  
  /* votes is an object mapping candidate names -> number of votes */
  var votes = get_votes(results_range, candidates, keys_range, valid_keys);
  
  /* winner is candidate name (string) or null */
  var winner = get_winner(votes, candidates);

  while (winner == null) {
    /* Modify candidates to only include remaining candidates */
    var prev_candidates =  copy_candidates(candidates);
    get_remaining_candidates(votes, candidates);
    //Browser.msgBox(prev_candidates);
    if (candidates.length == 0) {
      if (missing_keys_used_sheet_alert) {
        Browser.msgBox("Unable to record keys used. Looking for sheet: " + USED_KEYS_SHEET_NAME);    
      }
      //Browser.msgBox("Tie");
      //return;
      return  {
        winner: prev_candidates,
        outcome: 0,
    };
    }
    votes = get_votes(results_range, candidates, keys_range, valid_keys);
    winner = get_winner(votes, candidates);
  }
  
  if (USING_KEYS) {
    if (missing_keys_used_sheet_alert) {
      Browser.msgBox("Unable to record keys used. Looking for sheet: " + USED_KEYS_SHEET_NAME);    
    }
    var used_keys_range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USED_KEYS_SHEET_NAME).getRange(1,2,1,1);
    used_keys_range.setValue(winner_message);
  }
  var winner_message = "Winner: " + winner + ".\nDate and time: " +  Utilities.formatDate(new Date(), "PST", "yyyy-MM-dd HH:mm:ss");
        return  {
        winner: winner,
        outcome: 1,
    };
  /*return winner_message;
  Browser.msgBox(winner_message);*/
}


function get_range_with_values(sheet_string, base_row, base_column, num_columns) {
  var results_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_string);
  if (results_sheet == null) {
    return null;
  }
  var a1string = String.fromCharCode(65 + base_column - 1) +
      base_row + ':' + 
      String.fromCharCode(65 + base_column + num_columns - 2);
  var results_range = results_sheet.getRange(a1string);
  // results_range contains the whole columns all the way to
  // the bottom of the spreadsheet. We only want the rows
  // with votes in them, so we're going to count how many
  // there are and then just return those.
  var num_rows = get_num_rows_with_values(results_range);
  if (num_rows == 0) {
    return null;
  }
  results_range = results_sheet.getRange(base_row, base_column, num_rows, num_columns);
  return results_range;
}


function range_to_array(results_range) {
        if (CHANGE_COLORS){
        results_range.setBackground("#eeeeee");
      }
  
  var candidates = [];
  var num_rows = results_range.getNumRows();
  var num_columns = results_range.getNumColumns();
  for (var row = num_rows; row >= 1; row--) {
    var first_is_blank = results_range.getCell(row, 1).isBlank();
    if (first_is_blank) {
      continue;
    }
    for (var column = 1; column <= num_columns; column++) {
      var cell = results_range.getCell(row, column);
      if (cell.isBlank()) {
        break;
      }
      var cell_value = cell.getValue();
      if (CHANGE_COLORS){
        cell.setBackground("#ffff00");
      }
      if (!include(candidates, cell_value)) {
        candidates.push(cell_value);
      }
    }
  }
  return candidates;
}

function copy_candidates(candidates){
var prev_candidates = [];
for (var c = 0; c < candidates.length; c++) {
    var name = candidates[c];
   prev_candidates.push(name);
}
  return prev_candidates
}

function get_all_candidates(results_range) {
  if (CHANGE_COLORS){
    results_range.setBackground("#eeeeee");
  }
  
  var candidates = [];
  var num_rows = results_range.getNumRows();
  var num_columns = results_range.getNumColumns();
  for (var row = num_rows; row >= 1; row--) {
    var first_is_blank = results_range.getCell(row, 1).isBlank();
    if (first_is_blank) {
      continue;
    }
    for (var column = 1; column <= num_columns; column++) {
      var cell = results_range.getCell(row, column);
      if (cell.isBlank()) {
        break;
      }
      var cell_value = cell.getValue();
      if (CHANGE_COLORS){
        cell.setBackground("#ffff00");
      }
      if (!include(candidates, cell_value)) {
        candidates.push(cell_value);
      }
    }
  }
  return candidates;
}


function get_votes(results_range, candidates, keys_range, valid_keys) {
  if (typeof keys_range === "undefined") {
    keys_range = null;
  }
  var votes = {};
  var keys_used = [];
  
  for (var c = 0; c < candidates.length; c++) {
    votes[candidates[c]] = 0;
  }
  
  var num_rows = results_range.getNumRows();
  var num_columns = results_range.getNumColumns();
  for (var row = num_rows; row >= 1; row--) {
    var first_is_blank = results_range.getCell(row, 1).isBlank();
    if (first_is_blank) {
      break;
    }
    
    if (keys_range != null) {
      // Only use key once.
      var key_cell = keys_range.getCell(row, 1);
      var key_cell_value = key_cell.getValue();
      if (!include(valid_keys, key_cell_value) ||
          include(keys_used, key_cell_value)) {
        if (CHANGE_COLORS){
          key_cell.setBackground('#ffaaaa');}
        continue;
      } else {
        if (CHANGE_COLORS){
          key_cell.setBackground('#aaffaa');}
        keys_used.push(key_cell_value);
      }
    }
    
    for (var column = 1; column <= num_columns; column++) {
      var cell = results_range.getCell(row, column);
      if (cell.isBlank()) {
        break;
      }
      
      var cell_value = cell.getValue();
      if (include(candidates, cell_value)) {
        votes[cell_value] += 1;
        if (CHANGE_COLORS){
          cell.setBackground("#aaffaa");}
        break;
      }
      if (CHANGE_COLORS){
        cell.setBackground("#aaaaaa");}
    }
  }
  if (keys_range != null) {
    update_keys_used(keys_used);
  }
  return votes;
}


function update_keys_used(keys_used) {
  var keys_used_range = get_range_with_values(USED_KEYS_SHEET_NAME, BASE_ROW, 1, 1);
  if (keys_used_range != null) {
    if (CHANGE_COLORS){
      keys_used_range.setBackground('#ffffff');}
    if (keys_used_range != null) {
      var num_rows = keys_used_range.getNumRows();
      for (var row = num_rows; row >= 1; row--) {
        keys_used_range.getCell(row, 1).setValue('');
      }
    }
  }
  
  var keys_used_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USED_KEYS_SHEET_NAME);
  if (keys_used_sheet == null) {
    missing_keys_used_sheet_alert = true;
    return;
  }
  var a1string = String.fromCharCode(65 + 1 - 1) +
    BASE_ROW + ':' + 
    String.fromCharCode(65 + 1 + 1 - 2);
  var keys_used_range = keys_used_sheet.getRange(a1string);
  for (var k = 0; k < keys_used.length; k++) {
    var cell = keys_used_range.getCell(k+1, 1);
    cell.setValue(keys_used[k]);
    if (CHANGE_COLORS){
      cell.setBackground('#eeeeee');}
  }
}




function get_winner(votes, candidates) {
  var total = 0;
  var winning = null;
  var max = 0;
  for (var c = 0; c < candidates.length; c++) {
    //Browser.msgBox("length " + c)
    var name = candidates[c];
    var count = votes[name];
    total += count;
    if (count > max) {
      winning = name;
      max = count;
    }
  }
  
  if (max * 2 > total) {
    return winning;
  }
  return null;
}


function get_remaining_candidates(votes, candidates) {
  var min = -1;
  for (var c = 0; c < candidates.length; c++) {
    var name = candidates[c];
    var count = votes[name];
    if (count < min || min == -1) {
      min = count;
    }
  }
  
  var c = 0;
  while (c < candidates.length) {
    var name = candidates[c];
    var count = votes[name];
    if (count == min) {
      candidates.splice(c, 1);
    } else {
      c++;
    }
  }
  return candidates;
}

function remove_candidate(candidates,badcandidate) {
   var c = 0;
  while (c < candidates.length) {
    var name = candidates[c];
    if (name == badcandidate) {
      candidates.splice(c, 1);
    } else {
      c++;
    }
  }
  return candidates
}

/*
http://stackoverflow.com/questions/143847/best-way-to-find-an-item-in-a-javascript-array
*/
function include(arr,obj) {
    return (arr.indexOf(obj) != -1);
}


/*
Returns the number of consecutive rows that do not have blank values in the first column.
http://stackoverflow.com/questions/4169914/selecting-the-last-value-of-a-column
*/
function get_num_rows_with_values(results_range) {
  var num_rows_with_votes = 0;
  var num_rows = results_range.getNumRows();
  for (var row = 1; row <= num_rows; row++) {
    var first_is_blank = results_range.getCell(row, 1).isBlank();
    if (first_is_blank) {
      break;
    }
    num_rows_with_votes += 1;
  }
  num_rows_with_votes = Math.min(num_rows_with_votes,MAX_ROW);
  //Browser.msgBox(num_rows_with_votes)
  return num_rows_with_votes;
}


/*
Returns the number of consecutive columns that do not have blank values in the first row.
http://stackoverflow.com/questions/4169914/selecting-the-last-value-of-a-column
*/
function get_num_columns_with_values(results_range) {
  var num_columns_with_values = 0;
  var num_columns = results_range.getNumColumns();
  for (var col = 1; col <= num_columns; col++) {
    var first_is_blank = results_range.getCell(1, col).isBlank();
    if (first_is_blank) {
      break;
    }
    num_columns_with_values += 1;
  }
  return num_columns_with_values;
}


function clear_background_color(VOTE_SHEET_NAME) {
  var BASE_COLUMN = how_many_candidates(VOTE_SHEET_NAME)+1;
  var results_range = get_range_with_values(VOTE_SHEET_NAME, BASE_ROW, BASE_COLUMN, NUM_COLUMNS);
  if (results_range == null) {
    return;
  }
  if (CHANGE_COLORS){
    results_range.setBackground('#eeeeee');
  }
  
  if (USING_KEYS) {
    var keys_range = get_range_with_values(VOTE_SHEET_NAME, BASE_ROW, VOTE_SHEET_KEYS_COLUMN, 1);
    if (CHANGE_COLORS){
      //keys_range.setBackground('#eeeeee');
    }
  }
}

