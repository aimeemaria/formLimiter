// This code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit to report our impact to funders
// For original source see http://www.edcode.org

function formLimiter_logTimeLimitSet() {
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Time%20Limit%20Set", scriptName, scriptTrackingId, systemName)
}


function formLimiter_logNumLimitSet() {
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Max%20Response%20Limit%20Set", scriptName, scriptTrackingId, systemName)
}


function formLimiter_logSSValueLimitSet() {
  NVSL.log("Cell%20Value%20Limit%20Set", scriptName, scriptTrackingId, systemName)
}

function formLimiter_logFormClosed() {
  NVSL.log("Form%20Closed", scriptName, scriptTrackingId, systemName)
}

//This function makes a call to the correct installation function.
//Embed this in the function that creates first actively loaded UI panel within the script
function setSid() { 
  var scriptNameLower = scriptName.toLowerCase();
  var sid = ScriptProperties.getProperty(scriptNameLower + "_sid");
  if (sid == null || sid == "")
  {
    var dt = new Date();
    var ms = dt.getTime();
    var ms_str = ms.toString();
    ScriptProperties.setProperty(scriptNameLower + "_sid", ms_str);
    var uid = UserProperties.getProperty(scriptNameLower + "_uid");
    if (uid) {
      logRepeatInstall();
    } else {
      logFirstInstall();
      UserProperties.setProperty(scriptNameLower + "_uid", ms_str);
    }      
  }
}

function logRepeatInstall() {
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Repeat%20Install", scriptName, scriptTrackingId, systemName)
}

function logFirstInstall() {
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("First%20Install", scriptName, scriptTrackingId, systemName)
}



function formLimiter_extractorWindow () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var propertyString = '';
  for (var key in properties) {
    if ((properties[key]!='')&&(key!="preconfigStatus")&&(key!="formLimiter_sid")&&(key!="ssId")) {
      var keyProperty = properties[key].replace(/[/\\*]/g, "\\\\");                                     
      propertyString += "   ScriptProperties.setProperty('" + key + "','" + keyProperty + "');\n";
    }
  }
  var app = UiApp.createApplication().setHeight(500).setTitle("Export preconfig() settings");
  var panel = app.createVerticalPanel().setWidth("100%").setHeight("100%");
  var labelText = "Copying a Google Spreadsheet copies scripts along with it, but without any of the script settings saved.  This normally makes it hard to share full, script-enabled Spreadsheet systems. ";
  labelText += " You can solve this problem by pasting the code below into a script function called \"formLimiter_preconfig\" (go to Script Editor) prior to publishing your Spreadsheet for others to copy. \n";
  labelText += " After a user copies your spreadsheet, they will select \"Run initial configuration.\"  This will preconfigure all needed script settings.  If you got this workflow from someone as a copy of a spreadsheet, this has probably already been done for you.";
  var label = app.createLabel(labelText);
  var window = app.createTextArea();
  var codeString = "//This section sets all script properties associated with this formLimiter profile \n";
  codeString +=  "var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();\n";
  codeString +=  "ScriptProperties.setProperty('ssId', ssId);\n";
  codeString += "var preconfigStatus = ScriptProperties.getProperty('preconfigStatus');\n";
  codeString += "if (preconfigStatus!='true') {\n";
  codeString += propertyString; 
  codeString += "};\n";
  codeString += "ScriptProperties.setProperty('preconfigStatus','true');\n";
  codeString += "var ss = SpreadsheetApp.getActiveSpreadsheet();\n";
  codeString += "if (ss.getSheetByName('Forms in same folder')) { \n";
  codeString += "  formLimiter_retrieveformurls(); \n";
  codeString += "} \n";
  codeString += "var properties = ScriptProperties.getProperties();\n";
  codeString += "if (properties.limiterType == 'date and time') {\n"
  codeString += "   formLimiter_deleteFormTrigger();\n";
  codeString += "   var date = new Date(e.parameter.date);\n";
  codeString += "   var amPm = e.parameter.amPm;\n";
  codeString += "   var hour = parseInt(e.parameter.hour);\n";
  codeString += "   var min = parseInt(e.parameter.min);\n";
  codeString += "   var year = date.getYear();\n";
  codeString += "   var month = date.getMonth();\n";
  codeString += "   var day = date.getDate();\n";
  codeString += "   var hour24 = hour;\n";
  codeString += " if ((amPm==\"PM\")&&(hour!=12)) {\n";
  codeString += "   hour24 = hour + 12;\n";
  codeString += " }\n";
  codeString += " if ((amPm==\"AM\")&&(hour==12)) {\n";
  codeString += "   hour24 = 0;\n";
  codeString += " }\n";
  codeString += "   var dateTime = new Date(year, month, day, hour24, min);\n";
  codeString += "   formLimiter_checkSetTimeTrigger(dateTime);\n";
  codeString += "} else if ((properties.limiterType == 'max number of form responses')||(properties.limiterType == 'spreadsheet cell value')) {\n";
  codeString += "   formLimiter_deleteTimeTrigger();\n";
  codeString += "   formLimiter_checkSetFormTrigger();\n";
  codeString += "}\n";
  codeString += "ss.toast('Custom formLimiter preconfiguration ran successfully. Please check formLimiter menu options to confirm system settings.');\n";
  window.setText(codeString).setWidth("100%").setHeight("350px");
  app.add(label);
  panel.add(window);
  app.add(panel);
  ss.show(app);
  return app;
}

