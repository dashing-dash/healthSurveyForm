function myFunction() {
  setFormColumnHeaders();
}

function setFormColumnHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var headers = [
    "Timestamp", "Participant ID", "Age", "Gender", "Residential Status", "Marital Status", 
    "Type of Diet", "Mixed Diet Frequency", "Junk Food Frequency", "Skip Breakfast", 
    "Fruit Consumption", "Vegetable Consumption", "Sugary Drinks", "Supplements Intake", 
    "Supplement Details", "Custom Supplement Text", "Weight Medications", "Weight Medication Details", 
    "Physical Activity Frequency", "Physical Activity Types", "Custom Activity Text", "Session Duration", 
    "Total Screen Time", "Educational Screen Time", "Recreational Screen Time", "Screen Time Before Sleep", 
    "Sleep Duration", "Feel Refreshed", "PSS1 Uncontrolled Life", "PSS2 Problem Confidence", 
    "PSS3 Things Going Your Way", "PSS4 Difficulties Piling Up", "Tobacco Use Profile", "Tobacco Mode", 
    "Tobacco Frequency", "Tobacco Quantity", "Alcohol Use Profile", "Alcohol Type", "Alcohol Amount", 
    "Alcohol Sessions", "Other Substances Status", "Substance Selection Details", "Custom Substance Text", 
    "Chronic Illness Status", "Chronic Illness Details", "Chronic Treatment Timeline", "Family Disease History", 
    "Custom Family Disease Text", "Premature Family Death", "Family Income", "Unexpected Weight Gain", 
    "Unintentional Weight Loss", "Proactive Health Checkup", "Menarche Age", "Cycle Regularity", 
    "Cycle Length", "Delayed Periods Experience", "Menstrual Bleeding Flow", "Body Hair Changes", 
    "Acne Severity Increase", "PCOS Diagnosis Status", "Family History of PCOS", "Maternity Profile", 
    "Youngest Child Age", "History of Abortion"
  ];
  
  // Clear any existing contents on the first row and drop the clean matrix headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e8f0fe");
  
  Logger.log("Headers initialized cleanly for 63 tracking coordinates!");
}