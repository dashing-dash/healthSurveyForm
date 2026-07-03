function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Health Survey Questionnaire')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getNewToken() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // Clean alphanumeric string
  var token = '';
  for (var i = 0; i < 6; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

function processFormData(formData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Explicitly mapping out every incoming variable to match its column position sequentially
    var rowData = [
      new Date(), // Column A: Timestamp
      formData.participantId,
      formData.age,
      formData.gender,
      formData.residing,
      formData.maritalStatus,
      formData.dietType,
      formData.mixedDietFreq || 'N/A',
      formData.junkFoodFreq,
      formData.skipBreakfast,
      formData.fruitConsumption,
      formData.vegConsumption,
      formData.sugaryDrinks,
      formData.supplements,
      formData.supplementDetails || 'N/A',
      formData.customSupplement || 'N/A',
      formData.weightMeds,
      formData.weightMedDetails || 'N/A',
      formData.physicalActivityFreq,
      formData.activityTypes || 'N/A',
      formData.customActivity || 'N/A',
      formData.sessionDuration,
      formData.totalScreenTime,
      formData.eduScreenTime,
      formData.recScreenTime,
      formData.sleepScreenTime,
      formData.sleepDuration,
      formData.feelRefreshed,
      formData.pss1, formData.pss2, formData.pss3, formData.pss4,
      formData.tobaccoUse,
      formData.tobaccoMode || 'N/A', formData.tobaccoFreq || 'N/A', formData.tobaccoQty || 'N/A',
      formData.alcoholUse,
      formData.alcoholType || 'N/A', formData.alcoholQty || 'N/A', formData.alcoholSessions || 'N/A',
      formData.otherSubstances || 'N/A',
      formData.customSubstance || 'N/A',
      formData.chronicIllness,
      formData.chronicIllnessDetails || 'N/A',
      formData.chronicTreatment || 'N/A',
      formData.familyHistory || 'N/A',
      formData.customFamilyHistory || 'N/A',
      formData.familyDeath,
      formData.familyIncome,
      formData.weightGain,
      formData.weightLoss,
      formData.proactiveCheckup,
      // Section 7: Female Assessment Columns (Will hold 'N/A' defaults for non-females)
      formData.periodAge || 'N/A',
      formData.cycleRegularity || 'N/A',
      formData.cycleLength || 'N/A',
      formData.delayedPeriods || 'N/A',
      formData.menstrualFlow || 'N/A',
      formData.hairChange || 'N/A',
      formData.acneIncrease || 'N/A',
      formData.pcosStatus || 'N/A',
      formData.familyPcos || 'N/A',
      formData.hasChildren || 'N/A',
      formData.youngestChildAge || 'N/A',
      formData.hadAbortion || 'N/A'
    ];
    
    sheet.appendRow(rowData);
    return "SUCCESS";
  } catch(err) {
    return "ERROR: " + err.toString();
  }
}