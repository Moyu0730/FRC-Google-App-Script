var m_Sheet = SpreadsheetApp.getActiveSpreadsheet();
var m_firstSheet = m_Sheet.getSheets()[0];

var m_A1Data = m_firstSheet.getRange("A1").getValue();
Logger.log(m_A1Data);