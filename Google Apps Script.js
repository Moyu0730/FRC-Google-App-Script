var m_Sheet = SpreadsheetApp.getActiveSpreadsheet();
var m_BaseSheet = m_Sheet.getSheets()[0];
var m_IsUpdatedButton = m_BaseSheet.getRange("C15");
var m_IsClearedButton = m_BaseSheet.getRange("E15");
var m_StatusCell = m_BaseSheet.getRange("E16");
var m_IsUpdated = m_IsUpdatedButton.getValue();
var m_IsCleared = m_IsClearedButton.getValue();
var m_flag = 0;

function SubmitButtomChecked(){
  if( m_IsUpdated == true ){
    var m_UpdatedSheetName = m_BaseSheet.getRange("D14").getValue();

    if( m_UpdatedSheetName == "Update WorkSheet" ){
      m_StatusCell.setBackground('#FF7575');
      m_StatusCell.setValue('Submit To Wrong WorkSheet');
      m_IsUpdatedButton.setValue(false);
      m_flag++;
    }else{
      UpdateSpecifyWorkSheetData(m_UpdatedSheetName);
      m_StatusCell.setBackground('#77FF00');
      m_StatusCell.setValue('Submit Successful');
      m_IsUpdatedButton.setValue(false);
      m_flag++;
    }

    
  }else if( m_flag % 2 == 0 ){
    m_flag++;
    m_StatusCell.setBackground('#DDDDDD');
    m_StatusCell.setValue("There Isn't Any Request");
  }
}

function ClearButtomChecked(){
  if( m_IsCleared == true ){
    ClearUpdateWorkSheet();
    m_StatusCell.setBackground('#77FF00');
    m_StatusCell.setValue('Clear Successful');
    m_IsClearedButton.setValue(false);
    m_flag++;
  }else if( m_flag % 2 == 0 ){
    m_flag++; 
    m_StatusCell.setBackground('#DDDDDD');
    m_StatusCell.setValue("There Isn't Any Request");
  }
}

function ClearUpdateWorkSheet(){
  var m_GridFalse 
    = [
        [false, false, false], // 0
        [false, false, false], // 1
        [false, false, false], // 2
        [false, false, false], // 3
        [false, false, false], // 4
        [false, false, false], // 5
        [false, false, false], // 6
        [false, false, false], // 7
        [false, false, false], // 8
      ];
  
  var m_DefFalse 
    = [ 
        [false], // 0
        [false], // 1
        [false], // 2
        [false], // 3
        [false], // 4
        [false], // 5
        [false], // 6
        [false], // 7
        [false], // 8
      ];

  var m_AtkFalse 
    = [ 
        [false], // 0
        [false], // 1
        [false], // 2
        [false], // 3
        [false], // 4
        [false], // 5
      ];

  var m_MoveFalse 
    = [
        [false], // 0
        [false], // 1
        [false], // 2
      ];

  m_BaseSheet.getRangeList(["A4", "A9", "E4", "E9", "N3", "N7", "N11"]).setValue(false);
  m_BaseSheet.getRange("B3:D11").setValues(m_GridFalse);
  m_BaseSheet.getRange("F3:H11").setValues(m_GridFalse);
  m_BaseSheet.getRange("J3:J11").setValues(m_DefFalse);
  m_BaseSheet.getRange("M3:M8").setValues(m_AtkFalse);
  m_BaseSheet.getRange("M10:M12").setValues(m_MoveFalse);
  m_BaseSheet.getRange("O7").setValue(0);
}

function ClearSpecifyWorkSheetData( m_UpdatedSheetName ) {
  var m_ctrSheet = m_Sheet.getSheetByName(m_UpdatedSheetName);
  // var m_ctrSheet = m_Sheet.getSheets()[1];
  var m_DataUpdateFrequency = m_ctrSheet.getRange("D1").getValue();
  var m_UpdatedAutoGridData = m_ctrSheet.getRange("C15:E23").getValues();
  var m_UpdatedTeleopGridData = m_ctrSheet.getRange("H15:J23").getValues();
  var m_BaseAutoGridData = m_BaseSheet.getRange("B3:D11").getValues();
  var m_BaseTeleopGridData = m_BaseSheet.getRange("F3:H11").getValues();

  for( var i = 0 ; i < 3 ; i++ ){
    for( var j = 0 ; j < 9 ; j++ ){
      var m_TeleopCellValue = Math.ceil(m_DataUpdateFrequency * m_UpdatedTeleopGridData[j][i]);
      var m_newTeleopCellValue = ( m_TeleopCellValue + ( m_BaseTeleopGridData[j][i] == true ? 1 : 0 ) ) / (m_DataUpdateFrequency+1);
      m_UpdatedTeleopGridData[j][i] = m_newTeleopCellValue;

      var m_AutoCellValue = Math.ceil(m_DataUpdateFrequency * m_UpdatedAutoGridData[j][i]);
      var m_newAutoCellValue = ( m_AutoCellValue + ( m_BaseAutoGridData[j][i] == true ? 1 : 0 ) ) / (m_DataUpdateFrequency+1);
      m_UpdatedAutoGridData[j][i] = m_newAutoCellValue;
    }
  }

  m_ctrSheet.getRange("H15:J23").setValue(m_UpdatedTeleopGridData);
  m_ctrSheet.getRange("C15:E23").setValue(m_UpdatedAutoGridData);
  m_ctrSheet.getRange("D1").setValue(m_DataUpdateFrequency+1);
}

function UpdateSpecifyWorkSheetData( m_UpdatedSheetName ) {
  var m_ctrSheet = m_Sheet.getSheetByName(m_UpdatedSheetName);
  // var m_ctrSheet = m_Sheet.getSheets()[1];
  var m_DataUpdateFrequency = m_ctrSheet.getRange("D1").getValue();
  
  // Grid
  var m_UpdatedAutoGridData = m_ctrSheet.getRange("C15:E23").getValues();
  var m_UpdatedTeleopGridData = m_ctrSheet.getRange("H15:J23").getValues();
  var m_BaseAutoGridData = m_BaseSheet.getRange("B3:D11").getValues();
  var m_BaseTeleopGridData = m_BaseSheet.getRange("F3:H11").getValues();

  for( var i = 0 ; i < 3 ; i++ ){
    for( var j = 0 ; j < 9 ; j++ ){
      m_UpdatedTeleopGridData[j][i] = getPercentData(m_DataUpdateFrequency, m_UpdatedTeleopGridData[j][i], m_BaseTeleopGridData[j][i]);
      m_UpdatedAutoGridData[j][i] = getPercentData(m_DataUpdateFrequency, m_UpdatedAutoGridData[j][i], m_BaseAutoGridData[j][i]);
    }
  }

  // Charge Station
    // Auto
    var m_UpdatedAutoEngageData = m_ctrSheet.getRange("B16").getValue();
    var m_UpdatedAutoDockData = m_ctrSheet.getRange("B21").getValue();
    var m_BaseAutoEngageData = m_BaseSheet.getRange("A4").getValue();
    var m_BaseAutoDockData = m_BaseSheet.getRange("A9").getValue();

    // Teleop
    var m_UpdatedTeleopEngageData = m_ctrSheet.getRange("G16").getValue();
    var m_UpdatedTeleopDockData = m_ctrSheet.getRange("G21").getValue();
    var m_BaseTeleopEngageData = m_BaseSheet.getRange("E4").getValue();
    var m_BaseTeleopDockData = m_BaseSheet.getRange("E9").getValue();

    // Auto Engage
    m_UpdatedAutoEngageData = getPercentData(m_DataUpdateFrequency, m_UpdatedAutoEngageData, m_BaseAutoEngageData);

    // Auto Dock
    m_UpdatedAutoDockData = getPercentData(m_DataUpdateFrequency, m_UpdatedAutoDockData, m_BaseAutoDockData);

    // Teleop Engage
    m_UpdatedTeleopEngageData = getPercentData(m_DataUpdateFrequency, m_UpdatedTeleopEngageData, m_BaseTeleopEngageData);

    // Teleop Dock
    m_UpdatedTeleopDockData = getPercentData(m_DataUpdateFrequency, m_UpdatedTeleopDockData, m_BaseTeleopDockData);
  
  // 3-Dim
    // Defence
    var m_UpdatedDefenceA = m_ctrSheet.getRange("G3").getValue();
    var m_UpdatedDefenceB = m_ctrSheet.getRange("G5").getValue();
    var m_UpdatedDefenceC = m_ctrSheet.getRange("G7").getValue();
    var m_UpdatedDefenceD = m_ctrSheet.getRange("G9").getValue();
    var m_UpdatedDefenceE = m_ctrSheet.getRange("G11").getValue();
    var m_BaseDefenceA = m_BaseSheet.getRange("J3").getValue();
    var m_BaseDefenceB = m_BaseSheet.getRange("J5").getValue();
    var m_BaseDefenceC = m_BaseSheet.getRange("J7").getValue();
    var m_BaseDefenceD = m_BaseSheet.getRange("J9").getValue();
    var m_BaseDefenceE = m_BaseSheet.getRange("J11").getValue();

    // Attack
    var m_UpdatedAttackA = m_ctrSheet.getRange("J3").getValue();
    var m_UpdatedAttackB = m_ctrSheet.getRange("J4").getValue();
    var m_UpdatedAttackC = m_ctrSheet.getRange("J5").getValue();
    var m_UpdatedAttackD = m_ctrSheet.getRange("J6").getValue();
    var m_UpdatedAttackE = m_ctrSheet.getRange("J7").getValue();
    var m_UpdatedAttackF = m_ctrSheet.getRange("J8").getValue();
    var m_BaseAttackA = m_BaseSheet.getRange("M3").getValue();
    var m_BaseAttackB = m_BaseSheet.getRange("M4").getValue();
    var m_BaseAttackC = m_BaseSheet.getRange("M5").getValue();
    var m_BaseAttackD = m_BaseSheet.getRange("M6").getValue();
    var m_BaseAttackE = m_BaseSheet.getRange("M7").getValue();
    var m_BaseAttackF = m_BaseSheet.getRange("M8").getValue();

    // Transport
    var m_UpdatedTransportA = m_ctrSheet.getRange("J10").getValue();
    var m_UpdatedTransportB = m_ctrSheet.getRange("J11").getValue();
    var m_UpdatedTransportC = m_ctrSheet.getRange("J12").getValue();
    var m_BaseTransportA = m_BaseSheet.getRange("M10").getValue();
    var m_BaseTransportB = m_BaseSheet.getRange("M11").getValue();
    var m_BaseTransportC = m_BaseSheet.getRange("M12").getValue();

  // Mobility & Park & Flexibility
  m_BaseMobilityData = m_BaseSheet.getRange("N3").getValue();
  m_UpdatedMobilityData = m_ctrSheet.getRange("F15").getValue();
  m_BaseParkData = m_BaseSheet.getRange("N7").getValue();
  m_UpdatedParkData = m_ctrSheet.getRange("F19").getValue();
  m_BaseFlexibilityData = m_BaseSheet.getRange("N11").getValue();
  m_UpdatedFlexibilityData = m_ctrSheet.getRange("F23").getValue();

    // Mobility
    m_UpdatedMobilityData = getPercentData(m_DataUpdateFrequency, m_UpdatedMobilityData, m_BaseMobilityData);

    // Park       
    m_UpdatedParkData = getPercentData(m_DataUpdateFrequency, m_UpdatedParkData, m_BaseParkData);

    // Flexibility     
    m_UpdatedFlexibilityData = getPercentData(m_DataUpdateFrequency, m_UpdatedFlexibilityData, m_BaseFlexibilityData);
  
  // Penalty
  var m_UpdatedTeamPenalty = m_ctrSheet.getRange("D2").getValue();
  var m_BaseTeamPenalty = m_BaseSheet.getRange("O7").getValue();
  var m_newTeamPenalty = ( m_BaseTeamPenalty + ( m_UpdatedTeamPenalty * m_DataUpdateFrequency ) ) / (m_DataUpdateFrequency+1);

  // Penalty
  m_ctrSheet.getRange("D2").setValue(m_newTeamPenalty);

  // Mobility & Park & Flexibility
  m_ctrSheet.getRange("F15").setValue(m_UpdatedMobilityData);
  m_ctrSheet.getRange("F19").setValue(m_UpdatedParkData);
  m_ctrSheet.getRange("F23").setValue(m_UpdatedFlexibilityData);

  // 3-Dim
    // Defence
    m_ctrSheet.getRange("G3").setValue( m_UpdatedDefenceA + ( m_BaseDefenceA == true ? 1 : 0 ) ); // A
    m_ctrSheet.getRange("G5").setValue( m_UpdatedDefenceB + ( m_BaseDefenceB == true ? 1 : 0 ) ); // B
    m_ctrSheet.getRange("G7").setValue( m_UpdatedDefenceC + ( m_BaseDefenceC == true ? 1 : 0 ) ); // C
    m_ctrSheet.getRange("G9").setValue( m_UpdatedDefenceD + ( m_BaseDefenceD == true ? 1 : 0 ) ); // D
    m_ctrSheet.getRange("G11").setValue( m_UpdatedDefenceE + ( m_BaseDefenceE == true ? 1 : 0 ) ); // E

    // Attack
    m_ctrSheet.getRange("J3").setValue( m_UpdatedAttackA + ( m_BaseAttackA == true ? 1 : 0 ) ); // A
    m_ctrSheet.getRange("J4").setValue( m_UpdatedAttackB + ( m_BaseAttackB == true ? 1 : 0 ) ); // B
    m_ctrSheet.getRange("J5").setValue( m_UpdatedAttackC + ( m_BaseAttackC == true ? 1 : 0 ) ); // C
    m_ctrSheet.getRange("J6").setValue( m_UpdatedAttackD + ( m_BaseAttackD == true ? 1 : 0 ) ); // D
    m_ctrSheet.getRange("J7").setValue( m_UpdatedAttackE + ( m_BaseAttackE == true ? 1 : 0 ) ); // E
    m_ctrSheet.getRange("J8").setValue( m_UpdatedAttackF + ( m_BaseAttackF == true ? 1 : 0 ) ); // F

    // Transport
    m_ctrSheet.getRange("J10").setValue( m_UpdatedTransportA + ( m_BaseTransportA == true ? 1 : 0 ) ); // A
    m_ctrSheet.getRange("J11").setValue( m_UpdatedTransportB + ( m_BaseTransportB == true ? 1 : 0 ) ); // B
    m_ctrSheet.getRange("J12").setValue( m_UpdatedTransportC + ( m_BaseTransportC == true ? 1 : 0 ) ); // C

  // Charge Station
  m_ctrSheet.getRange("G16").setValue(m_UpdatedTeleopEngageData);
  m_ctrSheet.getRange("G21").setValue(m_UpdatedTeleopDockData);
  m_ctrSheet.getRange("B16").setValue(m_UpdatedAutoEngageData);
  m_ctrSheet.getRange("B21").setValue(m_UpdatedAutoDockData);

  // Grid
  m_ctrSheet.getRange("H15:J23").setValues(m_UpdatedTeleopGridData);
  m_ctrSheet.getRange("C15:E23").setValues(m_UpdatedAutoGridData);

  // Data Frequency
  m_ctrSheet.getRange("D1").setValue(m_DataUpdateFrequency+1);
}

function getPercentData( dataUpdateFrequency, inputData, baseData ) {
  var m_ctrCellValue = Math.ceil( dataUpdateFrequency * inputData );
  return ( m_ctrCellValue + ( baseData == true ? 1 : 0 ) ) / (dataUpdateFrequency+1);
}