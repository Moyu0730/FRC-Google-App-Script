var m_Sheet = SpreadsheetApp.getActiveSpreadsheet();
var m_BaseSheet = m_Sheet.getSheets()[0];
var m_IsUpdatedButtom = m_BaseSheet.getRange("C15");
var m_IsClearedButtom = m_BaseSheet.getRange("E15");
var m_StatusCell = m_BaseSheet.getRange("E16");
var m_IsUpdated = m_IsUpdatedButtom.getValue();
var m_IsCleared = m_IsClearedButtom.getValue();
var m_flag = 0;

function SubmitButtomChecked(){
  
  // m_IsUpdated = true;
  if( m_IsUpdated == true ){
    var m_UpdatedSheetName = m_BaseSheet.getRange("D14").getValue();
    UpdateSpecifyWorkSheetData(m_UpdatedSheetName);
    m_StatusCell.setBackground('#77FF00');
    m_StatusCell.setValue('Submit Successful');
    m_IsUpdatedButtom.setValue(false);
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
    m_IsClearedButtom.setValue(false);
  }else if( m_flag % 2 == 0 ){
    m_flag++;
    m_StatusCell.setBackground('#DDDDDD');
    m_StatusCell.setValue("There Isn't Any Request");
  }
}

function ClearUpdateWorkSheet(){
  var m_GridFalse = [
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
  
  var m_DefFalse = [ 
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

  var m_AtkFalse = [ 
                      [false], // 0
                      [false], // 1
                      [false], // 2
                      [false], // 3
                      [false], // 4
                      [false], // 5
                   ];

  var m_MoveFalse = [ 
                      [false], // 0
                      [false], // 1
                      [false], // 2
                    ];

  m_BaseSheet.getRangeList(["A4", "A9",  "E4", "E9"]).setValue(false);
  m_BaseSheet.getRange("B3:D11").setValues(m_GridFalse);
  m_BaseSheet.getRange("F3:H11").setValues(m_GridFalse);
  m_BaseSheet.getRange("J3:J11").setValues(m_DefFalse);
  m_BaseSheet.getRange("M3:M8").setValues(m_AtkFalse);
  m_BaseSheet.getRange("M10:M12").setValues(m_MoveFalse);
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
      var m_TeleopCellValue = Math.ceil(m_DataUpdateFrequency * m_UpdatedTeleopGridData[j][i]);
      var m_newTeleopCellValue = ( m_TeleopCellValue + ( m_BaseTeleopGridData[j][i] == true ? 1 : 0 ) ) / (m_DataUpdateFrequency+1);
      m_UpdatedTeleopGridData[j][i] = m_newTeleopCellValue;

      var m_AutoCellValue = Math.ceil(m_DataUpdateFrequency * m_UpdatedAutoGridData[j][i]);
      var m_newAutoCellValue = ( m_AutoCellValue + ( m_BaseAutoGridData[j][i] == true ? 1 : 0 ) ) / (m_DataUpdateFrequency+1);
      m_UpdatedAutoGridData[j][i] = m_newAutoCellValue;
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
    var m_AutoEngageCellValue = Math.ceil(m_DataUpdateFrequency * m_UpdatedAutoEngageData);
    var m_newAutoEngageCellValue = ( m_AutoEngageCellValue + ( m_BaseAutoEngageData == true ? 1 : 0 ) ) / (m_DataUpdateFrequency+1);         
    m_UpdatedAutoEngageData = m_newAutoEngageCellValue;

    // Auto Dock
    var m_AutoDockCellValue = Math.ceil(m_DataUpdateFrequency * m_UpdatedAutoDockData);
    var m_newAutoDockCellValue = ( m_AutoDockCellValue + ( m_BaseAutoDockData == true ? 1 : 0 ) ) / (m_DataUpdateFrequency+1);
    m_UpdatedAutoDockData = m_newAutoDockCellValue;

    // Teleop Engage
    var m_TeleopEngageCellValue = Math.ceil(m_DataUpdateFrequency * m_UpdatedTeleopEngageData);
    var m_newTeleopEngageCellValue = ( m_TeleopEngageCellValue + ( m_BaseTeleopEngageData == true ? 1 : 0 ) ) / (m_DataUpdateFrequency+1);         
    m_UpdatedTeleopEngageData = m_newTeleopEngageCellValue;

    // Teleop Dock
    var m_TeleopDockCellValue = Math.ceil(m_DataUpdateFrequency * m_UpdatedTeleopDockData);
    var m_newTeleopDockCellValue = ( m_TeleopDockCellValue + ( m_BaseTeleopDockData == true ? 1 : 0 ) ) / (m_DataUpdateFrequency+1);
    m_UpdatedTeleopDockData = m_newTeleopDockCellValue;
  
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
    var m_UpdatedTransportA = m_UpdatedSheet.getRange("J10").getValue();
    var m_UpdatedTransportB = m_UpdatedSheet.getRange("J11").getValue();
    var m_UpdatedTransportC = m_UpdatedSheet.getRange("J12").getValue();
    var m_BaseTransportA = m_BaseSheet.getRange("M10").getValue();
    var m_BaseTransportB = m_BaseSheet.getRange("M11").getValue();
    var m_BaseTransportC = m_BaseSheet.getRange("M12").getValue();

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

function testfunction(){
  var m_ctrSheet = m_Sheet.getSheets()[1];
  var m_DataUpdateFrequency = m_ctrSheet.getRange("D1").getValue();

  var m_CellValue = Math.ceil(m_DataUpdateFrequency * m_ctrSheet.getRange("C15").getValue());
  var m_newCellValue = (m_CellValue + ( m_BaseSheet.getRange("B3") == true ? 0 : 1 )) / (m_DataUpdateFrequency+1);

  m_ctrSheet.getRange("C15").setValue(m_newCellValue);
  m_ctrSheet.getRange("D1").setValue(m_DataUpdateFrequency+1);
}