/**
 * 생산관리시스템 - 백엔드
 * 공정명/설비명 기반의 간단한 생산관리 시스템
 */

function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('index_new')
      .evaluate()
      .setTitle('생산관리시스템')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput(`
      <html>
      <body>
        <h1>오류 발생</h1>
        <p>오류 내용: ${error.toString()}</p>
      </body>
      </html>
    `);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * process_machine 시트에서 공정명/설비명 데이터를 가져옵니다.
 * @returns {Array} 공정명과 설비명 데이터 배열
 */
function getProcessMachineData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('process_machine');

    if (!sheet) {
      Logger.log('process_machine 시트를 찾을 수 없습니다.');
      return [];
    }

    const data = sheet.getDataRange().getValues();

    // 헤더 제외하고 데이터만 반환 (첫 번째 행은 헤더로 가정)
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) { // 공정명과 설비명이 모두 있는 경우만
        result.push({
          processName: data[i][0],  // 첫 번째 열: 공정명
          machineName: data[i][1]    // 두 번째 열: 설비명
        });
      }
    }

    return result;

  } catch (error) {
    Logger.log('getProcessMachineData 오류: ' + error.toString());
    return [];
  }
}

/**
 * 고유한 공정명 목록을 가져옵니다.
 * @returns {Array} 중복 제거된 공정명 배열
 */
function getUniqueProcesses() {
  try {
    const data = getProcessMachineData();
    const processSet = new Set();

    data.forEach(item => {
      processSet.add(item.processName);
    });

    return Array.from(processSet).sort();

  } catch (error) {
    Logger.log('getUniqueProcesses 오류: ' + error.toString());
    return [];
  }
}

/**
 * 특정 공정에 속한 설비 목록을 가져옵니다.
 * @param {string} processName - 공정명
 * @returns {Array} 해당 공정의 설비명 배열
 */
function getMachinesByProcess(processName) {
  try {
    const data = getProcessMachineData();
    const machines = [];

    data.forEach(item => {
      if (item.processName === processName) {
        machines.push(item.machineName);
      }
    });

    return machines.sort();

  } catch (error) {
    Logger.log('getMachinesByProcess 오류: ' + error.toString());
    return [];
  }
}

/**
 * process_machine 시트 샘플 데이터 생성 (테스트용)
 */
function createSampleProcessMachineSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('process_machine');

  if (!sheet) {
    sheet = ss.insertSheet('process_machine');
  } else {
    sheet.clear();
  }

  // 헤더 및 샘플 데이터
  const data = [
    ['공정명', '설비명'],
    ['사출', '사출기-01'],
    ['사출', '사출기-02'],
    ['사출', '사출기-03'],
    ['조립', '조립라인-A'],
    ['조립', '조립라인-B'],
    ['검사', '검사대-01'],
    ['검사', '검사대-02'],
    ['포장', '포장라인-01']
  ];

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#4A86E8').setFontColor('white');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);

  Logger.log('process_machine 시트 생성 완료');
}
