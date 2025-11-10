function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
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

// ===== 시스템 초기화 및 시트 생성 =====

function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 1. 설비 마스터 생성
    createEquipmentMaster(ss);
    
    // 2. 제품 마스터 생성
    createProductMaster(ss);
    
    // 3. 작업자 마스터 생성
    createWorkerMaster(ss);
    
    // 4. 근무 스케줄 생성
    createWorkSchedule(ss);
    
    // 5. LOSS 마스터 생성 (기종교환 제외)
    createLossMaster(ss);
    
    // 6. 기종교환 마스터 생성
    createChangeoverMaster(ss);
    
    // 7. 불량 마스터 생성 (설정/공정 구분)
    createSettingDefectMaster(ss);
    createProcessDefectMaster(ss);
    
    // 8. 작업일지 시트 생성
    createWorkReport(ss);
    
    // 9. 불량 보고서 시트 생성
    createDefectReport(ss);
    
    // 10. 교대 인수인계 시트 생성
    createShiftHandover(ss);

    // 11. DAILY_SUMMARY 시트 생성
    createDailySummary(ss);  // ← 이 줄 추가
    
    Logger.log('모든 시트 초기화 완료');
    return '시스템 초기화 완료';
  } catch (error) {
    Logger.log('초기화 오류: ' + error.toString());
    return '초기화 중 오류 발생: ' + error.toString();
  }
}

// ===== DAILY_SUMMARY 관련 함수들 =====

// DAILY_SUMMARY 시트 생성
function createDailySummary(ss) {
  let sheet = ss.getSheetByName('DAILY_SUMMARY');
  if (!sheet) {
    sheet = ss.insertSheet('DAILY_SUMMARY');
    const headers = [
      'Date', 'Shift', 'EquipmentCode', 'ProductCode', 'WorkType', 
      'StartTime', 'EndTime', 'Quantity', 'TotalShiftTime', 'WorkTime', 'DowntimeTime', 
      'LossCode', 'DowntimeDetail', 'Remark', 'Status'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange('D:D').setNumberFormat('@');
  }
}

// 조 구분 함수 (StartTime 기준)
function getShiftByStartTime(startTime) {
  if (!startTime) return 'Unknown';
  
  const date = new Date(startTime);
  const hour = date.getHours();
  
  if (hour >= 0 && hour < 8) {
    return 'C조';
  } else if (hour >= 8 && hour < 16) {
    return 'A조';
  } else {
    return 'B조';
  }
}

// 조별 총 근무시간 계산
function calculateTotalShiftTime(shift, breakWork1, mealWork, breakWork2) {
  
  let totalTime = 480; // 기본 8시간
  
  if (breakWork1 === 'N') totalTime -= 10;
  if (mealWork === 'N') totalTime -= 40;
  if (breakWork2 === 'N') totalTime -= 10;
  
  return totalTime;
}

// WORKREPORT 데이터를 DAILY_SUMMARY로 변환
function processDailySummary(targetDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    createDailySummary(ss);
    
    const workReportSheet = ss.getSheetByName('WorkReport');
    const summarySheet = ss.getSheetByName('DAILY_SUMMARY');
    
    if (!workReportSheet) {
      Logger.log('WorkReport 시트를 찾을 수 없습니다.');
      return { success: false, message: 'WorkReport 시트를 찾을 수 없습니다.' };
    }
    
    const scriptTimeZone = Session.getScriptTimeZone();
    const data = workReportSheet.getDataRange().getValues();
    
    const targetDateStr = Utilities.formatDate(new Date(targetDate), scriptTimeZone, "yyyy-MM-dd");
    
    // 기존 해당 날짜 데이터 삭제
    const summaryData = summarySheet.getDataRange().getValues();
    for (let i = summaryData.length - 1; i >= 1; i--) {
      const summaryDate = (summaryData[i][0] instanceof Date) 
        ? Utilities.formatDate(summaryData[i][0], scriptTimeZone, "yyyy-MM-dd")
        : summaryData[i][0];
      
      if (summaryDate === targetDateStr) {
        summarySheet.deleteRow(i + 1);
      }
    }
    
    // 조별/설비별로 데이터 그룹화
    const shiftGroups = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const reportDate = (row[1] instanceof Date)
        ? Utilities.formatDate(row[1], scriptTimeZone, "yyyy-MM-dd")
        : row[1];
      
      if (reportDate !== targetDateStr) continue;
      
      const isOldFormat = row.length <= 15;
      
      const shift = getShiftByStartTime(row[6]);
      const equipmentCode = row[2];
      
      const groupKey = `${shift}|${equipmentCode}`;
      
      if (!shiftGroups[groupKey]) {
        shiftGroups[groupKey] = {
          date: reportDate,
          shift: shift,
          equipmentCode: equipmentCode,
          breakWork1: 'N',
          mealWork: 'N',
          breakWork2: 'N',
          records: []
        };
      }
      
      // WORK 타입일 때만 휴식근무 정보 업데이트
      if (row[5] === 'WORK') {
        shiftGroups[groupKey].breakWork1 = isOldFormat ? 'N' : (row[13] || 'N');
        shiftGroups[groupKey].mealWork = isOldFormat ? 'N' : (row[14] || 'N');
        shiftGroups[groupKey].breakWork2 = isOldFormat ? 'N' : (row[15] || 'N');
      }
      
      // 제품코드 날짜 오인식 방지 처리
      let productCode = '';
      if (row[4] instanceof Date) {
        const year = row[4].getFullYear();
        const month = row[4].getMonth() + 1;
        const monthStr = month.toString().padStart(2, '0');
        productCode = `${year}-${monthStr}`;
      } else {
        productCode = String(row[4] || '');
      }
      
      shiftGroups[groupKey].records.push({
        workType: row[5],
        productCode: productCode,
        startTime: row[6],
        endTime: row[7],
        quantity: row[8] || 0,
        lossCode: row[9] || '',
        downtimeDetail: isOldFormat ? '' : (row[11] || ''),
        remark: row[10] || '',
        status: isOldFormat ? 'COMPLETED' : (row[16] || 'COMPLETED')
      });
    }
    
    // 각 조별/설비별 처리
    const summaryRows = [];
    
    for (const groupKey in shiftGroups) {
      const group = shiftGroups[groupKey];
      
      // 총 근무시간 계산
      const totalShiftTime = calculateTotalShiftTime(
        group.shift,
        group.breakWork1,
        group.mealWork,
        group.breakWork2
      );
      
      // 휴식시간 정보
      const breakTimes = getBreakTimes(group.shift, targetDateStr);
      
      // 레코드를 시간순 정렬
      group.records.sort((a, b) => new Date(a.startTime) - new Date(b.startTime));
      
      // 각 레코드별 처리
      for (const record of group.records) {
        const startTime = new Date(record.startTime);
        const endTime = new Date(record.endTime);
        const duration = (endTime - startTime) / (1000 * 60);
        
        if (record.workType === 'WORK') {
          // WORK: 순수 작업시간만 계산
          let pureWorkTime = duration;
          
          // 이 작업 구간 내의 비가동시간 차감
          for (const otherRecord of group.records) {
            if (otherRecord.workType === 'DOWNTIME') {
              const downtimeStart = new Date(otherRecord.startTime);
              const downtimeEnd = new Date(otherRecord.endTime);
              
              if (downtimeStart >= startTime && downtimeEnd <= endTime) {
                pureWorkTime -= (downtimeEnd - downtimeStart) / (1000 * 60);
              }
            }
          }
          
          // 휴식시간 차감 (근무 안한 경우만)
          for (const breakInfo of breakTimes) {
            const breakStart = breakInfo.time;
            const breakEnd = new Date(breakStart.getTime() + breakInfo.duration * 60000);
            
            if (breakStart >= startTime && breakEnd <= endTime) {
              if ((breakInfo.type === 'break1' && group.breakWork1 === 'N') ||
                  (breakInfo.type === 'meal' && group.mealWork === 'N') ||
                  (breakInfo.type === 'break2' && group.breakWork2 === 'N')) {
                pureWorkTime -= breakInfo.duration;
              }
            }
          }
          
          // WORK 행 추가
          summaryRows.push([
            group.date,
            group.shift,
            group.equipmentCode,
            record.productCode,
            'WORK',
            formatUnifiedTime(startTime),
            formatUnifiedTime(endTime),
            record.quantity || 0,
            totalShiftTime,
            Math.round(pureWorkTime),
            0,
            '',
            '',
            record.remark,
            record.status
          ]);
          
        } else if (record.workType === 'DOWNTIME') {
          // DOWNTIME: 각 비가동을 별도 행으로
          summaryRows.push([
            group.date,
            group.shift,
            group.equipmentCode,
            record.productCode,
            'DOWNTIME',
            formatUnifiedTime(startTime),
            formatUnifiedTime(endTime),
            0,
            totalShiftTime,
            0,
            Math.round(duration),
            record.lossCode,
            record.downtimeDetail,
            record.remark,
            record.status
          ]);
          
        } else if (record.workType === 'CHANGEOVER') {
          // CHANGEOVER: 별도 행으로
          summaryRows.push([
            group.date,
            group.shift,
            group.equipmentCode,
            '',
            'CHANGEOVER',
            formatUnifiedTime(startTime),
            formatUnifiedTime(endTime),
            0,
            totalShiftTime,
            0,
            Math.round(duration),
            'CO-001',
            '',
            record.remark,
            record.status
          ]);
        }
      }
    }
    
    // DAILY_SUMMARY 시트에 데이터 추가
    if (summaryRows.length > 0) {
      summarySheet.getRange(summarySheet.getLastRow() + 1, 1, summaryRows.length, summaryRows[0].length)
        .setValues(summaryRows);
    }
    
    Logger.log(`${targetDateStr} DAILY_SUMMARY 생성 완료: ${summaryRows.length}개 레코드`);
    
    return { 
      success: true, 
      message: `${targetDateStr} 요약 완료 (${summaryRows.length}개 레코드)` 
    };
    
  } catch (error) {
    Logger.log('processDailySummary 오류: ' + error.toString());
    return { 
      success: false, 
      message: '요약 생성 중 오류 발생: ' + error.toString() 
    };
  }
}

// 특정 조/설비의 DAILY_SUMMARY만 생성 (중복 방지)
function processDailySummaryByShift(targetDate, shift, equipmentCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    createDailySummary(ss);
    
    const workReportSheet = ss.getSheetByName('WorkReport');
    const summarySheet = ss.getSheetByName('DAILY_SUMMARY');
    
    if (!workReportSheet) {
      Logger.log('WorkReport 시트를 찾을 수 없습니다.');
      return { success: false, message: 'WorkReport 시트를 찾을 수 없습니다.' };
    }
    
    const scriptTimeZone = Session.getScriptTimeZone();
    const data = workReportSheet.getDataRange().getValues();
    
    const targetDateStr = Utilities.formatDate(new Date(targetDate), scriptTimeZone, "yyyy-MM-dd");
    
    // ★ 중복 방지: 해당 날짜 + 조 + 설비 조합만 삭제
    const summaryData = summarySheet.getDataRange().getValues();
    for (let i = summaryData.length - 1; i >= 1; i--) {
      const summaryDate = (summaryData[i][0] instanceof Date) 
        ? Utilities.formatDate(summaryData[i][0], scriptTimeZone, "yyyy-MM-dd")
        : summaryData[i][0];
      const summaryShift = summaryData[i][1];
      const summaryEquipment = summaryData[i][2];
      
      // 같은 날짜 + 같은 조 + 같은 설비만 삭제
      if (summaryDate === targetDateStr && 
          summaryShift === shift && 
          summaryEquipment === equipmentCode) {
        summarySheet.deleteRow(i + 1);
      }
    }
    
    // 해당 조/설비만 그룹화
    const shiftGroup = {
      date: targetDateStr,
      shift: shift,
      equipmentCode: equipmentCode,
      breakWork1: 'N',
      mealWork: 'N',
      breakWork2: 'N',
      records: []
    };
    
    // WorkReport에서 해당 날짜 + 조 + 설비 데이터만 수집
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const reportDate = (row[1] instanceof Date)
        ? Utilities.formatDate(row[1], scriptTimeZone, "yyyy-MM-dd")
        : row[1];
      
      if (reportDate !== targetDateStr) continue;
      
      const isOldFormat = row.length <= 15;
      const rowShift = getShiftByStartTime(row[6]);
      const rowEquipment = row[2];
      
      // 해당 조 + 설비만 필터링
      if (rowShift !== shift || rowEquipment !== equipmentCode) continue;
      
      // WORK 타입일 때만 휴식근무 정보 업데이트
      if (row[5] === 'WORK') {
        shiftGroup.breakWork1 = isOldFormat ? 'N' : (row[13] || 'N');
        shiftGroup.mealWork = isOldFormat ? 'N' : (row[14] || 'N');
        shiftGroup.breakWork2 = isOldFormat ? 'N' : (row[15] || 'N');
      }
      
      // 제품코드 날짜 오인 방지 처리
      let productCode = '';
      if (row[4] instanceof Date) {
        const year = row[4].getFullYear();
        const month = row[4].getMonth() + 1;
        const monthStr = month.toString().padStart(2, '0');
        productCode = `${year}-${monthStr}`;
      } else {
        productCode = String(row[4] || '');
      }
      
      shiftGroup.records.push({
        workType: row[5],
        productCode: productCode,
        startTime: row[6],
        endTime: row[7],
        quantity: row[8] || 0,
        lossCode: row[9] || '',
        downtimeDetail: isOldFormat ? '' : (row[11] || ''),
        remark: row[10] || '',
        status: isOldFormat ? 'COMPLETED' : (row[16] || 'COMPLETED')
      });
    }
    
    // 레코드가 없으면 종료
    if (shiftGroup.records.length === 0) {
      Logger.log(`${targetDateStr} ${shift} ${equipmentCode} - 데이터 없음`);
      return { 
        success: true, 
        message: `${targetDateStr} ${shift} ${equipmentCode} - 작업 기록 없음` 
      };
    }
    
    // 총 근무시간 계산
    const totalShiftTime = calculateTotalShiftTime(
      shiftGroup.shift,
      shiftGroup.breakWork1,
      shiftGroup.mealWork,
      shiftGroup.breakWork2
    );
    
    // 휴식시간 정보
    const breakTimes = getBreakTimes(shiftGroup.shift, targetDateStr);
    
    // 레코드를 시간순 정렬
    shiftGroup.records.sort((a, b) => new Date(a.startTime) - new Date(b.startTime));
    
    // 각 레코드별 처리
    const summaryRows = [];
    
    for (const record of shiftGroup.records) {
      const startTime = new Date(record.startTime);
      const endTime = new Date(record.endTime);
      const duration = (endTime - startTime) / (1000 * 60);
      
      if (record.workType === 'WORK') {
        // WORK: 순수 작업시간만 계산
        let pureWorkTime = duration;
        
        // 이 작업 구간 내의 비가동시간 차감
        for (const otherRecord of shiftGroup.records) {
          if (otherRecord.workType === 'DOWNTIME') {
            const downtimeStart = new Date(otherRecord.startTime);
            const downtimeEnd = new Date(otherRecord.endTime);
            
            if (downtimeStart >= startTime && downtimeEnd <= endTime) {
              pureWorkTime -= (downtimeEnd - downtimeStart) / (1000 * 60);
            }
          }
        }
        
        // 휴식시간 차감 (근무 안한 경우만)
        for (const breakInfo of breakTimes) {
          const breakStart = breakInfo.time;
          const breakEnd = new Date(breakStart.getTime() + breakInfo.duration * 60000);
          
          if (breakStart >= startTime && breakEnd <= endTime) {
            if ((breakInfo.type === 'break1' && shiftGroup.breakWork1 === 'N') ||
                (breakInfo.type === 'meal' && shiftGroup.mealWork === 'N') ||
                (breakInfo.type === 'break2' && shiftGroup.breakWork2 === 'N')) {
              pureWorkTime -= breakInfo.duration;
            }
          }
        }
        
        // WORK 행 추가
        summaryRows.push([
          shiftGroup.date,
          shiftGroup.shift,
          shiftGroup.equipmentCode,
          record.productCode,
          'WORK',
          formatUnifiedTime(startTime),
          formatUnifiedTime(endTime),
          record.quantity || 0,
          totalShiftTime,
          Math.round(pureWorkTime),
          0,
          '',
          '',
          record.remark,
          record.status
        ]);
        
      } else if (record.workType === 'DOWNTIME') {
        // DOWNTIME: 각 비가동을 별도 행으로
        summaryRows.push([
          shiftGroup.date,
          shiftGroup.shift,
          shiftGroup.equipmentCode,
          record.productCode,
          'DOWNTIME',
          formatUnifiedTime(startTime),
          formatUnifiedTime(endTime),
          0,
          totalShiftTime,
          0,
          Math.round(duration),
          record.lossCode,
          record.downtimeDetail,
          record.remark,
          record.status
        ]);
        
      } else if (record.workType === 'CHANGEOVER') {
        // CHANGEOVER: 별도 행으로
        summaryRows.push([
          shiftGroup.date,
          shiftGroup.shift,
          shiftGroup.equipmentCode,
          '',
          'CHANGEOVER',
          formatUnifiedTime(startTime),
          formatUnifiedTime(endTime),
          0,
          totalShiftTime,
          0,
          Math.round(duration),
          'CO-001',
          '',
          record.remark,
          record.status
        ]);
      }
    }
    
    // DAILY_SUMMARY 시트에 데이터 추가
    if (summaryRows.length > 0) {
      summarySheet.getRange(summarySheet.getLastRow() + 1, 1, summaryRows.length, summaryRows[0].length)
        .setValues(summaryRows);
    }
    
    Logger.log(`${targetDateStr} ${shift} ${equipmentCode} DAILY_SUMMARY 생성 완료: ${summaryRows.length}개 레코드`);
    
    return { 
      success: true, 
      message: `${targetDateStr} ${shift} ${equipmentCode} 요약 완료 (${summaryRows.length}개 레코드)` 
    };
    
  } catch (error) {
    Logger.log('processDailySummaryByShift 오류: ' + error.toString());
    return { 
      success: false, 
      message: '요약 생성 중 오류 발생: ' + error.toString() 
    };
  }
}

// 조별 휴식시간 정보 반환 (날짜 기준 추가)
function getBreakTimes(shift, dateStr) {
  if (shift === 'A조') {
    return [
      { type: 'break1', time: new Date(`${dateStr}T10:00:00`), duration: 10 },
      { type: 'meal', time: new Date(`${dateStr}T12:00:00`), duration: 40 },
      { type: 'break2', time: new Date(`${dateStr}T14:00:00`), duration: 10 }
    ];
  } else if (shift === 'B조') {
    return [
      { type: 'break1', time: new Date(`${dateStr}T18:00:00`), duration: 10 },
      { type: 'meal', time: new Date(`${dateStr}T20:00:00`), duration: 40 },
      { type: 'break2', time: new Date(`${dateStr}T22:00:00`), duration: 10 }
    ];
  } else if (shift === 'C조') {
    return [
      { type: 'break1', time: new Date(`${dateStr}T02:00:00`), duration: 10 },
      { type: 'meal', time: new Date(`${dateStr}T04:00:00`), duration: 40 },
      { type: 'break2', time: new Date(`${dateStr}T06:00:00`), duration: 10 }
    ];
  }
  return [];
}

// 특정 날짜의 SUMMARY 생성 (수동 실행용)
function generateSummaryForDate() {
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  
  const result = processDailySummary(yesterday);
  Logger.log(result.message);
}

// 매일 자동 실행용 (트리거 설정 필요)
function dailyAutoSummary() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  
  processDailySummary(yesterday);
}

// 교대 인수인계 시트 생성
function createShiftHandover(ss) {
  let sheet = ss.getSheetByName('ShiftHandover');
  if (!sheet) {
    sheet = ss.insertSheet('ShiftHandover');
    const headers = ['Date', 'EquipmentCode', 'ShiftTime', 'ShiftType', 'WorkerCode', 'HandoverNote', 'CreatedAt'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

// 설비 마스터 생성 (실제 설비 데이터로 수정)
function createEquipmentMaster(ss) {
  let sheet = ss.getSheetByName('EquipmentMaster');
  if (!sheet) {
    sheet = ss.insertSheet('EquipmentMaster');
    const headers = ['EquipmentCode', 'EquipmentName', 'LineCode', 'IsActive'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const sampleData = [
      ['EQ-001', 'TPA800-1호기', 'LINE-A', 'Y'],
      ['EQ-002', 'TPA800-2호기', 'LINE-A', 'Y'],
      ['EQ-003', 'TPA800-3호기', 'LINE-A', 'Y'],
      ['EQ-004', 'TPA800-4호기', 'LINE-B', 'Y'],
      ['EQ-005', 'TPA300', 'LINE-B', 'Y'],
      ['EQ-006', 'Dorst750', 'LINE-B', 'Y']
    ];
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }
}

// TM_Master 시트 생성 (TM 접두사 추가)
function createProductMaster(ss) {
  let sheet = ss.getSheetByName('TM_Master');
  if (!sheet) {
    sheet = ss.insertSheet('TM_Master');
    const headers = ['ProductCode', 'ERP_Code', 'ProductName', 'Group', 'Sub_Group', 'IsActive'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const sampleData = [
      ['TM1360-00', 'ERP001', '성형-MOTOR PULLEY', 'BrakePad', 'Passenger', 'Y'],
      ['TM1361-00J', 'ERP002', '성형-BALL NUT PULLEY', 'BrakePad', 'Commercial', 'Y'],
      ['TM1361-01J', 'ERP003', '성형-BALL NUT PULLEY', 'ClutchPad', 'Passenger', 'Y'],
      ['TM1361-02J', 'ERP004', '성형-BALL NUT PULLEY', 'BrakePad', 'Industrial', 'Y'],
      ['TM1464-00', 'ERP005', '성형-PULLEY', 'Friction', 'Industrial', 'Y'],
      ['TM1464-01', 'ERP006', '성형-PULLEY', 'Friction', 'Industrial', 'Y']
    ];
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }
}

// 작업자 마스터 생성
function createWorkerMaster(ss) {
  let sheet = ss.getSheetByName('WorkerMaster');
  if (!sheet) {
    sheet = ss.insertSheet('WorkerMaster');
    const headers = ['WorkerCode', 'WorkerName', 'LineCode', 'IsActive'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const sampleData = [
      ['W001', '김철수', 'LINE-A', 'Y'],
      ['W002', '박영희', 'LINE-A', 'Y'],
      ['W003', '이정훈', 'LINE-B', 'Y'],
      ['W004', '최민수', 'LINE-B', 'Y']
    ];
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }
}

// 근무 스케줄 생성
function createWorkSchedule(ss) {
  let sheet = ss.getSheetByName('WorkSchedule');
  if (!sheet) {
    sheet = ss.insertSheet('WorkSchedule');
    const headers = ['ShiftName', 'StartTime', 'EndTime', 'IsActive'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const sampleData = [
      ['A조', '08:00', '16:00', 'Y'],
      ['B조', '16:00', '00:00', 'Y'],
      ['C조', '00:00', '08:00', 'Y']
    ];
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }
}

// 기종교환 마스터 생성
function createChangeoverMaster(ss) {
  let sheet = ss.getSheetByName('ChangeoverMaster');
  if (!sheet) {
    sheet = ss.insertSheet('ChangeoverMaster');
    const headers = ['ChangeoverCode', 'ChangeoverName', 'IsActive'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const sampleData = [
      ['CO-001', '기종교환', 'Y']
    ];
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }
}

// LOSS 마스터 생성 (기종교환 제외)
function createLossMaster(ss) {
  let sheet = ss.getSheetByName('LossMaster');
  if (!sheet) {
    sheet = ss.insertSheet('LossMaster');
    const headers = ['LossCode', 'LossName', 'LossType', 'IsActive'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const sampleData = [
      ['L002', '점검정비', 'Planned', 'Y'],
      ['L003', '청소', 'Planned', 'Y'],
      ['L004', '설비고장', 'Unplanned', 'Y'],
      ['L005', '금형교체', 'Unplanned', 'Y'],
      ['L006', '원료부족', 'Unplanned', 'Y']
    ];
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }
}

// 셋팅불량 마스터 생성
function createSettingDefectMaster(ss) {
  let sheet = ss.getSheetByName('SettingDefectMaster');
  if (!sheet) {
    sheet = ss.insertSheet('SettingDefectMaster');
    const headers = ['DefectCode', 'DefectName', 'DefectType', 'ProcessType', 'IsActive'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const sampleData = [
      ['SD001', '온도설정오류', '셋팅불량', 'Pressing', 'Y'],
      ['SD002', '압력설정오류', '셋팅불량', 'Pressing', 'Y'],
      ['SD003', '시간설정오류', '셋팅불량', 'Pressing', 'Y'],
      ['SD004', '금형온도이상', '셋팅불량', 'Pressing', 'Y'],
      ['SD005', '투입량오류', '셋팅불량', 'Mixing', 'Y']
    ];
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }
}

// 공정불량 마스터 생성
function createProcessDefectMaster(ss) {
  let sheet = ss.getSheetByName('ProcessDefectMaster');
  if (!sheet) {
    sheet = ss.insertSheet('ProcessDefectMaster');
    const headers = ['DefectCode', 'DefectName', 'DefectType', 'ProcessType', 'IsActive'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const sampleData = [
      ['PD001', '크랙발생', '공정불량', 'Pressing', 'Y'],
      ['PD002', '치수불량', '공정불량', 'Pressing', 'Y'],
      ['PD003', '표면불량', '공정불량', 'Pressing', 'Y'],
      ['PD004', '밀도불량', '공정불량', 'Pressing', 'Y'],
      ['PD005', '기포발생', '공정불량', 'Mixing', 'Y'],
      ['PD006', '이물혼입', '공정불량', 'Mixing', 'Y'],
      ['PD007', '경도불량', '공정불량', 'Curing', 'Y']
    ];
    sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  }
}

// 작업일지 시트 생성 (Status 컬럼 추가)
function createWorkReport(ss) {
  let sheet = ss.getSheetByName('WorkReport');
  if (!sheet) {
    sheet = ss.insertSheet('WorkReport');
    const headers = [
      'ID', 'Date', 'EquipmentCode', 'WorkerCode', 'ProductCode', 
      'WorkType', 'StartTime', 'EndTime', 'Quantity', 'LossCode', 
      'Remark', 'DowntimeDetail', 'CreatedAt', 'BreakWork1', 'MealWork', 'BreakWork2', 'Status'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange('E:E').setNumberFormat('@');
  }
}

function createDefectReport(ss) {
  let sheet = ss.getSheetByName('DefectReport');
  if (!sheet) {
    sheet = ss.insertSheet('DefectReport');
    const headers = [
      'ID', 'Date', 'EquipmentCode', 'ShiftName', 'ProductCode', 'DefectCode', 'DefectQty', 'Remark'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange('E:E').setNumberFormat('@');
  }
}

// ===== 마스터 데이터 조회 함수들 =====

// 마스터 데이터를 한 번에 모두 가져오는 함수
function getAllMasterData() {
  try {
    const products = getProductList();
    const workers = getWorkerList();
    const lossCodes = getLossCodeList();
    
    Logger.log('getAllMasterData - products: ' + products.length);
    Logger.log('getAllMasterData - workers: ' + workers.length);
    Logger.log('getAllMasterData - lossCodes: ' + lossCodes.length);
    
    // 배열 순서: [제품, 작업자, LOSS 코드]
    return [products, workers, lossCodes]; 
  } catch (error) {
    Logger.log('getAllMasterData 오류: ' + error.toString());
    return [[], [], []]; // 오류 발생 시 빈 배열 반환
  }
}


// 설비 목록 조회
function getEquipmentList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('EquipmentMaster');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const equipments = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === 'Y') { // IsActive
        equipments.push({
          EquipmentCode: data[i][0],
          EquipmentName: data[i][1],
          LineCode: data[i][2]
        });
      }
    }
    
    return equipments;
  } catch (error) {
    Logger.log('getEquipmentList 오류: ' + error.toString());
    return [];
  }
}

// 제품 목록 조회 (TM 접두사 제거)
function getProductList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TM_Master');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const products = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][5] === 'Y') { // IsActive
        let productCode = String(data[i][0]);
        
        // TM 접두사 제거
        if (productCode.startsWith('TM')) {
          productCode = productCode.substring(2);
        }
        
        products.push({
          ProductCode: productCode,
          ERP_Code: String(data[i][1]),
          ProductName: String(data[i][2]),
          Group: String(data[i][3]),
          Sub_Group: String(data[i][4])
        });
      }
    }
    
    return products;
  } catch (error) {
    Logger.log('getProductList 오류: ' + error.toString());
    return [];
  }
}

// 작업자 목록 조회
function getWorkerList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('WorkerMaster');
  
  if (!sheet) {
    Logger.log('WorkerMaster 시트가 없습니다.');
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const workers = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'Y') { // IsActive 확인
      workers.push({
        WorkerCode: data[i][0],
        WorkerName: data[i][1]
      });
    }
  }
  return workers;
}


// LOSS 코드 목록 조회
function getLossCodeList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('LossMaster');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const lossCodes = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === 'Y') { // IsActive
        lossCodes.push({
          LossCode: data[i][0],
          LossName: data[i][1],
          LossType: data[i][2]
        });
      }
    }
    
    return lossCodes;
  } catch (error) {
    Logger.log('getLossCodeList 오류: ' + error.toString());
    return [];
  }
}

// 시간을 통일된 형식으로 변환 (YYYY.MM.DD HH:mm)
function formatUnifiedTime(dateInput) {
  if (!dateInput) return '';
  
  let date;
  if (typeof dateInput === 'string') {
    date = new Date(dateInput);
  } else if (dateInput instanceof Date) {
    date = dateInput;
  } else {
    return '';
  }
  
  if (isNaN(date.getTime())) return '';
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  
  return `${year}.${month}.${day} ${hours}:${minutes}`;
}

// ===== 작업일지 관련 함수들 =====

// 상세한 설비 상태 정보 조회 (INCOMPLETE 기종교환 감지)
function getDetailedEquipmentStatus(equipmentCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('WorkReport');
    
    if (!sheet) {
      return {
        hasOngoing: false,
        ongoingType: '',
        details: '',
        lastActivity: '상태 확인 중 오류 발생',
        lastProductCode: ''
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const scriptTimeZone = Session.getScriptTimeZone();
    
    // 해당 설비의 기록들을 최신순으로 확인
    let ongoingRecord = null;
    let incompleteChangeover = null;
    let lastCompletedRecord = null;
    
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (row[2] === equipmentCode) { // 같은 설비
        if (row[16] === 'ONGOING') { // 진행중인 것
          // 제품코드 날짜 오인 처리
          let productCode = '';
          if (row[4] instanceof Date) {
            const year = row[4].getFullYear();
            const month = row[4].getMonth() + 1;
            const monthStr = month.toString().padStart(2, '0');
            productCode = `${year}-${monthStr}`;
          } else {
            productCode = String(row[4] || '');
          }
          
          ongoingRecord = {
            workType: row[5],
            productCode: productCode,
            startTime: row[6],
            endTime: row[7]
          };
          break;
        } else if (row[16] === 'INCOMPLETE' && row[5] === 'CHANGEOVER' && !incompleteChangeover) {
          // 미완료 기종교환 발견
          incompleteChangeover = {
            workType: row[5],
            startTime: row[6],
            endTime: row[7]
          };
        } else if (row[16] === 'COMPLETED' && !lastCompletedRecord) {
          // 제품코드 날짜 오인 처리
          let productCode = '';
          if (row[4] instanceof Date) {
            const year = row[4].getFullYear();
            const month = row[4].getMonth() + 1;
            const monthStr = month.toString().padStart(2, '0');
            productCode = `${year}-${monthStr}`;
          } else {
            productCode = String(row[4] || '');
          }
          
          lastCompletedRecord = {
            workType: row[5],
            productCode: productCode,
            endTime: row[7]
          };
        }
      }
    }
    
    // 1순위: 진행중인 작업이 있는 경우
    if (ongoingRecord) {
      let workType = '';
      let details = '';
      
      if (ongoingRecord.workType === 'WORK') {
        workType = '진행중인 작업';
        details = `${ongoingRecord.productCode} (${Utilities.formatDate(new Date(ongoingRecord.startTime), scriptTimeZone, "M/d HH:mm")})`;
      } else if (ongoingRecord.workType === 'CHANGEOVER') {
        workType = '진행중인 기종교환';
        details = `기종교환 (${Utilities.formatDate(new Date(ongoingRecord.startTime), scriptTimeZone, "M/d HH:mm")})`;
      }
      
      return {
        hasOngoing: true,
        ongoingType: workType,
        details: details,
        lastActivity: null,
        lastProductCode: ongoingRecord.productCode
      };
    }
    
    // 2순위: 미완료 기종교환이 있는 경우
    if (incompleteChangeover) {
      return {
        hasOngoing: true,
        ongoingType: '⚠️ 미완료 기종교환',
        details: `이전 교대에서 기종교환이 미완료 상태입니다 (${Utilities.formatDate(new Date(incompleteChangeover.endTime), scriptTimeZone, "M/d HH:mm")})`,
        lastActivity: null,
        lastProductCode: ''
      };
    }
    
    // 3순위: 완료된 상태
    if (lastCompletedRecord) {
      let activityType = '';
      if (lastCompletedRecord.workType === 'WORK') {
        activityType = 'WORK 완료';
      } else if (lastCompletedRecord.workType === 'CHANGEOVER') {
        activityType = '기종교환 완료';
      } else if (lastCompletedRecord.workType === 'DOWNTIME') {
        activityType = '비가동 완료';
      }
      
      const endTime = lastCompletedRecord.endTime ? 
        Utilities.formatDate(new Date(lastCompletedRecord.endTime), scriptTimeZone, "M/d HH:mm") : 
        '시간 미확인';
      
      return {
        hasOngoing: false,
        ongoingType: '',
        details: '',
        lastActivity: `${activityType} (${endTime})`,
        lastProductCode: lastCompletedRecord.productCode
      };
    }
    
    // 4순위: 기록이 없는 경우
    return {
      hasOngoing: false,
      ongoingType: '',
      details: '',
      lastActivity: null,
      lastProductCode: ''
    };
    
  } catch (error) {
    Logger.log('getDetailedEquipmentStatus 오류: ' + error.toString());
    return {
      hasOngoing: false,
      ongoingType: '',
      details: '',
      lastActivity: '상태 확인 중 오류 발생: ' + error.toString(),
      lastProductCode: ''
    };
  }
}

// 작업 시작 (시간 형식 통일)
function startWork(workData) {
  Logger.log('=== startWork 시작 ===');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('WorkReport');
    
    if (!sheet) {
      return JSON.stringify({ success: false, message: 'WorkReport 시트를 찾을 수 없습니다.' });
    }
    
    // 미완료 기종교환 체크
    const hasIncompleteChangeover = checkIncompleteChangeover(workData.equipmentCode, workData.date);
    if (hasIncompleteChangeover) {
      return JSON.stringify({ 
        success: false, 
        message: '이전 교대에서 기종교환이 미완료 상태입니다.\n먼저 기종교환을 등록해주세요.' 
      });
    }
    
    // 제품 변경 검증 로직... (기존과 동일)
    const data = sheet.getDataRange().getValues();
    let lastWorkProductCode = null;
    
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (row[2] === workData.equipmentCode && 
          row[5] === 'WORK' && 
          row[16] === 'COMPLETED' && 
          row[4] !== '') {
        
        // 제품코드 날짜 오인 처리
        if (row[4] instanceof Date) {
          const year = row[4].getFullYear();
          const month = row[4].getMonth() + 1;
          const monthStr = month.toString().padStart(2, '0');
          lastWorkProductCode = `${year}-${monthStr}`;
        } else {
          lastWorkProductCode = String(row[4]);
        }
        break;
      }
    }
    
    if (lastWorkProductCode && lastWorkProductCode !== workData.productCode) {
      let hasRecentChangeover = false;
      for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        if (row[2] === workData.equipmentCode) {
          if (row[5] === 'CHANGEOVER') {
            hasRecentChangeover = true;
            break;
          } else if (row[5] === 'WORK') {
            break;
          }
        }
      }
      
      if (!hasRecentChangeover) {
        return JSON.stringify ({ 
          success: false, 
          message: `제품이 변경되었습니다. (${lastWorkProductCode} → ${workData.productCode})\n먼저 기종교환을 등록해주세요.` 
        });
      }
    }
    
    // 진행중인 기종교환 완료 처리
    const sheetData = sheet.getDataRange().getValues();
    let ongoingChangeoverRowIndex = -1;

    for (let i = sheetData.length - 1; i >= 1; i--) {
      const row = sheetData[i];
      if (row[2] === workData.equipmentCode && 
          row[5] === 'CHANGEOVER' && 
          row[16] === 'ONGOING') {
        ongoingChangeoverRowIndex = i + 1;
        break;
      }
    }

    if (ongoingChangeoverRowIndex !== -1) {
      const startDateTime = formatUnifiedTime(workData.startTime);
      sheet.getRange(ongoingChangeoverRowIndex, 8).setValue(startDateTime); // EndTime
      sheet.getRange(ongoingChangeoverRowIndex, 17).setValue('COMPLETED'); // Status
    }
    
    // 새로운 작업 등록
    const id = generateId();
    const startTimeFormatted = formatUnifiedTime(workData.startTime);
    const createdAtFormatted = formatUnifiedTime(new Date());
    
    let productCodeToSave = '';
    if (workData.productCode instanceof Date) {
      const year = workData.productCode.getFullYear();
      const month = workData.productCode.getMonth() + 1;
      const monthStr = month.toString().padStart(2, '0');
      productCodeToSave = `${year}-${monthStr}`;
    } else {
      productCodeToSave = String(workData.productCode);
    }
    
    const rowData = [
      id,                          // A: ID
      workData.date,               // B: Date
      workData.equipmentCode,      // C: EquipmentCode
      workData.workerCode,         // D: WorkerCode
      productCodeToSave,        // E: ProductCode
      'WORK',                      // F: WorkType
      startTimeFormatted,          // G: StartTime
      '',                          // H: EndTime
      0,                           // I: Quantity
      '',                          // J: LossCode
      workData.remark || '',       // K: Remark
      '',                          // L: DowntimeDetail (작업은 빈값)
      createdAtFormatted,          // M: CreatedAt
      'N',                         // N: BreakWork1
      'N',                         // O: MealWork
      'N',                         // P: BreakWork2
      'ONGOING'                    // Q: Status
    ];

    sheet.appendRow(rowData);
    
    return JSON.stringify({
      success: true,
      message: '작업이 시작되었습니다.',
      workId: id
    });
  } catch (error) {
    Logger.log('startWork 치명적 오류: ' + error.toString());
    return JSON.stringify ({
      success: false,
      message: '작업 시작 중 오류가 발생했습니다: ' + error.toString()
    });
  }
}

// 비가동 등록 (JSON 문자열 반환)
function addDowntime(downtimeData) {
  Logger.log('=== addDowntime 시작 ===');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('WorkReport');
    
    if (!sheet) {
      const errorResponse = { success: false, message: 'WorkReport 시트를 찾을 수 없습니다.' };
      return JSON.stringify(errorResponse);
    }
    
    const id = generateId();
    const startTimeFormatted = formatUnifiedTime(downtimeData.startTime);
    const endTimeFormatted = formatUnifiedTime(downtimeData.endTime);
    const createdAtFormatted = formatUnifiedTime(new Date());
    
    // TM-NO 자동 매핑
    const productCodeAtTime = getProductCodeByTime(
      downtimeData.equipmentCode, 
      downtimeData.startTime, 
      downtimeData.date
    );

    const rowData = [
      id,                          // A: ID
      downtimeData.date,           // B: Date
      downtimeData.equipmentCode,  // C: EquipmentCode
      downtimeData.workerCode,     // D: WorkerCode
      productCodeAtTime,           // E: ProductCode (TM-NO 자동 매핑)
      'DOWNTIME',                  // F: WorkType
      startTimeFormatted,          // G: StartTime
      endTimeFormatted,            // H: EndTime
      0,                           // I: Quantity
      downtimeData.lossCode,       // J: LossCode
      '',                          // K: Remark (비가동은 빈값)
      downtimeData.downtimeDetail || '', // L: DowntimeDetail
      createdAtFormatted,          // M: CreatedAt
      'N',                         // N: BreakWork1
      'N',                         // O: MealWork
      'N',                         // P: BreakWork2
      'COMPLETED'                  // Q: Status
    ];

    sheet.appendRow(rowData);
    
    const successResponse = {
      success: true,
      message: '비가동이 등록되었습니다.'
    };
    return JSON.stringify(successResponse);
    
  } catch (error) {
    Logger.log('addDowntime 오류: ' + error.toString());
    const errorResponse = {
      success: false,
      message: '비가동 등록 중 오류가 발생했습니다: ' + error.toString()
    };
    return JSON.stringify(errorResponse);
  }
}

// 비가동 시작시간 기준 진행중인 제품코드 찾기
function getProductCodeByTime(equipmentCode, downtimeStartTime, downtimeDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('WorkReport');
    
    if (!sheet) return '';
    
    const data = sheet.getDataRange().getValues();
    const scriptTimeZone = Session.getScriptTimeZone();
    const downtimeStartDateTime = new Date(downtimeStartTime);
    
    let foundProductCode = '';
    
    // 해당 설비의 모든 작업 기록을 시간순으로 확인
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const reportDate = (row[1] instanceof Date) 
        ? Utilities.formatDate(row[1], scriptTimeZone, "yyyy-MM-dd")
        : row[1];
      
      if (row[2] === equipmentCode && // 같은 설비
          row[5] === 'WORK' && // 작업만
          row[4] !== '') { // 제품코드가 있는 것만
        
        const startTime = row[6] ? new Date(row[6]) : null;
        const endTime = row[7] ? new Date(row[7]) : null;
        
        // 비가동 시작시간에 진행중인 작업 찾기
        if (startTime && startTime <= downtimeStartDateTime) {
          if (!endTime || endTime > downtimeStartDateTime) {
            // 진행중인 작업 발견 - 제품코드 날짜 오인 처리
            if (row[4] instanceof Date) {
              const year = row[4].getFullYear();
              const month = row[4].getMonth() + 1;
              const monthStr = month.toString().padStart(2, '0');
              foundProductCode = `${year}-${monthStr}`;
            } else {
              foundProductCode = String(row[4]);
            }
            break;
          } else if (endTime <= downtimeStartDateTime) {
            // 가장 최근 완료된 작업
            if (!foundProductCode) {
              if (row[4] instanceof Date) {
                const year = row[4].getFullYear();
                const month = row[4].getMonth() + 1;
                const monthStr = month.toString().padStart(2, '0');
                foundProductCode = `${year}-${monthStr}`;
              } else {
                foundProductCode = String(row[4]);
              }
            }
          }
        }
      }
    }
    
    return foundProductCode;
  } catch (error) {
    Logger.log('getProductCodeByTime 오류: ' + error.toString());
    return '';
  }
}

// 작업 종료 (시간 형식 통일)
function endWork(workEndData) {
  Logger.log('=== endWork 시작 ===');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('WorkReport');
    
    if (!sheet) {
      return JSON.stringify ({ success: false, message: 'WorkReport 시트를 찾을 수 없습니다.' });
    }
    
    const data = sheet.getDataRange().getValues();
    let targetRowIndex = -1;
    
    for (let i = data.length - 1; i >= 1; i--) {
      const rowDateStr = (data[i][1] instanceof Date) 
        ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : data[i][1];
        
      if (rowDateStr === workEndData.date &&
          data[i][2] === workEndData.equipmentCode &&
          data[i][5] === 'WORK' &&
          data[i][16] === 'ONGOING') {
        targetRowIndex = i + 1;
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      return JSON.stringify ({ success: false, message: '진행중인 작업을 찾을 수 없습니다.' });
    }
    
    const endTimeFormatted = formatUnifiedTime(workEndData.endTime);
    
    sheet.getRange(targetRowIndex, 8).setValue(endTimeFormatted); // EndTime
    sheet.getRange(targetRowIndex, 9).setValue(workEndData.quantity); // Quantity
    sheet.getRange(targetRowIndex, 17).setValue('COMPLETED'); // Status
    
    return JSON.stringify({
      success: true,
      message: `작업이 완료되었습니다. (생산량: ${workEndData.quantity}개)`
    });
  } catch (error) {
    Logger.log('endWork 오류: ' + error.toString());
    return JSON.stringify ({
      success: false,
      message: '작업 종료 중 오류가 발생했습니다: ' + error.toString()
    });
  }
}

// 기종교환 시작 (시간 형식 통일)
function addChangeover(changeoverData) {
  Logger.log('=== addChangeover 시작 ===');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('WorkReport');

    if (!sheet) {
      const errorResponse = { success: false, message: 'WorkReport 시트를 찾을 수 없습니다.' };
      return JSON.stringify(errorResponse);
    }

    const id = generateId();

    // 1) 수동 시작시간이 있으면 최우선 적용
    let startTimeDate = null;
    if (changeoverData.startTimeManual) {
      startTimeDate = new Date(changeoverData.startTimeManual);
    }

    // 2) 없으면 기존 규칙(미완료 교대/마지막 작업 종료시간/지정시간) 적용
    if (!startTimeDate) {
      const hasIncompleteChangeover = checkIncompleteChangeover(changeoverData.equipmentCode, changeoverData.date);

      if (hasIncompleteChangeover) {
        const data = sheet.getDataRange().getValues();
        let incompleteEndTime = null;

        for (let i = data.length - 1; i >= 1; i--) {
          const row = data[i];
          if (row[2] === changeoverData.equipmentCode &&
              row[5] === 'CHANGEOVER' &&
              row[16] === 'INCOMPLETE' &&
              row[7] !== '') {
            incompleteEndTime = new Date(row[7]);
            break;
          }
        }

        if (incompleteEndTime) {
          startTimeDate = incompleteEndTime;
        } else {
          const currentShift = getCurrentShiftInfo();
          const today = changeoverData.date;
          let startTimeStr = String(currentShift.StartTime);
          if (currentShift.StartTime instanceof Date) {
            startTimeStr = Utilities.formatDate(currentShift.StartTime, Session.getScriptTimeZone(), "HH:mm");
          } else {
            const timeParts = startTimeStr.split(':');
            const hours = timeParts[0].padStart(2, '0');
            const minutes = timeParts[1].padStart(2, '0');
            startTimeStr = `${hours}:${minutes}`;
          }
          startTimeDate = new Date(`${today}T${startTimeStr}:00`);
        }
      } else {
        const data = sheet.getDataRange().getValues();
        let lastWorkEndTime = null;

        for (let i = data.length - 1; i >= 1; i--) {
          const row = data[i];
          if (row[2] === changeoverData.equipmentCode &&
              row[5] === 'WORK' &&
              row[16] === 'COMPLETED' &&
              row[7] !== '') {
            lastWorkEndTime = new Date(row[7]);
            break;
          }
        }

        if (lastWorkEndTime) {
          startTimeDate = lastWorkEndTime;
        } else {
          startTimeDate = new Date(changeoverData.startTime); // 요청시각/기존 fallback
        }
      }
    }

    const startTimeFormatted = formatUnifiedTime(startTimeDate);
    const createdAtFormatted = formatUnifiedTime(new Date());

    const rowData = [
      id,
      changeoverData.date,
      changeoverData.equipmentCode,
      changeoverData.workerCode,
      '',
      'CHANGEOVER',
      startTimeFormatted,
      '',
      0,
      'CO-001',
      changeoverData.remark || '',
      '',
      createdAtFormatted,
      'N',
      'N', 
      'N',
      'ONGOING'
    ];

    sheet.appendRow(rowData);

    const successResponse = {
      success: true,
      message: changeoverData.startTimeManual
        ? '수동 입력한 시작시간으로 기종교환이 시작되었습니다.'
        : '기종교환이 시작되었습니다.'
    };
    return JSON.stringify(successResponse);

  } catch (error) {
    Logger.log('addChangeover 치명적 오류: ' + error.toString());
    const errorResponse = {
      success: false,
      message: '기종교환 시작 중 오류가 발생했습니다: ' + error.toString()
    };
    return JSON.stringify(errorResponse);
  }
}

// 근무종료 (시간 형식 통일)
function endShift(shiftData) {
  Logger.log('=== endShift 시작 ===');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const workSheet = ss.getSheetByName('WorkReport');
    const handoverSheet = ss.getSheetByName('ShiftHandover');
    
    if (!workSheet) {
      return { success: false, message: 'WorkReport 시트를 찾을 수 없습니다.' };
    }
    
    const data = workSheet.getDataRange().getValues();
    const endTimeFormatted = formatUnifiedTime(shiftData.endTime);
    let updatedCount = 0;
    let incompleteChangeover = false;
    
    // 해당 설비의 진행중인 작업들을 모두 완료 처리
    for (let i = data.length - 1; i >= 1; i--) {
      const rowDateStr = (data[i][1] instanceof Date) 
        ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : data[i][1];
        
      if (rowDateStr === shiftData.date &&
          data[i][2] === shiftData.equipmentCode &&
          data[i][16] === 'ONGOING') {
        
        const rowIndex = i + 1;
        
        if (data[i][5] === 'WORK') {
          workSheet.getRange(rowIndex, 8).setValue(endTimeFormatted); // EndTime
          workSheet.getRange(rowIndex, 9).setValue(shiftData.quantity || 0); // Quantity
          workSheet.getRange(rowIndex, 17).setValue('COMPLETED'); // Status
        } else if (data[i][5] === 'CHANGEOVER') {
          workSheet.getRange(rowIndex, 8).setValue(endTimeFormatted); // EndTime
          workSheet.getRange(rowIndex, 17).setValue('INCOMPLETE'); // Status
          incompleteChangeover = true;
        }
        
        updatedCount++;
      }
    }
    
    // 휴식근무 정보 일괄 적용
    const dataRefresh = workSheet.getDataRange().getValues();
    let breakWorkUpdatedCount = 0;
    
    for (let i = dataRefresh.length - 1; i >= 1; i--) {
      const rowDateStr = (dataRefresh[i][1] instanceof Date) 
        ? Utilities.formatDate(dataRefresh[i][1], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : dataRefresh[i][1];
        
      if (rowDateStr === shiftData.date &&
          dataRefresh[i][2] === shiftData.equipmentCode &&
          dataRefresh[i][5] === 'WORK' &&
          dataRefresh[i][16] === 'COMPLETED') {
        
        const rowIndex = i + 1;
        workSheet.getRange(rowIndex, 14).setValue(shiftData.breakWork1 || 'N');
        workSheet.getRange(rowIndex, 15).setValue(shiftData.mealWork || 'N');
        workSheet.getRange(rowIndex, 16).setValue(shiftData.breakWork2 || 'N');
        
        breakWorkUpdatedCount++;
      }
    }
    
    // 인계사항 저장
    let finalHandoverNote = shiftData.handoverNote ? shiftData.handoverNote.trim() : '';
    
    if (incompleteChangeover) {
      const autoNote = '⚠️ 기종교환이 미완료 상태입니다. 다음 교대에서 기종교환을 먼저 등록해주세요.';
      finalHandoverNote = finalHandoverNote ? finalHandoverNote + '\n\n' + autoNote : autoNote;
    }
    
    if (handoverSheet && finalHandoverNote !== '') {
      const createdAtFormatted = formatUnifiedTime(new Date());
      const handoverData = [
        shiftData.date,
        shiftData.equipmentCode,
        shiftData.endTime,
        getShiftTypeName(shiftData.endTime),
        shiftData.workerCode || 'UNKNOWN',
        finalHandoverNote,
        createdAtFormatted
      ];
      
      handoverSheet.appendRow(handoverData);
    }
    
    // ★ DAILY_SUMMARY 자동 업로드 추가
    const currentShift = getCurrentShiftInfo();
    processDailySummaryByShift(shiftData.date, currentShift.ShiftName, shiftData.equipmentCode);
    
    return JSON.stringify({
      success: true,
      message: `근무가 종료되었습니다. (완료된 항목: ${updatedCount}개, 휴식근무 적용: ${breakWorkUpdatedCount}개 작업)${incompleteChangeover ? ' ⚠️ 미완료 기종교환이 다음 교대로 인계되었습니다.' : ''}`
    });
    
  } catch (error) {
    Logger.log('endShift 오류: ' + error.toString());
    return JSON.stringify ({
      success: false,
      message: '근무종료 중 오류가 발생했습니다: ' + error.toString()
    });
  }
}

// 이전 교대의 인계사항 조회
function getPreviousShiftHandover(equipmentCode, currentDate, currentShiftTime) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ShiftHandover');
    
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const scriptTimeZone = Session.getScriptTimeZone();
    
    // 이전 교대 시간 계산
    let previousShiftTime;
    if (currentShiftTime === '08:00') {
      previousShiftTime = '00:00'; // A조 시작 전 = B조 종료
    } else if (currentShiftTime === '16:00') {
      previousShiftTime = '08:00'; // B조 시작 전 = A조 종료  
    } else if (currentShiftTime === '00:00') {
      previousShiftTime = '16:00'; // C조 시작 전 = B조 종료
    }
    
    // 해당 설비의 이전 교대 인계사항 찾기 (최근순)
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const handoverDate = (row[0] instanceof Date) 
        ? Utilities.formatDate(row[0], scriptTimeZone, "yyyy-MM-dd")
        : row[0];
      const handoverShiftTime = (row[2] instanceof Date)
        ? Utilities.formatDate(row[2], scriptTimeZone, "HH:mm")
        : row[2];
        
      if (row[1] === equipmentCode && // 같은 설비
          handoverDate === currentDate && // 같은 날짜
          handoverShiftTime === previousShiftTime) { // 이전 교대
        
        return JSON.stringify ({
          date: handoverDate,
          shiftType: row[3],
          workerCode: row[4],
          handoverNote: row[5],
          createdAt: row[6]
        });
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log('getPreviousShiftHandover 오류: ' + error.toString());
    return null;
  }
}

// 근무시작 기록
function startShift(startShiftData) {
  Logger.log('=== startShift 시작 ===');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 근무시작 기록을 별도 시트에 저장하거나 WorkReport에 특별 타입으로 저장
    // 여기서는 간단히 로그만 남기고 성공 응답
    Logger.log('근무시작 - 설비: ' + startShiftData.equipmentCode + ', 시간: ' + startShiftData.startTime);
    
    return {
      success: true,
      message: '근무가 시작되었습니다. 오늘도 안전한 근무 되세요!'
    };
    
  } catch (error) {
    Logger.log('startShift 오류: ' + error.toString());
    return {
      success: false,
      message: '근무시작 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

// 교대 시간에 따른 교대명 반환 (WorkSchedule 기반)
function getShiftTypeName(timeStr) {
  try {
    const schedules = getWorkScheduleList();
    
    for (const schedule of schedules) {
      if (schedule.StartTime === timeStr) {
        return schedule.ShiftName;
      }
    }
    
    return '기타';
  } catch (error) {
    Logger.log('getShiftTypeName 오류: ' + error.toString());
    return '기타';
  }
}

// 특정 설비/2일간의 작업일지 조회 (통일된 시간 형식 파싱)
function getWorkReportByEquipment(equipmentCode, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('WorkReport');

  if (!sheet) {
    console.log("WorkReport 시트를 찾을 수 없습니다.");
    return JSON.stringify([]);
  }
  
  const data = sheet.getDataRange().getValues();
  const productMap = getProductMap();
  const workerMap = getWorkerMap();
  const lossMap = getLossMap();
  const scriptTimeZone = Session.getScriptTimeZone(); 
  const results = [];
  
  // 5일 전 날짜 계산
  const targetDate = new Date(date);
  const fiveDaysAgo = new Date(targetDate);
  fiveDaysAgo.setDate(targetDate.getDate() - 4);
  const fiveDaysAgoStr = Utilities.formatDate(fiveDaysAgo, scriptTimeZone, "yyyy-MM-dd");
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;

    const isOldFormat = row.length <= 15;
    
    const reportDate = (row[1] instanceof Date)
      ? Utilities.formatDate(row[1], scriptTimeZone, "yyyy-MM-dd")
      : row[1];
    
    if (row[2] === equipmentCode && reportDate >= fiveDaysAgoStr && reportDate <= date) {
      
      // 제품코드 안전 처리
      let productCode = '';
      if (row[4] instanceof Date) {
        const year = row[4].getFullYear();
        const month = row[4].getMonth() + 1; // 0부터 시작하므로 +1
        
        // 년도-월 형태로 제품코드 복원
        const monthStr = month.toString().padStart(2, '0');
        productCode = `${year}-${monthStr}`;
      } else {
        productCode = String(row[4]);
      }
      
      // 시간 파싱 함수
      function parseTime(timeValue) {
        if (!timeValue) return null;
        
        if (timeValue instanceof Date) {
          return timeValue;
        }
        
        if (typeof timeValue === 'string') {
          // 새로운 형식: "2025.09.27 15:10"
          if (timeValue.includes('.')) {
            const cleanTime = timeValue.replace(/\./g, '-');
            return new Date(cleanTime);
          }
          // 기존 ISO 형식
          return new Date(timeValue);
        }
        
        return null;
      }
      
      const report = {
          id: row[0],
          date: reportDate,
          equipmentCode: row[2],
          workerCode: row[3],
          workerName: workerMap[row[3]] || row[3],
          productCode: productCode,
          productName: productMap[productCode] || productCode,
          workType: row[5],
          startTime: parseTime(row[6]),
          endTime: parseTime(row[7]),
          quantity: row[8],
          lossCode: row[9],
          lossName: lossMap[row[9]] || row[9] || '',
          remark: row[10],
          downtimeDetail: isOldFormat ? '' : (row[11] || ''),
          createdAt: isOldFormat ? 
            ((row[11] instanceof Date) ? Utilities.formatDate(row[11], scriptTimeZone, "yyyy-MM-dd HH:mm:ss") : row[11]) :
            ((row[12] instanceof Date) ? Utilities.formatDate(row[12], scriptTimeZone, "yyyy-MM-dd HH:mm:ss") : row[12]),
          breakWork1: isOldFormat ? 'N' : (row[13] || 'N'),
          mealWork: isOldFormat ? 'N' : (row[14] || 'N'),
          breakWork2: isOldFormat ? 'N' : (row[15] || 'N'),
          status: isOldFormat ? 'COMPLETED' : (row[16] || 'COMPLETED')
      };
      
      results.push(report);
    }
  }

  // 등록순 정렬을 최신순으로 변경
  results.sort((a, b) => {
    if (a.date !== b.date) {
      return b.date.localeCompare(a.date); // 날짜 내림차순
    }
    // 같은 날이면 등록시간으로 내림차순 (최신이 위)
    const aCreated = new Date(a.createdAt);
    const bCreated = new Date(b.createdAt);
    return bCreated - aCreated;
  });

  console.log("최종 결과 (통일된 시간 파싱):", results);
  return JSON.stringify(results);
}

// ===== 헬퍼 함수들 =====

// 고유 ID 생성
function generateId() {
  return 'ID' + new Date().getTime() + Math.random().toString(36).substring(2, 7);
}

// 제품 매핑 데이터 생성 (TM 접두사 처리)
function getProductMap() {
  const products = getProductList();
  const map = {};
  products.forEach(product => {
    map[product.ProductCode] = product.ProductName;
  });
  return map;
}

// 작업자 매핑 데이터 생성
function getWorkerMap() {
  const workers = getWorkerList();
  const map = {};
  workers.forEach(worker => {
    map[worker.WorkerCode] = worker.WorkerName;
  });
  return map;
}

// Loss 매핑 데이터 생성
function getLossMap() {
  const lossCodes = getLossCodeList();
  const map = {};
  lossCodes.forEach(loss => {
    map[loss.LossCode] = loss.LossName;
  });
  return map;
}

// 설비의 마지막 상태 확인
function getEquipmentLastStatus(equipmentCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('WorkReport');
    
    if (!sheet) return { status: 'NONE', record: null };
    
    const data = sheet.getDataRange().getValues();
    const scriptTimeZone = Session.getScriptTimeZone();
    
    // 해당 설비의 모든 기록을 시간순으로 정렬
    const equipmentRecords = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[2] === equipmentCode) { // EquipmentCode
        equipmentRecords.push({
          id: row[0],
          date: (row[1] instanceof Date) ? Utilities.formatDate(row[1], scriptTimeZone, "yyyy-MM-dd") : row[1],
          workType: row[5],
          startTime: row[6],
          endTime: row[7],
          lossCode: row[9],
          status: row[16] || 'COMPLETED',
          rowIndex: i + 1
        });
      }
    }
    
    if (equipmentRecords.length === 0) {
      return { status: 'NONE', record: null };
    }
    
    // 가장 최근 기록 찾기 (CreatedAt 기준으로 정렬)
    const lastRecord = equipmentRecords.sort((a, b) => {
      const aCreatedAt = new Date(data[equipmentRecords.findIndex(r => r.id === a.id) + 1][11]);
      const bCreatedAt = new Date(data[equipmentRecords.findIndex(r => r.id === b.id) + 1][11]);
      return bCreatedAt - aCreatedAt;
    })[0];
    
    Logger.log('마지막 기록: ' + JSON.stringify(lastRecord));
    
    return {
      status: lastRecord.status,
      record: lastRecord
    };
    
  } catch (error) {
    Logger.log('getEquipmentLastStatus 오류: ' + error.toString());
    return { status: 'ERROR', record: null };
  }
}

// 근무 스케줄 목록 조회
function getWorkScheduleList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('WorkSchedule');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const schedules = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === 'Y') { // IsActive
        schedules.push({
          ShiftName: data[i][0],
          StartTime: data[i][1],
          EndTime: data[i][2]
        });
      }
    }
    
    return schedules;
  } catch (error) {
    Logger.log('getWorkScheduleList 오류: ' + error.toString());
    return [];
  }
}

// 현재 시간 기준으로 교대 판별
function getCurrentShiftInfo() {
  try {
    const schedules = getWorkScheduleList();
    if (schedules.length === 0) {
      return { ShiftName: '교대 정보 없음', StartTime: '00:00', EndTime: '00:00' };
    }
    
    const now = new Date();
    const currentHour = now.getHours();
    const currentMinute = now.getMinutes();
    const currentTimeMinutes = currentHour * 60 + currentMinute;
    
    for (const schedule of schedules) {
      const startTime = schedule.StartTime.split(':');
      const endTime = schedule.EndTime.split(':');
      const startMinutes = parseInt(startTime[0]) * 60 + parseInt(startTime[1]);
      let endMinutes = parseInt(endTime[0]) * 60 + parseInt(endTime[1]);
      
      // 야간 교대 처리 (종료시간이 시작시간보다 작은 경우)
      if (endMinutes <= startMinutes) {
        endMinutes += 24 * 60; // 다음날로 계산
        
        if (currentTimeMinutes >= startMinutes || currentTimeMinutes < (endMinutes - 24 * 60)) {
          return schedule;
        }
      } else {
        if (currentTimeMinutes >= startMinutes && currentTimeMinutes < endMinutes) {
          return schedule;
        }
      }
    }
    
    // 해당하는 교대가 없으면 첫 번째 교대 반환
    return schedules[0];
    
  } catch (error) {
    Logger.log('getCurrentShiftInfo 오류: ' + error.toString());
    return { ShiftName: '오류', StartTime: '00:00', EndTime: '00:00' };
  }
}

// 이전 교대의 인계사항 조회 (WorkSchedule 기반)
function getPreviousShiftHandoverDynamic(equipmentCode, currentDate) {
  try {
    const currentShift = getCurrentShiftInfo();
    const schedules = getWorkScheduleList();
    
    // 이전 교대 찾기
    let previousShift = null;
    for (let i = 0; i < schedules.length; i++) {
      if (schedules[i].ShiftName === currentShift.ShiftName) {
        previousShift = schedules[i - 1] || schedules[schedules.length - 1]; // 첫 번째면 마지막 교대
        break;
      }
    }
    
    if (!previousShift) return null;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ShiftHandover');
    
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const scriptTimeZone = Session.getScriptTimeZone();
    
    // 해당 설비의 이전 교대 인계사항 찾기 (최근순)
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const handoverDate = (row[0] instanceof Date) 
        ? Utilities.formatDate(row[0], scriptTimeZone, "yyyy-MM-dd")
        : row[0];
      const handoverShiftTime = (row[2] instanceof Date)
        ? Utilities.formatDate(row[2], scriptTimeZone, "HH:mm")
        : row[2];
        
      if (row[1] === equipmentCode && // 같은 설비
          handoverDate === currentDate && // 같은 날짜
          handoverShiftTime === previousShift.StartTime) { // 이전 교대
        
        return {
          date: handoverDate,
          shiftType: row[3],
          workerCode: row[4],
          handoverNote: row[5],
          createdAt: row[6]
        };
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log('getPreviousShiftHandoverDynamic 오류: ' + error.toString());
    return null;
  }
}

// 설비의 미완료 기종교환 체크 (최신 기종교환만 확인)
function checkIncompleteChangeover(equipmentCode, currentDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('WorkReport');
    
    if (!sheet) return false;
    
    const data = sheet.getDataRange().getValues();
    const scriptTimeZone = Session.getScriptTimeZone();
    
    // 해당 설비의 가장 최근 기종교환 기록 찾기
    let latestChangeoverRow = null;
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const reportDate = (row[1] instanceof Date)
        ? Utilities.formatDate(row[1], scriptTimeZone, "yyyy-MM-dd")
        : row[1];
      
      if (row[2] === equipmentCode && // 같은 설비
          row[5] === 'CHANGEOVER') { // 기종교환
        latestChangeoverRow = row;
        break;
      }
    }
    
    // 가장 최근 기종교환이 INCOMPLETE인지 확인
    if (latestChangeoverRow && latestChangeoverRow[16] === 'INCOMPLETE') {
      Logger.log('미완료 기종교환 발견');
      return true;
    }
    
    return false;
    
  } catch (error) {
    Logger.log('checkIncompleteChangeover 오류: ' + error.toString());
    return false;
  }
}

// 현재 제품 목록 상태 확인
function debugCurrentProducts() {
  const products = getProductList();
  Logger.log('=== 현재 제품 목록 ===');
  Logger.log('제품 개수: ' + products.length);
  products.forEach(product => {
    Logger.log('제품코드: ' + product.ProductCode + ', 제품명: ' + product.ProductName);
  });
  
  return products;
}

function testSummaryToday() {
  const today = new Date();
  const result = processDailySummary(today);
  Logger.log(result.message);
}

// 셋팅불량 목록 조회
function getSettingDefectList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('SettingDefectMaster');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const defects = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] === 'Y') { // IsActive
        defects.push({
          DefectCode: data[i][0],
          DefectName: data[i][1],
          DefectType: data[i][2],
          ProcessType: data[i][3]
        });
      }
    }
    
    return defects;
  } catch (error) {
    Logger.log('getSettingDefectList 오류: ' + error.toString());
    return [];
  }
}

// 공정불량 목록 조회
function getProcessDefectList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ProcessDefectMaster');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const defects = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] === 'Y') { // IsActive
        defects.push({
          DefectCode: data[i][0],
          DefectName: data[i][1],
          DefectType: data[i][2],
          ProcessType: data[i][3]
        });
      }
    }
    
    return defects;
  } catch (error) {
    Logger.log('getProcessDefectList 오류: ' + error.toString());
    return [];
  }
}

// 현재 교대의 불량 데이터 조회
function getCurrentShiftDefects(equipmentCode, date) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DefectReport');
    
    if (!sheet) {
      return JSON.stringify([]);
    }
    
    const data = sheet.getDataRange().getValues();
    const scriptTimeZone = Session.getScriptTimeZone();
    const results = [];
    
    // 현재 교대 정보 가져오기
    const currentShift = getCurrentShiftInfo();
    const currentShiftName = currentShift.ShiftName;
    
    // 불량 마스터 매핑
    const settingDefects = getSettingDefectList();
    const processDefects = getProcessDefectList();
    const defectMap = {};
    const defectTypeMap = {}; // 셋팅/공정 구분
    
    settingDefects.forEach(d => {
      defectMap[d.DefectCode] = d.DefectName;
      defectTypeMap[d.DefectCode] = 'setting';
    });
    processDefects.forEach(d => {
      defectMap[d.DefectCode] = d.DefectName;
      defectTypeMap[d.DefectCode] = 'process';
    });
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      const reportDate = (row[1] instanceof Date)
        ? Utilities.formatDate(row[1], scriptTimeZone, "yyyy-MM-dd")
        : row[1];
      
      // 오늘 날짜 + 해당 설비 + 현재 교대만 필터링
      // 오늘 날짜 + 해당 설비만 필터링 (교대 필터 제거)
      if (row[2] === equipmentCode && reportDate === date) {
        
        let productCode = '';
        if (row[4] instanceof Date) {
          const year = row[4].getFullYear();
          const month = row[4].getMonth() + 1;
          const monthStr = month.toString().padStart(2, '0');
          productCode = `${year}-${monthStr}`;
        } else {
          productCode = String(row[4]);
        }
        
        const defectReport = {
          id: row[0],
          date: reportDate,
          equipmentCode: row[2],
          shiftName: row[3],
          productCode: productCode,
          defectCode: row[5],
          defectName: defectMap[row[5]] || row[5],
          defectType: defectTypeMap[row[5]] || 'unknown',
          defectQty: row[6],
          remark: row[7]
        };
        
        results.push(defectReport);
      }
    }
    
    return JSON.stringify(results);
    
  } catch (error) {
    Logger.log('getCurrentShiftDefects 오류: ' + error.toString());
    return JSON.stringify([]);
  }
}

// ===== 불량 등록 함수 =====

// 불량 보고서 등록 (구조 변경)
function addDefectReport(defectData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DefectReport');
    
    if (!sheet) {
      return JSON.stringify({ success: false, message: 'DefectReport 시트를 찾을 수 없습니다.' });
    }
    
    const id = generateId();
    
    // 발생 시간 기준으로 교대 판별
    const shiftName = getShiftByStartTime(defectData.defectTime);
    
    const rowData = [
      id,                           // A: ID
      defectData.date,              // B: Date
      defectData.equipmentCode,     // C: EquipmentCode
      shiftName,                    // D: ShiftName (A조/B조/C조)
      defectData.productCode,       // E: ProductCode
      defectData.defectCode,        // F: DefectCode
      defectData.defectQty,         // G: DefectQty
      defectData.remark || ''       // H: Remark
    ];
    
    sheet.appendRow(rowData);
    
    return JSON.stringify({
      success: true,
      message: '불량이 등록되었습니다.'
    });
    
  } catch (error) {
    Logger.log('addDefectReport 오류: ' + error.toString());
    return JSON.stringify({
      success: false,
      message: '불량 등록 중 오류가 발생했습니다: ' + error.toString()
    });
  }
}

// 불량 보고서 배치 등록 (여러 건을 한 번에)
function addDefectReportBatch(defectDataArray) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DefectReport');
    
    if (!sheet) {
      return JSON.stringify({ 
        success: false, 
        message: 'DefectReport 시트를 찾을 수 없습니다.' 
      });
    }
    
    if (!defectDataArray || defectDataArray.length === 0) {
      return JSON.stringify({ 
        success: false, 
        message: '등록할 불량 데이터가 없습니다.' 
      });
    }
    
    const rows = [];
    
    // 모든 불량 데이터를 행 배열로 변환
    for (let i = 0; i < defectDataArray.length; i++) {
      const defectData = defectDataArray[i];
      const id = generateId();
      const shiftName = getShiftByStartTime(defectData.defectTime);
      
      const rowData = [
        id,                           // A: ID
        defectData.date,              // B: Date
        defectData.equipmentCode,     // C: EquipmentCode
        shiftName,                    // D: ShiftName
        defectData.productCode,       // E: ProductCode
        defectData.defectCode,        // F: DefectCode
        defectData.defectQty,         // G: DefectQty
        defectData.remark || ''       // H: Remark
      ];
      
      rows.push(rowData);
    }
    
    // 한 번에 모든 행 삽입
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
    
    return JSON.stringify({
      success: true,
      message: `${rows.length}건의 불량이 등록되었습니다.`,
      count: rows.length
    });
    
  } catch (error) {
    Logger.log('addDefectReportBatch 오류: ' + error.toString());
    return JSON.stringify({
      success: false,
      message: '불량 일괄 등록 중 오류가 발생했습니다: ' + error.toString()
    });
  }
}


// 불량 보고서 조회 (구조 변경)
function getDefectReportByEquipment(equipmentCode, date) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DefectReport');
    
    if (!sheet) {
      return JSON.stringify([]);
    }
    
    const data = sheet.getDataRange().getValues();
    const scriptTimeZone = Session.getScriptTimeZone();
    const results = [];
    
    // 5일 전 날짜 계산
    const targetDate = new Date(date);
    const fiveDaysAgo = new Date(targetDate);
    fiveDaysAgo.setDate(targetDate.getDate() - 1);
    const fiveDaysAgoStr = Utilities.formatDate(fiveDaysAgo, scriptTimeZone, "yyyy-MM-dd");
    
    // 불량 마스터 매핑
    const settingDefects = getSettingDefectList();
    const processDefects = getProcessDefectList();
    const defectMap = {};
    
    settingDefects.forEach(d => {
      defectMap[d.DefectCode] = d.DefectName;
    });
    processDefects.forEach(d => {
      defectMap[d.DefectCode] = d.DefectName;
    });
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      const reportDate = (row[1] instanceof Date)
        ? Utilities.formatDate(row[1], scriptTimeZone, "yyyy-MM-dd")
        : row[1];
      
      // 설비코드 및 날짜 필터 적용
      if (row[2] === equipmentCode && reportDate >= fiveDaysAgoStr && reportDate <= date) {
        
        // 제품코드 안전 처리
        let productCode = '';
        if (row[4] instanceof Date) {
          const year = row[4].getFullYear();
          const month = row[4].getMonth() + 1;
          const monthStr = month.toString().padStart(2, '0');
          productCode = `${year}-${monthStr}`;
        } else {
          productCode = String(row[4]);
        }

        const defectReport = {
          id: row[0],
          date: reportDate,
          equipmentCode: row[2],
          shiftName: row[3],
          productCode: productCode,
          defectCode: row[5],
          defectName: defectMap[row[5]] || row[5],
          defectQty: row[6],
          remark: row[7]
        };
        
        results.push(defectReport);
      }
    }
    
    // 최신순 정렬
    results.sort((a, b) => {
      if (a.date !== b.date) {
        return b.date.localeCompare(a.date);
      }
      return 0;
    });
    
    return JSON.stringify(results);
    
  } catch (error) {
    Logger.log('getDefectReportByEquipment 오류: ' + error.toString());
    return JSON.stringify([]);
  }
}

// 과거 ONGOING 존재 여부 및 목록 조회
function getStaleOngoingInfo(equipmentCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('WorkReport');
    if (!sheet) return JSON.stringify({ hasStale: false, items: [] });

    const data = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    const items = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[2] !== equipmentCode) continue;
      const status = row[16];
      if (status !== 'ONGOING') continue;

      const rowDateStr = (row[1] instanceof Date)
        ? Utilities.formatDate(row[1], tz, 'yyyy-MM-dd')
        : row[1];

      // 오늘 이전 날짜의 ONGOING을 '과거 ONGOING'으로 간주
      if (rowDateStr < todayStr) {
        items.push({
          id: row[0],
          date: rowDateStr,
          workType: row[5],
          startTime: row[6],
          productCode: row[4] || ''
        });
      }
    }

    return JSON.stringify({ hasStale: items.length > 0, items });
  } catch (err) {
    Logger.log('getStaleOngoingInfo 오류: ' + err);
    return JSON.stringify({ hasStale: false, items: [], error: err.toString() });
  }
}

// 과거 ONGOING 일괄 정리
function resolveStaleOngoing(equipmentCode, nowISOString) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('WorkReport');
    if (!sheet) return JSON.stringify({ success: false, message: 'WorkReport 시트를 찾을 수 없습니다.' });

    const data = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(nowISOString || new Date()), tz, 'yyyy-MM-dd');
    const endTimeFormatted = formatUnifiedTime(nowISOString ? new Date(nowISOString) : new Date());

    let updated = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[2] !== equipmentCode) continue;
      if (row[16] !== 'ONGOING') continue;

      const rowDateStr = (row[1] instanceof Date)
        ? Utilities.formatDate(row[1], tz, 'yyyy-MM-dd')
        : row[1];

      if (rowDateStr < todayStr) {
        const rowIndex = i + 1;
        // EndTime 채우기
        sheet.getRange(rowIndex, 8).setValue(endTimeFormatted);
        // 상태 전환: WORK -> COMPLETED, CHANGEOVER -> INCOMPLETE
        if (row[5] === 'WORK') {
          sheet.getRange(rowIndex, 17).setValue('COMPLETED');
          // 수량은 변경하지 않음(미입력 상태면 0 그대로)
        } else if (row[5] === 'CHANGEOVER') {
          sheet.getRange(rowIndex, 17).setValue('INCOMPLETE');
        } else {
          // 기타 타입은 보수적으로 COMPLETED 처리
          sheet.getRange(rowIndex, 17).setValue('COMPLETED');
        }
        updated++;
      }
    }

    return JSON.stringify({ success: true, message: `과거 ONGOING ${updated}건 정리 완료`, updated });
  } catch (err) {
    Logger.log('resolveStaleOngoing 오류: ' + err);
    return JSON.stringify({ success: false, message: '정리 중 오류: ' + err.toString() });
  }
}
