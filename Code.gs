/**
 * 우리 반 독서 기록장 - Google Apps Script 백엔드
 *
 * 이 파일을 구글 시트의 [확장 프로그램 > Apps Script]에 붙여넣고
 * "웹 앱으로 배포"하면 학생용 HTML 앱에서 이 시트로 데이터를 저장/조회합니다.
 *
 * 시트 구성 (자동 생성됨, 미리 만들 필요 없음)
 *  - "기록"     : 학생이 저장 버튼을 누를 때마다 한 줄씩 쌓이는 원본 기록
 *  - "누적현황" : 학생별 누적 독서시간을 자동으로 요약해주는 표 (기록이 추가될 때마다 갱신)
 */

const SHEET_LOG = '기록';
const SHEET_SUMMARY = '누적현황';

// 학생 앱이 "읽기 종료 → 저장"을 누르면 호출되는 함수 (기록 1건 추가)
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (!data.studentId || !data.name || !data.minutes) {
      return jsonOutput_({ ok: false, error: '필수 값 누락 (studentId/name/minutes)' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_LOG);
      sheet.appendRow(['날짜', '학번', '이름', '시작시각', '종료시각', '읽은시간(분)', '기록시각']);
      sheet.setFrozenRows(1);
    }

    const now = new Date();
    sheet.appendRow([
      data.date || Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd'),
      data.studentId,
      data.name,
      data.startTime || '',
      data.endTime || '',
      Number(data.minutes),
      Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
    ]);

    updateSummary_();

    return jsonOutput_({ ok: true });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  }
}

// 학생 앱이 로그인 직후 "나의 누적 독서시간"을 물어볼 때 호출되는 함수
function doGet(e) {
  try {
    const studentId = (e.parameter.studentId || '').toString();
    const name = (e.parameter.name || '').toString();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LOG);
    if (!sheet || sheet.getLastRow() < 2) {
      return jsonOutput_({ ok: true, totalMinutes: 0, sessionCount: 0, days: [] });
    }

    const values = sheet.getDataRange().getValues();
    let totalMinutes = 0;
    let sessionCount = 0;
    const dayMap = {};

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (String(row[1]) === studentId && String(row[2]) === name) {
        const mins = Number(row[5]) || 0;
        totalMinutes += mins;
        sessionCount++;
        const dayKey = formatDateCell_(row[0]);
        dayMap[dayKey] = (dayMap[dayKey] || 0) + mins;
      }
    }

    const days = Object.keys(dayMap).sort().map(function (d) {
      return { date: d, minutes: dayMap[d] };
    });

    return jsonOutput_({ ok: true, totalMinutes: totalMinutes, sessionCount: sessionCount, days: days });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err) });
  }
}

// "기록" 시트를 훑어서 "누적현황" 시트를 새로 만듦 (교사가 시트를 열어 바로 확인 가능)
function updateSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_LOG);
  let sumSheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sumSheet) sumSheet = ss.insertSheet(SHEET_SUMMARY);

  const values = logSheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const key = row[1] + '__' + row[2];
    if (!map[key]) map[key] = { id: row[1], name: row[2], minutes: 0, count: 0 };
    map[key].minutes += Number(row[5]) || 0;
    map[key].count++;
  }

  sumSheet.clearContents();
  sumSheet.appendRow(['학번', '이름', '누적 읽은시간(분)', '누적 시간(시:분)', '누적 횟수']);
  sumSheet.setFrozenRows(1);

  const rows = Object.keys(map).map(function (k) { return map[k]; })
    .sort(function (a, b) { return String(a.id).localeCompare(String(b.id), 'ko'); });

  rows.forEach(function (r) {
    const h = Math.floor(r.minutes / 60);
    const m = r.minutes % 60;
    sumSheet.appendRow([r.id, r.name, r.minutes, h + '시간 ' + m + '분', r.count]);
  });
}

function formatDateCell_(cell) {
  if (Object.prototype.toString.call(cell) === '[object Date]') {
    return Utilities.formatDate(cell, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  return String(cell);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
