/**
 * AI CAMP 웹사이트 진단 설문 및 상담신청 데이터 수집용 Google Apps Script
 * 설문조사 시트: 13kcd9Hs1BGaNosfm3hjyd3KIIsJ7Jc-ib4Jawcvwa1Q
 * 상담신청 시트: 1LQNeT0abhMHXktrNjRbxl2XEFWVCwcYr5kVTAcRvpfM
 */

// 스프레드시트 ID 설정
const SURVEY_SPREADSHEET_ID = '13kcd9Hs1BGaNosfm3hjyd3KIIsJ7Jc-ib4Jawcvwa1Q';
const CONSULTATION_SPREADSHEET_ID = '1LQNeT0abhMHXktrNjRbxl2XEFWVCwcYr5kVTAcRvpfM';

function doPost(e) {
  try {
    // 요청 데이터 파싱
    const requestData = JSON.parse(e.postData.contents);
    
    // 요청 타입 확인
    if (requestData.type === 'consultation') {
      return handleConsultationRequest(requestData.data);
    } else {
      // 기존 설문 데이터 처리
      return handleSurveyRequest(requestData);
    }
  } catch (error) {
    // 오류 응답
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: '데이터 저장 중 오류가 발생했습니다: ' + error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 설문 데이터 처리 함수 (별도 스프레드시트)
function handleSurveyRequest(data) {
  // 설문조사 전용 스프레드시트 접근
  const spreadsheet = SpreadsheetApp.openById(SURVEY_SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName('설문조사');
  
  // 설문조사 시트가 없으면 생성
  if (!sheet) {
    sheet = spreadsheet.insertSheet('설문조사');
    const headers = [
      '접수일시', '진단유형', '기업명', '담당자명', '지역', '전화번호', '이메일',
      '총점', '진단등급', '백분율', '질문1', '답변1', '질문2', '답변2', 
      '질문3', '답변3', '질문4', '답변4', '질문5', '답변5', 
      '질문6', '답변6', '질문7', '답변7', '질문8', '답변8', 
      '질문9', '답변9', '질문10', '답변10', '추가의견', '개인정보동의', '마케팅동의'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 헤더 스타일링
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2d5a27')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
      
    // 열 너비 자동 조정
    sheet.autoResizeColumns(1, headers.length);
  }
  
  // 등급 계산
  let grade = '';
  if (data.percentage >= 90) grade = '매우 우수';
  else if (data.percentage >= 70) grade = '우수';
  else if (data.percentage >= 50) grade = '보통';
  else if (data.percentage >= 30) grade = '미흡';
  else grade = '매우 미흡';
  
  // 데이터 행 구성
  const rowData = [
    data.timestamp,
    data.serviceType,
    data.companyInfo.companyName,
    data.companyInfo.contactName,
    data.companyInfo.region,
    data.companyInfo.phoneNumber,
    data.companyInfo.email,
    data.totalScore,
    grade,
    data.percentage + '%'
  ];
  
  // 질문별 답변 추가 (최대 10개)
  for (let i = 0; i < 10; i++) {
    if (i < data.questions.length) {
      rowData.push(data.questions[i].question);
      rowData.push(data.questions[i].answer);
    } else {
      rowData.push('');
      rowData.push('');
    }
  }
  
  // 추가 의견
  rowData.push(data.additionalComments);
  
  // 개인정보 동의 정보 추가
  rowData.push(data.privacyConsent || 'Y');
  rowData.push(data.marketingConsent || 'N');
  
  // 시트에 데이터 추가
  sheet.appendRow(rowData);
  
  // 방금 추가된 행 스타일링
  const lastRow = sheet.getLastRow();
  
  // 등급별 색상 적용
  let gradeColor = '';
  if (grade === '매우 우수') gradeColor = '#d4edda';
  else if (grade === '우수') gradeColor = '#d1ecf1';
  else if (grade === '보통') gradeColor = '#fff3cd';
  else if (grade === '미흡') gradeColor = '#f8d7da';
  else gradeColor = '#f5c6cb';
  
  sheet.getRange(lastRow, 9, 1, 1).setBackground(gradeColor); // 등급 칸 색상
  
  // 데이터 행에 테두리 적용
  sheet.getRange(lastRow, 1, 1, sheet.getLastColumn())
    .setBorder(true, true, true, true, true, true);
  
  // 성공 응답
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      message: '설문 데이터가 설문조사 전용 시트에 성공적으로 저장되었습니다.',
      timestamp: data.timestamp,
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${SURVEY_SPREADSHEET_ID}/edit`
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 상담신청 데이터 처리 함수 (별도 스프레드시트)
function handleConsultationRequest(data) {
  // 상담신청 전용 스프레드시트 접근
  const spreadsheet = SpreadsheetApp.openById(CONSULTATION_SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName('상담신청');
  
  // 상담신청 시트가 없으면 생성
  if (!sheet) {
    sheet = spreadsheet.insertSheet('상담신청');
    const headers = [
      '접수일시', '상담유형', '연락방법', '연락정보', '기업명', '담당자명', 
      '전화번호', '이메일', '상담분야', '시급성', '추가요청사항', 
      '개인정보동의', '마케팅동의', '참조URL', 'UserAgent'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // 헤더 스타일링
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4CAF50')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
      
    // 열 너비 자동 조정
    sheet.autoResizeColumns(1, headers.length);
  }
  
  // 데이터 행 구성
  const rowData = [
    data.timestamp,
    data.consultationType,
    data.contactMethod,
    data.contactInfo,
    data.companyName,
    data.contactName,
    data.phoneNumber,
    data.email,
    data.consultationArea,
    data.urgency,
    data.additionalRequest,
    data.privacyConsent || 'Y',
    data.marketingConsent || 'N',
    data.referenceUrl,
    data.userAgent
  ];
  
  // 시트에 데이터 추가
  sheet.appendRow(rowData);
  
  // 방금 추가된 행 스타일링
  const lastRow = sheet.getLastRow();
  
  // 시급성별 색상 적용
  let urgencyColor = '';
  if (data.urgency === '매우급함') urgencyColor = '#ffebee';
  else if (data.urgency === '급함') urgencyColor = '#fff3e0';
  else if (data.urgency === '보통') urgencyColor = '#f3e5f5';
  else if (data.urgency === '여유있음') urgencyColor = '#e8f5e8';
  else urgencyColor = '#e1f5fe';
  
  sheet.getRange(lastRow, 10, 1, 1).setBackground(urgencyColor); // 시급성 칸 색상
  
  // 데이터 행에 테두리 적용
  sheet.getRange(lastRow, 1, 1, sheet.getLastColumn())
    .setBorder(true, true, true, true, true, true);
  
  // 성공 응답
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      message: '상담신청이 상담신청 전용 시트에 성공적으로 저장되었습니다.',
      timestamp: data.timestamp,
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${CONSULTATION_SPREADSHEET_ID}/edit`
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // GET 요청용 (테스트)
  return ContentService
    .createTextOutput('AI CAMP 진단 설문 및 상담신청 데이터 수집 API가 정상 작동 중입니다.\n\n설문조사 시트: ' + SURVEY_SPREADSHEET_ID + '\n상담신청 시트: ' + CONSULTATION_SPREADSHEET_ID)
    .setMimeType(ContentService.MimeType.TEXT);
} 