// 설문 데이터 처리용 Google Apps Script
// 구글시트: https://docs.google.com/spreadsheets/d/13kcd9Hs1BGaNosfm3hjyd3KIIsJ7Jc-ib4Jawcvwa1Q/edit?usp=sharing

// 기본 실행 함수 (Google Apps Script 오류 방지)
function myFunction() {
  console.log('📋 AI CAMP 설문 시스템이 정상 작동 중입니다.');
  console.log('📅 현재 시간:', new Date().toLocaleString('ko-KR'));
  
  // 시스템 테스트 실행
  console.log('🧪 시스템 테스트를 실행합니다...');
  const testResult = testPost();
  
  console.log('✅ 시스템 테스트 완료');
  return 'AI CAMP 설문 시스템 - 정상 작동 중 (테스트 완료)';
}

function doPost(e) {
  try {
    console.log('📥 POST 요청 수신');
    console.log('📋 전체 이벤트 객체:', e);
    
    // 이벤트 객체 검증 및 테스트 모드 대응
    if (!e) {
      console.log('⚠️ 이벤트 객체가 없습니다 - 테스트 모드로 실행합니다');
      // 테스트 데이터로 실행
      return testPost();
    }
    
    // postData 검증
    if (!e.postData) {
      console.log('⚠️ postData가 없습니다');
      console.log('📋 사용 가능한 속성들:', Object.keys(e));
      console.log('🔄 테스트 모드로 전환합니다');
      return testPost();
    }
    
    // contents 검증
    if (!e.postData.contents) {
      console.log('⚠️ postData.contents가 없습니다');
      console.log('📋 postData 속성들:', Object.keys(e.postData));
      console.log('🔄 테스트 모드로 전환합니다');
      return testPost();
    }
    
    console.log('📥 설문 데이터 수신:', e.postData.contents);
    
    // JSON 파싱 시도
    let requestData;
    try {
      requestData = JSON.parse(e.postData.contents);
    } catch (parseError) {
      console.error('❌ JSON 파싱 오류:', parseError);
      console.log('📋 원본 데이터:', e.postData.contents);
      return createErrorResponse('JSON 파싱 오류: ' + parseError.toString());
    }
    
    console.log('📋 파싱된 데이터:', requestData);
    
    // survey 타입인지 확인
    if (requestData.type !== 'survey') {
      console.log('❌ 잘못된 요청 타입:', requestData.type);
      return createErrorResponse('잘못된 요청 타입입니다: ' + requestData.type);
    }
    
    const data = requestData.data || {};
    console.log('📊 설문 데이터:', data);
    
    // 스프레드시트 처리
    const result = saveToSheet(data);
    
    if (result.success) {
      console.log('✅ 설문 데이터 저장 완료');
      return createSuccessResponse('설문 데이터가 성공적으로 저장되었습니다.');
    } else {
      console.error('❌ 시트 저장 실패:', result.error);
      return createErrorResponse('데이터 저장 실패: ' + result.error);
    }
      
  } catch (error) {
    console.error('❌ 설문 데이터 처리 오류:', error);
    console.error('오류 스택:', error.stack);
    return createErrorResponse('처리 중 오류 발생: ' + error.toString());
  }
}

// 스프레드시트에 데이터 저장
function saveToSheet(data) {
  try {
    // 스프레드시트 열기
    const spreadsheet = SpreadsheetApp.openById('13kcd9Hs1BGaNosfm3hjyd3KIIsJ7Jc-ib4Jawcvwa1Q');
    let sheet = spreadsheet.getSheetByName('설문결과');
    
    // 시트가 없으면 생성
    if (!sheet) {
      sheet = spreadsheet.insertSheet('설문결과');
      // 헤더 추가
      const headers = [
        '제출일시', '서비스타입', '기업명', '담당자명', '지역', '전화번호', '이메일',
        '총점수', '백분율', '등급', '추가의견', '개인정보동의', '마케팅동의'
      ];
      
      // 답변 항목들을 동적으로 추가 (최대 15개 질문 대응)
      for (let i = 1; i <= 15; i++) {
        headers.push(`답변${i}`);
      }
      
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // 헤더 스타일링
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#4285F4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
    }
    
    // 등급 계산
    const percentage = data.percentage || 0;
    let grade = '';
    if (percentage >= 90) grade = '매우우수';
    else if (percentage >= 70) grade = '우수';
    else if (percentage >= 50) grade = '보통';
    else if (percentage >= 30) grade = '미흡';
    else grade = '매우미흡';
    
    // 기본 데이터 행 구성
    const row = [
      data.timestamp || new Date().toLocaleString('ko-KR'),
      data.serviceType || '',
      (data.companyInfo && data.companyInfo.companyName) || '',
      (data.companyInfo && data.companyInfo.contactName) || '',
      (data.companyInfo && data.companyInfo.region) || '',
      (data.companyInfo && data.companyInfo.phoneNumber) || '',
      (data.companyInfo && data.companyInfo.email) || '',
      data.totalScore || 0,
      percentage,
      grade,
      data.additionalComments || '',
      data.privacyConsent || 'N',
      data.marketingConsent || 'N'
    ];
    
    // 답변 점수들 추가 (최대 15개)
    const answers = data.answers || [];
    for (let i = 0; i < 15; i++) {
      row.push(answers[i] || '');
    }
    
    sheet.appendRow(row);
    
    // 질문과 답변을 별도 시트에 상세 저장
    if (data.questions && data.questions.length > 0) {
      saveQuestionDetails(spreadsheet, data);
    }
    
    return { success: true };
    
  } catch (error) {
    console.error('❌ 시트 저장 오류:', error);
    return { success: false, error: error.toString() };
  }
}

// 질문과 답변 상세 저장
function saveQuestionDetails(spreadsheet, data) {
  try {
    let detailSheet = spreadsheet.getSheetByName('설문상세');
    
    if (!detailSheet) {
      detailSheet = spreadsheet.insertSheet('설문상세');
      const headers = ['제출일시', '기업명', '서비스타입', '질문번호', '질문내용', '답변점수'];
      detailSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // 헤더 스타일링
      const headerRange = detailSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#FF9800');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
    }
    
    const timestamp = data.timestamp || new Date().toLocaleString('ko-KR');
    const companyName = (data.companyInfo && data.companyInfo.companyName) || '';
    const serviceType = data.serviceType || '';
    
    data.questions.forEach((item, index) => {
      const row = [
        timestamp,
        companyName,
        serviceType,
        index + 1,
        item.question || '',
        item.answer || 0
      ];
      detailSheet.appendRow(row);
    });
    
    console.log('📋 설문 상세 정보 저장 완료');
  } catch (error) {
    console.error('❌ 설문 상세 저장 오류:', error);
  }
}

// 성공 응답 생성
function createSuccessResponse(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: true, 
      message: message,
      timestamp: new Date().toLocaleString('ko-KR')
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 오류 응답 생성
function createErrorResponse(errorMessage) {
  return ContentService
    .createTextOutput(JSON.stringify({ 
      success: false, 
      error: errorMessage,
      timestamp: new Date().toLocaleString('ko-KR')
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET 요청 처리 (테스트용)
function doGet(e) {
  try {
    console.log('📥 GET 요청 수신');
    console.log('📋 파라미터:', e ? e.parameter : 'e 객체 없음');
    
    return ContentService
      .createTextOutput('AI CAMP 설문 API가 정상 작동 중입니다. 현재 시간: ' + new Date().toLocaleString('ko-KR'))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    console.error('❌ GET 요청 처리 오류:', error);
    return ContentService
      .createTextOutput('GET 요청 처리 중 오류 발생: ' + error.toString())
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// 테스트 함수
function test() {
  console.log('📋 테스트 함수 실행');
  return 'Apps Script가 정상 작동 중입니다.';
}

// 수동 테스트용 함수
function testPost() {
  try {
    console.log('🧪 테스트 모드 실행');
    
    const testData = {
      timestamp: new Date().toLocaleString('ko-KR'),
      serviceType: '테스트 진단',
      companyInfo: {
        companyName: '테스트 회사',
        contactName: '테스트 담당자',
        region: '서울특별시',
        phoneNumber: '010-1234-5678',
        email: 'test@test.com'
      },
      answers: [4, 3, 5, 2, 4, 3, 5],
      totalScore: 26,
      percentage: 74,
      additionalComments: '테스트 의견',
      privacyConsent: 'Y',
      marketingConsent: 'N',
      questions: [
        { question: '테스트 질문 1', answer: 4 },
        { question: '테스트 질문 2', answer: 3 },
        { question: '테스트 질문 3', answer: 5 }
      ]
    };
    
    console.log('📊 테스트 데이터:', testData);
    
    // 스프레드시트에 직접 저장
    const result = saveToSheet(testData);
    
    if (result.success) {
      console.log('✅ 테스트 데이터 저장 완료');
      return createSuccessResponse('테스트 데이터가 성공적으로 저장되었습니다.');
    } else {
      console.error('❌ 테스트 데이터 저장 실패:', result.error);
      return createErrorResponse('테스트 데이터 저장 실패: ' + result.error);
    }
    
  } catch (error) {
    console.error('❌ 테스트 실행 오류:', error);
    return createErrorResponse('테스트 실행 중 오류 발생: ' + error.toString());
  }
} 