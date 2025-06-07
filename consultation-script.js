// 상담신청 데이터 처리용 Google Apps Script (안정화 버전 v3.0)
// 구글시트: https://docs.google.com/spreadsheets/d/1LQNeT0abhMHXktrNjRbxl2XEFWVCwcYr5kVTAcRvpfM/edit?usp=sharing

// GET 요청 처리 (테스트용)
function doGet(e) {
  console.log('📥 GET 요청 수신');
  console.log('📋 GET 이벤트 객체:', typeof e, e ? 'exists' : 'undefined');
  
  const response = {
    status: 'success',
    message: 'AI CAMP 상담신청 API가 정상 작동 중입니다.',
    timestamp: new Date().toLocaleString('ko-KR'),
    version: '3.0',
    test_info: {
      event_object: typeof e,
      has_parameters: e && e.parameter ? Object.keys(e.parameter) : 'none'
    }
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST 요청 처리 (실제 데이터 처리) - 강화된 버전
function doPost(e) {
  // 실행 시작 로그
  console.log('🚀 doPost 함수 실행 시작');
  console.log('📅 실행 시간:', new Date().toLocaleString('ko-KR'));
  
  try {
    // 이벤트 객체 상세 검사
    console.log('🔍 이벤트 객체 분석:');
    console.log('  - typeof e:', typeof e);
    console.log('  - e === null:', e === null);
    console.log('  - e === undefined:', e === undefined);
    console.log('  - Boolean(e):', Boolean(e));
    
    if (e) {
      console.log('  - Object.keys(e):', Object.keys(e));
      console.log('  - e.toString():', e.toString());
    }
    
    // 이벤트 객체가 없는 경우의 처리
    if (!e || typeof e === 'undefined') {
      console.error('❌ 이벤트 객체가 전달되지 않았습니다');
      console.log('🔧 문제해결 가이드:');
      console.log('  1. 웹앱을 새로 배포해주세요');
      console.log('  2. 권한을 다시 승인해주세요');
      console.log('  3. 스크립트를 저장한 후 재배포해주세요');
      
      return createSafeResponse({
        success: false,
        error: 'EVENT_OBJECT_MISSING',
        message: '이벤트 객체가 전달되지 않았습니다.',
        troubleshooting: {
          step1: '스크립트 편집기에서 Ctrl+S로 저장',
          step2: '배포 > 배포 관리 > 새 배포 생성',
          step3: '유형: 웹 앱, 실행 권한: 본인, 액세스: 모든 사용자',
          step4: '권한 승인 후 새로운 URL로 테스트',
          note: '이 오류는 Google Apps Script 서비스 일시적 문제일 수 있습니다.'
        }
      });
    }
    
    // postData 검사
    console.log('📋 postData 검사:');
    if (!e.postData) {
      console.error('❌ e.postData가 없습니다');
      console.log('📋 이벤트 객체 속성들:', Object.keys(e));
      
      return createSafeResponse({
        success: false,
        error: 'POSTDATA_MISSING',
        message: 'POST 데이터가 없습니다.',
        received_properties: Object.keys(e),
        help: 'fetch 요청에서 method: POST와 body를 포함해서 전송해주세요.'
      });
    }
    
    console.log('  - typeof e.postData:', typeof e.postData);
    console.log('  - postData keys:', Object.keys(e.postData));
    
    // contents 검사
    if (!e.postData.contents) {
      console.error('❌ e.postData.contents가 없습니다');
      console.log('📋 postData 속성들:', Object.keys(e.postData));
      
      return createSafeResponse({
        success: false,
        error: 'POSTDATA_CONTENTS_MISSING',
        message: 'POST 데이터 내용이 없습니다.',
        postData_properties: Object.keys(e.postData),
        help: 'body에 JSON 데이터를 포함해서 전송해주세요.'
      });
    }
    
    console.log('📥 POST 데이터 수신 성공');
    console.log('📝 데이터 내용 (처음 100자):', e.postData.contents.substring(0, 100));
    
    // JSON 파싱
    let requestData;
    try {
      requestData = JSON.parse(e.postData.contents);
      console.log('✅ JSON 파싱 성공');
      console.log('📊 요청 데이터 구조:', Object.keys(requestData));
    } catch (parseError) {
      console.error('❌ JSON 파싱 실패:', parseError.toString());
      console.log('📋 원본 데이터:', e.postData.contents);
      
      return createSafeResponse({
        success: false,
        error: 'JSON_PARSE_ERROR',
        message: 'JSON 파싱에 실패했습니다.',
        parse_error: parseError.toString(),
        received_data: e.postData.contents.substring(0, 200)
      });
    }
    
    // 요청 타입 확인
    if (!requestData.type) {
      console.error('❌ 요청 타입이 없습니다');
      return createSafeResponse({
        success: false,
        error: 'REQUEST_TYPE_MISSING',
        message: '요청 타입이 지정되지 않았습니다.',
        received_data: requestData
      });
    }
    
    if (requestData.type !== 'consultation') {
      console.log('⚠️ 다른 요청 타입:', requestData.type);
      return createSafeResponse({
        success: false,
        error: 'INVALID_REQUEST_TYPE',
        message: '잘못된 요청 타입입니다.',
        expected: 'consultation',
        received: requestData.type
      });
    }
    
    console.log('✅ 요청 타입 확인 완료: consultation');
    
    // 데이터 추출
    const data = requestData.data;
    if (!data) {
      console.error('❌ 요청 데이터가 없습니다');
      return createSafeResponse({
        success: false,
        error: 'REQUEST_DATA_MISSING',
        message: '요청 데이터가 없습니다.'
      });
    }
    
    console.log('📋 상담신청 데이터 처리 시작');
    console.log('🏢 기업명:', data.companyName || 'N/A');
    console.log('👤 담당자:', data.contactName || 'N/A');
    
    // 스프레드시트 처리
    const result = processConsultationDataSafely(data);
    
    console.log('✅ 처리 완료:', result.success ? '성공' : '실패');
    return createSafeResponse(result);
    
  } catch (error) {
    console.error('❌ doPost 전체 처리 오류:', error.toString());
    console.error('📋 오류 스택:', error.stack);
    
    return createSafeResponse({
      success: false,
      error: 'UNEXPECTED_ERROR',
      message: '예상치 못한 오류가 발생했습니다.',
      error_details: error.toString(),
      error_stack: error.stack,
      timestamp: new Date().toLocaleString('ko-KR')
    });
  }
}

// 안전한 응답 생성 함수
function createSafeResponse(data) {
  try {
    const response = {
      ...data,
      timestamp: data.timestamp || new Date().toLocaleString('ko-KR'),
      server_info: {
        script_version: '3.0',
        execution_time: new Date().toISOString()
      }
    };
    
    return ContentService
      .createTextOutput(JSON.stringify(response, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error('❌ 응답 생성 오류:', error);
    
    // 최후의 안전장치
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: 'RESPONSE_CREATION_ERROR',
        message: '응답 생성 중 오류가 발생했습니다.',
        original_error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 안전한 상담신청 데이터 처리 함수
function processConsultationDataSafely(data) {
  try {
    console.log('📊 구글시트 접근 시작');
    
    // 스프레드시트 ID 확인
    const SPREADSHEET_ID = '1LQNeT0abhMHXktrNjRbxl2XEFWVCwcYr5kVTAcRvpfM';
    console.log('📋 스프레드시트 ID:', SPREADSHEET_ID);
    
    // 스프레드시트 열기
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
      console.log('✅ 스프레드시트 접근 성공');
      console.log('📋 스프레드시트 이름:', spreadsheet.getName());
    } catch (spreadsheetError) {
      console.error('❌ 스프레드시트 접근 실패:', spreadsheetError.toString());
      return {
        success: false,
        error: 'SPREADSHEET_ACCESS_ERROR',
        message: '스프레드시트에 접근할 수 없습니다.',
        spreadsheet_id: SPREADSHEET_ID,
        error_details: spreadsheetError.toString()
      };
    }
    
    // 시트 찾기 또는 생성
    const SHEET_NAME = '상담신청';
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      console.log('📋 시트가 없어서 새로 생성:', SHEET_NAME);
      try {
        sheet = spreadsheet.insertSheet(SHEET_NAME);
        
        // 헤더 행 생성
        const headers = [
          '접수일시', '상담유형', '연락방법', '연락처정보',
          '기업명', '담당자명', '전화번호', '이메일',
          '상담분야', '시급성', '추가요청사항', '첨부파일',
          '개인정보동의', '마케팅동의', '참조URL', '브라우저정보'
        ];
        
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        // 헤더 스타일 적용
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground('#2d5a27');
        headerRange.setFontColor('#ffffff');
        headerRange.setFontWeight('bold');
        headerRange.setHorizontalAlignment('center');
        
        console.log('✅ 새 시트 생성 및 헤더 설정 완료');
      } catch (sheetError) {
        console.error('❌ 시트 생성 실패:', sheetError.toString());
        return {
          success: false,
          error: 'SHEET_CREATION_ERROR',
          message: '시트 생성에 실패했습니다.',
          error_details: sheetError.toString()
        };
      }
    } else {
      console.log('✅ 기존 시트 찾음:', SHEET_NAME);
    }
    
    // 첨부파일 정보 처리
    let attachmentInfo = '';
    if (data.attachments && Array.isArray(data.attachments) && data.attachments.length > 0) {
      attachmentInfo = data.attachments.map(file => 
        `${file.name} (${formatFileSize(file.size)})`
      ).join(', ');
      console.log('📎 첨부파일 정보:', attachmentInfo);
    }
    
    // 데이터 행 준비
    const currentTime = new Date().toLocaleString('ko-KR');
    const rowData = [
      data.timestamp || currentTime,
      data.consultationType || '상담신청',
      data.contactMethod || '온라인폼',
      data.contactInfo || '',
      data.companyName || '',
      data.contactName || '',
      data.phoneNumber || '',
      data.email || '',
      data.consultationArea || '',
      data.urgency || '',
      data.additionalRequest || '',
      attachmentInfo,
      data.privacyConsent || 'N',
      data.marketingConsent || 'N',
      data.referenceUrl || '',
      data.userAgent || ''
    ];
    
    console.log('📝 저장할 데이터 준비 완료');
    console.log('📊 데이터 개수:', rowData.length);
    
    // 다음 빈 행에 데이터 추가
    try {
      const lastRow = sheet.getLastRow();
      const newRow = lastRow + 1;
      
      sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
      
      // 새로 추가된 행 스타일 적용
      const newRowRange = sheet.getRange(newRow, 1, 1, rowData.length);
      newRowRange.setBorder(true, true, true, true, true, true);
      
      console.log('✅ 구글시트 저장 완료');
      console.log('📍 저장된 행 번호:', newRow);
      
      return {
        success: true,
        message: '상담신청이 성공적으로 저장되었습니다.',
        details: {
          sheet_name: SHEET_NAME,
          row_number: newRow,
          data_count: rowData.length,
          company_name: data.companyName,
          contact_name: data.contactName,
          has_attachments: attachmentInfo ? true : false
        },
        timestamp: currentTime
      };
      
    } catch (saveError) {
      console.error('❌ 데이터 저장 실패:', saveError.toString());
      return {
        success: false,
        error: 'DATA_SAVE_ERROR',
        message: '데이터 저장 중 오류가 발생했습니다.',
        error_details: saveError.toString()
      };
    }
    
  } catch (error) {
    console.error('❌ processConsultationDataSafely 오류:', error.toString());
    console.error('📋 오류 스택:', error.stack);
    
    return {
      success: false,
      error: 'PROCESS_ERROR',
      message: '데이터 처리 중 오류가 발생했습니다.',
      error_details: error.toString(),
      error_stack: error.stack
    };
  }
}

// 파일 크기 포맷 함수
function formatFileSize(bytes) {
  if (!bytes || bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 수동 테스트 함수
function testConsultationManually() {
  console.log('🧪 수동 테스트 시작');
  
  const testData = {
    timestamp: new Date().toLocaleString('ko-KR'),
    consultationType: '테스트상담',
    contactMethod: '수동테스트',
    companyName: '테스트회사',
    contactName: '홍길동',
    phoneNumber: '010-1234-5678',
    email: 'test@example.com',
    consultationArea: 'AI생산성혁신',
    urgency: '보통',
    additionalRequest: '테스트 요청사항입니다.',
    privacyConsent: 'Y',
    marketingConsent: 'N'
  };
  
  try {
    const result = processConsultationDataSafely(testData);
    console.log('✅ 수동 테스트 결과:', JSON.stringify(result, null, 2));
    return result;
  } catch (error) {
    console.error('❌ 수동 테스트 실패:', error);
    return { success: false, error: error.toString() };
  }
}

// 스프레드시트 권한 테스트 함수
function testSpreadsheetAccess() {
  console.log('🔐 스프레드시트 접근 권한 테스트');
  
  try {
    const spreadsheet = SpreadsheetApp.openById('1LQNeT0abhMHXktrNjRbxl2XEFWVCwcYr5kVTAcRvpfM');
    console.log('✅ 스프레드시트 접근 성공:', spreadsheet.getName());
    
    const sheets = spreadsheet.getSheets();
    console.log('📋 시트 목록:', sheets.map(s => s.getName()));
    
    return {
      success: true,
      spreadsheet_name: spreadsheet.getName(),
      sheets: sheets.map(s => s.getName())
    };
  } catch (error) {
    console.error('❌ 스프레드시트 접근 실패:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
} 