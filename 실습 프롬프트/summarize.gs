const API_KEY = 'AIXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'; 
const MODEL_NAME = 'gemini-2.5-flash';

/**
 * Google Sheets에서 선택된 셀 범위의 데이터를 Gemini API를 사용하여 요약합니다.
 */
function summarizeSelectedRange() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  // 1. 초기 점검
  if (!range || range.getNumRows() === 0) {
    Browser.msgBox('오류', '데이터를 요약할 셀 범위를 먼저 선택해주세요.', Browser.Buttons.OK);
    return;
  }
  if (API_KEY === 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXX') {
    Browser.msgBox('오류', 'API_KEY를 유효한 키로 변경해야 스크립트가 실행됩니다.', Browser.Buttons.OK);
    return;
  }
  
  // 2. 선택된 셀들을 하나의 텍스트로 변환
  const values = range.getDisplayValues(); 
  // 탭(\t)으로 열을 구분하고, 개행 문자(\n)로 행을 구분하는 텍스트로 변환
  let textData = values.map(row => row.join('\t')).join('\n');

  // 3. 맞춤형 프롬프트 정의
  const prompt = `
너는 데이터 분석 요약 도우미이다.
아래는 스프레드시트에서 복사한 판매 데이터이다. 이 데이터를 분석하고 요청된 형식에 맞춰 요약 보고서를 작성해라.

데이터:
---
${textData}
---

요약 규칙:
1. 전체 매출 규모와 특징을 분석하여 3줄로 요약한다.
2. 데이터에 나타난 상위 지역 또는 제품/카테고리 3개를 분석하여 bullet 형태로 정리한다.
3. 최종 요약은 경영진에게 보고하는 톤으로 5줄 이내로 작성한다.

출력은 오직 요약 내용만 포함해야 한다.
`;

  // 4. Gemini API 요청 payload 구성 (⭐ 오류 수정 부분)
  const payload = {
    contents: [
      {
        parts: [
          { text: prompt }
        ]
      }
    ],
    //  수정: 모델 설정은 'generationConfig' 필드로 전달해야 합니다.
    generationConfig: { 
      temperature: 0.1 // 창의성 최소화 (사실 기반 분석에 유리)
    }
  };

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${API_KEY}`;

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    // 5. 응답 텍스트 추출 및 오류 처리
    const text = result.candidates &&
                 result.candidates[0] &&
                 result.candidates[0].content &&
                 result.candidates[0].content.parts &&
                 result.candidates[0].content.parts[0].text
                 ? result.candidates[0].content.parts[0].text
                 : `⚠️ 응답을 읽어오지 못했습니다. API 오류: ${JSON.stringify(result.error || result)}`;

    // 6. 요약 결과를 시트에 출력
    const lastRow = range.getLastRow();
    const startColumn = range.getColumn();
    const outputCell = sheet.getRange(lastRow + 2, startColumn);  // 두 줄 띄우고 출력
    
    outputCell.setValue(text);
    outputCell.setWrap(true); // 텍스트 줄바꿈 설정
    outputCell.setFontWeight(text.startsWith('⚠️') ? 'normal' : 'bold'); // 오류 메시지면 볼드 해제
    
    Browser.msgBox('보고서 자동 생성 완료', '선택한 데이터의 요약 보고서가 시트에 출력되었습니다.', Browser.Buttons.OK);
    
  } catch (e) {
    Logger.log("스크립트 실행 오류: " + e.toString());
    Browser.msgBox('스크립트 오류', '스크립트 실행 중 문제가 발생했습니다: ' + e.toString(), Browser.Buttons.OK);
  }

}
