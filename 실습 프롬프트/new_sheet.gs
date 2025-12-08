// ⭐ [필수] 유효한 키로 교체해야 합니다.
const GEMINI_API_KEY = 'AIXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
// const MODEL_NAME = 'gemini-2.5-flash';

/**
 * Google Sheets에서 데이터를 가져와 CSV 형식의 문자열로 변환합니다.
 * @returns {string} 분석할 데이터가 담긴 CSV 문자열
 */
function getSheetDataAsCsv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheet = ss.getSheetByName("시트1");

  // 🛡️ 방어 코드 추가: 시트가 존재하지 않을 경우 오류 처리
  if (!sheet) {
    throw new Error("오류: '판매 실적 자동 보고서 스크립트'라는 이름의 시트를 찾을 수 없습니다. 시트 이름을 확인하거나 시트 이름을 현재 시트로 변경하세요.");
  }
  
  // 데이터 범위: A1부터 D9까지 (헤더 포함)
  const range = sheet.getRange("A1:D9");
  const values = range.getValues();
  
  // CSV 문자열로 변환 
  let csvString = values.map(row => row.join(",")).join("\n");
  
  return csvString;
}
// ----------------------------------------------------------------------


/**
 * 보고서 작성을 위한 프롬프트를 구성합니다.
 * @param {string} dataCsv 분석할 데이터 (CSV 형식)
 * @returns {string} Gemini에게 전달할 전체 프롬프트
 */
function buildReportPrompt(dataCsv) {
  const systemPrompt = `
당신은 주어진 데이터를 분석하고 핵심 인사이트를 추출하여 명확한 보고서를 작성하는 전문 데이터 분석가입니다.
사용자에게 제공되는 데이터는 '지역별/제품별 3분기 판매 실적' 데이터이며, CSV 텍스트 형식으로 제공됩니다.

다음 지침에 따라 보고서를 작성하세요:

1.  **데이터 분석:**
    * 총 판매량이 가장 높은 **지역 Top 3**와 가장 낮은 **지역 Bottom 3**를 분석합니다.
    * 총 판매량이 가장 높은 **제품(A, B, C) 순위**를 분석합니다.
2.  **차트 생성 제안 (MarkDown):**
    * 위 분석 결과를 가장 효과적으로 시각화할 수 있는 **차트 종류(예: 막대 그래프, 원 그래프, 꺾은선 그래프 등)**를 제안합니다. 제안 이유도 간략하게 포함합니다.
3.  **요약 보고서:**
    * 분석 결과를 바탕으로 **경영진을 위한 요약 보고서**를 200자 이내로 작성합니다. 핵심적인 발견과 간단한 전략적 제언을 포함해야 합니다.

출력은 다음 3가지 섹션으로 구성되어야 하며, 각 섹션 제목을 명확하게 표시하세요.
`;

  const userPrompt = `
분석할 데이터:
---
${dataCsv}
---
`;
  
  return systemPrompt + userPrompt;
}

/**
 * Gemini API를 호출하여 데이터 분석 보고서를 생성합니다.
 */
function generateSalesReport() {
  //  GEMINI_API_KEY 체크 문자열을 실제 키에 맞게 수정해야 합니다.
  if (GEMINI_API_KEY === "YOUR_ACTUAL_GEMINI_API_KEY_HERE") {
    Browser.msgBox("오류", "GEMINI_API_KEY를 유효한 키로 변경해주세요.", Browser.Buttons.OK);
    return;
  }
  
  try {
    const dataCsv = getSheetDataAsCsv();
    const fullPrompt = buildReportPrompt(dataCsv);
    
    // Gemini API 엔드포인트 (v1beta와 MODEL_NAME 사용)
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${GEMINI_API_KEY}`;

    // 요청 본문 (Payload) 구성: generationConfig 누락 문제 해결을 위한 방어적 설계
    const payload = {
      contents: [
        {
          role: "user",
          parts: [{ text: fullPrompt }]
        }
      ],
      //  API 호출 안정성을 위한 generationConfig 추가 (optional, but recommended)
      generationConfig: {
          temperature: 0.1
      }
    };
    
    // 요청 옵션 설정
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true // 오류 발생 시 스크립트 중단 방지
    };

    Logger.log("Gemini API 호출 시작...");

    const response = UrlFetchApp.fetch(apiUrl, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.candidates && result.candidates.length > 0) {
      const generatedText = result.candidates[0].content.parts[0].text;
      
      outputReportToSheet(generatedText);
      Browser.msgBox("성공", "판매 실적 보고서가 성공적으로 생성되어 Sheets에 추가되었습니다.", Browser.Buttons.OK);
      
    } else if (result.error) {
       Logger.log("API 오류: " + JSON.stringify(result.error));
       Browser.msgBox("API 오류", "Gemini API 호출 중 오류가 발생했습니다: " + result.error.message, Browser.Buttons.OK);
    } else {
       Logger.log("알 수 없는 API 응답: " + response.getContentText());
       Browser.msgBox("오류", "Gemini에서 유효한 응답을 받지 못했습니다. 로그를 확인하세요.", Browser.Buttons.OK);
    }
    
  } catch (e) {
    // getSheetDataAsCsv에서 발생한 오류 포함 모든 오류를 처리
    Logger.log("스크립트 실행 중 오류: " + e.toString());
    Browser.msgBox("스크립트 오류", "보고서 생성 중 오류가 발생했습니다: " + e.toString(), Browser.Buttons.OK);
  }
}

/**
 * 생성된 보고서 텍스트를 Google Sheets의 새 시트에 출력합니다.
 * @param {string} reportText Gemini가 생성한 보고서 텍스트
 */
function outputReportToSheet(reportText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "보고서_" + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMdd_HHmm");
  
  const newSheet = ss.insertSheet(sheetName);
  
  newSheet.getRange("A1").setValue(reportText);
  newSheet.getRange("A1").setWrap(true); 
  
  newSheet.getRange("A1").setFontWeight("bold");
  newSheet.setColumnWidth(1, 800); 

}
