// @ts-nocheck
const NEWS_SHEET_NAME = "설정";
const GEMINI_API_SHEET_NAME = "Gemini API 설정"; // 시트 이름은 유지하되, 내부에서 네이버 키만 가져옵니다.

/**
 * 1. 메뉴 생성
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🟢 뉴스 요약 실행')
      .addItem('📧 이메일 발송하기', 'sendWeeklyNewsEmail')
      .addToUi();
}

/**
 * 2. 메인 실행 함수
 */
function sendWeeklyNewsEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 설정 값 불러오기
  const keywordCombos = getKeywordCombos(ss, NEWS_SHEET_NAME);
  const recipientsRaw = getSettingValue(ss, NEWS_SHEET_NAME, "메일 수신자");
  const recipients = recipientsRaw ? recipientsRaw.split(',').map(r => r.trim()).filter(r => r.length > 0) : [];
  
  const searchDaysVal = getSettingValue(ss, NEWS_SHEET_NAME, "검색 기간");
  const searchDays = searchDaysVal ? parseInt(searchDaysVal) : 1;

  // [변경] Gemini API Key는 필요 없으므로 주석 처리하거나 제거
  // const geminiApiKey = getSettingValue(ss, GEMINI_API_SHEET_NAME, "Gemini API Key");
  
  const naverClientId = getSettingValue(ss, GEMINI_API_SHEET_NAME, "Naver Client ID");
  const naverClientSecret = getSettingValue(ss, GEMINI_API_SHEET_NAME, "Naver Client Secret");

  // 언론사 필터링
  const allowedPublishersRaw = getSettingValue(ss, NEWS_SHEET_NAME, "필터링할 언론사 (쉼표로 구분)");
  const allowedPublishers = allowedPublishersRaw ? allowedPublishersRaw.split(',').map(p => p.trim()).filter(p => p.length > 0) : [];

  // 필수 설정 값 검증 (Gemini 관련 검증 제거)
  if (!naverClientId || !naverClientSecret) {
    Browser.msgBox("오류", "네이버 Client ID/Secret 설정이 누락되었습니다. 시트를 확인해주세요.", Browser.Buttons.OK);
    return;
  }
  if (keywordCombos.length === 0) {
    Browser.msgBox("오류", "검색 키워드가 설정되지 않았습니다.", Browser.Buttons.OK);
    return;
  }

  let allNews = [];
  const maxArticlesPerKeyword = 6; 

  // 뉴스 검색 루프
  for (const combo of keywordCombos) {
    const mainKeyword = combo.mainKeyword;
    Logger.log(`"${mainKeyword}" 검색 중...`);
    
    // 네이버 뉴스 검색 (요약문 포함하여 가져오기)
    let newsArticles = searchNaverNewsAPI(mainKeyword, searchDays, 100, naverClientId, naverClientSecret);

    // 언론사 필터링
    if (allowedPublishers.length > 0) {
      newsArticles = newsArticles.filter(article => allowedPublishers.includes(article.publisher));
    }

    // 중복 제거 및 상위 N개 선택
    newsArticles = filterDuplicateNews(newsArticles);
    newsArticles = newsArticles.slice(0, maxArticlesPerKeyword);

    if (newsArticles.length > 0) {
      Logger.log(`"${mainKeyword}" - ${newsArticles.length}개 기사 처리 완료`);
      
      for (const article of newsArticles) {
        // [핵심 변경] Gemini 요약 함수 호출 제거
        // 대신 네이버 API에서 가져온 description을 summary로 사용
        
        allNews.push({
          keyword: mainKeyword,
          headline: article.headline,
          link: article.link,
          summary: article.description, // 네이버 제공 요약문 사용
          thumbnail: article.thumbnail,
          publisher: article.publisher
        });

        // [변경] Gemini를 사용하지 않으므로 대기 시간(Sleep) 제거
        // 속도가 매우 빨라집니다.
      }
    }
  }

  if (allNews.length === 0) {
    Browser.msgBox("알림", "새로운 뉴스가 없습니다.", Browser.Buttons.OK);
    return;
  }

  // 이메일 발송
  const emailSubject = `[Sleep상품기획팀] 주간 뉴스 요약 (${new Date().toLocaleDateString('ko-KR')})`;
  const emailBody = composeEmailContent(allNews, searchDays);

  if (recipients.length > 0) {
    for (const recipient of recipients) {
      MailApp.sendEmail({
        to: recipient,
        subject: emailSubject,
        htmlBody: emailBody
      });
    }
    Browser.msgBox("성공", "뉴스 요약 메일 발송이 완료되었습니다.", Browser.Buttons.OK);
  } else {
    Browser.msgBox("오류", "메일 수신자가 설정되지 않았습니다.", Browser.Buttons.OK);
  }
}

// ==========================================
// 3. 헬퍼 함수들 (기능 구현부)
// ==========================================

function getSettingValue(ss, sheetName, itemName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return '';
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === itemName) {
      return data[i][1] !== null ? data[i][1].toString().trim() : '';
    }
  }
  return '';
}

function getKeywordCombos(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headerRow = data[0];
  const mainIdx = headerRow.indexOf("메인 키워드");
  const subIdx = headerRow.indexOf("보조 키워드 (OR)");

  if (mainIdx === -1) return [];

  const combos = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const mainKw = (row[mainIdx] || '').toString().trim();
    const subKwStr = (subIdx !== -1 ? row[subIdx] : '').toString().trim();

    if (mainKw) {
      const subKws = subKwStr.split(',').map(k => k.trim()).filter(k => k.length > 0);
      combos.push({ mainKeyword: mainKw, subKeywords: subKws });
    }
  }
  return combos;
}

function searchNaverNewsAPI(keyword, days, limit, clientId, clientSecret) {
  const encodedKeyword = encodeURIComponent(keyword);
  const apiUrl = `https://openapi.naver.com/v1/search/news.json?query=${encodedKeyword}&display=${limit}&sort=sim`;

  try {
    const options = {
      method: "get",
      headers: { "X-Naver-Client-Id": clientId, "X-Naver-Client-Secret": clientSecret },
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(apiUrl, options);
    if (response.getResponseCode() !== 200) return [];

    const json = JSON.parse(response.getContentText());
    if (!json.items || json.items.length === 0) return [];

    let articles = [];
    const thresholdDate = new Date();
    thresholdDate.setHours(0, 0, 0, 0);
    thresholdDate.setDate(thresholdDate.getDate() - days);

    // HTML 태그 제거 및 특수문자 디코딩 함수
    const cleanText = (text) => {
      if (!text) return "";
      return text.replace(/<[^>]+>/g, '') // HTML 태그 제거
                 .replace(/&quot;/g, '"')
                 .replace(/&amp;/g, '&')
                 .replace(/&lt;/g, '<')
                 .replace(/&gt;/g, '>')
                 .replace(/&nbsp;/g, ' ');
    };

    for (const item of json.items) {
      const pubDate = new Date(item.pubDate);
      pubDate.setHours(0, 0, 0, 0);

      if (pubDate >= thresholdDate) {
        const headline = cleanText(item.title);
        // [추가] 네이버 API가 주는 description(요약문)을 가져와 정제
        const description = cleanText(item.description);

        articles.push({
          headline: headline,
          link: item.link,
          description: description, // 요약문 저장
          pubDate: pubDate,
          thumbnail: item.thumbnail || null,
          publisher: item.publisher || '알 수 없음'
        });
      }
      if (articles.length >= limit) break;
    }
    return articles;
  } catch (e) {
    Logger.log(`네이버 API 오류: ${e.message}`);
    return [];
  }
}

function filterDuplicateNews(newsList) {
  if (newsList.length <= 1) return newsList;
  const filtered = [];
  const seen = new Set();
  const normalize = (t) => t.toLowerCase().replace(/[^\w\sㄱ-힣]/g, '').replace(/\s+/g, '');
  
  for (const news of newsList) {
    const key = normalize(news.headline);
    if (!seen.has(key)) {
      seen.add(key);
      filtered.push(news);
    }
  }
  return filtered;
}

function composeEmailContent(newsData, searchDays) {
  const formatDate = (date) => {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  };
  const today = new Date();
  const endDate = formatDate(today);
  const startDate = new Date();
  startDate.setDate(today.getDate() - searchDays);

  let htmlBody = `<div style="font-family: Arial, sans-serif; line-height: 150%;">
    <p>안녕하세요. Sleep상품기획팀 원량진입니다.</p>
    <p>주간 뉴스 검색 결과 공유드립니다. (${formatDate(startDate)} ~ ${endDate})</p>
    <hr style="border: 0; height: 1px; background: #eee;">`;

  const newsByKeyword = newsData.reduce((acc, news) => {
    if (!acc[news.keyword]) acc[news.keyword] = [];
    acc[news.keyword].push(news);
    return acc;
  }, {});

  for (const keyword in newsByKeyword) {
    htmlBody += `<h2 style="color: #333;">${keyword} 뉴스</h2><ul style="list-style: none; padding: 0;">`;
    for (const news of newsByKeyword[keyword]) {
      htmlBody += `<li style="margin-bottom: 20px; border-bottom: 1px solid #f0f0f0; padding-bottom: 15px;">
        <h3 style="font-size: 14pt; margin: 0 0 7px 0;">
          <strong><a href="${news.link}" style="color: #0000FF; text-decoration: none;">${news.headline}</a></strong>
        </h3>`;
      if (news.thumbnail) {
        htmlBody += `<div style="margin-top: 10px;"><img src="${news.thumbnail}" alt="썸네일" style="max-width: 300px; border-radius: 5px;"></div>`;
      }
      // 요약문(description) 표시
      htmlBody += `<p style="margin-top: 10px; color: #555;">${news.summary}</p></li>`;
    }
    htmlBody += `</ul><hr style="border: 0; height: 1px; background: #eee;">`;
  }

  htmlBody += `<p>감사합니다.<br>원량진 드림</p></div>`;
  return htmlBody;
}
