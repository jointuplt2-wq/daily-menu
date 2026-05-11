const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const SAVE_DIR = path.join(__dirname, '엑셀파일저장');
const TARGET_URL = 'https://www.comwel.or.kr/care-gyeonggi/open/data.jsp';

async function downloadLatestExcel() {
  if (!fs.existsSync(SAVE_DIR)) {
    fs.mkdirSync(SAVE_DIR, { recursive: true });
  }

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ acceptDownloads: true, ignoreHTTPSErrors: true });
  const page = await context.newPage();

  console.log('페이지 열기:', TARGET_URL);
  await page.goto(TARGET_URL, { waitUntil: 'networkidle', timeout: 30000 });

  // 첫 번째 행의 다운로드 링크들 위치 파악
  const rowInfo = await page.evaluate(() => {
    const rows = document.querySelectorAll('table tbody tr');
    for (const row of rows) {
      const links = row.querySelectorAll('a[href*="download"]');
      if (links.length === 0) continue;
      const titleCell = row.querySelectorAll('td')[1];
      return {
        title: titleCell ? titleCell.innerText.trim() : '',
        count: links.length,
        // 링크 인덱스 (전체 페이지에서)
        firstHref: links[0].href,
      };
    }
    return null;
  });

  console.log('첫 번째 행 정보:', JSON.stringify(rowInfo, null, 2));

  if (!rowInfo) {
    console.error('행을 찾지 못했습니다.');
    await browser.close();
    return;
  }

  // 페이지의 모든 download 링크 중 첫 번째 행에 해당하는 것들 클릭
  let saved = false;

  for (let i = 0; i < rowInfo.count; i++) {
    // 페이지를 다시 로드해서 링크 클릭
    if (i > 0) {
      await page.goto(TARGET_URL, { waitUntil: 'networkidle', timeout: 30000 });
    }

    const downloadPromise = page.waitForEvent('download', { timeout: 20000 });

    // 첫 번째 행의 i번째 다운로드 링크 클릭
    await page.evaluate((idx) => {
      const rows = document.querySelectorAll('table tbody tr');
      for (const row of rows) {
        const links = row.querySelectorAll('a[href*="download"]');
        if (links.length === 0) continue;
        links[idx].click();
        return;
      }
    }, i);

    let download;
    try {
      download = await downloadPromise;
    } catch (e) {
      console.log(`  → 링크 ${i}: 다운로드 실패 -`, e.message);
      continue;
    }

    const name = download.suggestedFilename();
    console.log(`  → 링크 ${i} 파일명:`, name);

    const savePath = path.join(SAVE_DIR, name);
    await download.saveAs(savePath);
    console.log('저장 완료:', savePath);

    if (/\.xlsx?$/i.test(name)) {
      saved = true;
      break;
    }
  }

  if (!saved) {
    console.log('Excel 파일을 다운로드하지 못했습니다.');
  }

  await browser.close();
}

downloadLatestExcel().catch((err) => {
  console.error('오류:', err.message);
  process.exit(1);
});
