//import { CollectFixedImportExportInput } from './../../dto/collect-fixed-import-export.input';
import { chromium, Page, Browser } from 'playwright';
import * as xlsx from 'xlsx';
import * as fs from 'fs/promises';
import { rm, mkdir, access } from 'fs/promises';
import { constants } from 'fs';
import * as path from 'path';


  //&series=&mdayCnt=0#none
const 다트홈페이지 = 'https://dart.fss.or.kr';

const 최대응답대기시간ms = 10000 as const;

async function saveFile(Content: string, FilePath : string) {
  //console.log('saveFile');
  const fileExists = await checkFileExists(FilePath);
  if (!fileExists) {
    try {
      await fs.writeFile(FilePath, Content, { encoding: 'utf-8' });
      console.log(`파일이 성공적으로 저장되었습니다: ${FilePath}`);
    } catch (err) {
      console.error("파일 저장 중 오류가 발생했습니다", err);
    }
  }
}

export async function saveHTMLfile(equityList: any[], 지분공시폴더 : string) {
  await createNewFolder(지분공시폴더);
  for (let idx = 0; idx < equityList.length; idx++) {
    const 공시항목FilePath = path.join(지분공시폴더, `${idx+1}공시항목.txt`);

    const row = equityList[idx];
    //console.log(row.companyName + ', ' + row.reportType + ', ' + row.회사지분공시홈페이지 + ', ' + row.holder);

    const 공시항목Content = '회사명 : ' + row.companyName + '\n'
    + '제출인 : ' + row.holder + '\n'
    + '보고서명 : ' + row.reportType + '\n'
    + '공시대상회사 : ' + row.market;
    
    console.log(공시항목Content);
    await saveFile(공시항목Content, 공시항목FilePath);
    
    if (row.reportType.includes('특정증권등소유상황보고서')){ //임원주요주주 특정증권등 소유상황보고서 #&& !row.reportType.includes('기재정정')) 
      const 보고자에관한상황htmlFilePath = path.join(지분공시폴더, `${idx+1}보고자에관한상황.html`);
      const 소유특정증권등의수및소유비율htmlFilePath = path.join(지분공시폴더, `${idx+1}소유특정증권등의수및소유비율.html`);
      const 세부변동내역htmlFilePath = path.join(지분공시폴더, `${idx+1}세부변동내역.html`);

      // 파일이 존재한다면 스킾
      const fileExists = await checkFileExists(보고자에관한상황htmlFilePath);
      if (fileExists) {
        continue
      }

      const [보고자에관한상황html, 소유특정증권등의수및소유비율html, 세부변동내역html] = await 임원주요주주특정증권등소유상황보고서(row.회사지분공시홈페이지);

      await saveFile(보고자에관한상황html, 보고자에관한상황htmlFilePath);      
      await saveFile(소유특정증권등의수및소유비율html, 소유특정증권등의수및소유비율htmlFilePath);      
      await saveFile(세부변동내역html, 세부변동내역htmlFilePath);
    }
    else if (row.reportType.includes('주식등의대량보유상황보고서')){ //주식등의대량보유상황보고서 #&& !row.reportType.includes('기재정정')      
      const 대량보유자에관한사항htmlFilePath = path.join(지분공시폴더, `${idx+1}보고자에관한상황.html`);
      const 주식등의종류별보유내역htmlFilePath = path.join(지분공시폴더, `${idx+1}소유특정증권등의수및소유비율.html`);
      const 주식등의세부변동내역htmlFilePath = path.join(지분공시폴더, `${idx+1}세부변동내역.html`);

      // 파일이 존재한다면 스킾
      const fileExists = await checkFileExists(대량보유자에관한사항htmlFilePath);
      if (fileExists) {
        continue
      }
      
      const [대량보유자에관한사항html, 주식등의종류별보유내역html, 주식등의세부변동내역html] = await 주식등의대량보유상황보고서(row.회사지분공시홈페이지);

      await saveFile(대량보유자에관한사항html, 대량보유자에관한사항htmlFilePath);
      await saveFile(주식등의종류별보유내역html, 주식등의종류별보유내역htmlFilePath);      
      await saveFile(주식등의세부변동내역html, 주식등의세부변동내역htmlFilePath);

    }
    else {
      console.log('처음보는 리포트 타입입니다 : ' + row.reportType);
    }
  }
}

async function createNewFolder(folderPath: string) {
  try {
    // 폴더가 존재하는지 확인
    try {
      await access(folderPath, constants.F_OK);
      console.log(`Folder already exists at: ${folderPath}`);
    } catch {
      // 폴더가 존재하지 않으면 새 폴더 생성
      await mkdir(folderPath);
      console.log(`Folder created at: ${folderPath}`);
    }
  } catch (error) {
    console.error(`Error in creating folder: ${error}`);
  }
}


async function checkFileExists(filePath : string) {
  try {
    await fs.access(filePath, constants.F_OK);
    return true; // 파일이 존재함
  } catch {
    return false; // 파일이 존재하지 않음 또는 접근 불가
  }
}

async function clickByText(page: Page, text: string): Promise<void> {
  try {
    // '2. 세부변동내역' 텍스트를 포함하는 요소를 찾습니다.
    const element = await page.$(`text=${text}`);
    if (element) {
      // 요소를 클릭합니다.
      await element.click();
    } else {
      console.log(`Element with text "${text}" not found`);
    }
  } catch (error) {
    console.error(`Error occurred: ${error}`);
  }
}

async function tryClickByText(page: Page, text1: string, text2: string): Promise<string | null> {
  // 첫 번째 텍스트를 찾아 클릭 시도
  const foundText1 = await page.$(`text="${text1}"`);
  if (foundText1) {
    await foundText1.click();
    return text1; // 첫 번째 텍스트 페이지 방문
  }

  // 두 번째 텍스트를 찾아 클릭 시도
  const foundText2 = await page.$(`text="${text2}"`);
  if (foundText2) {
    await foundText2.click();
    return text2; // 두 번째 텍스트 페이지 방문
  }

  // 두 텍스트 모두 찾지 못한 경우
  return null; // 또는 오류 메시지 반환
}


async function 임원주요주주특정증권등소유상황보고서 ( 회사지분공시홈페이지 : string ) : Promise<[string, string, string]> {
  const browser: Browser = await chromium.launch({ headless: false });
  const 지분공시페이지 = await browser.newPage();

  await 지분공시페이지.goto(회사지분공시홈페이지);

  await 지분공시페이지.waitForTimeout(1000);

  //'2. 보고자에 관한 사항' 클릭
  await clickByText(지분공시페이지, '2. 보고자에 관한 사항');
  let iframeElement = await 지분공시페이지.$('iframe#ifrm'); // id가 'ifrm'인 iframe 요소 찾기
  const 보고자에관한상황페이지url = await iframeElement?.getAttribute('src'); // src 속성 가져오기

  const 보고자에관한상황페이지= await browser.newPage();
  await 보고자에관한상황페이지.goto(다트홈페이지 + 보고자에관한상황페이지url);

  await 보고자에관한상황페이지.waitForTimeout(1000);
  
  await 보고자에관한상황페이지.waitForSelector('text="2. 보고자에 관한 사항"');
  //const 보고자에관한상황table = await 보고자에관한상황페이지.$('text="2. 보고자에 관한 사항" >> xpath=following-sibling::table[1]');
  const 보고자에관한상황table = await 보고자에관한상황페이지.$('table');
  //const 보고자에관한상황table = await 보고자에관한상황페이지.$('xpath=//p[contains(text(), "2. 보고자에 관한 사항")]/following-sibling::*[1]/following-sibling::table[1]');

  //console.log(보고자에관한상황table)
  const 보고자에관한상황html = await 보고자에관한상황table.evaluate(node => node.outerHTML);
  
  //const 보고자에관한상황html = 보고자에관한상황table ? await 보고자에관한상황table.evaluate(node => node.outerHTML) : '테이블을 찾을 수 없음';

  // '3.특정증권등의 소유상황' 클릭
  await clickByText(지분공시페이지, '3. 특정증권등의 소유상황');

  iframeElement = await 지분공시페이지.$('iframe#ifrm'); // id가 'ifrm'인 iframe 요소 찾기
  const 특정증권등의소유상황url = await iframeElement?.getAttribute('src'); // src 속성 가져오기

  const 특정증권등의소유상황페이지 = await browser.newPage();
  await 특정증권등의소유상황페이지.goto(다트홈페이지 + 특정증권등의소유상황url);

  await 특정증권등의소유상황페이지.waitForTimeout(1000);

  // '가. 소유 특정증권등의 수 및 소유비율' 텍스트 바로 아래에 있는 table 찾기
  const 소유특정증권등의수및소유비율table = await 특정증권등의소유상황페이지.$('text="가. 소유 특정증권등의 수 및 소유비율" >> xpath=following-sibling::table[1]');
  const 소유특정증권등의수및소유비율html = await 소유특정증권등의수및소유비율table.evaluate(node => node.outerHTML);
  
  // '다. 세부변동내역' 텍스트 바로 아래에 있는 table 찾기
  const 세부변동내역table = await 특정증권등의소유상황페이지.$('text="다. 세부변동내역" >> xpath=following-sibling::table');
  const 세부변동내역html = await 세부변동내역table.evaluate(node => node.outerHTML);

  await browser.close();
  return [보고자에관한상황html, 소유특정증권등의수및소유비율html, 세부변동내역html];
}

async function 주식등의대량보유상황보고서( 회사지분공시홈페이지 : string ) : Promise<[string, string, string]> {
  
  const browser: Browser = await chromium.launch({ headless: false });
  const 지분공시페이지 = await browser.newPage();
  await 지분공시페이지.goto(회사지분공시홈페이지);

  // '2. 대량보유자에 관한 사항' 클릭
  await clickByText(지분공시페이지, '2. 대량보유자에 관한 사항');
  
  let iframeElement = await 지분공시페이지.$('iframe#ifrm'); // id가 'ifrm'인 iframe 요소 찾기
  const 대량보유자에관한사항페이지url = await iframeElement?.getAttribute('src'); // src 속성 가져오기
  
  const 대량보유자에관한사항페이지= await browser.newPage();
  await 대량보유자에관한사항페이지.goto(다트홈페이지 + 대량보유자에관한사항페이지url);
  
  await 대량보유자에관한사항페이지.waitForTimeout(1000);

  const 대량보유자에관한사항table = await 대량보유자에관한사항페이지.$('text="(1) 보고자 개요" >> xpath=following-sibling::table[1]');
  const 대량보유자에관한사항html = await 대량보유자에관한사항table.evaluate(node => node.outerHTML);

  // '1. 보고자 및 특별관계자별 보유내역' or  '1. 보고자 및 특별관계자의 주식등의 종류별 보유내역' 클릭
  //await clickByText(지분공시페이지, '1. 보고자 및 특별관계자별 보유내역');
  // 함수 사용 예시
  const tryPage1 = '1. 보고자 및 특별관계자별 보유내역'; // (일반)
  const tryPage2 = '1. 보고자 및 특별관계자의 주식등의 종류별 보유내역'; // (약식)

  let 보고자및특별관계자보유내역html: string = ''; // 변수 초기화
  const visitedPage = await tryClickByText(지분공시페이지, tryPage1, tryPage2);
  if (visitedPage == tryPage1) {
    iframeElement = await 지분공시페이지.$('iframe#ifrm'); // id가 'ifrm'인 iframe 요소 찾기
    const 보고자및특별관계자별보유내역url = await iframeElement?.getAttribute('src'); // src 속성 가져오기
  
    const 보고자및특별관계자별보유내역페이지= await browser.newPage();
    await 보고자및특별관계자별보유내역페이지.goto(다트홈페이지 + 보고자및특별관계자별보유내역url);
    
    await 보고자및특별관계자별보유내역페이지.waitForTimeout(1000);

    //보고자및특별관계자보유내역html = await 보고자및특별관계자별보유내역페이지.$eval('table', (table) => table.outerHTML);
    보고자및특별관계자보유내역html = await 보고자및특별관계자별보유내역페이지.$$eval('table', (tables) => {
      // 모든 table 요소 중에서 thead를 포함하는 첫 번째 table을 찾습니다.
      const tableWithThead = tables.find(table => table.querySelector('thead'));
      // 찾은 table의 outerHTML을 반환합니다. 만약 thead를 포함하는 table이 없다면, 빈 문자열을 반환합니다.
      return tableWithThead ? tableWithThead.outerHTML : '';
    });
  } else if (visitedPage == tryPage2) {
    iframeElement = await 지분공시페이지.$('iframe#ifrm'); // id가 'ifrm'인 iframe 요소 찾기
    const 보고자및특별관계자별보유내역url = await iframeElement?.getAttribute('src'); // src 속성 가져오기
  
    const 보고자및특별관계자별보유내역페이지= await browser.newPage();
    await 보고자및특별관계자별보유내역페이지.goto(다트홈페이지 + 보고자및특별관계자별보유내역url);
  
    await 보고자및특별관계자별보유내역페이지.waitForTimeout(1000);
  
    //const 보고자및특별관계자보유내역table = await 보고자및특별관계자별보유내역페이지.$('text="1. 보고자 및 특별관계자의 주식등의 종류별 보유내역" >> xpath=following-sibling::table[1]');
    // 첫 번째 테이블의 HTML 추출
    //보고자및특별관계자보유내역html = await 보고자및특별관계자별보유내역페이지.$eval('table', (table) => table.outerHTML);
    보고자및특별관계자보유내역html = await 보고자및특별관계자별보유내역페이지.$$eval('table', (tables) => {
      // 모든 table 요소 중에서 thead를 포함하는 첫 번째 table을 찾습니다.
      const tableWithThead = tables.find(table => table.querySelector('thead'));
      // 찾은 table의 outerHTML을 반환합니다. 만약 thead를 포함하는 table이 없다면, 빈 문자열을 반환합니다.
      return tableWithThead ? tableWithThead.outerHTML : '';
    });
  } else {
    console.log("'1. 보고자 및 특별관계자별 보유내역' or  '1. 보고자 및 특별관계자의 주식등의 종류별 보유내역' 둘 다 없음");
  }
  
  // '2. 세부변동내역' 클릭
  await clickByText(지분공시페이지, '2. 세부변동내역');
  iframeElement = await 지분공시페이지.$('iframe#ifrm'); // id가 'ifrm'인 iframe 요소 찾기
  const 세부변동내역url = await iframeElement?.getAttribute('src'); // src 속성 가져오기

  const 세부변동내역페이지= await browser.newPage();
  await 세부변동내역페이지.goto(다트홈페이지 + 세부변동내역url);

  await 세부변동내역페이지.waitForTimeout(1000);

  // 두 번째 테이블을 찾습니다.
  const tables = await 세부변동내역페이지.$$('table');

  // 블록 밖에서 변수를 선언하여 함수 전체에서 접근 가능하게 합니다.
  let 주식등의세부변동내역html;
  if (tables.length < 2) {
      console.log("테이블이 1개!");
      주식등의세부변동내역html = await tables[0].evaluate(node => node.outerHTML);
      //process.exit(1);
  } else {
    // 두 번째 테이블의 외부 HTML을 가져옵니다.
    주식등의세부변동내역html = await tables[1].evaluate(node => node.outerHTML);
  }
  await browser.close();
  //console.log(htmlContent);
  return [대량보유자에관한사항html, 보고자및특별관계자보유내역html, 주식등의세부변동내역html];
}

export async function collectEquityPages( 지분공시페이지 : string ) {
  const browser = await chromium.launch({ headless : false });
  const context = await browser.newContext();

  const page = await context.newPage();
  page.setDefaultTimeout(최대응답대기시간ms);

  // 쿠키 삭제
  await context.clearCookies();
  await page.goto(지분공시페이지); 
  //await page.goto(지분공시페이지, { waitUntil: 'networkidle' }); 

  let currentPage = 1;
  let lastPage = await page.$$eval('.pageSkip > ul > li', items => items.length);
  let equityDataList = []; // 데이터를 저장할 리스트

  while (currentPage <= lastPage) {
    console.log('currentPage:' + currentPage + '  lastPage:' + lastPage);
    const trElements = await page.$$('tbody > tr');
    
    for (const tr of trElements) {
      const time = await tr.$eval('td:nth-child(1)', node => node.textContent.trim());
      const market = await tr.$eval('td:nth-child(2) span[title]', node => node.getAttribute('title').trim());
      const companyName = await tr.$eval('td:nth-child(2) a', node => node.textContent.trim());
      const reportType = await tr.$eval('td:nth-child(3)', node => node.textContent.trim().split('\n')[0].trim());
      const href = await tr.$eval('td:nth-child(3) a', node => node.getAttribute('href'));
      const holder = await tr.$eval('td:nth-child(4)', node => node.textContent.trim());
      const date = await tr.$eval('td:nth-child(5)', node => node.textContent.trim());
      
      const 회사지분공시홈페이지 = 다트홈페이지 + href
      // 추출된 데이터를 리스트에 추가
      equityDataList.push({ time, market, companyName, reportType, 회사지분공시홈페이지, holder, date });
      console.log('equityDataList push');
    }

    // 페이지 이동 전, 다음 페이지 번호 확인
    if (currentPage < lastPage) {
      const nextPage = currentPage + 1;
      await page.click(`.pageSkip > ul > li:nth-child(${nextPage}) a`);
      // 다음 페이지의 내용이 로드될 때까지 대기 (예: 새로운 테이블 행이 나타날 때까지)
      await page.waitForSelector(`.pageSkip > ul > li:nth-child(${nextPage}).on`, { state: 'attached' });
      // 페이지 번호 업데이트
      currentPage = nextPage;
    } else {
      break; // 마지막 페이지에 도달하면 반복문 종료
    }
  } // while

  // 브라우저도 닫기
  await browser.close();
  //console.log(equityDataList)
  //saveToExcel(equityDataList,'지분공시.xlsx');
  return equityDataList;
}

export async function saveToExcel(data, filename) {
  const worksheet = xlsx.utils.json_to_sheet(data);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Data');
  xlsx.writeFile(workbook, filename);
}
