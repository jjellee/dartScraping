"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.saveToExcel = exports.collectEquityPages = exports.saveHTMLfile = void 0;
const playwright_1 = require("playwright");
const xlsx = require("xlsx");
const fs = require("fs/promises");
const promises_1 = require("fs/promises");
const fs_1 = require("fs");
const path = require("path");
const 다트홈페이지 = 'https://dart.fss.or.kr';
const 최대응답대기시간ms = 10000;
async function saveFile(Content, FilePath) {
    const fileExists = await checkFileExists(FilePath);
    if (!fileExists) {
        try {
            await fs.writeFile(FilePath, Content, { encoding: 'utf-8' });
            console.log(`파일이 성공적으로 저장되었습니다: ${FilePath}`);
        }
        catch (err) {
            console.error("파일 저장 중 오류가 발생했습니다", err);
        }
    }
}
async function saveHTMLfile(equityList, 지분공시폴더) {
    await createNewFolder(지분공시폴더);
    for (let idx = 0; idx < equityList.length; idx++) {
        const 공시항목FilePath = path.join(지분공시폴더, `${idx + 1}공시항목.txt`);
        const row = equityList[idx];
        const 공시항목Content = '회사명 : ' + row.companyName + '\n'
            + '제출인 : ' + row.holder + '\n'
            + '보고서명 : ' + row.reportType + '\n'
            + '공시대상회사 : ' + row.market;
        console.log(공시항목Content);
        await saveFile(공시항목Content, 공시항목FilePath);
        if (row.reportType.includes('특정증권등소유상황보고서')) {
            const 보고자에관한상황htmlFilePath = path.join(지분공시폴더, `${idx + 1}보고자에관한상황.html`);
            const 소유특정증권등의수및소유비율htmlFilePath = path.join(지분공시폴더, `${idx + 1}소유특정증권등의수및소유비율.html`);
            const 세부변동내역htmlFilePath = path.join(지분공시폴더, `${idx + 1}세부변동내역.html`);
            const fileExists = await checkFileExists(보고자에관한상황htmlFilePath);
            if (fileExists) {
                continue;
            }
            const [보고자에관한상황html, 소유특정증권등의수및소유비율html, 세부변동내역html] = await 임원주요주주특정증권등소유상황보고서(row.회사지분공시홈페이지);
            await saveFile(보고자에관한상황html, 보고자에관한상황htmlFilePath);
            await saveFile(소유특정증권등의수및소유비율html, 소유특정증권등의수및소유비율htmlFilePath);
            await saveFile(세부변동내역html, 세부변동내역htmlFilePath);
        }
        else if (row.reportType.includes('주식등의대량보유상황보고서')) {
            const 대량보유자에관한사항htmlFilePath = path.join(지분공시폴더, `${idx + 1}보고자에관한상황.html`);
            const 주식등의종류별보유내역htmlFilePath = path.join(지분공시폴더, `${idx + 1}소유특정증권등의수및소유비율.html`);
            const 주식등의세부변동내역htmlFilePath = path.join(지분공시폴더, `${idx + 1}세부변동내역.html`);
            const fileExists = await checkFileExists(대량보유자에관한사항htmlFilePath);
            if (fileExists) {
                continue;
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
exports.saveHTMLfile = saveHTMLfile;
async function createNewFolder(folderPath) {
    try {
        try {
            await (0, promises_1.access)(folderPath, fs_1.constants.F_OK);
            console.log(`Folder already exists at: ${folderPath}`);
        }
        catch {
            await (0, promises_1.mkdir)(folderPath);
            console.log(`Folder created at: ${folderPath}`);
        }
    }
    catch (error) {
        console.error(`Error in creating folder: ${error}`);
    }
}
async function checkFileExists(filePath) {
    try {
        await fs.access(filePath, fs_1.constants.F_OK);
        return true;
    }
    catch {
        return false;
    }
}
async function clickByText(page, text) {
    try {
        const element = await page.$(`text=${text}`);
        if (element) {
            await element.click();
        }
        else {
            console.log(`Element with text "${text}" not found`);
        }
    }
    catch (error) {
        console.error(`Error occurred: ${error}`);
    }
}
async function tryClickByText(page, text1, text2) {
    const foundText1 = await page.$(`text="${text1}"`);
    if (foundText1) {
        await foundText1.click();
        return text1;
    }
    const foundText2 = await page.$(`text="${text2}"`);
    if (foundText2) {
        await foundText2.click();
        return text2;
    }
    return null;
}
async function 임원주요주주특정증권등소유상황보고서(회사지분공시홈페이지) {
    const browser = await playwright_1.chromium.launch({ headless: false });
    const 지분공시페이지 = await browser.newPage();
    await 지분공시페이지.goto(회사지분공시홈페이지);
    await 지분공시페이지.waitForTimeout(1000);
    await clickByText(지분공시페이지, '2. 보고자에 관한 사항');
    let iframeElement = await 지분공시페이지.$('iframe#ifrm');
    const 보고자에관한상황페이지url = await iframeElement?.getAttribute('src');
    const 보고자에관한상황페이지 = await browser.newPage();
    await 보고자에관한상황페이지.goto(다트홈페이지 + 보고자에관한상황페이지url);
    await 보고자에관한상황페이지.waitForTimeout(1000);
    await 보고자에관한상황페이지.waitForSelector('text="2. 보고자에 관한 사항"');
    const 보고자에관한상황table = await 보고자에관한상황페이지.$('table');
    const 보고자에관한상황html = await 보고자에관한상황table.evaluate(node => node.outerHTML);
    await clickByText(지분공시페이지, '3. 특정증권등의 소유상황');
    iframeElement = await 지분공시페이지.$('iframe#ifrm');
    const 특정증권등의소유상황url = await iframeElement?.getAttribute('src');
    const 특정증권등의소유상황페이지 = await browser.newPage();
    await 특정증권등의소유상황페이지.goto(다트홈페이지 + 특정증권등의소유상황url);
    await 특정증권등의소유상황페이지.waitForTimeout(1000);
    const 소유특정증권등의수및소유비율table = await 특정증권등의소유상황페이지.$('text="가. 소유 특정증권등의 수 및 소유비율" >> xpath=following-sibling::table[1]');
    const 소유특정증권등의수및소유비율html = await 소유특정증권등의수및소유비율table.evaluate(node => node.outerHTML);
    const 세부변동내역table = await 특정증권등의소유상황페이지.$('text="다. 세부변동내역" >> xpath=following-sibling::table');
    const 세부변동내역html = await 세부변동내역table.evaluate(node => node.outerHTML);
    await browser.close();
    return [보고자에관한상황html, 소유특정증권등의수및소유비율html, 세부변동내역html];
}
async function 주식등의대량보유상황보고서(회사지분공시홈페이지) {
    const browser = await playwright_1.chromium.launch({ headless: false });
    const 지분공시페이지 = await browser.newPage();
    await 지분공시페이지.goto(회사지분공시홈페이지);
    await clickByText(지분공시페이지, '2. 대량보유자에 관한 사항');
    let iframeElement = await 지분공시페이지.$('iframe#ifrm');
    const 대량보유자에관한사항페이지url = await iframeElement?.getAttribute('src');
    const 대량보유자에관한사항페이지 = await browser.newPage();
    await 대량보유자에관한사항페이지.goto(다트홈페이지 + 대량보유자에관한사항페이지url);
    await 대량보유자에관한사항페이지.waitForTimeout(1000);
    const 대량보유자에관한사항table = await 대량보유자에관한사항페이지.$('text="(1) 보고자 개요" >> xpath=following-sibling::table[1]');
    const 대량보유자에관한사항html = await 대량보유자에관한사항table.evaluate(node => node.outerHTML);
    const tryPage1 = '1. 보고자 및 특별관계자별 보유내역';
    const tryPage2 = '1. 보고자 및 특별관계자의 주식등의 종류별 보유내역';
    let 보고자및특별관계자보유내역html = '';
    const visitedPage = await tryClickByText(지분공시페이지, tryPage1, tryPage2);
    if (visitedPage == tryPage1) {
        iframeElement = await 지분공시페이지.$('iframe#ifrm');
        const 보고자및특별관계자별보유내역url = await iframeElement?.getAttribute('src');
        const 보고자및특별관계자별보유내역페이지 = await browser.newPage();
        await 보고자및특별관계자별보유내역페이지.goto(다트홈페이지 + 보고자및특별관계자별보유내역url);
        console.log('tryPage!!!!!!!');
        await 보고자및특별관계자별보유내역페이지.waitForTimeout(1000);
        보고자및특별관계자보유내역html = await 보고자및특별관계자별보유내역페이지.$$eval('table', (tables) => {
            const tableWithThead = tables.find(table => table.querySelector('thead'));
            return tableWithThead ? tableWithThead.outerHTML : '';
        });
    }
    else if (visitedPage == tryPage2) {
        iframeElement = await 지분공시페이지.$('iframe#ifrm');
        const 보고자및특별관계자별보유내역url = await iframeElement?.getAttribute('src');
        const 보고자및특별관계자별보유내역페이지 = await browser.newPage();
        await 보고자및특별관계자별보유내역페이지.goto(다트홈페이지 + 보고자및특별관계자별보유내역url);
        await 보고자및특별관계자별보유내역페이지.waitForTimeout(1000);
        보고자및특별관계자보유내역html = await 보고자및특별관계자별보유내역페이지.$$eval('table', (tables) => {
            const tableWithThead = tables.find(table => table.querySelector('thead'));
            return tableWithThead ? tableWithThead.outerHTML : '';
        });
    }
    else {
        console.log("'1. 보고자 및 특별관계자별 보유내역' or  '1. 보고자 및 특별관계자의 주식등의 종류별 보유내역' 둘 다 없음");
    }
    await clickByText(지분공시페이지, '2. 세부변동내역');
    iframeElement = await 지분공시페이지.$('iframe#ifrm');
    const 세부변동내역url = await iframeElement?.getAttribute('src');
    const 세부변동내역페이지 = await browser.newPage();
    await 세부변동내역페이지.goto(다트홈페이지 + 세부변동내역url);
    await 세부변동내역페이지.waitForTimeout(1000);
    const tables = await 세부변동내역페이지.$$('table');
    let 주식등의세부변동내역html;
    if (tables.length < 2) {
        console.log("테이블이 1개!");
        주식등의세부변동내역html = await tables[0].evaluate(node => node.outerHTML);
    }
    else {
        주식등의세부변동내역html = await tables[1].evaluate(node => node.outerHTML);
    }
    await browser.close();
    return [대량보유자에관한사항html, 보고자및특별관계자보유내역html, 주식등의세부변동내역html];
}
async function collectEquityPages(지분공시페이지) {
    const browser = await playwright_1.chromium.launch({ headless: false });
    const context = await browser.newContext();
    const page = await context.newPage();
    page.setDefaultTimeout(최대응답대기시간ms);
    await context.clearCookies();
    await page.goto(지분공시페이지);
    let currentPage = 1;
    let lastPage = await page.$$eval('.pageSkip > ul > li', items => items.length);
    let equityDataList = [];
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
            const 회사지분공시홈페이지 = 다트홈페이지 + href;
            equityDataList.push({ time, market, companyName, reportType, 회사지분공시홈페이지, holder, date });
            console.log('equityDataList push');
        }
        if (currentPage < lastPage) {
            const nextPage = currentPage + 1;
            await page.click(`.pageSkip > ul > li:nth-child(${nextPage}) a`);
            await page.waitForSelector(`.pageSkip > ul > li:nth-child(${nextPage}).on`, { state: 'attached' });
            currentPage = nextPage;
        }
        else {
            break;
        }
    }
    await browser.close();
    return equityDataList;
}
exports.collectEquityPages = collectEquityPages;
async function saveToExcel(data, filename) {
    const worksheet = xlsx.utils.json_to_sheet(data);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Data');
    xlsx.writeFile(workbook, filename);
}
exports.saveToExcel = saveToExcel;
//# sourceMappingURL=collectEquity.js.map