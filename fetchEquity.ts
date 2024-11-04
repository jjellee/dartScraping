import { collectEquityPages, saveHTMLfile, saveToExcel } from './collectEquity';
//import { saveHTMLfile } from 'collectEquity';
//import { saveToExcel } from 'collectEquity';



async function main() {
  try {
    const SKIP_NUMBERS = []
      const date = '2024.11.01'
    const 지분공시파일명 = date + '_지분공시' + '.xlsx';
    const 지분공시폴더 = date + '_지분공시'
    const 지분공시페이지 = 'https://dart.fss.or.kr/dsac001/mainO.do?selectDate=' + date +  '&sort=time';
    //console.log(지분공시페이지);
    const equityList = await collectEquityPages(지분공시페이지);

    await saveToExcel(equityList, date + '_5%임원보고.xlsx');

    // SKIP_NUMBERS 순회하면서 해당 인덱스 - 1 위치의 항목 제거
    for (let i = SKIP_NUMBERS.length - 1; i >= 0; i--) {
      const indexToRemove = SKIP_NUMBERS[i] - 2
      if (indexToRemove >= 0 && indexToRemove < equityList.length) {
        equityList.splice(indexToRemove, 1);
      }
    }
    await saveHTMLfile(equityList, 지분공시폴더);

    //await getTransactionDetails('https://dart.fss.or.kr/dsaf001/main.do?rcpNo=20240115000355');
    
    //await test11();
    process.exit(0);
  } catch (error) {
    console.error(`🔴 오류 발생: ${error}`);
    process.exit(1);
  }
}

main();