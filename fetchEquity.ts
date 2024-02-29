import { collectEquityPages, saveHTMLfile, saveToExcel } from './collectEquity';
//import { saveHTMLfile } from 'collectEquity';
//import { saveToExcel } from 'collectEquity';


async function main() {
  try {
    const date = '2024.02.29';
    const 지분공시파일명 = date + '_지분공시' + '.xlsx';
    const 지분공시폴더 = date + '_지분공시'
    const 지분공시페이지 = 'https://dart.fss.or.kr/dsac001/mainO.do?selectDate=' + date +  '&sort=time';
    //console.log(지분공시페이지);
    const equityList = await collectEquityPages(지분공시페이지);
    await saveToExcel(equityList, date + '_5%임원보고.xlsx');
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