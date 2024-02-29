"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const collectEquity_1 = require("./collectEquity");
async function main() {
    try {
        const date = '2024.02.29';
        const 지분공시파일명 = date + '_지분공시' + '.xlsx';
        const 지분공시폴더 = date + '_지분공시';
        const 지분공시페이지 = 'https://dart.fss.or.kr/dsac001/mainO.do?selectDate=' + date + '&sort=time';
        const equityList = await (0, collectEquity_1.collectEquityPages)(지분공시페이지);
        await (0, collectEquity_1.saveToExcel)(equityList, date + '_5%임원보고.xlsx');
        await (0, collectEquity_1.saveHTMLfile)(equityList, 지분공시폴더);
        process.exit(0);
    }
    catch (error) {
        console.error(`🔴 오류 발생: ${error}`);
        process.exit(1);
    }
}
main();
//# sourceMappingURL=fetchEquity.js.map