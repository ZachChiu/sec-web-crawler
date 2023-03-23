import puppeteer from "puppeteer";
import cliProgress from "cli-progress";
import chalk from "chalk";
import { GoogleSpreadsheet } from "google-spreadsheet";
import dayjs from "dayjs";
import { readFile } from "fs/promises";

/**
 * 上傳發放狀態
 * @param {String} docID docID
 * @param {String} sheetID sheetID
 * @param {String} credentialsPath credentialsPath
 * @return {Array}
 */
const getSheet = async (
  docID,
  sheetID,
  credentialsPath = "./credentials.json"
) => {
  const doc = new GoogleSpreadsheet(docID);
  const credits = JSON.parse(
    await readFile(new URL(credentialsPath, import.meta.url))
  );
  await doc.useServiceAccountAuth(credits);
  await doc.loadInfo();
  const sheet = doc.sheetsById[sheetID];
  //   rows[0]["Disclosure"] = "1";
  //   save(rows[0]);
  //   rows[0]._rawData;
  //   for (row of rows) {
  //     result.push(row._rawData);
  //   }

  return sheet;
};

/**
 * 上傳發放狀態
 * @param {Array} row google 列資料
 */
async function save(row) {
  try {
    await row.save();
  } catch (error) {
    await save(row);
  }
}

/**
 * 找到 10-K 的 link
 * @param {Object} page page
 * @param {String} year year
 */
async function link10K(page, year) {
  try {
    const trs = await page.$$("#filingsTable tr");

    for (const tr of trs) {
      const tdFormType = await tr.$$("td.dtr-control");
      const tdFilingDate = await tr.$$("td.sorting_1");
      if (tdFormType.length === 1 && tdFilingDate.length === 1) {
        const formType = await tdFormType[0].evaluate((el) =>
          el.textContent.trim()
        );
        const filingDate = await tdFilingDate[0].evaluate((el) =>
          el.textContent.trim()
        );

        if (formType === "10-K" && year === dayjs(filingDate).format("YYYY")) {
          const href = await tr.$eval("a", (element) => element.href);

          return href;
        }
      }
    }
    return null;
  } catch (error) {
    return null;
  }
}

/**
 * 找到關鍵字
 * @param {Object} page page
 */
async function findKeywords(page) {
  try {
    const keywords = JSON.parse(
      await readFile(new URL("./findData.json", import.meta.url))
    );

    const content = await page.content();
    const match = [];
    for (let i = 0; i < keywords.length; i++) {
      const keyword = keywords[i];
      const regex = new RegExp(keyword, "gi");

      if (content.match(regex)) {
        match.push(i + 1);
      }
    }

    return match;
  } catch (error) {
    return [];
  }
}

/**
 * 延遲 X 秒
 * @param {Number} time time
 */
const delay = (time) => {
  return new Promise(function (resolve) {
    setTimeout(resolve, time);
  });
};

/**
 * 操作 Dom 元素
 * @param {Object} page page
 * @param {String} year year
 */
const domOperate = async (page, year) => {
  try {
    // 確認 View Filings 是否要打開以及點擊
    const isBtnVisible = await page.evaluate(() => {
      const btn = document.querySelector("#btnViewAllFilings");
      return btn ? !btn.classList.contains("hidden") : false;
    });
    if (isBtnVisible) {
      const btnViewAllFilings = await page.$('button[id="btnViewAllFilings"]');
      await btnViewAllFilings.click();
    }

    // 選擇表單類型為 10-K 的輸入框
    const formTypeSelect = await page.$('input[id="searchbox"]');
    await formTypeSelect.click({ clickCount: 3 });
    await formTypeSelect.type("10-K");
    // await delay(500);

    // 選擇填報日期為 year 年的輸入框
    const dateFromField = await page.$('input[id="filingDateFrom"]');
    await dateFromField.click({ clickCount: 3 });
    await dateFromField.type(year);
    await dateFromField.press("Enter");

    // 選擇填報日期為 year 年的輸入框
    const dateToField = await page.$('input[id="filingDateTo"]');
    await dateToField.click({ clickCount: 3 });
    await dateToField.type(year);
    await dateToField.press("Enter");
    await delay(500);
  } catch (error) {
    throw error;
  }
};

const main = async () => {
  const browser = await puppeteer.launch({
    executablePath:
      // "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", // MAC
      "C:/Users/awdx8/AppData/Local/Google/Chrome/Application/chrome.exe", // windows
    headless: false,
  });
  try {
    const sheet = await getSheet(
      "1l8Wtu7aLmfPMinx67Ja6EgJoKw8yELcStHyKTB8SCnY",
      "1838907097"
    );
    const rows = await sheet.getRows();

    await sheet.loadCells(`A2:H${rows.length + 1}`);

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowData = row._rawData;

      const secLink = sheet.getCell(i + 1, 6);
      const year = dayjs(rowData[3]).format("YYYY");
      const hyperlink = secLink?.hyperlink;

      if (hyperlink) {
        const page = await browser.newPage();
        await page.goto(hyperlink);
        await page.setViewport({ width: 1080, height: 1024 });

        await delay(500);
        await domOperate(page, year);

        const link = await link10K(page, year);
        console.log(link);

        if (link) {
          // const linkPage = await browser.newPage();
          await page.goto(link);

          const result = await findKeywords(page);
          console.log("result", result);
        }

        // 等待搜索結果頁面載入

        console.log(i, "end");

        await page.close();
      }
    }
  } catch (error) {
    console.log(error);
  } finally {
    await browser.close();
  }
};

main();
