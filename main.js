import puppeteer from "puppeteer";
import cliProgress from "cli-progress";
import chalk from "chalk";
import { GoogleSpreadsheet } from "google-spreadsheet";
import dayjs from "dayjs";
import { readFile } from "fs/promises";
import delay from "./utils/delay.js";
import waitPageVisible from "./utils/waitPageVisible.js";
import save from "./utils/save.js";

// sheet https://docs.google.com/spreadsheets/d/1l8Wtu7aLmfPMinx67Ja6EgJoKw8yELcStHyKTB8SCnY/edit#gid=1248092653
// node main
/**
 * 取得 sheet data
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

  return sheet;
};

/**
 * 找到 links
 * @param {Object} page page
 * @param {String} year year
 * @param {String} keyword keyword
 */
async function getLinks(page, year, keyword) {
  try {
    const trs = await page.$$("#filingsTable tr");
    let links = [];
    for (const tr of trs) {
      const dataArr = await tr.$$eval("td", async (tdElements) => {
        return tdElements.map((ele) => ele.textContent);
      });
      if (dataArr.length) {
        const formType = dataArr[0];
        const reportingDate = dataArr[3].substring(0, 4);
        if (formType === keyword && reportingDate === year) {
          const href = await tr.$eval("a", (element) => element.href);
          links.push(href);
        }
      }
    }
    return links;
  } catch (error) {
    return [];
  }
}

/**
 * 找到關鍵字
 * @param {Object} browser browser
 * @param {Array} links links
 */
async function findKeywords(browser, links) {
  try {
    const match = [];
    const keywords = JSON.parse(
      await readFile(new URL("./findData.json", import.meta.url))
    );
    const newPage = await browser.newPage();
    await newPage.setViewport({ width: 1080, height: 1024 });

    for (let i = 0; i < links.length; i++) {
      await newPage.goto(links[i]);
      await waitPageVisible(newPage, "#xbrl-form-loading", "d-none");

      const content = await newPage.content();
      for (let j = 0; j < keywords.length; j++) {
        const keyword = keywords[j];
        const regex = new RegExp(keyword, "gi");
        if (content.match(regex)) {
          match.push(j + 1);
        }
      }
    }
    await newPage.close();

    return match;
  } catch (error) {
    return [];
  }
}

/**
 * 操作 Dom 元素
 * @param {Object} page page
 * @param {String} year year
 * @param {String} keyword keyword
 */
const domOperate = async (page, year, keyword) => {
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
    await formTypeSelect.type(keyword);
    await delay(1000);

    const fromYear = String(Number(year) - 2);
    const dateFromField = await page.$('input[id="filingDateFrom"]');
    await dateFromField.click({ clickCount: 3 });
    await dateFromField.type(fromYear);
    await dateFromField.press("Enter");

    // 選擇填報日期為 year 年的輸入框
    const toYear = Number(year) + 2;
    const dateToField = await page.$('input[id="filingDateTo"]');
    await dateToField.click({ clickCount: 3 });
    await dateToField.type(toYear <= 2023 ? String(toYear) : "2023");
    await dateToField.press("Enter");
    await delay(1000);
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
  const sheet = await getSheet(
    "1_peHmn2NjX_4R40aonPqiaytR8dVGei6OBRojKbgL1I",
    "1287819765"
  );
  const rows = await sheet.getRows();

  const bar1 = new cliProgress.SingleBar(
    {},
    cliProgress.Presets.shades_classic
  );

  bar1.start(rows.length, 0);

  for (let i = 0; i < 2500; i++) {
    try {
      const row = rows[i];
      const rowData = row._rawData;

      // const secLink = sheet.getCell(i + 1, 6);
      const year = dayjs(rowData[7]).format("YYYY");
      const CIKNumber = rowData[34];
      const secLink = `https://www.sec.gov/edgar/browse/?CIK=${CIKNumber}`;
      const hasProcessed = !!rowData[62];

      if (!hasProcessed && CIKNumber && secLink) {
        const page = await browser.newPage();
        await page.setViewport({ width: 1080, height: 1024 });
        await page.goto(secLink);
        await delay(100);
        await waitPageVisible(page, "#loading", "hidden");

        let link10KList = [];
        let link10QList = [];
        let result10K = [];
        let result10Q = [];

        const isError = await page.evaluate(() => {
          const errorDom = document.querySelector("#errorLoading");
          return errorDom ? !errorDom.classList.contains("hidden") : false;
        });

        if (isError) {
          row["10-K"] = "error page";
        } else {
          // 找 10-K
          await domOperate(page, year, "10-K");
          link10KList = await getLinks(page, year, "10-K");
          if (link10KList.length) {
            result10K = await findKeywords(browser, link10KList);
          }
          if (result10K.length) {
            row["10-K"] = String(result10K.join("、"));
          } else if (link10KList.length === 0) {
            row["10-K"] = "no file";
          } else {
            row["10-K"] = "no match";
          }

          if (result10K.length === 0) {
            // 找 10-Q
            await domOperate(page, year, "10-Q");
            link10QList = await getLinks(page, year, "10-Q");
            if (link10QList.length) {
              result10Q = await findKeywords(browser, link10QList);
            } else {
              row["10-Q"] = "no file";
            }
            if (result10Q.length) {
              row["10-Q"] = String(result10Q.join("、"));
            } else if (link10QList.length === 0) {
              row["10-Q"] = "no file";
            } else {
              row["10-Q"] = "no match";
            }
          }
        }
        await save(row);

        await delay(2000);
        await page.close();
      }
      bar1.increment();
    } catch (error) {
      console.log(error);
    }
  }
  bar1.stop();
  return;
};

main();
