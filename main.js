import puppeteer from "puppeteer";
import cliProgress from "cli-progress";
import chalk from "chalk";
import { GoogleSpreadsheet } from "google-spreadsheet";
import dayjs from "dayjs";
import { readFile } from "fs/promises";

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
 * 延遲 X 秒
 * @param {Number} time time
 */
const delay = (time) => {
  return new Promise(function (resolve) {
    setTimeout(resolve, time);
  });
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
      const tdFormType = await tr.$$("td.dtr-control");
      const tdFilingDate = await tr.$$("td.sorting_1");
      if (tdFormType.length === 1 && tdFilingDate.length === 1) {
        const formType = await tdFormType[0].evaluate((el) =>
          el.textContent.trim()
        );
        const filingDate = await tdFilingDate[0].evaluate((el) =>
          el.textContent.trim()
        );

        if (formType === keyword && year === dayjs(filingDate).format("YYYY")) {
          const href = await tr.$eval("a", (element) => element.href);
          links.push(href);
        }
      }
    }
    return links;
  } catch (error) {
    return null;
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
  const sheet = await getSheet(
    "1l8Wtu7aLmfPMinx67Ja6EgJoKw8yELcStHyKTB8SCnY",
    "1248092653"
  );
  const rows = await sheet.getRows();

  const bar1 = new cliProgress.SingleBar(
    {},
    cliProgress.Presets.shades_classic
  );

  bar1.start(rows.length, 0);

  for (let i = 0; i < rows.length; i++) {
    try {
      const row = rows[i];
      const rowData = row._rawData;

      // const secLink = sheet.getCell(i + 1, 6);
      const year = dayjs(rowData[3]).format("YYYY");
      const CIKNumber = rowData[4];
      const secLink = `https://www.sec.gov/edgar/browse/?CIK=${CIKNumber}`;
      const hasProcessed = !!rowData[7];

      if (!hasProcessed && CIKNumber && secLink) {
        const page = await browser.newPage();
        await page.goto(secLink);
        await page.setViewport({ width: 1080, height: 1024 });
        await delay(500);

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
          // console.log(chalk.bgBlue("link10KList", link10KList));
          // console.log(chalk.bgCyan("result10K", result10K));

          if (result10K.length) {
            row["10-K"] = result10K.join("、");
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
            // console.log(chalk.bgBlueBright("link10QList", link10QList));
            // console.log(chalk.bgGreenBright("result10Q", result10Q));

            if (result10Q.length) {
              row["10-Q"] = result10Q.join("、");
            } else if (link10QList.length === 0) {
              row["10-Q"] = "no file";
            } else {
              row["10-Q"] = "no match";
            }
          }
        }
        await save(row);

        await page.close();
      }
      bar1.increment();
    } catch (error) {
      console.log(error);
    }
  }
  bar1.stop();
};

main();
