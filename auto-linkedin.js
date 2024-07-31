const puppeteer = require("puppeteer");
const fs = require("fs");
const XLSX = require("xlsx");

const chromedriver_path = "path_to_chromedriver";

async function loginLinkedIn(page, email, password) {
  await page.goto("https://www.linkedin.com/login");
  await page.type('input[name="session_key"]', email);
  await page.type('input[name="session_password"]', password);
  await page.click('button[type="submit"]');
  await page.waitForNavigation();
}

async function searchLinkedInProfile(page, name) {
  const searchQuery = `${name} organization_name`;

  await page.waitForSelector(".search-global-typeahead__input");
  const searchBox = await page.$(".search-global-typeahead__input");
  await searchBox.click({ clickCount: 3 });
  await searchBox.type(searchQuery);
  await searchBox.press("Enter");
  await page.waitForTimeout(2000);

  const noResults = await page.$x("//h2[text()='No results found']");
  if (noResults.length) {
    console.log("No results found for:", name);
    return null;
  }

  const noAccess = await page.$x(
    "//h2[@id='out-of-network-modal-header' and text()='You don’t have access to this profile']"
  );
  if (noAccess.length) {
    console.log("No access to profile for:", name);
    const gotItButton = await page.$x("//button[contains(., 'Got it')]");
    await page.waitForTimeout(2000);
    await gotItButton[0].click();
    await page.waitForTimeout(4000);
    return null;
  }

  const firstResult = await page.$x(
    '//span[contains(@class, "entity-result__title-text")]//a[contains(@class, "app-aware-link")]'
  );
  await firstResult[0].click();
  await page.waitForTimeout(2000);

  const currentUrl = page.url();
  if (currentUrl.includes("linkedin.com/jobs")) {
    await page.goto("https://www.linkedin.com/feed/");
    console.log("No results found for:", name);
    console.log("Current page is LinkedIn Jobs. Redirecting to LinkedIn feed.");
    return null;
  } else if (currentUrl.includes("linkedin.com/in")) {
    console.log("Profile URL:", currentUrl);
    return currentUrl;
  }

  return null;
}

(async () => {
  const workbook = XLSX.readFile("load_file.xlsx");
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const headers = data[0];
  const rows = data.slice(1);

  const browser = await puppeteer.launch({
    headless: false,
    args: ["--ignore-certificate-errors", "--ignore-ssl-errors"],
  });
  const page = await browser.newPage();
  page.setDefaultNavigationTimeout(60000);

  await loginLinkedIn(page, "email", "password");

  for (const [index, row] of rows.entries()) {
    const name = row[headers.indexOf("NAME")];

    if (!name) {
      console.log(`Name at index ${index} is empty or not valid.`);
      continue;
    }

    const profileUrl = await searchLinkedInProfile(page, name);
    if (profileUrl) {
      row[headers.indexOf("LINK")] = profileUrl;
    }
  }

  const newSheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
  XLSX.writeFile(newWorkbook, "save_to_file.xlsx");

  await browser.close();
})();
