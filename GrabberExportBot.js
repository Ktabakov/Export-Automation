const { chromium } = require(`playwright-chromium`);
const xlsx = require("xlsx");
const spreadsheet = xlsx.readFile('./Alte Bekannte.xlsx');
const sheets = spreadsheet.SheetNames;
const secondSheet = spreadsheet.Sheets[sheets[0]]; //sheet 1 is index 0
const expect = require(`mocha`);
const fs = require('fs');



(async () => {

    // const { test, expect } = require('@playwright/test');
    let names = [];

    for (let i = 2; ; i++) {
        const firstColumn = secondSheet['A' + i];
        if (!firstColumn) {
            break;
        }

        names.push(firstColumn.h);
    }


    // const browser = await chromium.launch({ headless: false });
    const browser = await chromium.launch({
        headless: false,
        downloadsPath: './downloads',
        acceptDownloads: true
    });
    const crContext = await browser.newContext({ acceptDownloads: true });

    const page = await crContext.newPage();
    page.setDefaultTimeout(6000)

    await page.goto('***********************');

    await page.click('[placeholder="Username"]');

    await page.fill('[placeholder="Username"]', '************');

    await page.click('[placeholder="Password"]');

    await page.fill('[placeholder="Password"]', '*************');

    await page.click('text=Login');

    await page.waitForTimeout(1000);

    // await page.goto(`****************`)
    // await page.click(1286, 36)



    for (let index = 0; index < names.length; index++) {
        try {
            await page.goto(`https:*****************`)
            await page.waitForTimeout(3000);
            await page.click(`#frame_wrapper > div > div.row.mx-auto.w-100.bg-light.action_button_row > div > div > div.col.px-0 > i`)
            await page.waitForTimeout(1000);

            await page.click('#select2-id_sub_accounts-container >> text=Search...');

            await page.fill('input[role="searchbox"]', `${names[index]}`);
            await page.waitForTimeout(3000);
            await page.click('li:nth-child(1)');

            await page.click('text=Filter');
            await page.waitForTimeout(1000);

            // const [page1] = await Promise.all([
            //     page.waitForEvent('popup'),
            //     page.waitForEvent('download'),
            //     page.click(``)
            // ]);
            await page.waitForTimeout(5000);


            const [download] = await Promise.all([
                page.waitForEvent("download"), // wait for download to start
                page.click("#frame_wrapper > div > div.row.mx-auto.w-100.bg-light.action_button_row > div > div > div:nth-child(1) > button"),   //vda_qmc_eReader.exe
            ]);

            await page.waitForTimeout(3000);

            const path = await download.path();

            await download.saveAs(`./downloads`);

            await page.waitForTimeout(2000);

        } catch (error) {
            await page.waitForTimeout(1000);
            continue;
        }
    }

    // browser.close();

})();