const arrLinks = require('./links')
const puppeteer = require('puppeteer');
const excelJS = require("exceljs")

async function getData(link) {
    const browser = await puppeteer.launch({ executablePath: 'C:/Program Files/Google/Chrome/Application/chrome', headless: false, args: ['--no-sandbox'] });
    const page = await browser.newPage();
    await page.setDefaultNavigationTimeout(0);
    await page.goto(link, { waitUntil: 'domcontentloaded' })
    const tags = await page.evaluate(async () => {
        const tag_title = document.title;
        const tag_h1 = document.querySelectorAll('h1');
        const tag_meta = document.querySelector('meta[name="description"]').content
        const tag_short = document.querySelectorAll('.vtex-rich-text-0-x-paragraph--seo-short-description')
        const tag_long = document.querySelectorAll('.vtex-rich-text-0-x-wrapper--seo-long-description')

        const h1 = [];
        let short = '';
        let long = '';

        for (let e of tag_h1) {
            h1.push(e.innerText);
        }
        for (let e of tag_short) {
            short = e.innerText;
        }
        for (let e of tag_long) {
            long = e.innerText;
        }
        return { title: tag_title, h1: h1, meta: tag_meta, short: short, long: long };
    });

    allData.push({ url: link, ...tags })
    await browser.close();
}

let allData = [];
process.setMaxListeners(0);
const bar = new Promise(async(resolve, reject) => {
    await arrLinks.reduce(async (a, link, index, array) => {
        await a;
        await getData(link);
        if (index === array.length - 1) resolve();
    }, Promise.resolve());
});
bar.then(async () => {
    console.log('rows =>', allData.length)
    const workbook = new excelJS.Workbook();  // Create a new workbook
    const worksheet = workbook.addWorksheet("SEO TAGS"); // New Worksheet
    const path = "./files";  // Path to download excel
    // Column for data in excel. key must match data key
    worksheet.columns = [
        { header: "#", key: "count", width: 10 },
        { header: "URL", key: "url", width: 10 },
        { header: "TITULO", key: "title", width: 10 },
        { header: "H1", key: "h1", width: 10 },
        { header: "META DESCRIPCIÓN", key: "meta", width: 10 },
        { header: "DESCRIPCIÓN CORTA", key: "short", width: 10 },
        { header: "DESCRIPCIÓN LARGA", key: "long", width: 10 },
    ];
    // Looping through User data
    let counter = 1;
    allData.forEach((data) => {
        data = { count: counter, ...data };
        worksheet.addRow(data); // Add data in worksheet
        counter++;
    });
    // Making first line in excel bold
    worksheet.getRow(1).eachCell((cell) => {
        cell.font = { bold: true };
    });
    try {
        const data = await workbook.xlsx.writeFile(`${path}/seo.xlsx`)
            .then(() => {
                console.log({
                    status: "success",
                    message: "file successfully downloaded",
                    path: `${path}/seo.xlsx`,
                });
            });
    } catch (err) {
        console.log({
            status: "error",
            message: "Something went wrong",
            description: err
        });
    }
});