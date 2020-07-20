const xlsx = require('xlsx')
const puppeteer = require('puppeteer');
const cheerio  = require('cheerio');

const wkbk = xlsx.readFile('./files/file.xlsx')
const ws = wkbk.Sheets.Book
const range = wkbk.Sheets.Book['!ref']
const r = /:\w\d+/g
const endRange = range.match(r)[0].slice(2)


function getRandomInt(max) {
      return Math.floor(Math.random() * Math.floor(max));
}

async function excel() {
    let browser, page;
    [browser, page] = await getPage()
    let id
    let title
    let link

    for(var x = 2; x <= endRange; x++) {
        await setTimeout(() => {},getRandomInt(15000))
        id = ws[`A${x}`].w
        title = (ws[`C${x}`].w)
        link = ws[`H${x}`].w
        console.log(`Looking up ${title}`)
        pageTitle = await retrieveTitle(link, page)
        console.log(`\tFound ${pageTitle}`)
        title = title.toLowerCase().replace(/\s+/g, '');
        processedPageTitle = pageTitle.toLowerCase().replace(/\s+/g, '');

        if(title === processedPageTitle) {
            console.log(`\tTitles matched`)
            ws[`L${x}`].v = 'MATCH'
        }else if(processedPageTitle.slice(0, title.length) === title){
            console.log('\tShortened titles matched')
                ws[`L${x}`].v = 'SHORT MATCH'
        } else {
            console.error(`\tTitles did not match`)
            ws[`L${x}`].v = pageTitle
        }
        if(x%100 == 0) {
            console.log('Multiple of 100 hit. Saving work, and refreshing browser instance')
            xlsx.writeFile(wkbk, './files/test.xlsx')
            await browser.close();
            [browser,page] = await getPage();
        }
    }
    xlsx.writeFile(wkbk, './files/test.xlsx')
    await browser.close();
    console.log('Exiting Process')

}
 
async function getPage() {
    let browser = await puppeteer.launch();
    let page = await browser.newPage();
    return [browser,page];
}

async function retrieveTitle(link, page) {
    try {
        await page.goto(link);
        await page.waitForSelector('#fullView', { timeout: 4000 });

        const body = await page.evaluate(() => {
                return document.querySelector('body').innerHTML;
                });
        const $ = cheerio.load(body)
        const title = $('h3.item-title').text().slice(1)
        return title

    } catch (error) {
        console.log(error);
        return 'Unable to retrieve title'
    }
}
excel()

