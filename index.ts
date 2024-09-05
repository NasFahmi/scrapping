import puppeteer from 'puppeteer';
import * as XLSX from 'xlsx';
import * as path from 'path';
import * as fs from 'fs';
async function scrapeTokopedia() {
    const browser = await puppeteer.launch({ 
        executablePath: 'C:/Program Files/Google/Chrome/Application/chrome.exe',
        headless: false
    });
    const page = await browser.newPage();
    // const url = 'https://www.tokopedia.com/search?st=&q=flex%20skincare&srp_component_id=02.01.00.00&srp_page_id=&srp_page_title=&navsource=%27';
    const url = 'https://www.tokopedia.com/';
    const xlsxName = 'Handphone';
    const searchQuery = 'Handphone';
    // const url = 'file:///E:/Media/Project/Javascript/scrapping/index.html';
    
    await page.goto(url, { waitUntil: 'networkidle2' });
    await page.setViewport({ width: 1280, height: 1024 });

     // Wait for the search input to be available and fill it
     await page.waitForSelector('input[data-unify="Search"]');
     await page.type('input[data-unify="Search"]', searchQuery);
     await page.keyboard.press('Enter');
    
    await smoothScroll(page);
    


    // Scrape the content after loading
    const items = await page.evaluate(() => {
        // document.querySelector('[class="css-3017qm exxxdg63"]');
        const products: any[] = [];
        document.querySelectorAll('.css-5wh65g').forEach(product => {
            const link = product.querySelector('a[class*="Nq8NlC5Hk9KgVBJzMYBUsg=="]')?.getAttribute('href');
            const gambar = product.querySelector('img[class*="css-1c345mg"]')?.getAttribute('src');
            const nama_brand = product.querySelector('[class="OWkG6oHwAppMn1hIBsC3pQ=="]')?.textContent?.trim();
            
            // Check for discounted price first
            // ! diskon
            let harga = product.querySelector('[class="en+9Xhk5rmGNLiUfSuIuqg=="]')?.textContent?.trim(); //div, span ,dll. a. 
            // If no discounted price, get the regular price
            if (!harga) {
                harga = product.querySelector('[class="_8cR53N0JqdRc+mQCckhS0g== "]')?.textContent?.trim();
            }
            
            const rating = product.querySelector('[class="nBBbPk9MrELbIUbobepKbQ=="]')?.textContent?.trim();
            const terjual = product.querySelector('[class="eLOomHl6J3IWAcdRU8M08A=="]')?.textContent?.trim();
            const toko = product.querySelector('[class*="X6c-fdwuofj6zGvLKVUaNQ=="]')?.textContent?.trim();
            // console.log(`nama brand : ${nama_brand}, gambar : ${gambar}, link : ${link}, harga : ${harga}, rating : ${rating}, terjual : ${terjual} toko : ${toko}`)
            products.push({
                nama_brand: nama_brand || 'N/A',
                harga: harga || 'N/A',
                rating: rating || '0.0',
                terjual: terjual || '0',
                link: link || 'N/A',
                toko: toko || 'N/A',
                gambar: gambar || 'N/A'
            });
        });
        return products;
    });

    // await browser.close();
    console.log(`Number of data = ${items.length}`);
    const allNACount = items.filter(item => 
        Object.values(item).every(value => value === 'N/A')
    ).length;

    console.log(`Number of items with all values 'N/A': ${allNACount}`);
    // Create Excel file
    const worksheet = XLSX.utils.json_to_sheet(items);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, `Products ${xlsxName}`);
     // Create output folder if it doesn't exist
     const outputFolder = path.join(__dirname, 'output');
     if (!fs.existsSync(outputFolder)) {
         fs.mkdirSync(outputFolder);
     }
 
     // Generate filename and full path
     const filename = `tokopedia_products_${searchQuery.replace(/\s+/g, '_')}_${new Date().toISOString().split('T')[0]}.xlsx`;
     const fullPath = path.join(outputFolder, filename);

    // Write to file
    XLSX.writeFile(workbook, fullPath);

    console.log(`Scraping completed. Data saved to tokopedia_products_${xlsxName}.xlsx`);

    await browser.close();
    
}
async function smoothScroll(page: any) {
    await page.evaluate(async () => {
        await new Promise<void>((resolve) => {
            let totalHeight = 0;
            const distance = 100; // Scroll 100px at a time
            const timer = setInterval(() => {
                const scrollHeight = document.body.scrollHeight;
                window.scrollBy(0, distance);
                totalHeight += distance;

                if (totalHeight >= scrollHeight) {
                    clearInterval(timer);
                    resolve();
                }
            }, 200); // Scroll every 200ms
        });
    });
}

scrapeTokopedia();