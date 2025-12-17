// scrape_statsguru_puppeteer.js
// Usage: node scrape_statsguru_puppeteer.js
// npm i puppeteer cheerio csv-writer

const puppeteer = require('puppeteer');
const cheerio = require('cheerio');
const fs = require('fs');
const path = require('path');
const createCsvWriter = require('csv-writer').createObjectCsvWriter;

// --- CONFIGURATION ---
const OUTPUT_DIR = 'F:\\Data Analytics\\Projects\\csk_analysis\\data\\raw_temp';

// Array defining the full, correct URLs for each record type,
// including all necessary parameters like team ID, class, results, and date span.
const RECORDS_TO_SCRAPE = [
    { 
        name: 'team_results', 
        url: 'https://stats.espncricinfo.com/ci/engine/team/335974.html?class=6;home_or_away=1;home_or_away=2;home_or_away=3;result=1;result=2;result=3;result=5;template=results;type=team;view=results' 
    },
    { 
        name: 'partnership', 
        url: 'https://stats.espncricinfo.com/ci/engine/stats/index.html?class=6;home_or_away=1;home_or_away=2;home_or_away=3;home_or_away=4;result=1;result=2;result=3;result=5;spanmin1=13+mar+2008;spanval1=span;team=4343;template=results;type=fow' 
    }
];
// ---------------------

// Utility function to introduce a delay (Fixes potential TypeError: page.waitForTimeout)
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function extractLargestTableFromHtml(html) {
  const $ = cheerio.load(html);
  const tables = $('table');
  if (tables.length === 0) return null;

  let largest = null;
  let maxCells = 0;
  tables.each((i, t) => {
    const $t = $(t);
    const rows = $t.find('tr').length;
    const cols = Math.max(...$t.find('tr').map((i, r) => $(r).find('td,th').length).get(), 0);
    const cells = rows * (cols || 1);
    if (cells > maxCells) {
      maxCells = cells;
      largest = $t;
    }
  });

  if (!largest) return null;

  // build rows
  const rows = [];
  largest.find('tr').each((i, tr) => {
    const cols = $(tr).find('th,td').map((j, td) => $(td).text().trim()).get();
    if (cols.length) rows.push(cols);
  });
  if (rows.length === 0) return null;

  const header = rows[0];
  const data = rows.slice(1).map(r => {
    const obj = {};
    // Fallback for empty header cells
    header.forEach((h, idx) => obj[h || `col${idx+1}`] = r[idx] || '');
    return obj;
  });

  return { header, data };
}

async function saveJsonCsv(data, outPrefix) {
  if (!data || !data.data) return;

  // Ensure the directory exists
  try {
      if (!fs.existsSync(OUTPUT_DIR)) {
          fs.mkdirSync(OUTPUT_DIR, { recursive: true });
          console.log(`Created directory: ${OUTPUT_DIR}`);
      }
  } catch (err) {
      console.error(`Failed to create directory ${OUTPUT_DIR}:`, err);
      return;
  }

  // Construct full paths
  const jsonPath = path.join(OUTPUT_DIR, `${outPrefix}.json`);
  const csvPath = path.join(OUTPUT_DIR, `${outPrefix}.csv`);

  // Write JSON
  fs.writeFileSync(jsonPath, JSON.stringify(data.data, null, 2), 'utf8');

  // Write CSV
  const csvWriter = createCsvWriter({
    path: csvPath,
    // Ensure header object is correctly mapped for csv-writer
    header: data.header.map(h => ({ id: h || `col${data.header.indexOf(h)+1}`, title: h || `col${data.header.indexOf(h)+1}` }))
  });
  await csvWriter.writeRecords(data.data);
  console.log(`Saved ${path.basename(jsonPath)} and ${path.basename(csvPath)} to ${OUTPUT_DIR} (${data.data.length} rows)`);
}

async function scrapeRecord(page, record) {
    // FIX: Use the complete URL directly from the RECORDS_TO_SCRAPE array
    let url = record.url;
    const allRows = [];
    let pageIndex = 0;

    console.log(`\n======================================================`);
    console.log(`STARTING SCRAPE for: ${record.name.toUpperCase()}`);
    console.log(`======================================================`);

    while (url) {
        console.log(`Loading page ${pageIndex + 1} for ${record.name}: ${url}`);
        
        await page.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });

        // Wait a tick for DOM to settle
        await delay(700);

        const html = await page.content();
        const extracted = await extractLargestTableFromHtml(html);
        
        if (extracted && extracted.data.length) {
            // Add to aggregator
            extracted.data.forEach(r => allRows.push(r));
        } else {
            console.log(`No main table detected on page ${pageIndex + 1} for ${record.name}. This often means the table is empty or the URL is incorrect.`);
        }

        // Check for "Next" or pagination link
        const nextHref = await page.evaluate(() => {
            // Search for anchor containing 'next', '›', or '»' text
            const anchors = Array.from(document.querySelectorAll('a'));
            // Refined search to only look at pagination links
            const next = anchors.find(a => (/next|›|»/i.test(a.textContent) && a.href.includes('stats.espncricinfo.com/ci/engine/')));
            return next ? next.href : null;
        });

        if (nextHref && nextHref !== url) {
            url = nextHref;
            pageIndex += 1;
            // Polite random delay between pages
            await delay(1000 + Math.floor(Math.random()*1000));
        } else {
            url = null; // No next page found, stop loop
        }
    }

    // Final Step: Save aggregated CSV/JSON file for this record type
    if (allRows.length) {
        const header = Object.keys(allRows[0]);
        const aggregated = { header, data: allRows };
        await saveJsonCsv(aggregated, `${record.name}_allpages`);
        console.log(`*** Successfully aggregated ALL ${allRows.length} rows for ${record.name} into ${record.name}_allpages.csv ***`);
    } else {
        console.log(`*** Completed ${record.name} scrape. No data was found or aggregated. ***`);
    }
}


async function run() {
  const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox'] });
  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64)');
  
  // Loop through all defined record types
  for (const record of RECORDS_TO_SCRAPE) {
      await scrapeRecord(page, record);
  }

  await browser.close();
  console.log('\n✅ All scraping tasks completed and files saved.');
}

run().catch(err => {
  console.error('An error occurred during the scraping process:', err);
});