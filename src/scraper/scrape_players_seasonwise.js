// scrape_csk_seasonwise_v3.js
// Usage: node scrape_csk_seasonwise_v3.js
// Status: Uses temporary JSON files to preserve BBI format and consolidates data into 3 final Excel files.

const puppeteer = require('puppeteer');
const cheerio = require('cheerio');
const fs = require('fs');
const path = require('path');
// Removed csv-writer as we're switching to JSON for temp storage
const XLSX = require('xlsx');

// --- CONFIGURATION ---
const OUTPUT_BASE_DIR = 'F:\\Data Analytics\\Projects\\csk_analysis\\data\\';
const RAW_DIR = path.join(OUTPUT_BASE_DIR, 'raw_temp'); // Temp directory for individual JSONs
const FINAL_DIR = path.join(OUTPUT_BASE_DIR, 'raw_temp'); // Final output directory
const TEAM_ID = 4343; // Chennai Super Kings
const START_YEAR = 2007;
const END_YEAR = 2025;
const RECORD_TYPES = ['batting', 'bowling', 'fielding']; // Now includes all three types
// ----------------------

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

/**
 * Builds the ESPN Cricinfo URL, checking both single year and year-range formats.
 */
/**
 * Builds the ESPN Cricinfo URL, checking both single year and year-range formats,
 * with a specific exception for the 2010 season to ensure IPL data is captured.
 */
function buildSeasonUrl(type, season, use_range) {
    const TEAM_ID = 4343; // CSK ID, keeping it local for clarity in this function

    // --- 🚨 2010 Season EXCEPTION ---
    if (season === 2010) {
        console.log(`\n🚨 Using specific 2010 URL structure for ${type.toUpperCase()}.`);
        // Use the highly specific URL provided for 2009/10 season stats
        return `https://stats.espncricinfo.com/ci/engine/stats/index.html?class=6;home_or_away=1;home_or_away=2;home_or_away=3;home_or_away=4;orderby=runs;result=1;result=2;result=3;result=5;season=2009%2F10;team=${TEAM_ID};template=results;type=${type}`;
    }
    // --- END EXCEPTION ---

    // Standard URL construction for all other years (2007-2009, 2011-2025)
    let seasonParam;
    if (use_range) {
        const endYear = (season + 1).toString().slice(-2);
        seasonParam = `${season}%2F${endYear}`;
    } else {
        seasonParam = season;
    }

    // Fixed URL (without spanmin1) for general seasons
    return `https://stats.espncricinfo.com/ci/engine/stats/index.html?class=6;season=${seasonParam};team=${TEAM_ID};template=results;type=${type}`;
}


async function extractLargestTableFromHtml(html) {
  const $ = cheerio.load(html);
  const tables = $('table');
  if (tables.length === 0) return null;

  let largest = null, maxCells = 0;
  tables.each((i, t) => {
    const $t = $(t);
    const rows = $t.find('tr').length;
    const dataRows = $t.find('tr:has(td)');
    const cols = Math.max(...dataRows.map((i, r) => $(r).find('td,th').length).get(), 0);
    const cells = rows * (cols || 1);
    
    // Heuristic: Ensure the table is large enough to be a stats table
    if (cells > maxCells && rows > 5 && cols > 5) { 
      maxCells = cells;
      largest = $t;
    }
  });

  if (!largest) return null;
  const rows = [];
  const header = largest.find('tr:has(th)').first().find('th,td').map((j, td) => $(td).text().trim()).get();
  
  largest.find('tr:has(td)').each((i, tr) => {
    const cols = $(tr).find('td').map((j, td) => $(td).text().trim()).get();
    if (Math.abs(cols.length - header.length) <= 1 && cols.length > 0) {
      rows.push(cols);
    }
  });

  if (rows.length === 0 || header.length === 0) return null;

  const data = rows.map(r => {
    const obj = {};
    header.forEach((h, idx) => obj[h || `col${idx+1}`] = r[idx] || '');
    return obj;
  });
  return { header, data };
}

/**
 * Saves data to a temporary JSON file, and implements the BBI text fix.
 * @param {object} data - The extracted table data { header, data }.
 * @param {string} filename - The name for the output JSON file.
 */
async function saveCsv(data, filename) {
    if (!data || !data.data || data.data.length === 0) return;

    // 1. Identify the 'BBI' column index (case-insensitive)
    const bbiIndex = data.header.findIndex(h => h.toUpperCase() === 'BBI');

    // 2. Data Cleaning: Convert BBI column to explicit strings (The fix)
    if (bbiIndex !== -1) {
        // console.log(`⚙️ Cleaning data: Found 'BBI' column. Converting values to string...`);
        data.data.forEach(record => {
            const headerKeys = Object.keys(record);
            const bbiKey = headerKeys[bbiIndex];
            if (bbiKey && record[bbiKey]) {
                // Ensure the value is explicitly a string and trim whitespace
                record[bbiKey] = String(record[bbiKey]).trim();
            }
        });
    }

    // 3. Save as JSON instead of CSV to avoid conversion issues before XLSX
    const jsonPath = path.join(RAW_DIR, `${filename}.json`);
    fs.writeFileSync(jsonPath, JSON.stringify(data.data, null, 2));
    
    console.log(`✅ Saved temp JSON: ${filename}.json (${data.data.length} rows)`);
}

/**
 * Scrapes data, trying both URL format variations.
 */
async function scrapeSeason(page, type, season) {
  let extracted = null;
  
  // 1. TRY SINGLE YEAR FORMAT (e.g., ?season=2015)
  let url = buildSeasonUrl(type, season, false);
  console.log(`\n🔍 Attempt 1 (Single Year): ${type.toUpperCase()} for ${season}...`);
  await page.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });
  await delay(1000);
  extracted = await extractLargestTableFromHtml(await page.content());

  if (extracted && extracted.data.length > 0) {
    await saveCsv(extracted, `${type}_${season}`); // Now saves JSON
    return;
  }
  
  // 2. TRY YEAR RANGE FORMAT (e.g., ?season=2015%2F16)
  url = buildSeasonUrl(type, season, true);
  console.log(`\n🔍 Attempt 2 (Year Range): ${type.toUpperCase()} for ${season}-${season+1}...`);
  await page.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });
  await delay(1000);
  extracted = await extractLargestTableFromHtml(await page.content());

  if (extracted && extracted.data.length > 0) {
    await saveCsv(extracted, `${type}_${season}`); // Now saves JSON
    return;
  } 

  console.log(`⚠️ No ${type} data found for season ${season} in either format.`);
}


/**
 * Consolidates the scraped data into three separate Excel files (one per record type), 
 * with each season as a sheet. Reads from temporary JSON files.
 */
async function combineAndSave() {
  console.log('\n📘 Consolidating all JSON data into 3 final Excel files (one per type)...');
  
  for (const type of RECORD_TYPES) {
    const wb = XLSX.utils.book_new();
    let hasData = false;
    
    for (let season = START_YEAR; season <= END_YEAR; season++) {
      const jsonPath = path.join(RAW_DIR, `${type}_${season}.json`);
      
      if (fs.existsSync(jsonPath)) {
        // Read JSON data
        const jsonData = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
        
        if (jsonData.length > 0) {
            // Convert JSON array of objects directly to a worksheet
            const ws = XLSX.utils.json_to_sheet(jsonData);
            
            // Use a consistent sheet name format: "Season_2008"
            const sheetName = `Season_${season}`; 
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
            hasData = true;
        }
      }
    }

    if (hasData) {
      const outPath = path.join(FINAL_DIR, `${type}_records_csk.xlsx`);
      XLSX.writeFile(wb, outPath);
      console.log(`✅ Final Excel file created: ${outPath}`);
    } else {
      console.log(`\n--- WARNING: No data found for ${type.toUpperCase()} across all seasons. No Excel file created. ---`);
    }
  }
}

async function run() {
  console.log('--- CSK Season-wise Scraper Started ---');
  // Setup directories
  if (fs.existsSync(RAW_DIR)) fs.rmSync(RAW_DIR, { recursive: true, force: true });
  fs.mkdirSync(RAW_DIR, { recursive: true });
  if (!fs.existsSync(FINAL_DIR)) fs.mkdirSync(FINAL_DIR, { recursive: true });

  const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox'] });
  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');
  
  for (let season = START_YEAR; season <= END_YEAR; season++) {
    for (const type of RECORD_TYPES) {
      // NOTE: CSK was suspended in 2016 and 2017. 
      // The script will correctly return 'No data found' for these years.
      await scrapeSeason(page, type, season);
      await delay(1500); // Respectful delay
    }
  }

  await browser.close();
  await combineAndSave();
  console.log('\n🏁 All scraping and merging completed successfully.');
}

run().catch(err => console.error('❌ Error:', err));