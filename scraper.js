const puppeteer = require('puppeteer');
const Excel = require('exceljs');

async function scrapePage(url) {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.goto(url);

  let address, street, extAddress, locality, country;

  const [el] = await page.$x('//*[@id="taplc_resp_rr_top_info_rr_resp_0"]/div/div[4]/div[1]/div/div/div[1]/span[2]/span[1]');
  if(el) {
    const streetRaw = await el.getProperty('textContent');
    street = await streetRaw.jsonValue();
  }

  const [el2] = await page.$x('//*[@id="taplc_resp_rr_top_info_rr_resp_0"]/div/div[4]/div[1]/div/div/div[1]/span[2]/span[2]');
  if(el2) {
    const extAddressRaw = await el2.getProperty('textContent');
    extAddress = await extAddressRaw.jsonValue();
  }

  const [el3] = await page.$x('//*[@id="taplc_resp_rr_top_info_rr_resp_0"]/div/div[4]/div[1]/div/div/div[1]/span[2]/span[3]');
  if(el3) {
    const localityRaw = await el3.getProperty('textContent');
    locality = await localityRaw.jsonValue();
  }
  
  const [el4] = await page.$x('//*[@id="taplc_resp_rr_top_info_rr_resp_0"]/div/div[4]/div[1]/div/div/div[1]/span[2]/text()[3]');
  const [el5] = await page.$x('//*[@id="taplc_resp_rr_top_info_rr_resp_0"]/div/div[4]/div[1]/div/div/div[1]/span[2]/text()[2]');
  if(el4) {
    const countryRaw = await el4.getProperty('textContent');
    country = await countryRaw.jsonValue();
  } else if(el5) {    
    const countryRaw = await el5.getProperty('textContent');
    country = await countryRaw.jsonValue();
  }

  await browser.close();

  if (locality) {
    address = street + ' ' + extAddress + ' ' + locality + country;
  } else {
    address = street + ' ' + extAddress + country;
  }

  if (!street && !extAddress && !locality && !country) {
    address = "N/A";
  }

  return address;
}

async function updatedLondonListingsAddress() {
  const baseURL = 'https://www.tripadvisor.co.uk';
  const workbook = new Excel.Workbook();
  const londonListings = await workbook.xlsx.readFile('./london_listings.xlsx');
  // const londonListings = await workbook.xlsx.readFile('./london_listings_sample.xlsx');
  // const londonListings = await workbook.xlsx.readFile('./sample_london_listings.xlsx');
  
  const worksheet = londonListings.getWorksheet(1);

  const rowCount = worksheet.lastRow._number + 1;
  console.log('******************************');
  console.log('Listing address update started');
  console.log('******************************');
  for (let i = 2; i < rowCount; i++) {
    let row = worksheet.getRow(i);
    let name = worksheet.getCell(`B${i}`);
    let path = worksheet.getCell(`J${i}`);
    let cell = worksheet.getCell(`L${i}`);
    const address = await scrapePage(baseURL + path);

    cell.value = address;
    row.commit();
    
    console.log('Restaurant name:', name.value);
    console.log('Restaurant address:', cell.value);
    console.log('===');
  }

  console.log('****** Writing to file ******');
  workbook.xlsx.writeFile('./london_listings_updated.xlsx');
  console.log('****** Writing complete ******');

  console.log('********************************');
  console.log('Listing address update completed');
  console.log('********************************');
}

updatedLondonListingsAddress();
