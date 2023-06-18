const puppeteer=require('puppeteer');
const xlsx = require('xlsx')

/*
=========================================================
=========Chỉnh sửa tham số chương trình==================
*/
let username="thuocthanthien";
let password="0985447007";
let fileName = 'Excel.xlsx'
let resultFileName = 'result.xlsx'
/*
=========================================================
=========================================================
*/
let crawlURL = 'https://giathuochapu.com/san-pham/page/';
// let redirectedURL = 'https://giathuochapu.com/dat-hang-nhanh/?swcfpc=1'

(async () => {
  console.log("Remember all production must be handle on " + fileName)
  console.log("Starting program...")
  const products = await getProd()
  console.log("Going to order " + products.length + " products!")

  const browser = await puppeteer.launch({headless: false});

  const page = await browser.newPage();
  await page.goto('https://giathuochapu.com/dang-nhap/');

  console.log("Opening Page...")

  await loggin(page)

  console.log("Redirecting to product page...")
  console.log("Start order...")

  for (const product of products) {
    await order(page, product.production, product.quantity)
    await order(page, product.production, product.quantity)
  }

  await handleResult()

  console.log("Shutting down...")
  await browser.close();
})();

async function getProd() {
  let data = []
  const buffer = xlsx.readFile(fileName)

  buffer.SheetNames.map(sheetName => {
    data = data.concat(xlsx.utils.sheet_to_json(buffer.Sheets[sheetName]))
  });

  return data.map(c => ({
    production: c.production,
    quantity: c.quantity || 0
  }))
}
const successProd = []
const failProd = []

async function order(page, production, quantity) {
  try{
    await page.goto("https://giathuochapu.com/san-pham/?swcfpc=1");
    await page.waitForNavigation({ waitUntil: "domcontentloaded" })
    console.log('Starting order ' + production)
    await page.type('#quick-search', production);
    await page.keyboard.press('Enter');
    await page.waitForNavigation()
    console.log('Make order ' + quantity + ' products!')
    // await page.type('.quantity_products', '1')
    await page.type('.quantity_products', quantity.toString())
    await page.keyboard.press('ArrowRight')
    await page.keyboard.press('Backspace')

    await page.keyboard.press('Enter');
    successProd.push(production)
    console.log('Order successful')
  }catch(error){
    failProd.push(production)
    console.log('Order fail ' + production)
  }
}

function delay(time) {
  return new Promise(function(resolve) {
    setTimeout(resolve, time)
  });
}

async function handleResult() {
  console.log('Handle order result...')
  const result = {
    successProd: successProd.map(c => ({production: c})),
    failProd: failProd.map(c => ({production: c}))
  }
  
  const workbook = xlsx.utils.book_new();
  const workSheetSuccess = xlsx.utils.json_to_sheet(result.successProd)
  const workSheetFail = xlsx.utils.json_to_sheet(result.failProd)
  await xlsx.utils.book_append_sheet(workbook,workSheetSuccess,'Thanh cong')
  await xlsx.utils.book_append_sheet(workbook,workSheetFail,'That bai')
  console.log('Success! Result save on ' + resultFileName)
  await xlsx.writeFileXLSX(workbook, resultFileName)
}

async function setTimeout(timeout) {
  return await new Promise(c => setTimeout(c, timeout));
}

async function loggin(page) {
  console.log("logging In...")
  await	page.type('#user_login',username);
  await	page.type('#user_pass',password);
  await page.click("label[id='remember_description']")
  await	page.click("button[value='1']");
  await	page.waitForNavigation();

  let current_url = page.url()
  if(current_url === 'https://giathuochapu.com/dang-nhap/'){
    console.log("Loggin fail!")
    console.log('Retry')
    await loggin(page)
  }else{
    console.log("Successful!")
  }
}