const puppeteer=require('puppeteer');
const xl = require('excel4node');
const wb = new xl.Workbook();

/*
=========================================================
=========Chỉnh sửa tham số chương trình==================
*/
let username="thuocthanthien";
let password="0985447007";
let fileName = 'Excel.xlsx'
let ENDPAGE = 0;
/*
=========================================================
=========================================================
*/
let crawlURL = 'https://giathuochapu.com/san-pham/page/';
let redirectedURL = 'https://giathuochapu.com/dat-hang-nhanh/?swcfpc=1'

let globalProducts = [];

const ws = wb.addWorksheet('Sheet 1');
const style = wb.createStyle({
  font: {
    color: '#000000',
    size: 12,
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -',
});

(async () => {
  console.log("Starting program...")
  const browser = await puppeteer.launch({headless:true});

  const page = await browser.newPage();
  await page.goto('https://giathuochapu.com/dang-nhap/');

  console.log("Opening Page...")

  await loggin(page)

  console.log("Redirecting to product page...")
  console.log("Get EndPage...")
  await page.goto("https://giathuochapu.com/san-pham/?swcfpc=1");
  await page.waitForSelector(".page-numbers").then(async ()=>{
    ENDPAGE = await page.evaluate(()=>{
      return Number(document.querySelectorAll(".page-numbers")[3].innerText)
    })
  })
  console.log("Successful! We have " + ENDPAGE + " pages here!")

  console.log("Crawling data...")
  for(let i=1; i<=ENDPAGE; i++){
    await page.goto(`${crawlURL}${i}/?swcfpc=1`);
    console.log("Crawling page " + i + "/" + ENDPAGE)
    await page.waitForSelector(".product-item", {timeout: 60000}).then(async ()=>{
        const products = await page.evaluate(() => {
          let items = document.querySelectorAll(".product-item");
          let product = [];

          items.forEach(item => {
            product.push({
            title: item.querySelector(".entry-title") ? item.querySelector(".entry-title").innerText : "",
            type: item.querySelector(".product-type") ? item.querySelector(".product-type").innerText.replace("Nhóm: ","") : "",
            price: item.querySelector(".price") ? item.querySelector(".price").innerText.slice(0, -2) : "",
            tag: item.querySelector(".product-tag") ? item.querySelector(".product-tag").innerText : "",
            date: item.querySelector(".product-expire-date") ?  item.querySelector(".product-expire-date").innerText : ""
          });
          });
          return product;
        });

        if(globalProducts.length === 0){
          globalProducts = products
        }else{
          globalProducts = globalProducts.concat(products)
        }
      })
  }
  console.log(globalProducts.length + " items was crawled!")
  console.log("Creating Excel File...")

  ws.cell(1, 1)
    .string("Nhãn")
    .style(style);
  ws.cell(1, 2)
    .string("Nhóm")
    .style(style);
  ws.cell(1, 3)
    .string("Giá tiền")
    .style(style);
  ws.cell(1, 4)
      .string("Product Tag")
      .style(style);
  ws.cell(1, 5)
      .string("Date")
      .style(style);

  globalProducts.forEach((item, i) => {
    ws.cell(i+2, 1)
      .string(item.title)
      .style(style);
    ws.cell(i+2, 2)
      .string(item.type)
      .style(style);
    ws.cell(i+2, 3)
      .string(item.price)
      .style(style);
    ws.cell(i+2, 4)
        .string(item.tag)
        .style(style);
    ws.cell(i+2, 5)
        .string(item.date)
        .style(style);
  });

  wb.write(fileName);
  console.log("Done! File named: " + fileName)
  console.log("Shutting down...")

  await browser.close();

})();

async function loggin(page) {
  console.log("logging In...")
  await	page.type('#user_login',username);
  await	page.type('#user_pass',password);
  await page.click("label[id='remember_description']")
  await	page.click("button[value='1']");
  await	page.waitForNavigation({
    timeout: '100000'
  });

  let current_url = page.url()
  if(current_url === 'https://giathuochapu.com/dang-nhap/'){
    console.log("Loggin fail!")
    console.log('Retry')
    await loggin(page)
  }else{
    console.log("Successful!")
  }
}