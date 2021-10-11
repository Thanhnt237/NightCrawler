const puppeteer=require('puppeteer');
const readline = require('readline');
const xl = require('excel4node');
const wb = new xl.Workbook();
const fs = require('fs');

/*
=========================================================
=========Chỉnh sửa tham số chương trình==================
*/
var username="";
var password="";
var fileName = 'Excel.xlsx'
const ENDPAGE = 218;
/*
=========================================================
=========================================================
*/

var URL = 'https://giathuochapu.com/dang-nhap/';
var crawlURL = 'https://giathuochapu.com/san-pham/page/';


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
  const browser=await puppeteer.launch({headless:true});

  const page=await browser.newPage();
  await page.goto(URL);

  console.log("Opening Page...")
  console.log("loging In...")
  await	page.type('#user_login',username);
  await	page.type('#user_pass',password);
  await	page.click("button[value='1']");
  console.log("Successful!")
  await	page.waitForNavigation();

  console.log("Crawling data...")
  for(var i=1;i<=ENDPAGE;i++){
    await page.goto(crawlURL + i);
    console.log("Crawling page " + i + "/" + ENDPAGE)
    await page.waitForSelector(".product-item", {timeout: 60000}).then(async ()=>{
        const products = await page.evaluate(() => {
          let items = document.querySelectorAll(".product-item");
          let product = [];
          items.forEach(item => {
          if(item.children[1].childElementCount == 3){
            product.push({
              title: item.children[1].children[0].innerText,
              type: item.children[1].children[2].innerText.replace("Nhóm: ",""),
              price: item.children[2].innerText.slice(0,-2)
            });
          }else
          if(item.children[1].childElementCount == 4){
            product.push({
              title: item.children[1].children[0].innerText,
              type: item.children[1].children[3].innerText.replace("Nhóm: ",""),
              price: item.children[2].innerText.slice(0,-2)
            });
          }
          else{
            product.push({
              title: item.children[1].children[0].innerText,
              type: item.children[1].children[1].innerText.replace("Nhóm: ",""),
              price: item.children[2].innerText.slice(0,-2)
            });
          }
          });
          return product;
        });
        if(globalProducts.length == 0){
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

  globalProducts.forEach((item, i) => {
    ws.cell(i+2, 1)
      .string(item.title)
      .style(style);
    ws.cell(i+2, 2)
      .string(item.type)
      .style(style);
    ws.cell(i+2, 3)
      .string(item.price.replace(" đ", ""))
      .style(style);
  });

  wb.write(fileName);
  console.log("Done! File named: " + fileName)
  console.log("Shutting down...")
  //console.log(globalProducts)
  await browser.close();

})();
