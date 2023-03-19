const xl = require('excel4node');
const wb = new xl.Workbook();

const axios = require('axios')

const filename = 'thuocsivn.xlsx'

async function craw(){
    console.log("Program started!")
    console.log("Logging in!")
    const {data} = await login()
    console.log("Succeed!")
    const token = await data.data[0].bearerToken

    const prod = await getProd(token)

    console.log("Preparing to get all products!")
    const products = prod.data.map(c => ({
        productID: c.product?.productID?.toString() || "",
        productCode: c.product?.code || "",
        name: c.product?.name || "",
        unit: c.product?.unit || "",
        volume: c.product?.volume || "",
        price: c.sku?.retailPriceValue?.toString() || ""
    }))
    console.log("Succeed!")

    console.log("Preparing excel file!")
    await exportExcel(products)

    console.log("Done! File named: " + filename)
    console.log("Shutting down...")

}

async function login(){
    const url = `https://thuocsi.vn/backend/marketplace/customer/v1/authentication`

    const body = {
        username: "0985447007",
        password: "12345678",
        type: "CUSTOMER"
    }

    try{
        return await axios.post(url, body)
    }catch(error){
        console.log(error)
    }
}

async function getProd(token) {
    const header = {
        authorization: `Bearer ${token}`
    }

    const body = {
        "text": null,
        "offset": 0,
        "limit": 12222,
        "getTotal": true,
        "filter": {},
        "sort": "",
        "isAvailable": false,
        "searchStrategy": {
            "text": true,
            "keyword": true,
            "ingredient": true
        }
    }

    const url = `https://thuocsi.vn/backend/marketplace/product/v2/search/fuzzy`

    try {
        const {data} = await axios.post(url, body, {
            headers: header
        })

        return data
    }catch(error){
        console.log(error)
    }
}

async function exportExcel(products){
    const ws = wb.addWorksheet('Sheet 1');
    const style = wb.createStyle({
      font: {
        color: '#000000',
        size: 12,
      },
      numberFormat: '$#,##0.00; ($#,##0.00); -',
    });

    ws.cell(1, 1)
        .string("productId")
        .style(style);
    ws.cell(1, 2)
        .string("productCode")
        .style(style);
    ws.cell(1, 3)
        .string("name")
        .style(style);
    ws.cell(1, 4)
        .string("unit")
        .style(style);
    ws.cell(1, 5)
        .string("volume")
        .style(style);
    ws.cell(1, 6)
        .string("price")
        .style(style)

    products.forEach((item, i) => {
        ws.cell(i+2, 1)
            .string(item.productID)
            .style(style);
        ws.cell(i+2, 2)
            .string(item.productCode)
            .style(style);
        ws.cell(i+2, 3)
            .string(item.name)
            .style(style);
        ws.cell(i+2, 4)
            .string(item.unit)
            .style(style);
        ws.cell(i+2, 5)
            .string(item.volume)
            .style(style);
        ws.cell(i+2, 6)
            .string(item.price)
            .style(style);
    });

    wb.write(filename);
}

craw()
