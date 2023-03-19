const xl = require('excel4node');
const wb = new xl.Workbook();

const axios = require('axios')

const filename = 'thuocsivn.xlsx'

async function craw(){
    console.log("Program started! 🚀🚀")
    console.log("Logging in!")
    const {data} = await login()
    console.log("Succeed!")
    const token = await data.data[0].bearerToken

    console.log("Preparing to get all products!")
    const prod = await getProd(token)

    const products = prod.map(c => ({
        productID: c.product?.productID?.toString() || "",
        productCode: c.product?.code || "",
        registrationNumber: c.product?.registrationNumber || "",
        name: c.product?.name || "",
        unit: c.product?.unit || "",
        volume: c.product?.volume || "",
        price: c.sku?.retailPriceValue?.toString() || "",
        stock: c.sku.status || ""
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
        "limit": 1000,
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
        let totalProd = []

        for (let i = 0; i < 12; i++) {
            const {data} = await axios.post(url, {
                ...body,
                offset: i*1000,
                limit: 1000
            }, {
                headers: header
            })

            totalProd = [...totalProd, ...data.data]
            console.log("Crawled " + Number(totalProd.length) + " products...")
        }

        console.log("Craw total Prod: " + totalProd.length)

        return totalProd
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
        .string("registrationNumber")
        .style(style);
    ws.cell(1, 4)
        .string("name")
        .style(style);
    ws.cell(1, 5)
        .string("unit")
        .style(style);
    ws.cell(1, 6)
        .string("volume")
        .style(style);
    ws.cell(1, 7)
        .string("price")
        .style(style)
    ws.cell(1, 8)
        .string("stock")
        .style(style)

    products.forEach((item, i) => {
        ws.cell(i+2, 1)
            .string(item.productID)
            .style(style);
        ws.cell(i+2, 2)
            .string(item.productCode)
            .style(style);
        ws.cell(i+2, 3)
            .string(item.registrationNumber)
            .style(style);
        ws.cell(i+2, 4)
            .string(item.name)
            .style(style);
        ws.cell(i+2, 5)
            .string(item.unit)
            .style(style);
        ws.cell(i+2, 6)
            .string(item.volume)
            .style(style);
        ws.cell(i+2, 7)
            .string(item.price)
            .style(style);
        ws.cell(i+2, 8)
            .string(item.stock === "OUT_OF_STOCK" ? "Hết hàng" : "")
            .style(style);
    });

    wb.write(filename);
}

craw()
