/***
Author: Liem Pham
Tool: Get link admin
Version: 1.0
***/
const puppeteer = require('puppeteer');
const fs = require('fs');
var csvWriter = require('csv-write-stream');
var writer = csvWriter({ sendHeaders: true });
var colors = require('colors/safe');
var csvFilename = "myfile.csv";
// Require excel library
var excel = require('excel4node');
var file = "read.xlsx";
// writing to xlsx
function exportToExcel(data) {
    // Create a new instance of a Workbook class
    var workbook = new excel.Workbook();

    // Add Worksheets to the workbook
    var worksheet = workbook.addWorksheet('Sheet 1');

    const headings = ['Link'];//, '', 'Gift2','Gift3', 'Gift4', 'Gift5', 'Gift6', 'Gift7'];;

    // Writing from cell A1 to I1
    headings.forEach((heading, index) => {
        worksheet.cell(1, index + 1).string(heading);
    })

    // Writing from cell A2 to I2 , A3 to I3, .....
    data.forEach((item, index) => {
        worksheet.cell(index + 2, 1).string(item.link);
    });
    var today = new Date();
    var filename = "FileSai" + '-' + today.getDate() + '-' +(today.getMonth()+1) + '-' + today.getFullYear() + '.xlsx';
    workbook.write(filename);
}


    function getLink(file){
    const XLSX = require('xlsx');
    var workbook = XLSX.readFile(file);
    var sheet_name_list = workbook.SheetNames;
    var urls = [];
    data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]); 
    // for(var key in data){
    //     urls.push(data[key]['Article']); //Cho lon
    // }
    return data;
    };
(async() => {
    const browser = await puppeteer.launch({ headless: false, args:['--start-maximized' ] })
    const page = await browser.newPage();
    await page.setViewport({width: 1366, height: 1200});
    const urls = getLink(file);

    console.log(urls);

    await page.goto('https://dienmaycholon.vn/admin/users/login',{waitUntil: 'load'}); 
    await page.type('#username', 'thanhliem');
    await page.type('#passwords', 'beo4356143');
    await Promise.all([
        page.click('.button')
    ]);
    /***
    *
    ******************** add link pro text / tool sticker xuống link dưới 
    *
    **/
    //await page.setViewport({ width: 1366, height: 768});
    await page.goto('https://dienmaycholon.vn/admin/promocidcode/edit/id/16843',{waitUntil: 'load'});
    let arrInfo = [];
    for (var rs of urls) {
        try {
            // await Promise.all([
            //     page.click('.name_product')
            // ]);
            function sleep(ms) {
              return new Promise(resolve => setTimeout(resolve, ms))
            }
            let osap = await page.$('div.input_field #name_product')
            const art = "'" + rs.Article + "'" 
            await osap.click({clickCount: 3})
            await sleep(2000);
            await osap.type(art.replace(/[']/g,""))
            await sleep(1000);
            await Promise.all([
                page.click("div.input_field .loc_product",{clickCount: 10})

            ]);
            await sleep(1000)
            let co = await page.$(".bt_deal_false tbody input[type='checkbox']")
            await sleep(1000)
            if(co == null){
               console.log(colors.red("Lỗi => ") + rs.Article)
            }else{


               await page.click(".bt_deal_false thead input[type='checkbox']")
               await sleep(2000)
               await Promise.all([
                page.click(".bt_deal_false input[name='add']")
                ]);
               let idproduct = await page.evaluate('document.querySelector(".bt_deal_false tbody td input").getAttribute("idxx")')
               await sleep(1000)
               await page.click(".bt_deal_false thead input[type='checkbox']")
               // khai báo id ô nhập giá, nhập comment, nhập loại
               let ogiathitruong = "#thitruong"+idproduct; // ô giá thị trường
               let ogiany = "#price"+idproduct; // ô giá
               let ocomment = "#cke_comments-"+idproduct; // ô comment 
               let oqua = "#cke_descriptions-"+idproduct; // ô quà tặng
               let oloai = "#oftype-"+idproduct; // ô loai
               let ogiaqua = "#giaqua"+idproduct; // ô giá quà

               // chọn ô nhập giá, ô nhập comment
               let onhapgiathitruong = await page.$(ogiathitruong);
               let onhapgia = await page.$(ogiany);
               let onhapcomment  = await page.$(ocomment);
               let onhapqua  = await page.$(oqua);
               let onhaploai = await page.$(oloai);
               let onhapgiaqua = await page.$(ogiaqua);

               // data nhập vào km
               let giathitruong = "'" + rs.MarketPrice + "'";
               let gia = "'" + rs.Price + "'";
               let comment = "'" + rs.Note + ": '" + rs.Price + " - " + rs.Loai;
               let loai = rs.Loai;
               let giaqua = "'" + rs.Tonggiatri + "'";
               let qua1 = "'" + rs.Gift1 + "'";
               let qua2 = "'" + rs.Gift2 + "'"; 
               let qua3 = "'" + rs.Gift3 + "'"; 
               let qua4 = "Tặng Bộ Phiếu Mua Hàng 7.000.000đ";//rs.Gift4 
               let qua5 = "Tặng Combo Dịch Vụ Phiếu Mua Hàng 7000.000đ";//rs.Gift5 
               let qua6 = "Hỗ Trợ Trả Góp 0% Qua Các Công Ty Tài Chính Và Ngân Hàng";//rs.Gift6 
               //var qua = [qua1,qua2,qua3,qua4,qua5,qua6];

               // khu vực giả lập user
               // nhập giá thị trường
               let demnhapgiathitruong = await onhapgiathitruong.click({clickCount: 3})
               let nhapgiathitruong = await onhapgiathitruong.type(giathitruong.replace(/[']/g,""))
               // nhập giá NY
               let demnhapgia = await onhapgia.click({clickCount: 3})
               let nhapgia = await onhapgia.type(gia.replace(/[']/g,""))
               // nhập giá quà
               let demnhapgiaqua = await onhapgiaqua.click({clickCount: 3})
               let nhapgiaqua = await onhapgiaqua.type(giaqua.replace(/[']/g,""))
               // nhập loại
               let nhaploai = await onhaploai.select(loai);
                
               // for(var i = 0; i < qua.length; i++){
               //      const data = []
               //      var demqua = await onhapqua.click({clickCount: 3})
               //      var nhapqua = await onhapqua.type(data[i],"");
               // }
               if(qua1 !== null & qua1 !== ""){
                     await Promise.all([
                        onhapqua.click({clickCount: 1}),
                        onhapqua.type(qua1.replace(/[']/g,"")),
                    ]).catch((error) => {console.log(error)});
               }
               if(qua2 !== null & qua2 !== ""){
                     await Promise.all([
                        page.keyboard.down('Enter'),
                        onhapqua.type(qua2.replace(/[']/g,"")),
                    ]).catch((error) => {console.log(error)});
               }
               if(qua3 !== null & qua3 !== ""){
                     await Promise.all([
                        page.keyboard.down('Enter'),
                        onhapqua.type(qua3.replace(/[']/g,"")),
                    ]).catch((error) => {console.log(error)});
               }
               if(qua4 !== null & qua4 !== ""){
                     await Promise.all([
                        page.keyboard.down('Enter'),
                        onhapqua.type(qua4,""),
                    ]).catch((error) => {console.log(error)});
               }
               if(qua5 !== null & qua5 !== ""){
                     await Promise.all([
                        page.keyboard.down('Enter'),
                        onhapqua.type(qua5,""),
                    ]).catch((error) => {console.log(error)});
               }
               if(qua6 !== null & qua6 !== ""){
                     await Promise.all([
                        page.keyboard.down('Enter'),
                        onhapqua.type(qua6,""),
                    ]).catch((error) => {console.log(error)});
               }             

               // add comment user
               await Promise.all([
                    onhapcomment.click({clickCount: 1}),
                    onhapcomment.type(comment.replace(/[']/g,"")),
                    console.log("Xong Comment! => " + rs.Article)
                ]).catch((error) => {
                    console.log(error)
                });
            }
        }catch(err){
            console.log("Có lỗi xảy ra", err + rs.Article)
        }

    }
//await page.click("input[value='Lưu']")
//await browser.close()
})();

