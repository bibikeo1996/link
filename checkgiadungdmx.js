const puppeteer = require('puppeteer');
const fs = require('fs');
var csvWriter = require('csv-write-stream');
var writer = csvWriter({ sendHeaders: true });
var csvFilename = "myfile.csv";
const inquirer = require('inquirer')
var excel = require('excel4node');
function expxortLink(data){
	var workbook = new excel.Workbook();
	var worksheet = workbook.addWorksheet('DMCL');
	const headings = ['Link'];
	headings.forEach((heading, index) => {
        worksheet.cell(1, index + 1).string(heading);
    })
    data.forEach((item, index) => {
    	worksheet.cell(index + 2, 1).string(item.Link);
    });
    var today1 = new Date();
    var filename1 = "Link2" + '-' + today1.getDate() + '-' +(today1.getMonth()+1) + '-' + today1.getFullYear() +'.xlsx';
    workbook.write(filename1);
}
// writing to xlsx
function exportToExcel(data) {
    // Create a new instance of a Workbook class
    var workbook = new excel.Workbook();
    // Add Worksheets to the workbook
    var worksheet = workbook.addWorksheet('DMCL');
    var style = workbook.createStyle({
      font: {
        size: 12
      },
      alignment: {
        horizontal: 'center'
      },
    });

    worksheet.row(1).freeze();
    worksheet.column(7).hide();
    worksheet.column(8).hide();
    worksheet.column(9).hide();
    worksheet.column(10).hide();
    worksheet.column(11).hide();
    worksheet.column(12).hide();
    const headings = ['SAP', 'Model', 'NY','Giá Cuối','Note','Loại','Góp 0%',"Bài Viết","Slide","Sticker","Icon Giảm Thêm","Icon Big Sale","Quà","Quà 2","LayOut","Format","Giảm miệng","%"]//,"LoạiSP","LoạiSP2"];//, 'Buộc trừ 2','Trừ tiền 1', 'Trừ tiền 2', 'Tổng Giá Quà'];//, 'Gift6', 'Gift7'];
    //const headings = ['Article', 'Cmt', 'Loại'];
    // Writing from cell A1 to I1
    headings.forEach((heading, index) => {
        worksheet.cell(1, index + 1).string(heading).style(style);
    })
    
    // Writing from cell A2 to I2 , A3 to I3, .....
    data.forEach((item, index) => {
        worksheet.cell(index + 2, 1).string(item.Sap);
        worksheet.cell(index + 2, 2).string(item.Name);
        worksheet.cell(index + 2, 3).string(item.Price1);
       	worksheet.cell(index + 2, 4).string(item.Price2);
        worksheet.cell(index + 2, 5).string(item.Comment);
        worksheet.cell(index + 2, 6).string(item.Loai);
        worksheet.cell(index + 2, 7).string(item.Gop);
        worksheet.cell(index + 2, 8).string(item.Baiviet);
        worksheet.cell(index + 2, 9).string(item.Slide);
        worksheet.cell(index + 2, 10).string(item.Sticker);
        worksheet.cell(index + 2, 11).string(item.Giamthem);
        worksheet.cell(index + 2, 12).string(item.SaleBig);
        worksheet.cell(index + 2, 13).string(item.Gift);
        worksheet.cell(index + 2, 14).string(item.Gift2);
        worksheet.cell(index + 2, 15).string(item.LayOut);
        worksheet.cell(index + 2, 16).formula('IF(AND(C2=D2),"Show",IF(AND(C2>D2),"Giấu","Sai"))');
        worksheet.cell(index + 2, 17).formula('C2-D2');
        worksheet.cell(index + 2, 18).formula('ROUND((Q2*100)/C2,0)');
        // worksheet.cell(index + 2, 15).string(item.LoaiSP2);
    });
    var today = new Date();
    var filename = "CheckGiaDCML" + '-' + today.getDate() + '-' +(today.getMonth()+1) + '-' + today.getFullYear() +'.xlsx';
    workbook.write(filename);
}
var file = "test.xlsx";
function getLink(file){
    const XLSX = require('xlsx');
    var workbook = XLSX.readFile(file);
    var sheet_name_list = workbook.SheetNames;
    var urls = [];
    data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]); 
    return data;
};
(async () => {
	const browser = await puppeteer.launch({ headless: false })
    const page = await browser.newPage()
    const urls = getLink(file);
    console.log(urls);
    let arrInfo = [];
    let result = [];
    for (var rs of urls) { // open for 1
        try { // open try 1
            // await page.goto("https://www.google.com/search?rlz=1C1NDCM_enVN894VN894&ei=5z3kX_mfLo-Rr7wPjdS2wAU&q="+rs.Model+"&oq="+rs.Model+"&gs_lcp=CgZwc3ktYWIQAzIHCAAQyQMQEzoJCAAQyQMQFhAeUPTxD1j08Q9g8_QPaAFwAHgAgAFbiAGzAZIBATKYAQCgAQKgAQGqAQdnd3Mtd2l6wAEB&sclient=psy-ab&ved=0ahUKEwi5r-aMh-btAhWPyIsBHQ2qDVgQ4dUDCA0&uact=5",{waitUntil: 'load', timeout: 100000});
            await page.goto("https://www.bing.com/search?q="+rs.Model+" điện máy xanh",{waitUntil: 'load', timeout: 100000});
                
                	const info1 = await page.evaluate(() => {
                		const checkweb = document.querySelector("#b_content");
                        if(checkweb !== null){
                            function CheckSieuThi(){
                                const checklink = document.querySelectorAll("div.b_attribution");
                                const res = [];
                                for(var i = 0; i < checklink.length; i++){
                                    const data = checklink[i].querySelector("div.b_attribution cite").innerText;
                                    if(data.includes("dienmayxanh")){
                                        res.push(data);
                                    }else{
                                        res.push("https://www.dienmayxanh.com/");
                                    }
                                }return res;
                            } // end function
                            const kiemtralink = CheckSieuThi();
                            const link = kiemtralink[0];
                            return {
                                Link: link ? link : "https://www.dienmayxanh.com/",
                            }
                        }
                        return {
                            Link: "https://www.dienmayxanh.com/",
                        };
		            })
		            if(info1){
		                result.push(info1)
		                console.log(info1.Link + " => Done!")
		                expxortLink(result)
		            }
        } // end try 1
        catch (err) {
            console.log("Có lỗi xảy ra", err);
        }
    } // end for 1
    console.log("Get Link Done!")
    await browser.close();
})();