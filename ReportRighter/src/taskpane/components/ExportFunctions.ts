import { ContextualMenuItemType, getPlaceholderStyles } from "@fluentui/react";
import { ExportSettings, SheetExportSettings } from "../../GlobalDeclarations";

function createCell_JSON(address: string, value: any){
    return `{"address":"${address}","value":"${value}"}`
}

function createSheet_JSON(name: string, cells: any){
    return `{"name":"${name}","cells":[${cells}]}`
}

function createChart_JSON(name: string, base64: string){
    return `{"name":"${name}","base64":"${base64}"}`
}

function createWorkbook_JSON(name: string, sheets: any, charts: any){
    return `{"name":"${name}", "sheets":[${sheets}], "charts":[${charts}]}`
}

function download(content, filename, contentType){
    if(!contentType) contentType = 'application/octet-stream';
        var a = document.createElement('a');
        var blob = new Blob([content], {'type':contentType});
        a.href = window.URL.createObjectURL(blob);
        a.download = filename;
        a.click();
}

function GetSettings(settingName: string){
    return Office.context.document.settings.get(settingName);
}

function SetSettings(settingName: string, value: any){
    Office.context.document.settings.set(settingName, value);
    Office.context.document.settings.saveAsync();
}

export async function ExcelToJSON(){
    //Get Name
    let name: string = await GetName();

    // Get Sheets
    let sheets = await GetSheets();

    //Get Charts    
    let charts  = await GetCharts();

    //Create Workbook
    let workbook = createWorkbook_JSON(name, sheets, charts);
    console.log(workbook);
    console.log(JSON.parse(workbook));

    download(workbook, `${name}.json`, "application/json")
};


async function GetSheets(){
    let sheetNames: string[] = [];
    let sheets: string = "";
    
    let isFirstSheet: boolean = true;

    
        const cont = await Excel.run(async (context) => {
            let worksheets = context.workbook.worksheets;
            worksheets.load('items');
            await context.sync();
    
            for(const worksheet of worksheets.items){
                worksheet.load('name');
                await context.sync();

                let newSheet = await GetCells(worksheet.name);
                if(isFirstSheet){
                    sheets += `${newSheet}`;
                    isFirstSheet = false;
                }else{
                    sheets += `,${newSheet}`;
                }
            }
        });
        return sheets;
    
}

async function GetCells(sheetName: string){
    let cells: string = "";
    let cont;
    let settings: SheetExportSettings[] = GetSettings(ExportSettings);

    return new Promise(async (resolve) => {
        cont = await Excel.run(async (context) => {
            let worksheet: Excel.Worksheet = context.workbook.worksheets.getItem(sheetName);
            let range: Excel.Range = worksheet.getRange(settings.find(setting => setting.Name === sheetName).ExportRange);
            range.load('values');
            range.load("columnIndex");
            range.load("rowIndex");
            await context.sync();

            let firstColumn = range.columnIndex;
            let firstRow = range.rowIndex + 1;

            let isTheFirstCell: boolean = true;

            let rowIndex = firstRow;
            range.values.forEach(row => {
                let columnIndex = firstColumn;
                row.forEach(value => {
                    if(value !== ""){
                        let address = `${String.fromCharCode(97+columnIndex)}${rowIndex}`;
                        if(isTheFirstCell){
                            cells += createCell_JSON(address, value);
                            isTheFirstCell = false;
                        }else{
                            cells += `,${createCell_JSON(address, value)}`;
                        }
                    }
                    columnIndex += 1;
                });
                rowIndex += 1;
            })
        });
        let sheet = createSheet_JSON(sheetName, cells)

        resolve(sheet);
    })
    
}

async function GetCharts(){
    let chartsString: string = "";
    let isFirstChart: boolean = true;

    await Excel.run(async (context) => {
        let worksheets: Excel.WorksheetCollection = context.workbook.worksheets;
        worksheets.load('items');
        await context.sync();

        for(const worksheet of worksheets.items){
            let charts = worksheet.shapes;
            charts.load('items');
            await context.sync();

            for(const chart of charts.items){
                let img = chart.getAsImage(Excel.PictureFormat.jpeg);
                chart.load('name');
                await context.sync();

                if(isFirstChart){
                    chartsString += `${createChart_JSON(chart.name, img.value)}`;
                    isFirstChart = false;
                }else{
                    chartsString += `,${createChart_JSON(chart.name, img.value)}`;
                }
            }
        }
    })
    return chartsString;
}

async function GetName(){
    let name;
    await Excel.run(async (context) => {
        let workbook = context.workbook;
        workbook.load('name');
        
        context.application.load();
        await context.sync();

        console.log('props', context.application)

        name = workbook.name.replace(".xlsx" || ".xlsm" || ".csv", "");
    })

    return name;
}