export class SheetExportSettings{
    constructor(name: string, exportRange: string){
        this.name = name;
        this.exportRange = exportRange;
    }

    name: string;
    exportRange: string;
}

export const ExportSettings: string = 'ExportSettings'

function ClearSettings(settingName: string){
    Office.context.document.settings.remove(settingName);
}

export function SetSettings(settingName: string, value: any){
    ClearSettings(settingName);
    Office.context.document.settings.set(settingName, value);
    Office.context.document.settings.saveAsync();
}

export function GetSettings(settingName: string){
    let setting =  Office.context.document.settings.get(settingName);
    return setting;
}


// Array methods
export function removeElement(array: any[], callback){
    let index: number = array.findIndex(callback);
    array.splice(index, 1);
}

export let tempSettings: SheetExportSettings[]

export async function LoadSettings(){
  await Excel.run(async (context) => {
    let worksheets: Excel.WorksheetCollection = context.workbook.worksheets;
    worksheets.load('items');
    await context.sync();

    tempSettings = GetSettings(ExportSettings);

    if(tempSettings === null){
      tempSettings = []
    }

    // CheckForDeletes(worksheets.items, tempSettings);

    for(const ws of worksheets.items){
      if(!tempSettings.find(setting => setting.name === ws.name)){
        console.log(ws.name)
        tempSettings.push(new SheetExportSettings(ws.name, "A1:AX10000"));
      }
      
    }

    SetSettings(ExportSettings, tempSettings);
    return tempSettings;
  })  
}

async function CheckForDeletes(worksheets, settingsArray: SheetExportSettings[]){
  for(const setting of settingsArray){
    if(!worksheets.find(ws => ws.name === setting.name)){
      removeElement(settingsArray, s => s === setting)
    }
  }
}

function download(filename, content) {
  var element = document.createElement('a');
  element.setAttribute('href', URL.createObjectURL(new Blob([content], {type: "application/json"}))); 
  element.download = filename;

  element.click();
}

export async function ExcelToJSON(){
  //Get Name
  let name: string = await GetName();

  // Get Sheets
  let sheets = await GetSheets();

  //Get Charts    
  let charts  = await GetCharts();

  //Create Workbook
  let newName = name.replace(".xlsm", "");
  newName = newName.replace(".xls", "");
  newName = newName.replace(".xlsx", "");
  newName = newName.replace(".csv", "");
  console.log('name:', newName);

  let workbook = createWorkbook_JSON(newName, sheets, charts);
  console.log(workbook);

  await Excel.run(async (context) => {
    let ws = context.workbook.getActiveCell();
    ws.load();

    await context.sync();
    ws.values = [[workbook]];
  })
  download(newName, workbook)
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
          let range: Excel.Range = worksheet.getRange(settings.find(setting => setting.name === sheetName).exportRange);
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