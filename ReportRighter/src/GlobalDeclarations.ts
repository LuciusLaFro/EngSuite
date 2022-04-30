export class SheetExportSettings{
    constructor(Name: string, ExportRange: string){
        this.Name = Name;
        this.ExportRange = ExportRange;
    }

    Name: string;
    ExportRange: string;
}

export const ExportSettings = "ExportSettings";

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

export async function LoadSettings(){
    let tempSettings: SheetExportSettings[] = GetSettings(ExportSettings); 
    await Excel.run(async (context) => {
      let worksheets: Excel.WorksheetCollection = context.workbook.worksheets;
      worksheets.load('items');
      await context.sync();
  
      for(const worksheet of worksheets.items){
        let name = worksheet.name;
        if(!tempSettings){
          tempSettings = [];
          tempSettings.push(new SheetExportSettings(name, "A1:A10000"));
        }
  
        if(!tempSettings.find(setting => setting.Name === name)){
          tempSettings.push(new SheetExportSettings(name, "A1:A10000"));
        }
        
      }
    });
    SetSettings(ExportSettings, tempSettings);
    return tempSettings
  }