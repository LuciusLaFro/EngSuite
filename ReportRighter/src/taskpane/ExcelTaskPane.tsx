import { DefaultButton, mergeStyles } from '@fluentui/react';
import React from 'react'
import { ExportSettings, GetSettings, SetSettings, SheetExportSettings } from '../GlobalDeclarations';
import { ExcelToJSON } from './components/ExportFunctions';
import SheetSettingsInputBox from './components/SheetSettingsInputBox';

export default function ExcelTaskPane({ isOfficeInitialized, settings }) {
    if(!isOfficeInitialized){
        return(
            <div>Loading...</div>
        )
    }
    
    return (
        <div>
            <div className={headerContainer}>
                <h1 className={RRLabel}>ReportRighter</h1>
                <p className={tipParagraph}>To optimize performance adjust the export range of each worksheet to include the minimum amount of cells</p>
                <DefaultButton className={`ms-welcome__action ${exportButton}`} iconProps={{ iconName: "ChevronRight" }} onClick={ExcelToJSON}>
                Export Data To File
                </DefaultButton>
            </div>
            <div className={divider}></div>
          {settings.map(c => (
            <SheetSettingsInputBox setting={c} key={GetKey(c.Name)} />
          ))}
            
          
        </div>
      );
    
}
  
const tipParagraph = mergeStyles({
    marginLeft: '2.5%',
})

const RRLabel = mergeStyles({
    marginLeft: '2.5%',
    marginTop:'0',
})

const divider = mergeStyles({
    backgroundColor: 'grey',
    height: '1px',
    marginBottom: '5px',
  })
  
const exportButton = mergeStyles({
    width: '95%',
    marginLeft: '2.5%',
    marginBottom: '8px',
})
const headerContainer = mergeStyles({
    backgroundColor: '#d8d8d8',
    margin: '0px',
    top: '0',
})

  let key = 0;
  function GetKey(Name){
    let newKey = `${Name}_${key}`
    key += 1;
    return newKey
  }