import { mergeStyles, TextField } from '@fluentui/react';
import React, { useState } from 'react'
import { useDebounce } from '../../CustomHooks';
import { ExportSettings, GetSettings, SetSettings, SheetExportSettings } from '../../GlobalDeclarations';

export default function ExcelSheetSettingsMenu({ setting }) {
    const [ExportRange, SetExportRange] = useState(setting.ExportRange);

    function ChangeChartName(event){
        SetExportRange(event.target.value);
    }

    useDebounce(UpdateChartSettings, 500, ExportRange)

    function UpdateChartSettings(){
        let tempSettings: SheetExportSettings[] = GetSettings(ExportSettings);
        let current = tempSettings.find(chart => chart.Name === setting.Name);
        current.ExportRange = ExportRange;
        SetSettings(ExportSettings, tempSettings);
    }

  return (
    <div className={`${ContainerDiv}`}>
        <div className={`${sheetNameDiv} ms-font-bold ms-fontSize-18`}>{setting.Name}</div>
        <TextField className={`${TextFieldStyle}`} label='Sheet:' placeholder={ExportRange} onChange={ChangeChartName}></TextField>
        <div className={divider}></div>
    </div>
  )
}

const sheetNameDiv = mergeStyles({
  
})
const TextFieldStyle = mergeStyles({
  marginBottom: '10px'
})

const ContainerDiv = mergeStyles({
  padding: '15',
  margin: '2.5%'
});

const divider = mergeStyles({
  marginTop: '0px',
  backgroundColor: 'grey',
  height: '1px',
  marginBottom: '5px',
})