import { merge, mergeStyles } from '@fluentui/react'
import React, { useState } from 'react'
import { useDebounce } from '../../functions/CustomHooks'
import { ExportSettings, GetSettings, SetSettings, SheetExportSettings } from '../../functions/GlobalDeclarations'
import TextInput from './TextInput'

export default function SheetSettingsInputBox({ sheet }) {
  
  const [exportRangeSetting, setExportRangeSetting] = useState(sheet.exportRange)

  function changeExportSettings(event){
    setExportRangeSetting(event.target.value);
  }

  function updateSettings(){
    let tempSettings: SheetExportSettings[] = GetSettings(ExportSettings);
    let current = tempSettings.find(setting => setting.name === sheet.name);
    let text = `${exportRangeSetting}`;
    current.exportRange = text.toUpperCase();
    SetSettings(ExportSettings, tempSettings);
  }

  useDebounce(updateSettings, 100, exportRangeSetting);

  return (
    <div className={settings__wrapper}>
        <h1 className={sheetName}>{sheet.name}</h1>
        <TextInput className={input_box} placeholder={sheet.exportRange} label="Export Range:" onChange={changeExportSettings}></TextInput>
        <div className={divider}></div>
    </div>
  )
}

const divider = mergeStyles({
    marginTop: '10px',
    backgroundColor: 'grey',
    height: '1px',
    marginBottom: '5px',
})

const settings__wrapper = mergeStyles({
  padding: '0px 20px 0px 20px',
})

const sheetName = mergeStyles({
  margin: '0px 5px',
})

const input_box = mergeStyles({
  marginLeft: '20px',
})
