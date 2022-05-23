import * as React from "react";
import { DefaultButton, values } from "@fluentui/react";
import Header from "./Header";
import SheetSettingsInputBox from "./SheetSettingsInputBox";
import { ExcelToJSON, ExportSettings, GetSettings, LoadSettings, removeElement, SheetExportSettings, tempSettings } from "../../functions/GlobalDeclarations";

/* global console, Excel, require  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
  }

  click = async () => {
    try {
      await Excel.run(async (context) => {
        await context.sync();
        console.log('log');
      });
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <>
          <p>loading...</p>
        </>
      );
    }

    return (
      <div>
        <Header label='ReportRighter' tipText={tipParagraph} buttonText="Export Data" onClick={ExcelToJSON} imgSrc='../../../assets/RR-logo.png' imgAlt='ReportRighter Logo.png'></Header>
        {tempSettings.map(c => (
          <SheetSettingsInputBox sheet={c} key={GetKey()}></SheetSettingsInputBox>
        ))}     
      </div>
    );
  }
}

const testSheet = new SheetExportSettings('Sheet 1', 'A1:Ax10000');
let key_ID = 0;

function GetKey(){
  key_ID += 1;
  return `key__${key_ID}`;
}

const tipParagraph = "To optimize performance, adjust the export range of each worksheet to include the minimum range of cells"

async function ExportData(){
  console.log('hi');
  download('Test', "test");
}

function download(filename, text) {
  var element = document.createElement('a');
  element.setAttribute('href', URL.createObjectURL(new Blob([text], {type: "text/plain"}))); 
  element.download = filename;

  element.click();
}