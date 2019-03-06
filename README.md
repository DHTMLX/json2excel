JSON2Excel 
--------------

JSON2Excel is a Rust and WebAssembly-based library that allows converting JSON files into Excel ones at ease.

[![npm version](https://badge.fury.io/js/json2excel-wasm.svg)](https://badge.fury.io/js/json2excel-wasm) 

### How to build

```
yarn install
// build js code
yarn build
// rebuild wasm code
// rust toolchain is required
yarn build-wasm
```

### How to use via npm

- install the module

```js
yarn add json2excel-wasm
```
- import the module

```js
// worker.js
import "json2excel-wasm";
```

- use the module in the app

```js
// app.js
const worker = new Worker("worker.js");

function doConvert(){
    worker.postMessage({ 
        type: "convert",
        data: json_data_to_export
    });
}

worker.addEventListener("message", e => {
    if (e.data.type === "ready"){
        const blob = e.data.blob;
        //excel file is ready to be downloaded!

        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "data.xlsx";
        document.body.append(a);
        a.click();
        document.body.remove(a);
    }
});
```
### How to use from CDN

CDN links are the following:

- https://cdn.dhtmlx.com/libs/json2excel/1.0/worker.js 
- https://cdn.dhtmlx.com/libs/json2excel/1.0/lib.wasm

In case you use build system like webpack, it is advised to wrap the link to CDN source into a blob object to avoid possible breakdowns:

```js
var url = window.URL.createObjectURL(new Blob([
    "importScripts('https://cdn.dhtmlx.com/libs/json2excel/1.0/worker.js');"
], { type: "text/javascript" }));

var worker = new Worker(url);
```

### Input format

```ts
interface IConvertMessageData {
    uid?: string;
    data: ISheetData;
    styles?: IStyles[];
    wasmPath?: string; // use cdn by default
}

interface IReadyMessageData {
    uid: string; // same as incoming uid
    blob: Blob;
}

interface ISheetData {
    name?: string;
    cols?: IColumnData[];
    rows?: IRowData[];
    cells?: IDataCell[][]; // if cells mising, use plain
    plain?: string[][];

    merged?: IMergedCells;
}

interface IMergedCells {
    from: IDataPoint;
    to: IDataPoint;
}

interface IDataPoint {
    column: number; 
    row: number;
}

interface IColumnData {
    width: number;
}

interface IRowData {
    height: number;
}

interface IDataCell{
    v: string;
    s: number;
}

interface IStyle {
    fontSize?: string;
    fontFamily?: string;

    borderLeft?: string;
    borderTop?: string;
    borderBottom?: string;
    borderRight?: string;

    format?: string;
}
```


### License 

MIT
