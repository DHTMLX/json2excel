JSON to Excell 
--------------

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

### CDN links

- https://cdn.dhtmlx.com/libs/json2excel/1.0/worker.js 
- https://cdn.dhtmlx.com/libs/json2excel/1.0/lib.wasm

### How to use

```js
// worker.js
import "json2excel-wasm";
```

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

### License 

MIT

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