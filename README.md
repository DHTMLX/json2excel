JSON2Excel 
--------------

JSON2Excel is a Rust and WebAssembly-based library that allows converting JSON files into Excel ones at ease.

[![npm version](https://badge.fury.io/js/json2excel-wasm.svg)](https://badge.fury.io/js/json2excel-wasm) 

### How to build

```

cargo install wasm-pack
wasm-pack build
```

### How to use via npm

- install the module

```js
yarn add json2excel-wasm
```
- import and use the module

```js
// worker.js
import {convert} "json2excel-wasm";
const blob = convert(json_data_to_export);
```

you can use code like next to force download of the result blob

```js
const a = document.createElement("a");
a.href = URL.createObjectURL(blob);
a.download = "data.xlsx";
document.body.append(a);
a.click();
document.body.remove(a);
```

### How to use from CDN

CDN links are the following:

- https://cdn.dhtmlx.com/libs/json2excel/1.2/worker.js 
- https://cdn.dhtmlx.com/libs/json2excel/1.2/module.js 
- https://cdn.dhtmlx.com/libs/json2excel/1.2/json2excel_wasm_bg.wasm

You can import and use lib dynamically like 

```js
const convert = import("https://cdn.dhtmlx.com/libs/json2excel/1.2/module.js");
const blob = convert(json_data_to_export);
```

or use it as web worker

```js
// you need to server worker from the same domain as the main script
var worker = new Worker("./worker.js"); 
worker.addEventListener("message", ev => {
    if (ev.data.type === "ready"){
        const blob = ev.data.blob;
        // do something with result
    }
});
worker.postMessage({
    type:"convert",
    data: raw_json_data
});
```

if you want to load worker script from CDN and not from your domain it requires a more complicated approach, as you need to catch the moment when service inside of the worker will be fully initialized

```js
var url = window.URL.createObjectURL(new Blob([
    "importScripts('https://cdn.dhtmlx.com/libs/json2excel/1.2/worker.js');"
], { type: "text/javascript" }));

var worker = new Promise((res) => {
    const x = Worker(url); 
    worker.addEventListener("message", ev => {
        if (ev.data.type === "ready"){
            const json = ev.data.data;
            // do something with result
        } else if (ev.data.type === "init"){
            // service is ready
            res(x);
        }
    });
});

worker.then(x => x.postMessage({
    type:"convert",
    data: raw_json_data
}));
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
    align?: string;

    background?: string;
    color?: string;

    fontWeight?: string;
    fontStyle?: string;
    textDecoration?: string;

    format?: string;
}
```


### License 

MIT
