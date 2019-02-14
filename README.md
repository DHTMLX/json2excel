JSON to Excell 
--------------

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