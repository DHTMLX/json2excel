import { import_to_xlsx } from "../pkg/json2excel_wasm.js";

export { import_to_xlsx as toExcel };
export function convert(data){
    if (typeof data === "string")
        data = JSON.parse(data);

    const result = import_to_xlsx(data);
    const blob = new Blob([result], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"
    });

    return blob;
}
