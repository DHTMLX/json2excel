import "../node_modules/fast-text-encoding/text";
import "../wasm/xlsx_import";

onmessage = function(e) {
    if (e.data.type === "convert"){
        doConvert(e.data);
    }
}


let import_to_xlsx = null;
function doConvert(config){
    if (import_to_xlsx) {
        const result = import_to_xlsx(config.data);
        const blob = new Blob([result], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"
        });

        postMessage({
            uid: config.uid || (new Date()).valueOf(),
            type: "ready",
            blob
        });
    } else {
        const path = config.wasmPath || "https://cdn.dhtmlx.com/libs/json2excel/1.0/lib.wasm";

        wasm_bindgen(path).then(() => {
            import_to_xlsx = wasm_bindgen.import_to_xlsx;
            doConvert(config);
        }).catch(e => console.log(e));
    }
}