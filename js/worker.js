import { import_to_xlsx } from '../pkg/json2excel_wasm';


onmessage = function(e) {
    if (e.data.type === "convert") {
        let data = e.data.data;
        if (typeof data === "string")
            data = JSON.parse(data);
        doConvert(data);
    }
}

function doConvert(data, config = {}){
    const result = import_to_xlsx(data);
    const blob = new Blob([result], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"
    });

    postMessage({
        uid: config.uid || (new Date()).valueOf(),
        type: "ready",
        blob
    });
}

postMessage({ type:"init" });
