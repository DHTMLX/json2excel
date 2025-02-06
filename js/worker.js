import init, { import_to_xlsx } from '../pkg/json2excel_wasm.js';


onmessage = function(e) {
    if (e.data.type === "convert") {
        let data = e.data.data;
        if (typeof data === "string")
            data = JSON.parse(data);
        doConvert(data);
    }
}

async function doConvert(data, config = {}){
    await init();

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
