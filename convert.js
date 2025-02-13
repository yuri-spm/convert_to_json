let file     = document.querySelector("#file");
let page     = document.querySelector("#page");
let result   = document.querySelector("#result");
let download = document.querySelector("#download");
let allSheet;

file.addEventListener("change", () => {
    file.files[0].arrayBuffer().then((buffer) => {
        allSheet = XLSX.read(buffer);
        let forSelect = allSheet.SheetNames.reduce((acum, cur) => {
            return acum + `<option value="${cur}">${cur}</option>`;
        }, "");
        page.innerHTML = forSelect;
        let jsonObj = XLSX.utils.sheet_to_json(allSheet.Sheets[page.value]);
        let jsn = JSON.stringify({data: jsonObj}, null, 4);
        result.value = jsn;
        download.href = "data:application/json;charset=utf-8,"+encodeURIComponent(result.value);
        download.download = page.value;
        sendJsonToApi(jsonObj);
        console.log(allSheet);
    }).catch((error) => {
        console.log(error);
    });
});

function sendJsonToApi(jsonObj){
    let endpoint = "url";
    let token = "token" 
    jsonObj.forEach(async (item) => {
        try {
            let response = await fetch(endpoint, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${token}`,
                },
                body: JSON.stringify(item)
            });

            if (response.ok) {
                console.log("Item enviado com sucesso:", item);
            } else {
                console.error("Erro ao enviar o item:", item, await response.text());
            }
        } catch (error) {
            console.error("Erro na requisição:", error);
        }
    });
}
