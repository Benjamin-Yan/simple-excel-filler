// get data to let user select
function handleSubmit(event) {
    event.preventDefault();

    const formData = new FormData(event.target);
    fetch('/upload', {
        method: 'POST',
        enctype: 'multipart/form-data',
        body: formData
    })
    .then(response => response.json())
    .then(result => {
        const colL = result.columns, rowL = result.rows;
        const colList = document.getElementById('columnList');
        colList.innerHTML = "";

        for(let i = 0; i < colL.length; i++) {
            const newCB = document.createElement("input");
            newCB.type = "checkbox";
            newCB.name = "CB";
            newCB.id = "CB" + i;
            newCB.value = i;

            const newLB = document.createElement("label"); // 顯示文字
            newLB.setAttribute("for", newCB.id);
            newLB.textContent = colL[i];

            colList.appendChild(newCB);
            colList.appendChild(newLB);
            colList.appendChild(document.createElement("br"));
        }

        const newNB = document.createElement("input");
        newNB.type = "number";
        newNB.id = "quantity";
        newNB.name = "quantity";
        newNB.min = 2;
        newNB.max = rowL+1;
        newNB.required = true;
        const newLB = document.createElement("label");
        newLB.setAttribute("for", newNB.id);
        newLB.textContent = "要填寫幾筆資料(最大行數):";

        colList.appendChild(newLB);
        colList.appendChild(newNB);
        colList.appendChild(document.createElement("br"));

        const newSB = document.createElement("input");
        newSB.type = "submit";
        newSB.value = "Submit";
        colList.appendChild(newSB);
    })
    .catch(error => {
        document.getElementById('result').innerText = 'error: ' + error;
    });
}

