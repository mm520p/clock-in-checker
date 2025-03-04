const uploadInput = document.getElementById("upload");
const peopleList = document.getElementById("people");
const searchBtn = document.getElementById("search");
const nameInput = document.getElementById("name");
const resultTable = document.getElementById("result");
const resultBody = resultTable.getElementsByTagName("tbody")[0];
const dates = {};
const names = {};
const checkinTime = new Date("1995-12-17 " + "08:00:00");
const checkoutTime = new Date("1995-12-17 " + "17:30:00");
const earliestCheckoutTime = new Date("1995-12-17 " + "17:00:00");

uploadInput.addEventListener("change", () => {
    searchBtn.disabled = true;
    resultBody.innerHTML = "";

    const reader = new FileReader();
    reader.readAsArrayBuffer(uploadInput.files[0]);
    reader.addEventListener("load", () => {
        const file = reader.result;
        parseData(file);
        searchBtn.disabled = false;
    });
});

searchBtn.addEventListener("click", () => {
    const checkinStatus = {};
    const name = nameInput.value;
    for (const date in dates) {
        checkinStatus[date] = [];
        const rows = dates[date];
        let isCome = false;
        for (const row of rows) {
            if (row["名稱"] === name) {
                isCome = true;
                if (!checkinStatus[date][0]) {
                    if (row["時間2"] <= checkinTime) {
                        checkinStatus[date][0] = "上班打卡成功";
                    } else {
                        checkinStatus[date][0] = "遲到";
                        checkinStatus[date][2] = row["時間"];
                    }
                }
                if (row["時間2"] < checkoutTime) {
                    checkinStatus[date][1] = "早退";
                    if (row["時間2"] >= earliestCheckoutTime) {
                        checkinStatus[date][3] = row["時間"];
                    }
                } else {
                    checkinStatus[date][1] = "下班打卡成功";
                }
            }
        }
        if (!isCome) {
            checkinStatus[date][0] = "當日未到";
            checkinStatus[date][1] = "當日未到";
        }
    }
    resultBody.innerHTML = Object.entries(checkinStatus).map(([date, status]) => {
        const timeText1 = status[0] === "遲到" && status[2] ? "/" + status[2] : "";
        const timeText2 = status[1] === "早退" && status[3] ? "/" + status[3] : "";
        const statusText1 = `<td class="${getColor(status[0])}">${DOMPurify.sanitize(status[0] + timeText1)}</td>`;
        const statusText2 = `<td class="${getColor(status[1])}">${DOMPurify.sanitize(status[1] + timeText2)}</td>`;
        return `<tr><td>${DOMPurify.sanitize(date)}</td>${statusText1}${statusText2}</tr>`;
    }).join("\n");
});

function getColor(status) {
    switch (status) {
        case "上班打卡成功":
            return "text-success";
        case "下班打卡成功":
            return "text-success";
        case "當日未到":
            return "text-warning";
        case "遲到":
            return "text-danger";
        case "早退":
            return "text-danger";
    }
}

function parseData(file) {
    const workbook = XLSX.read(file);
    console.log(workbook.SheetNames);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);
    let date = "";
    json.forEach((row) => {
        if (!row["時間"].match(/:/)) {
            date = row["時間"].trim();
            dates[date] = [];
        } else {
            row["日期"] = date;
            row["時間2"] = new Date("1995-12-17 " + row["時間"]);
            dates[date].push(row);
            if (row["名稱"] !== "" && row["名稱"] !== "supervisor") {
                names[row["名稱"]] = true;
            }
        }
    });

    peopleList.innerHTML = Object.keys(names).map((name) => {
        return `<option>${DOMPurify.sanitize(name)}</option>`;
    }).join("\n");
}
