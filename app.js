const EXCEL_URL = "GTM 资料.xlsx";
let excelData = [];
const detailModal = new bootstrap.Modal(document.getElementById('detailModal'));

document.addEventListener('DOMContentLoaded', function(){
    document.getElementById("refreshBtn").onclick = refreshData;
    document.getElementById("categorySelect").onchange = function(){
        loadSeries();
        filterData();
    };
    document.getElementById("seriesSelect").onchange = filterData;
    document.getElementById("gtmSelect").onchange = filterData;
    refreshData();
});

async function refreshData(){
    setLoadingState();
    try{
        let res = await fetch(EXCEL_URL);
        if(!res.ok) throw "文件读取失败";
        let buf = await res.arrayBuffer();
        let wb = XLSX.read(buf, {type:"array"});
        excelData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:"未填写"});

        loadCategories();
        loadGtmTypes();
        filterData();
    }catch(e){
        resetTable("加载失败：" + e);
    }
}

function loadCategories(){
    let arr = [...new Set(excelData.map(x=>x["产品大类"]))].filter(x=>x&&x!="未填写").sort();
    let sel = document.getElementById("categorySelect");
    sel.innerHTML = "";
    if(arr.length===0){
        sel.innerHTML = "<option>无数据</option>";
        return;
    }
    arr.forEach(item=>{
        let op = document.createElement("option");
        op.value = item;
        op.innerText = item;
        sel.appendChild(op);
    });
    loadSeries();
}

function loadSeries(){
    let cat = document.getElementById("categorySelect").value;
    let arr = [...new Set(excelData.filter(x=>x["产品大类"]===cat).map(x=>x["系列名称"]))].filter(x=>x&&x!="未填写").sort();
    let sel = document.getElementById("seriesSelect");
    sel.innerHTML = "";
    if(arr.length===0){
        sel.innerHTML = "<option>无数据</option>";
        filterData();
        return;
    }
    arr.forEach(item=>{
        let op = document.createElement("option");
        op.value = item;
        op.innerText = item;
        sel.appendChild(op);
    });
    filterData();
}

function loadGtmTypes(){
    let arr = [...new Set(excelData.map(x=>x["GTM资料"]))].filter(x=>x&&x!="未填写").sort();
    let sel = document.getElementById("gtmSelect");
    sel.innerHTML = '<option value="all">全部</option>';
    arr.forEach(item=>{
        let op = document.createElement("option");
        op.value = item;
        op.innerText = item;
        sel.appendChild(op);
    });
}

function filterData(){
    let cat = document.getElementById("categorySelect").value;
    let ser = document.getElementById("seriesSelect").value;
    let gtm = document.getElementById("gtmSelect").value;

    let list = excelData.filter(x=>x["产品大类"]===cat && x["系列名称"]===ser);
    if(gtm !== "all"){
        list = list.filter(x=>x["GTM资料"] === gtm);
    }
    renderTable(list);
}

// 严格8列，和表头一一对应
function renderTable(list){
    let tb = document.querySelector("#dataTable tbody");
    tb.innerHTML = "";
    if(list.length===0){
        tb.innerHTML = '<tr><td colspan="8" class="text-center text-muted">暂无匹配数据</td></tr>';
        return;
    }
    list.forEach((row, idx)=>{
        let tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${row["GTM资料"]}</td>
            <td>${row["是否上传"]}</td>
            <td>${row["名称"]}</td>
            <td>${row["编号"]}</td>
            <td>${row["当前版本"]}</td>
            <td>${row["创建者"]}</td>
            <td>${row["数据状态"]}</td>
            <td>
                <button class="btn btn-sm btn-info" onclick='showDetail(${JSON.stringify(row).replace(/'/g, "\\'").replace(/"/g, "'")})'>详情</button>
                <button class="btn btn-sm btn-success" onclick="copyNumber('${row["编号"]}')">复制编号</button>
            </td>
        `;
        tb.appendChild(tr);
    });
}

function showDetail(rowData){
    let html = `
    <h6>${rowData["系列名称"]} 详情</h6><hr>
    <p>产品大类：${rowData["产品大类"]}</p>
    <p>GTM资料：${rowData["GTM资料"]}</p>
    <p>编号：${rowData["编号"]}</p>
    <p>名称：${rowData["名称"]}</p>
    <p>当前版本：${rowData["当前版本"]}</p>
    <p>创建者：${rowData["创建者"]}</p>
    <p>路径：${rowData["路径"]}</p>
    `;
    document.getElementById("detailContent").innerHTML = html;
    detailModal.show();
}

function copyNumber(num){
    if(num === "未填写"){
        alert("无编号");
        return;
    }
    navigator.clipboard.writeText(num).then(()=>{
        alert("已复制："+num);
    }).catch(()=>{
        prompt("手动复制", num);
    });
}

function setLoadingState(){
    document.getElementById("categorySelect").innerHTML = "<option>加载中</option>";
    document.getElementById("seriesSelect").innerHTML = "<option>加载中</option>";
    document.getElementById("gtmSelect").innerHTML = '<option value="all">加载中</option>';
    resetTable("加载中...");
}

function resetTable(txt){
    let tb = document.querySelector("#dataTable tbody");
    tb.innerHTML = `<tr><td colspan="8" class="text-center text-muted">${txt}</td></tr>`;
}
