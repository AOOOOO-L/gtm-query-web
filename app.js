// 跨域代理（固定用这个，别改）
const EXCEL_URL = "https://api.allorigins.win/raw?url=https://updatecmvrf.ks3-cn-beijing.ksyuncs.com/GTM%E8%B5%84%E6%96%99%E5%B7%A5%E5%85%B7/GTM%20%E8%B5%84%E6%96%99.xlsx";

let data = [];
let productTypes = new Set();
let seriesMap = new Map();

// 页面加载完成
window.onload = function () {
  document.getElementById("refreshBtn").addEventListener("click", loadExcel);
  loadExcel(); // 自动加载一次
};

async function loadExcel() {
  try {
    let res = await fetch(EXCEL_URL);
    if (!res.ok) throw new Error("请求失败：" + res.status);
    let buf = await res.arrayBuffer();
    let workbook = XLSX.read(buf, { type: "array" });
    let sheetName = workbook.SheetNames[0];
    let sheet = workbook.Sheets[sheetName];
    data = XLSX.utils.sheet_to_json(sheet);

    // 处理下拉
    productTypes.clear();
    seriesMap.clear();
    data.forEach(row => {
      let type = row["产品大类"];
      let series = row["产品系列"];
      if (type) productTypes.add(type);
      if (type && series) {
        if (!seriesMap.has(type)) seriesMap.set(type, new Set());
        seriesMap.get(type).add(series);
      }
    });

    // 填充大类
    let typeSelect = document.getElementById("productType");
    typeSelect.innerHTML = `<option value="">全部</option>`;
    [...productTypes].sort().forEach(t => {
      typeSelect.innerHTML += `<option value="${t}">${t}</option>`;
    });

    // 大类切换
    typeSelect.onchange = function () {
      let seriesSelect = document.getElementById("productSeries");
      seriesSelect.innerHTML = `<option value="">全部</option>`;
      let selectedType = this.value;
      if (selectedType && seriesMap.has(selectedType)) {
        [...seriesMap.get(selectedType)].sort().forEach(s => {
          seriesSelect.innerHTML += `<option value="${s}">${s}</option>`;
        });
      }
      renderTable();
    };

    document.getElementById("productSeries").onchange = renderTable;
    renderTable();
  } catch (e) {
    alert("加载失败：" + e.message);
    console.error(e);
  }
}

function renderTable() {
  let type = document.getElementById("productType").value;
  let series = document.getElementById("productSeries").value;
  let filtered = data.filter(row => {
    return (!type || row["产品大类"] === type) && (!series || row["产品系列"] === series);
  });

  let table = document.getElementById("dataTable");
  table.innerHTML = "";
  if (filtered.length === 0) {
    table.innerHTML = "<tr><td colspan='6'>无数据</td></tr>";
    return;
  }

  let headers = Object.keys(filtered[0]);
  let thead = "<tr>";
  headers.forEach(h => thead += `<th>${h}</th>`);
  thead += "</tr>";
  table.innerHTML = thead;

  filtered.forEach(row => {
    let tr = "<tr>";
    headers.forEach(h => tr += `<td>${row[h] || ""}</td>`);
    tr += "</tr>";
    table.innerHTML += tr;
  });
}
