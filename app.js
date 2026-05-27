const EXCEL_URL = "GTM 资料.xlsx";

let excelData = [];
const detailModal = new bootstrap.Modal(document.getElementById('detailModal'));

window.onload = function() {
    document.getElementById('refreshBtn').addEventListener('click', refreshData);
    refreshData();
};

async function refreshData() {
    try {
        setLoadingState();
        const response = await fetch(EXCEL_URL);
        if (!response.ok) throw new Error(`下载失败：${response.status}`);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        excelData = XLSX.utils.sheet_to_json(worksheet, { defval: "未填写" });
        
        excelData = excelData.map(row => {
            return {
                ...row,
                系列名称: (row.系列名称 || "未填写").toString().replace(/\n/g, " ").trim(),
                产品大类: (row.产品大类 || "未填写").toString().trim(),
                GTM资料: (row.GTM资料 || "未填写").toString().trim()
            };
        });

        loadCategories();
        loadGtmTypes();
        filterData();
    } catch (error) {
        alert(`数据加载失败：${error.message}\n请检查Excel文件是否上传到仓库`);
        resetTable("加载失败，请点击刷新重试");
    }
}

function loadCategories() {
    const categorySelect = document.getElementById('categorySelect');
    const categories = [...new Set(excelData.map(row => row.产品大类))]
        .filter(c => c && c !== "未填写").sort();
    categorySelect.innerHTML = "";
    if (categories.length === 0) {
        categorySelect.innerHTML = '<option value="">无可用产品大类</option>';
        return;
    }
    categories.forEach(cat => {
        const option = document.createElement('option');
        option.value = cat;
        option.textContent = cat;
        categorySelect.appendChild(option);
    });
    categorySelect.value = categories[0];
    loadSeries();
}

function loadSeries() {
    const categorySelect = document.getElementById('categorySelect');
    const seriesSelect = document.getElementById('seriesSelect');
    const selectedCategory = categorySelect.value;
    seriesSelect.innerHTML = "";
    if (!selectedCategory) {
        seriesSelect.innerHTML = '<option value="">无可用系列</option>';
        filterData();
        return;
    }
    const series = [...new Set(excelData.filter(row => row.产品大类 === selectedCategory).map(row => row.系列名称))]
        .filter(s => s && s !== "未填写").sort();
    if (series.length === 0) {
        seriesSelect.innerHTML = '<option value="">无可用系列</option>';
        filterData();
        return;
    }
    series.forEach(s => {
        const option = document.createElement('option');
        option.value = s;
        option.textContent = s;
        seriesSelect.appendChild(option);
    });
    seriesSelect.value = series[0];
}

function loadGtmTypes() {
    const gtmSelect = document.getElementById('gtmSelect');
    const gtmTypes = [...new Set(excelData.map(row => row.GTM资料))]
        .filter(g => g && g !== "未填写").sort();
    gtmSelect.innerHTML = '<option value="all">全部</option>';
    gtmTypes.forEach(type => {
        const option = document.createElement('option');
        option.value = type;
        option.textContent = type;
        gtmSelect.appendChild(option);
    });
}

function filterData() {
    const selectedCategory = document.getElementById('categorySelect').value;
    const selectedSeries = document.getElementById('seriesSelect').value;
    const selectedGtm = document.getElementById('gtmSelect').value;
    let filteredData = excelData.filter(row => row.产品大类 === selectedCategory && row.系列名称 === selectedSeries);
    if (selectedGtm !== "all") {
        filteredData = filteredData.filter(row => row.GTM资料 === selectedGtm);
    }
    renderTable(filteredData);
}

function renderTable(data) {
    const tableBody = document.querySelector('#dataTable tbody');
    tableBody.innerHTML = "";
    if (data.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td colspan="7" class="text-center text-muted">暂无匹配数据</td>';
        tableBody.appendChild(tr);
        return;
    }
    data.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.GTM资料 || "未填写"}</td>
            <td>${row.是否上传 || "未填写"}</td>
            <td>${row.名称 || "未填写"}</td>
            <td>${row.编号 || "未填写"}</td>
            <td>${row.创建者 || "未填写"}</td>
            <td>${row.数据状态 || "未填写"}</td>
            <td>
                <button class="btn btn-sm btn-info" onclick="showDetail(${index})">详情</button>
                <button class="btn btn-sm btn-success" onclick="copyNumber('${row.编号 || "未填写"}')">复制编号</button>
            </td>
        `;
        tableBody.appendChild(tr);
    });
}

function showDetail(index) {
    const row = excelData[index];
    const detailContent = document.getElementById('detailContent');
    const now = new Date().toLocaleString();
    detailContent.innerHTML = `
        <div class="card">
            <div class="card-body">
                <h6>【${row.系列名称 || "未填写"} - 资料详情】</h6>
                <hr>
                <div class="row mb-2">
                    <div class="col-md-6"><strong>产品大类：</strong>${row.产品大类 || "未填写"}</div>
                    <div class="col-md-6"><strong>系列名称：</strong>${row.系列名称 || "未填写"}</div>
                </div>
                <div class="row mb-2">
                    <div class="col-md-6"><strong>资料类型：</strong>${row.GTM资料 || "未填写"}</div>
                    <div class="col-md-6"><strong>是否上传：</strong>${row.是否上传 || "未填写"}</div>
                </div>
                <div class="row mb-2">
                    <div class="col-md-6"><strong>资料编号：</strong>${row.编号 || "未填写"}</div>
                    <div class="col-md-6"><strong>资料名称：</strong>${row.名称 || "未填写"}</div>
                </div>
                <div class="row mb-2">
                    <div class="col-md-6"><strong>创建者：</strong>${row.创建者 || "未填写"}</div>
                    <div class="col-md-6"><strong>创建时间：</strong>${row.创建时间 || "未填写"}</div>
                </div>
                <div class="row mb-2">
                    <div class="col-md-12"><strong>存储路径：</strong>${row.路径 || "未填写"}</div>
                </div>
                <hr>
                <div class="text-muted text-end">查询时间：${now}</div>
            </div>
        </div>
    `;
    detailModal.show();
}

function copyNumber(number) {
    if (number === "未填写") {
        alert("该条资料无编号可复制！");
        return;
    }
    navigator.clipboard.writeText(number).then(() => {
        alert(`资料编号已复制：\n${number}`);
    }).catch(() => {
        alert("复制失败，请手动复制！");
    });
}

function setLoadingState() {
    document.getElementById('categorySelect').innerHTML = '<option value="">加载中...</option>';
    document.getElementById('seriesSelect').innerHTML = '<option value="">加载中...</option>';
    document.getElementById('gtmSelect').innerHTML = '<option value="all">加载中...</option>';
    resetTable("加载中...");
}

function resetTable(text) {
    const tableBody = document.querySelector('#dataTable tbody');
    tableBody.innerHTML = `<tr><td colspan="7" class="text-center text-muted">${text}</td></tr>`;
}
