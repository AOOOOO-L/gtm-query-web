const EXCEL_URL = "https://updatecmvrf.ks3-cn-beijing.ksyuncs.com/GTM%E8%B5%84%E6%96%99%E5%B7%A5%E5%85%B7/GTM%20%E8%B5%84%E6%96%99.xlsx";

let excelData = [];
const detailModal = new bootstrap.Modal(document.getElementById('detailModal'));

window.onload = function() {
    refreshData();
};

async function refreshData() {
    try {
        setLoadingState();
        const response = await fetch(EXCEL_URL);
        if (!response.ok) throw new Error("下载失败");
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        excelData = XLSX.utils.sheet_to_json(worksheet, { defval: "未填写" });
        loadCategories();
        loadGtmTypes();
        filterData();
    } catch (error) {
        alert("加载失败：" + error.message);
        resetTable("加载失败，请重试");
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
        tableBody.innerHTML = '<tr><td colspan="7" class="text-center text-muted">暂无匹配数据</td></tr>';
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
    detailContent.innerHTML = `
        <h6>${row.系列名称 || "未填写"} - 资料详情</h6><hr>
        <p><strong>产品大类：</strong>${row.产品大类 || "未填写"}</p>
        <p><strong>资料类型：</strong>${row.GTM资料 || "未填写"}</p>
        <p><strong>资料编号：</strong>${row.编号 || "未填写"}</p>
        <p><strong>资料名称：</strong>${row.名称 || "未填写"}</p>
        <p><strong>创建者：</strong>${row.创建者 || "未填写"}</p>
        <p><strong>路径：</strong>${row.路径 || "未填写"}</p>
    `;
    detailModal.show();
}

function copyNumber(number) {
    if (number === "未填写") {
        alert("无编号可复制");
        return;
    }
    navigator.clipboard.writeText(number).then(() => alert("已复制：" + number));
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