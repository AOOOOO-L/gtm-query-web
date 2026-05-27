const EXCEL_URL = "GTM 资料.xlsx";

let excelData = [];
const detailModal = new bootstrap.Modal(document.getElementById('detailModal'));

// 页面加载完成后，一次性绑定所有事件
window.onload = function() {
    // 绑定刷新按钮
    document.getElementById('refreshBtn').addEventListener('click', refreshData);
    // 绑定产品大类下拉框事件
    document.getElementById('categorySelect').addEventListener('change', function() {
        loadSeries(); // 先加载对应系列
        filterData(); // 再触发筛选
    });
    // 绑定产品系列下拉框事件（修复点：直接绑定筛选函数）
    document.getElementById('seriesSelect').addEventListener('change', filterData);
    // 绑定GTM类型下拉框事件
    document.getElementById('gtmSelect').addEventListener('change', filterData);

    // 页面加载后自动刷新一次
    refreshData();
};

// 刷新数据：从仓库里的Excel文件加载数据
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
        
        // 数据格式清理
        excelData = excelData.map(row => {
            return {
                ...row,
                系列名称: (row.系列名称 || "未填写").toString().replace(/\n/g, " ").trim(),
                产品大类: (row.产品大类 || "未填写").toString().trim(),
                GTM资料: (row.GTM资料 || "未填写").toString().trim()
            };
        });

        // 加载产品大类下拉框
        loadCategories();
        // 加载GTM资料类型下拉框
        loadGtmTypes();
        // 初始化表格（加载大类下的第一个系列数据）
        filterData();
    } catch (error) {
        alert(`数据加载失败：${error.message}\n请检查Excel文件是否上传到仓库`);
        resetTable("加载失败，请点击刷新重试");
    }
}

// 加载产品大类下拉框
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
    // 加载对应系列
    loadSeries();
}

// 根据选中的产品大类，加载对应的产品系列
function loadSeries() {
    const categorySelect = document.getElementById('categorySelect');
    const seriesSelect = document.getElementById('seriesSelect');
    const selectedCategory = categorySelect.value;

    seriesSelect.innerHTML = "";
    if (!selectedCategory || selectedCategory === "无可用产品大类") {
        seriesSelect.innerHTML = '<option value="">无可用系列</option>';
        return;
    }

    const series = [...new Set(excelData.filter(row => row.产品大类 === selectedCategory).map(row => row.系列名称))]
        .filter(s => s && s !== "未填写").sort();
    
    if (series.length === 0) {
        seriesSelect.innerHTML = '<option value="">无可用系列</option>';
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

// 加载GTM资料类型下拉框
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

// 筛选数据并渲染表格
function filterData() {
    const selectedCategory = document.getElementById('categorySelect').value;
    const selectedSeries = document.getElementById('seriesSelect').value;
    const selectedGtm = document.getElementById('gtmSelect').value;

    // 基础筛选：产品大类 + 产品系列
    let filteredData = excelData.filter(row => {
        return row.产品大类 === selectedCategory && row.系列名称 === selectedSeries;
    });

    // 二次筛选：GTM资料类型
    if (selectedGtm !== "all" && selectedGtm) {
        filteredData = filteredData.filter(row => row.GTM资料 === selectedGtm);
    }

    // 渲染表格
    renderTable(filteredData);
}

// 渲染表格数据
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

// 显示资料详情弹窗
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

// 复制编号功能（完全保留）
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

// 设置加载状态
function setLoadingState() {
    document.getElementById('categorySelect').innerHTML = '<option value="">加载中...</option>';
    document.getElementById('seriesSelect').innerHTML = '<option value="">加载中...</option>';
    document.getElementById('gtmSelect').innerHTML = '<option value="all">加载中...</option>';
    resetTable("加载中...");
}

// 重置表格提示
function resetTable(text) {
    const tableBody = document.querySelector('#dataTable tbody');
    tableBody.innerHTML = `<tr><td colspan="7" class="text-center text-muted">${text}</td></tr>`;
}
