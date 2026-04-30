<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>企業數據視覺化儀表板</title>
    <!-- 引入 Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- 引入 Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- 引入 Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <style>
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        /* 隱藏未啟用的分頁內容 */
        .tab-content.hidden { display: none; }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 font-sans min-h-screen p-4 md:p-8">

    <div class="max-w-7xl mx-auto space-y-6">
        
        <!-- 頁首標題 -->
        <header class="flex items-center space-x-3 mb-4">
            <div class="p-3 bg-blue-600 rounded-lg text-white">
                <i data-lucide="bar-chart-2"></i>
            </div>
            <div>
                <h1 class="text-2xl font-bold text-slate-900">企業數據視覺化儀表板</h1>
                <p class="text-sm text-slate-500">輸入或貼上表格資料，自動生成圖表與統計數據</p>
            </div>
        </header>

        <!-- 分頁導覽列 -->
        <div class="border-b border-slate-200">
            <nav class="flex space-x-8" aria-label="Tabs">
                <button onclick="switchTab('input')" id="btn-tab-input" class="border-blue-500 text-blue-600 whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm flex items-center gap-2 transition-colors">
                    <i data-lucide="table" class="w-4 h-4"></i> 資料輸入
                </button>
                <button onclick="switchTab('charts')" id="btn-tab-charts" class="border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300 whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm flex items-center gap-2 transition-colors">
                    <i data-lucide="pie-chart" class="w-4 h-4"></i> 圖表分析
                </button>
                <button onclick="switchTab('stats')" id="btn-tab-stats" class="border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300 whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm flex items-center gap-2 transition-colors">
                    <i data-lucide="list-ordered" class="w-4 h-4"></i> 詳細數據
                </button>
            </nav>
        </div>

        <!-- 分頁 1：資料輸入 -->
        <div id="tab-input" class="tab-content block space-y-6">
            <section class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                <div class="p-6 border-b border-slate-100 bg-slate-50/50 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                    <h2 class="text-lg font-semibold flex items-center gap-2">
                        資料列表
                    </h2>
                    <div class="flex gap-2">
                        <button onclick="togglePasteArea()" class="px-4 py-2 text-sm font-medium text-blue-600 bg-blue-50 hover:bg-blue-100 rounded-lg transition-colors flex items-center gap-1">
                            <i data-lucide="clipboard-paste" class="w-4 h-4"></i> 從 Excel 貼上
                        </button>
                        <button onclick="addRow()" class="px-4 py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition-colors flex items-center gap-1">
                            <i data-lucide="plus" class="w-4 h-4"></i> 新增一列
                        </button>
                        <button onclick="resetData()" class="px-4 py-2 text-sm font-medium text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors flex items-center gap-1">
                            <i data-lucide="rotate-ccw" class="w-4 h-4"></i> 重設範例
                        </button>
                    </div>
                </div>

                <!-- 貼上資料區塊 -->
                <div id="pasteAreaContainer" class="hidden p-6 border-b border-slate-100 bg-blue-50/30">
                    <label class="block text-sm font-medium text-slate-700 mb-2">請將 Excel 或 CSV 資料貼在下方文字方塊中 (以 Tab 分隔)：</label>
                    <textarea id="pasteInput" rows="4" class="w-full p-3 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none text-sm font-mono" placeholder="公司&#9;個編帳號&#9;產業&#9;職務&#9;職位&#9;日期..."></textarea>
                    <div class="mt-3 flex justify-end gap-2">
                        <button onclick="togglePasteArea()" class="px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-100 rounded-lg transition-colors">取消</button>
                        <button onclick="processPastedData()" class="px-4 py-2 text-sm font-medium text-white bg-green-600 hover:bg-green-700 rounded-lg transition-colors">載入資料</button>
                    </div>
                </div>

                <!-- 表格主體 -->
                <div class="overflow-x-auto">
                    <table class="w-full text-sm text-left">
                        <thead class="text-xs text-slate-600 uppercase bg-slate-100/50">
                            <tr>
                                <th class="px-4 py-3 font-semibold w-1/6 whitespace-nowrap">公司</th>
                                <th class="px-4 py-3 font-semibold w-1/12 whitespace-nowrap">個編帳號</th>
                                <th class="px-4 py-3 font-semibold w-1/6 whitespace-nowrap">產業</th>
                                <th class="px-4 py-3 font-semibold w-1/6 whitespace-nowrap">職務</th>
                                <th class="px-4 py-3 font-semibold w-1/6 whitespace-nowrap">職位</th>
                                <th class="px-4 py-3 font-semibold w-1/6 whitespace-nowrap">日期</th>
                                <th class="px-4 py-3 font-semibold w-16 text-center whitespace-nowrap">操作</th>
                            </tr>
                        </thead>
                        <tbody id="tableBody" class="divide-y divide-slate-100"></tbody>
                    </table>
                </div>
                <div id="emptyState" class="hidden p-8 text-center text-slate-400">
                    <i data-lucide="inbox" class="w-12 h-12 mx-auto mb-3 opacity-50"></i>
                    <p>目前沒有資料，請新增資料或貼上 Excel 內容。</p>
                </div>
            </section>
        </div>

        <!-- 分頁 2：圖表分析 -->
        <div id="tab-charts" class="tab-content hidden space-y-6">
            
            <!-- 1. 折線圖 (最上方) -->
            <section>
                <h2 class="text-lg font-semibold flex items-center gap-2 px-1 mb-4">
                    <i data-lucide="trending-up" class="w-5 h-5 text-emerald-500"></i> 每日觀看/紀錄次數
                </h2>
                <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
                    <div class="relative w-full h-[300px] md:h-[400px]">
                        <canvas id="dateChart"></canvas>
                    </div>
                </div>
            </section>

            <!-- 2. 圓餅圖 (折線圖下方) -->
            <section>
                <h2 class="text-lg font-semibold flex items-center gap-2 px-1 mb-4 pt-4">
                    <i data-lucide="pie-chart" class="w-5 h-5 text-indigo-500"></i> 佔比分析 <span class="text-sm font-normal text-slate-400 ml-2">(佔比低於 2% 之項目已自動歸類為「其他」)</span>
                </h2>
                <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                    <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center">
                        <h3 class="text-sm font-bold text-slate-600 mb-4 w-full text-center">公司佔比</h3>
                        <div class="relative w-full aspect-square max-w-[280px]">
                            <canvas id="companyChart"></canvas>
                        </div>
                    </div>
                    <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center">
                        <h3 class="text-sm font-bold text-slate-600 mb-4 w-full text-center">產業佔比</h3>
                        <div class="relative w-full aspect-square max-w-[280px]">
                            <canvas id="industryChart"></canvas>
                        </div>
                    </div>
                    <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center">
                        <h3 class="text-sm font-bold text-slate-600 mb-4 w-full text-center">職務佔比</h3>
                        <div class="relative w-full aspect-square max-w-[280px]">
                            <canvas id="dutyChart"></canvas>
                        </div>
                    </div>
                    <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center">
                        <h3 class="text-sm font-bold text-slate-600 mb-4 w-full text-center">職位佔比</h3>
                        <div class="relative w-full aspect-square max-w-[280px]">
                            <canvas id="positionChart"></canvas>
                        </div>
                    </div>
                </div>
            </section>
        </div>

        <!-- 分頁 3：詳細數據 -->
        <div id="tab-stats" class="tab-content hidden space-y-6">
            <h2 class="text-lg font-semibold flex items-center gap-2 px-1 mb-4">
                <i data-lucide="list-ordered" class="w-5 h-5 text-purple-500"></i> 詳細統計數據報表
            </h2>
            
            <div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6">
                <!-- 公司統計表格 -->
                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden self-start">
                    <div class="bg-blue-50/50 p-3 border-b border-slate-200 font-semibold text-slate-800 text-center whitespace-nowrap">公司統計</div>
                    <div class="overflow-x-auto">
                        <table class="w-full text-sm text-left">
                            <thead class="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-100">
                                <tr>
                                    <th class="px-4 py-3 whitespace-nowrap">項目</th>
                                    <th class="px-4 py-3 text-right whitespace-nowrap">筆數</th>
                                    <th class="px-4 py-3 text-right whitespace-nowrap">佔比</th>
                                </tr>
                            </thead>
                            <tbody id="companyStatsTable" class="divide-y divide-slate-100"></tbody>
                        </table>
                    </div>
                </div>

                <!-- 產業統計表格 -->
                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden self-start">
                    <div class="bg-emerald-50/50 p-3 border-b border-slate-200 font-semibold text-slate-800 text-center whitespace-nowrap">產業統計</div>
                    <div class="overflow-x-auto">
                        <table class="w-full text-sm text-left">
                            <thead class="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-100">
                                <tr>
                                    <th class="px-4 py-3 whitespace-nowrap">項目</th>
                                    <th class="px-4 py-3 text-right whitespace-nowrap">筆數</th>
                                    <th class="px-4 py-3 text-right whitespace-nowrap">佔比</th>
                                </tr>
                            </thead>
                            <tbody id="industryStatsTable" class="divide-y divide-slate-100"></tbody>
                        </table>
                    </div>
                </div>

                <!-- 職位統計表格 -->
                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden self-start">
                    <div class="bg-amber-50/50 p-3 border-b border-slate-200 font-semibold text-slate-800 text-center whitespace-nowrap">職位統計</div>
                    <div class="overflow-x-auto">
                        <table class="w-full text-sm text-left">
                            <thead class="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-100">
                                <tr>
                                    <th class="px-4 py-3 whitespace-nowrap">項目</th>
                                    <th class="px-4 py-3 text-right whitespace-nowrap">筆數</th>
                                    <th class="px-4 py-3 text-right whitespace-nowrap">佔比</th>
                                </tr>
                            </thead>
                            <tbody id="positionStatsTable" class="divide-y divide-slate-100"></tbody>
                        </table>
                    </div>
                </div>

                <!-- 職務統計表格 -->
                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden self-start">
                    <div class="bg-cyan-50/50 p-3 border-b border-slate-200 font-semibold text-slate-800 text-center whitespace-nowrap">職務統計</div>
                    <div class="overflow-x-auto">
                        <table class="w-full text-sm text-left">
                            <thead class="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-100">
                                <tr>
                                    <th class="px-4 py-3 whitespace-nowrap">項目</th>
                                    <th class="px-4 py-3 text-right whitespace-nowrap">筆數</th>
                                    <th class="px-4 py-3 text-right whitespace-nowrap">佔比</th>
                                </tr>
                            </thead>
                            <tbody id="dutyStatsTable" class="divide-y divide-slate-100"></tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

    </div>

    <script>
        lucide.createIcons();

        const initialData = [
            { company: "鵬鼎科技(股)", a: "A0058", industry: "半導體", c: "研發設計", position: "主任", date: "2026/4/28" },
            { company: "鵬鼎科技(股)", a: "A0305", industry: "半導體", c: "研發設計", position: "經理", date: "2026/4/24" },
            { company: "鵬鼎科技(股)", a: "A0314", industry: "半導體", c: "研發設計", position: "研究員/分析師", date: "2026/4/24" },
            { company: "鴻海精密工業(股)", a: "A0042", industry: "其他", c: "研發設計", position: "工程師", date: "2026/4/29" },
            { company: "鴻海精密工業(股)", a: "A0071", industry: "電腦與週邊", c: "未分類", position: "未分類", date: "2026/4/28" }
        ];

        let tableData = JSON.parse(JSON.stringify(initialData));
        
        let chartInstances = { company: null, industry: null, position: null, duty: null, date: null };

        const chartColors = [
            '#3b82f6', '#10b981', '#f59e0b', '#ef4444', 
            '#8b5cf6', '#ec4899', '#06b6d4', '#f97316',
            '#64748b', '#84cc16'
        ];

        // --- 分頁切換邏輯 ---
        function switchTab(tabId) {
            // 隱藏所有分頁內容
            document.querySelectorAll('.tab-content').forEach(el => {
                el.classList.add('hidden');
                el.classList.remove('block');
            });
            // 顯示目標分頁內容
            document.getElementById(`tab-${tabId}`).classList.remove('hidden');
            document.getElementById(`tab-${tabId}`).classList.add('block');

            // 更新按鈕樣式
            const tabs = ['input', 'charts', 'stats'];
            tabs.forEach(t => {
                const btn = document.getElementById(`btn-tab-${t}`);
                if (t === tabId) {
                    btn.classList.add('border-blue-500', 'text-blue-600');
                    btn.classList.remove('border-transparent', 'text-slate-500', 'hover:text-slate-700', 'hover:border-slate-300');
                } else {
                    btn.classList.remove('border-blue-500', 'text-blue-600');
                    btn.classList.add('border-transparent', 'text-slate-500', 'hover:text-slate-700', 'hover:border-slate-300');
                }
            });

            // 如果切換到圖表頁，重新渲染一次以避免 Canvas 尺寸在隱藏時被算錯
            if(tabId === 'charts') {
                updateCharts();
            }
        }

        // --- 表格操作邏輯 ---
        function renderTable() {
            const tbody = document.getElementById('tableBody');
            const emptyState = document.getElementById('emptyState');
            tbody.innerHTML = '';

            if (tableData.length === 0) {
                emptyState.classList.remove('hidden');
            } else {
                emptyState.classList.add('hidden');
            }

            tableData.forEach((row, index) => {
                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-50 transition-colors group";
                
                const fields = ['company', 'a', 'industry', 'c', 'position', 'date'];
                fields.forEach(field => {
                    const td = document.createElement('td');
                    td.className = "p-2";
                    td.innerHTML = `
                        <input type="text" 
                               value="${row[field]}" 
                               oninput="updateData(${index}, '${field}', this.value)"
                               class="w-full px-2 py-1.5 bg-transparent border border-transparent hover:border-slate-300 focus:border-blue-500 focus:bg-white focus:ring-2 focus:ring-blue-100 rounded transition-all outline-none text-slate-700 whitespace-nowrap">
                    `;
                    tr.appendChild(td);
                });

                const tdAction = document.createElement('td');
                tdAction.className = "p-2 text-center whitespace-nowrap";
                tdAction.innerHTML = `
                    <button onclick="deleteRow(${index})" class="p-1.5 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded transition-colors opacity-0 group-hover:opacity-100 focus:opacity-100" title="刪除此列">
                        <i data-lucide="trash-2" class="w-4 h-4"></i>
                    </button>
                `;
                tr.appendChild(tdAction);
                tbody.appendChild(tr);
            });
            
            lucide.createIcons();
            updateAllData(); // 同步更新圖表與統計表格
        }

        function updateData(index, field, value) {
            tableData[index][field] = value;
            updateAllData();
        }

        function addRow() {
            const today = new Date();
            const dateString = `${today.getFullYear()}/${today.getMonth() + 1}/${today.getDate()}`;
            tableData.push({ company: "", a: "", industry: "", c: "", position: "", date: dateString });
            renderTable();
        }

        function deleteRow(index) {
            tableData.splice(index, 1);
            renderTable();
        }

        function resetData() {
            if(confirm("確定要放棄目前修改，還原成預設的範例資料嗎？")) {
                tableData = JSON.parse(JSON.stringify(initialData));
                renderTable();
            }
        }

        function togglePasteArea() {
            const area = document.getElementById('pasteAreaContainer');
            if (area.classList.contains('hidden')) {
                area.classList.remove('hidden');
                document.getElementById('pasteInput').focus();
            } else {
                area.classList.add('hidden');
                document.getElementById('pasteInput').value = '';
            }
        }

        function processPastedData() {
            const text = document.getElementById('pasteInput').value.trim();
            if (!text) return alert("請貼上資料！");

            const rows = text.split('\n');
            const newData = [];
            let startIndex = (rows[0].includes('公司') || rows[0].includes('產業')) ? 1 : 0;

            for (let i = startIndex; i < rows.length; i++) {
                const cols = rows[i].split('\t');
                const finalCols = cols.length === 1 ? rows[i].split(',') : cols;
                if (finalCols.length >= 6) {
                    newData.push({
                        company: finalCols[0]?.trim() || "",
                        a: finalCols[1]?.trim() || "",
                        industry: finalCols[2]?.trim() || "",
                        c: finalCols[3]?.trim() || "",
                        position: finalCols[4]?.trim() || "",
                        date: finalCols[5]?.trim() || ""
                    });
                }
            }

            if (newData.length > 0) {
                tableData = newData;
                renderTable();
                togglePasteArea();
            } else {
                alert("無法解析資料，請確認格式是否正確 (需包含 6 個欄位)。");
            }
        }

        // --- 核心資料計算與圖表邏輯 ---

        function aggregateData(data, key) {
            return data.reduce((acc, row) => {
                const val = row[key] || '未填寫';
                acc[val] = (acc[val] || 0) + 1;
                return acc;
            }, {});
        }

        // 將小於 2% 的資料合併為「其他」，並依據數值由大到小排序
        function processPieDataForUnder2Percent(countsObj) {
            const total = Object.values(countsObj).reduce((sum, val) => sum + val, 0);
            if (total === 0) return {};

            let validItems = [];
            let otherCount = 0;

            for (const [key, val] of Object.entries(countsObj)) {
                const percentage = val / total;
                if (percentage < 0.02) {
                    otherCount += val; // 低於2% 歸入計數
                } else {
                    if (key === '其他') {
                        otherCount += val; // 遇到原本就是「其他」的也一起歸入整合的計數
                    } else {
                        validItems.push({key, val});
                    }
                }
            }

            // 依據數值由大到小排序
            validItems.sort((a, b) => b.val - a.val);

            // 「其他」固定置於最後
            if (otherCount > 0) {
                validItems.push({key: '其他', val: otherCount});
            }

            // 重組為 Object，JavaScript 會保留字串 key 的插入順序
            const result = {};
            validItems.forEach(item => {
                result[item.key] = item.val;
            });

            return result;
        }

        function drawPieChart(chartId, dataObj, chartKey) {
            const ctx = document.getElementById(chartId).getContext('2d');
            const labels = Object.keys(dataObj);
            const dataValues = Object.values(dataObj);

            if (chartInstances[chartKey]) chartInstances[chartKey].destroy();

            if (labels.length === 0) {
                chartInstances[chartKey] = new Chart(ctx, { type: 'pie', data: { labels: ['無資料'], datasets: [{ data: [1], backgroundColor: ['#e2e8f0'] }] } });
                return;
            }

            chartInstances[chartKey] = new Chart(ctx, {
                type: 'pie',
                data: {
                    labels: labels,
                    datasets: [{
                        data: dataValues,
                        backgroundColor: chartColors.slice(0, labels.length),
                        borderWidth: 2,
                        borderColor: '#ffffff'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { position: 'bottom', labels: { padding: 15 } },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    const total = context.chart._metasets[context.datasetIndex].total;
                                    const value = context.raw;
                                    const percentage = ((value / total) * 100).toFixed(1);
                                    return `${context.label}: ${value} 筆 (${percentage}%)`;
                                }
                            }
                        }
                    }
                }
            });
        }

        function drawLineChart(chartId, dataObj, chartKey) {
            const ctx = document.getElementById(chartId).getContext('2d');
            const sortedDates = Object.keys(dataObj).sort((a, b) => new Date(a) - new Date(b));
            const dataValues = sortedDates.map(date => dataObj[date]);

            if (chartInstances[chartKey]) chartInstances[chartKey].destroy();

            if (sortedDates.length === 0) {
                 chartInstances[chartKey] = new Chart(ctx, { type: 'line', data: { labels: ['無資料'], datasets: [{ data: [] }] }});
                 return;
            }

            chartInstances[chartKey] = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: sortedDates,
                    datasets: [{
                        label: '每日次數',
                        data: dataValues,
                        borderColor: '#10b981',
                        backgroundColor: 'rgba(16, 185, 129, 0.1)',
                        borderWidth: 3,
                        pointBackgroundColor: '#ffffff',
                        pointBorderColor: '#10b981',
                        pointBorderWidth: 2,
                        pointRadius: 4,
                        fill: true,
                        tension: 0.3
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } }, x: { grid: { display: false } } },
                    plugins: { legend: { display: false }, tooltip: { intersect: false, mode: 'index' } },
                    interaction: { mode: 'nearest', axis: 'x', intersect: false }
                }
            });
        }

        // --- 分頁 3：詳細統計表格建立 ---
        function renderStatsTables(validData) {
            const companyCounts = aggregateData(validData, 'company');
            const industryCounts = aggregateData(validData, 'industry');
            const dutyCounts = aggregateData(validData, 'c');
            const positionCounts = aggregateData(validData, 'position');
            const total = validData.length;

            buildStatTable('companyStatsTable', companyCounts, total);
            buildStatTable('industryStatsTable', industryCounts, total);
            buildStatTable('dutyStatsTable', dutyCounts, total);
            buildStatTable('positionStatsTable', positionCounts, total);
        }

        function buildStatTable(containerId, dataObj, totalCount) {
            const container = document.getElementById(containerId);
            if(totalCount === 0) {
                container.innerHTML = '<tr><td colspan="3" class="text-center p-4 text-slate-400">無資料</td></tr>';
                return;
            }

            // 依數量由大到小排序
            const sortedData = Object.entries(dataObj).sort((a, b) => b[1] - a[1]);

            let html = '';
            sortedData.forEach(([key, count]) => {
                const percentage = ((count / totalCount) * 100).toFixed(1);
                html += `
                    <tr class="hover:bg-slate-50 transition-colors">
                        <td class="px-4 py-3 text-slate-700 whitespace-nowrap">${key}</td>
                        <td class="px-4 py-3 text-slate-700 text-right font-medium whitespace-nowrap">${count}</td>
                        <td class="px-4 py-3 text-slate-500 text-right whitespace-nowrap">${percentage}%</td>
                    </tr>
                `;
            });
            container.innerHTML = html;
        }

        // 統一更新流程
        function updateAllData() {
            const validData = tableData.filter(row => 
                row.company.trim() !== '' || row.industry.trim() !== '' || row.position.trim() !== '' || row.date.trim() !== ''
            );

            // 1. 更新圖表
            updateCharts(validData);
            
            // 2. 更新詳細數據表格
            renderStatsTables(validData);
        }

        function updateCharts(validData = null) {
            if(!validData) {
                validData = tableData.filter(row => 
                    row.company.trim() !== '' || row.industry.trim() !== '' || row.position.trim() !== '' || row.date.trim() !== ''
                );
            }

            // 原始計數
            const companyCounts = aggregateData(validData, 'company');
            const industryCounts = aggregateData(validData, 'industry');
            const dutyCounts = aggregateData(validData, 'c');
            const positionCounts = aggregateData(validData, 'position');
            const dateCounts = aggregateData(validData, 'date');

            // 處理低於 2% 的項目，並進行排序
            const companyPieData = processPieDataForUnder2Percent(companyCounts);
            const industryPieData = processPieDataForUnder2Percent(industryCounts);
            const dutyPieData = processPieDataForUnder2Percent(dutyCounts);
            const positionPieData = processPieDataForUnder2Percent(positionCounts);

            drawLineChart('dateChart', dateCounts, 'date');
            drawPieChart('companyChart', companyPieData, 'company');
            drawPieChart('industryChart', industryPieData, 'industry');
            drawPieChart('dutyChart', dutyPieData, 'duty');
            drawPieChart('positionChart', positionPieData, 'position');
        }

        // 頁面載入初始化
        window.onload = () => {
            renderTable();
        };

    </script>
</body>
</html>
