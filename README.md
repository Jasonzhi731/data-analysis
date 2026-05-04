
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
    <!-- 引入 Excel 匯出與截圖相關套件 -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
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
        <header class="flex flex-col md:flex-row items-start md:items-center justify-between gap-4 mb-4">
            <div class="flex items-center space-x-3">
                <div class="p-3 bg-blue-600 rounded-lg text-white">
                    <i data-lucide="bar-chart-2"></i>
                </div>
                <div>
                    <h1 class="text-2xl font-bold text-slate-900">企業數據視覺化儀表板</h1>
                    <p class="text-sm text-slate-500">輸入或貼上表格資料，自動生成圖表與統計數據</p>
                </div>
            </div>
            <!-- 匯出 Excel 按鈕 -->
            <button id="exportBtn" onclick="exportToExcel()" class="px-4 py-2.5 bg-green-600 text-white rounded-lg flex items-center gap-2 shadow-sm hover:bg-green-700 transition-colors whitespace-nowrap font-medium text-sm">
                <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> 匯出整份 Excel 報表
            </button>
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
                    <div class="flex flex-wrap gap-2">
                        <button onclick="togglePasteArea()" class="px-4 py-2 text-sm font-medium text-blue-600 bg-blue-50 hover:bg-blue-100 rounded-lg transition-colors flex items-center gap-1">
                            <i data-lucide="clipboard-paste" class="w-4 h-4"></i> 輸入原始資料
                        </button>
                        <button onclick="addRow()" class="px-4 py-2 text-sm font-medium text-white bg-blue-600 hover:bg-blue-700 rounded-lg transition-colors flex items-center gap-1">
                            <i data-lucide="plus" class="w-4 h-4"></i> 新增一列
                        </button>
                        <button onclick="copyInputTable()" class="px-4 py-2 text-sm font-medium text-slate-600 bg-white border border-slate-200 hover:bg-slate-50 rounded-lg transition-colors flex items-center gap-1 shadow-sm">
                            <i data-lucide="copy" class="w-4 h-4"></i> 複製表格
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
                    <p>目前沒有資料，請輸入原始資料。</p>
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
                <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 px-1 mb-4 pt-4">
                    <h2 class="text-lg font-semibold flex items-center gap-2">
                        <i data-lucide="pie-chart" class="w-5 h-5 text-indigo-500"></i> 佔比分析 
                        <span class="text-sm font-normal text-slate-400 ml-2 hidden lg:inline">(佔比低於 2% 之項目已自動歸類為「其他」)</span>
                    </h2>
                    
                    <!-- 日期篩選下拉選單 (支援複選) -->
                    <div class="relative inline-block text-left" id="dateFilterContainer">
                        <button id="dateFilterBtn" onclick="toggleDateDropdown()" class="flex items-center gap-2 bg-indigo-50 px-3 py-1.5 rounded-lg border border-indigo-100 hover:bg-indigo-100 transition-colors">
                            <i data-lucide="filter" class="w-4 h-4 text-indigo-500"></i>
                            <span id="dateFilterLabel" class="text-sm font-medium text-indigo-900 whitespace-nowrap">日期: 全部</span>
                            <i data-lucide="chevron-down" class="w-4 h-4 text-indigo-500"></i>
                        </button>
                        <div id="dateDropdown" class="hidden absolute right-0 mt-2 w-48 bg-white rounded-xl shadow-lg border border-slate-200 z-[60] max-h-64 overflow-y-auto p-2">
                            <!-- 選項將由 JavaScript 動態生成 -->
                        </div>
                    </div>
                </div>
                <p class="text-sm font-normal text-slate-400 mb-4 px-1 lg:hidden">(佔比低於 2% 之項目已自動歸類為「其他」)</p>
                
                <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                    <!-- 公司佔比 -->
                    <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center">
                        <h3 class="text-sm font-bold text-slate-600 mb-4 w-full text-center">公司佔比</h3>
                        <div class="relative w-full h-[200px]">
                            <canvas id="companyChart"></canvas>
                        </div>
                        <!-- 自訂對齊圖例的容器 -->
                        <div id="companyChartLegend" class="w-full mt-4 flex-1"></div>
                    </div>
                    <!-- 產業佔比 -->
                    <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center">
                        <h3 class="text-sm font-bold text-slate-600 mb-4 w-full text-center">產業佔比</h3>
                        <div class="relative w-full h-[200px]">
                            <canvas id="industryChart"></canvas>
                        </div>
                        <div id="industryChartLegend" class="w-full mt-4 flex-1"></div>
                    </div>
                    <!-- 職務佔比 -->
                    <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center">
                        <h3 class="text-sm font-bold text-slate-600 mb-4 w-full text-center">職務佔比</h3>
                        <div class="relative w-full h-[200px]">
                            <canvas id="dutyChart"></canvas>
                        </div>
                        <div id="dutyChartLegend" class="w-full mt-4 flex-1"></div>
                    </div>
                    <!-- 職位佔比 -->
                    <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col items-center">
                        <h3 class="text-sm font-bold text-slate-600 mb-4 w-full text-center">職位佔比</h3>
                        <div class="relative w-full h-[200px]">
                            <canvas id="positionChart"></canvas>
                        </div>
                        <div id="positionChartLegend" class="w-full mt-4 flex-1"></div>
                    </div>
                </div>
            </section>
        </div>

        <!-- 分頁 3：詳細數據 -->
        <div id="tab-stats" class="tab-content hidden space-y-6">
            <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2 px-1 mb-4">
                <h2 class="text-lg font-semibold flex items-center gap-2">
                    <i data-lucide="list-ordered" class="w-5 h-5 text-purple-500"></i> 詳細統計數據報表
                    <span id="statsDateLabel" class="text-sm font-normal text-slate-500 ml-2">(資料日期: 全部)</span>
                </h2>
            </div>
            
            <!-- 上半部：公司與產業統計 -->
            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                <!-- 公司統計表格 -->
                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden self-start">
                    <div class="bg-blue-50/50 p-3 border-b border-slate-200 flex justify-between items-center">
                        <div class="font-semibold text-slate-800 whitespace-nowrap">公司統計</div>
                        <button onclick="copyTable('companyTable')" class="text-slate-500 hover:text-blue-600 flex items-center gap-1 text-xs font-medium px-2 py-1 rounded bg-white/50 hover:bg-white transition-colors border border-transparent hover:border-slate-300" title="複製表格內容">
                            <i data-lucide="copy" class="w-3.5 h-3.5"></i> 複製
                        </button>
                    </div>
                    <div class="overflow-x-auto">
                        <table id="companyTable" class="w-full text-sm text-left">
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
                    <div class="bg-emerald-50/50 p-3 border-b border-slate-200 flex justify-between items-center">
                        <div class="font-semibold text-slate-800 whitespace-nowrap">產業統計</div>
                        <button onclick="copyTable('industryTable')" class="text-slate-500 hover:text-emerald-600 flex items-center gap-1 text-xs font-medium px-2 py-1 rounded bg-white/50 hover:bg-white transition-colors border border-transparent hover:border-slate-300" title="複製表格內容">
                            <i data-lucide="copy" class="w-3.5 h-3.5"></i> 複製
                        </button>
                    </div>
                    <div class="overflow-x-auto">
                        <table id="industryTable" class="w-full text-sm text-left">
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
            </div>

            <!-- 下半部：職位與職務統計 -->
            <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                <!-- 職位統計表格 -->
                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden self-start">
                    <div class="bg-amber-50/50 p-3 border-b border-slate-200 flex justify-between items-center">
                        <div class="font-semibold text-slate-800 whitespace-nowrap">職位統計</div>
                        <button onclick="copyTable('positionTable')" class="text-slate-500 hover:text-amber-600 flex items-center gap-1 text-xs font-medium px-2 py-1 rounded bg-white/50 hover:bg-white transition-colors border border-transparent hover:border-slate-300" title="複製表格內容">
                            <i data-lucide="copy" class="w-3.5 h-3.5"></i> 複製
                        </button>
                    </div>
                    <div class="overflow-x-auto">
                        <table id="positionTable" class="w-full text-sm text-left">
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
                    <div class="bg-cyan-50/50 p-3 border-b border-slate-200 flex justify-between items-center">
                        <div class="font-semibold text-slate-800 whitespace-nowrap">職務統計</div>
                        <button onclick="copyTable('dutyTable')" class="text-slate-500 hover:text-cyan-600 flex items-center gap-1 text-xs font-medium px-2 py-1 rounded bg-white/50 hover:bg-white transition-colors border border-transparent hover:border-slate-300" title="複製表格內容">
                            <i data-lucide="copy" class="w-3.5 h-3.5"></i> 複製
                        </button>
                    </div>
                    <div class="overflow-x-auto">
                        <table id="dutyTable" class="w-full text-sm text-left">
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

    <!-- 複製成功提示區塊 (預設隱藏) -->
    <div id="copyToast" class="fixed bottom-4 right-4 bg-slate-800 text-white px-4 py-2 rounded shadow-lg flex items-center gap-2 transform translate-y-20 opacity-0 transition-all duration-300 pointer-events-none z-50">
        <i data-lucide="check-circle" class="w-4 h-4 text-green-400"></i>
        <span class="text-sm">表格已複製到剪貼簿</span>
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
        
        // 支援複選的日期篩選條件
        let availableDates = [];
        let selectedDates = [];

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

        // 開關日期篩選下拉選單
        function toggleDateDropdown() {
            document.getElementById('dateDropdown').classList.toggle('hidden');
        }

        // 點擊外部自動關閉下拉選單
        document.addEventListener('click', (e) => {
            const container = document.getElementById('dateFilterContainer');
            const dropdown = document.getElementById('dateDropdown');
            if (container && dropdown && !container.contains(e.target)) {
                dropdown.classList.add('hidden');
            }
        });

        // 更新佔比圖表的下拉選單選項
        function updateDateFilterOptions(validData) {
            const uniqueDates = [...new Set(validData.map(row => row.date).filter(d => d.trim() !== ''))]
                .sort((a, b) => new Date(a) - new Date(b));

            // 初始化或更新選項
            if (availableDates.length === 0) {
                selectedDates = [...uniqueDates];
            } else {
                const wasAllSelected = selectedDates.length === availableDates.length;
                selectedDates = selectedDates.filter(d => uniqueDates.includes(d));
                if (wasAllSelected) {
                    selectedDates = [...uniqueDates];
                }
            }
            availableDates = uniqueDates;

            renderDateCheckboxes();
            updateDateFilterLabel();
        }

        // 渲染複選框
        function renderDateCheckboxes() {
            const dropdown = document.getElementById('dateDropdown');
            if (!dropdown) return;
            
            const isAll = selectedDates.length === availableDates.length && availableDates.length > 0;
            
            let html = `
                <label class="flex items-center gap-2 px-3 py-2 hover:bg-slate-50 rounded-lg cursor-pointer border-b border-slate-100 mb-1 transition-colors">
                    <input type="checkbox" id="selectAllDates" onchange="toggleAllDates(this.checked)" class="w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500 border-slate-300" ${isAll ? 'checked' : ''}>
                    <span class="text-sm font-bold text-slate-700">全部日期</span>
                </label>
            `;

            availableDates.forEach(date => {
                const isChecked = selectedDates.includes(date);
                html += `
                    <label class="flex items-center gap-2 px-3 py-2 hover:bg-slate-50 rounded-lg cursor-pointer transition-colors">
                        <input type="checkbox" value="${date}" onchange="toggleSingleDate(this)" class="date-checkbox w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500 border-slate-300" ${isChecked ? 'checked' : ''}>
                        <span class="text-sm text-slate-700">${date}</span>
                    </label>
                `;
            });
            
            dropdown.innerHTML = html;
        }

        // 全選 / 取消全選
        function toggleAllDates(checked) {
            if (checked) {
                selectedDates = [...availableDates];
            } else {
                selectedDates = [];
            }
            updateCheckboxesState();
            triggerFilterUpdate();
        }

        // 單選一個日期
        function toggleSingleDate(checkbox) {
            const val = checkbox.value;
            if (checkbox.checked) {
                if (!selectedDates.includes(val)) selectedDates.push(val);
            } else {
                selectedDates = selectedDates.filter(d => d !== val);
            }
            updateCheckboxesState();
            triggerFilterUpdate();
        }

        // 更新複選框畫面狀態
        function updateCheckboxesState() {
            const allCheckbox = document.getElementById('selectAllDates');
            const checkboxes = document.querySelectorAll('.date-checkbox');
            
            const isAll = selectedDates.length === availableDates.length && availableDates.length > 0;
            if(allCheckbox) allCheckbox.checked = isAll;
            
            checkboxes.forEach(cb => {
                cb.checked = selectedDates.includes(cb.value);
            });
        }

        // 更新選單按鈕的文字與詳細數據標題的日期提示
        function updateDateFilterLabel() {
            const label = document.getElementById('dateFilterLabel');
            const statsLabel = document.getElementById('statsDateLabel');
            
            let labelText = '';
            if (selectedDates.length === availableDates.length || availableDates.length === 0) {
                labelText = '全部';
            } else if (selectedDates.length === 0) {
                labelText = '未選擇';
            } else if (selectedDates.length === 1) {
                labelText = selectedDates[0];
            } else {
                labelText = `已選 ${selectedDates.length} 項`;
            }

            if (label) {
                label.textContent = `日期: ${labelText}`;
            }
            if (statsLabel) {
                statsLabel.textContent = `(資料日期: ${labelText})`;
            }
        }

        // 取得篩選後的資料
        function getFilteredPieData(validData) {
            if (selectedDates.length === availableDates.length || availableDates.length === 0) {
                return validData;
            }
            return validData.filter(row => selectedDates.includes(row.date));
        }

        // 觸發圖表與詳細數據更新
        function triggerFilterUpdate() {
            updateDateFilterLabel();
            const validData = tableData.filter(row => 
                row.company.trim() !== '' || row.industry.trim() !== '' || row.position.trim() !== '' || row.date.trim() !== ''
            );
            updateCharts(validData);
            renderStatsTables(getFilteredPieData(validData));
        }

        function drawPieChart(chartId, dataObj, chartKey) {
            const ctx = document.getElementById(chartId).getContext('2d');
            const labels = Object.keys(dataObj);
            const dataValues = Object.values(dataObj);

            if (chartInstances[chartKey]) chartInstances[chartKey].destroy();

            if (labels.length === 0) {
                chartInstances[chartKey] = new Chart(ctx, { 
                    type: 'pie', 
                    data: { labels: ['無資料'], datasets: [{ data: [1], backgroundColor: ['#e2e8f0'] }] },
                    options: { plugins: { legend: { display: false }, tooltip: { enabled: false } } }
                });
                
                const legendContainer = document.getElementById(chartId + 'Legend');
                if (legendContainer) legendContainer.innerHTML = '<div class="text-center text-slate-400 text-xs py-2">無資料</div>';
                
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
                        legend: { 
                            display: false 
                        },
                        tooltip: {
                            callbacks: {
                                // 隱藏原本的預設標題 (避免文字重複)
                                title: function() { return null; },
                                // 客製化顯示標籤格式為兩行:
                                // 第一行: 項目名稱:
                                // 第二行: ?筆 (??%)
                                label: function(context) {
                                    const total = context.chart._metasets[context.datasetIndex].total;
                                    const value = context.raw;
                                    const percentage = ((value / total) * 100).toFixed(1);
                                    // 透過回傳陣列，讓 Chart.js 在 Tooltip 中換行顯示
                                    return [
                                        `${context.label}:`,
                                        `${value}筆 (${percentage}%)`
                                    ];
                                }
                            }
                        }
                    }
                },
                // 自訂 HTML 圖例外掛程式
                plugins: [{
                    id: 'customHtmlLegend',
                    afterUpdate: function(chart) {
                        const container = document.getElementById(chart.canvas.id + 'Legend');
                        if (!container) return;
                        container.innerHTML = '';
                        
                        const ul = document.createElement('ul');
                        ul.className = 'grid grid-cols-[repeat(auto-fit,minmax(100px,1fr))] gap-x-2 gap-y-2.5 w-full text-xs text-slate-600 px-1';

                        const items = chart.options.plugins.legend.labels.generateLabels(chart);
                        items.forEach(item => {
                            const li = document.createElement('li');
                            li.className = 'flex items-start gap-1.5 cursor-pointer transition-all hover:opacity-80';
                            li.onclick = () => {
                                chart.toggleDataVisibility(item.index);
                                chart.update();
                            };
                            
                            if (item.hidden) {
                                li.classList.add('opacity-40', 'line-through');
                            }

                            const box = document.createElement('span');
                            box.className = 'w-3 h-3 shrink-0 rounded-[2px] mt-[2px] block border';
                            box.style.backgroundColor = item.fillStyle;
                            box.style.borderColor = item.strokeStyle || '#ffffff';
                            box.style.borderWidth = (item.lineWidth || 1) + 'px';

                            const text = document.createElement('span');
                            text.className = 'text-left leading-snug break-words flex-1';
                            text.textContent = item.text;

                            li.appendChild(box);
                            li.appendChild(text);
                            ul.appendChild(li);
                        });
                        container.appendChild(ul);
                    }
                }]
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

        function updateAllData() {
            const validData = tableData.filter(row => 
                row.company.trim() !== '' || row.industry.trim() !== '' || row.position.trim() !== '' || row.date.trim() !== ''
            );
            
            // 更新下拉選單裡的可用日期
            updateDateFilterOptions(validData);

            updateCharts(validData);
            
            // 圓餅圖與詳細數據表格使用「篩選後的資料」繪製
            const filteredPieData = getFilteredPieData(validData);
            renderStatsTables(filteredPieData);
        }

        function updateCharts(validData = null) {
            if(!validData) {
                validData = tableData.filter(row => 
                    row.company.trim() !== '' || row.industry.trim() !== '' || row.position.trim() !== '' || row.date.trim() !== ''
                );
            }

            // 折線圖永遠使用「全部資料」繪製
            const dateCounts = aggregateData(validData, 'date');
            drawLineChart('dateChart', dateCounts, 'date');

            // 圓餅圖使用「篩選後的資料」繪製
            const filteredPieData = getFilteredPieData(validData);

            const companyCounts = aggregateData(filteredPieData, 'company');
            const industryCounts = aggregateData(filteredPieData, 'industry');
            const dutyCounts = aggregateData(filteredPieData, 'c');
            const positionCounts = aggregateData(filteredPieData, 'position');

            const companyPieData = processPieDataForUnder2Percent(companyCounts);
            const industryPieData = processPieDataForUnder2Percent(industryCounts);
            const dutyPieData = processPieDataForUnder2Percent(dutyCounts);
            const positionPieData = processPieDataForUnder2Percent(positionCounts);

            drawPieChart('companyChart', companyPieData, 'company');
            drawPieChart('industryChart', industryPieData, 'industry');
            drawPieChart('dutyChart', dutyPieData, 'duty');
            drawPieChart('positionChart', positionPieData, 'position');
        }

        // --- 複製表格功能 ---
        function copyInputTable() {
            let textToCopy = '公司\t個編帳號\t產業\t職務\t職位\t日期\n';
            tableData.forEach(row => {
                textToCopy += `${row.company}\t${row.a}\t${row.industry}\t${row.c}\t${row.position}\t${row.date}\n`;
            });
            
            const textArea = document.createElement("textarea");
            textArea.value = textToCopy;
            document.body.appendChild(textArea);
            textArea.select();
            
            try {
                document.execCommand('copy');
                showToast();
            } catch (err) {
                alert('複製失敗，您的瀏覽器可能不支援此功能。');
            }
            document.body.removeChild(textArea);
        }

        function copyTable(tableId) {
            const table = document.getElementById(tableId);
            let textToCopy = '';
            const rows = table.querySelectorAll('tr');
            
            rows.forEach(row => {
                const cells = row.querySelectorAll('th, td');
                const rowData = Array.from(cells).map(cell => cell.innerText.trim());
                textToCopy += rowData.join('\t') + '\n';
            });

            const textArea = document.createElement("textarea");
            textArea.value = textToCopy;
            document.body.appendChild(textArea);
            textArea.select();
            
            try {
                document.execCommand('copy');
                showToast();
            } catch (err) {
                alert('複製失敗，您的瀏覽器可能不支援此功能。');
            }
            document.body.removeChild(textArea);
        }

        function showToast() {
            const toast = document.getElementById('copyToast');
            toast.classList.remove('translate-y-20', 'opacity-0');
            toast.classList.add('translate-y-0', 'opacity-100');
            
            setTimeout(() => {
                toast.classList.add('translate-y-20', 'opacity-0');
                toast.classList.remove('translate-y-0', 'opacity-100');
            }, 2500);
        }

        // --- 匯出 Excel 功能 ---
        async function exportToExcel() {
            // 提示使用者輸入自訂檔案名稱
            const fileNameInput = prompt("請輸入匯出檔案名稱：", "企業數據視覺化報表");
            
            // 若使用者點擊取消或未輸入名稱，則中斷匯出
            if (fileNameInput === null) return;
            
            // 處理副檔名確保匯出檔案格式正確
            const baseFileName = fileNameInput.trim() === "" ? "企業數據視覺化報表" : fileNameInput.trim();
            const exportFileName = baseFileName.endsWith(".xlsx") ? baseFileName : baseFileName + ".xlsx";

            const btn = document.getElementById('exportBtn');
            const originalBtnContent = btn.innerHTML;
            
            try {
                // 更新按鈕狀態
                btn.innerHTML = '<i data-lucide="loader" class="w-4 h-4 animate-spin"></i> 報表生成中...';
                lucide.createIcons();
                btn.disabled = true;

                const wb = new ExcelJS.Workbook();
                wb.creator = '企業數據視覺化儀表板';
                wb.created = new Date();

                // === Sheet 1: 資料輸入 ===
                const ws1 = wb.addWorksheet('資料輸入');
                ws1.columns = [
                    { header: '公司', key: 'company', width: 25 },
                    { header: '個編帳號', key: 'a', width: 15 },
                    { header: '產業', key: 'industry', width: 20 },
                    { header: '職務', key: 'c', width: 20 },
                    { header: '職位', key: 'position', width: 20 },
                    { header: '日期', key: 'date', width: 15 }
                ];
                ws1.getRow(1).font = { bold: true };
                tableData.forEach(row => {
                    ws1.addRow(row);
                });

                // === Sheet 2: 圖表分析 ===
                const ws2 = wb.addWorksheet('圖表分析');
                
                // 為了讓 html2canvas 能順利截圖，先切換到圖表頁面
                const activeTabBtn = document.querySelector('nav button.border-blue-500').id;
                if (activeTabBtn !== 'btn-tab-charts') {
                    switchTab('charts');
                    // 稍微等待 DOM 渲染與動畫
                    await new Promise(resolve => setTimeout(resolve, 300));
                }

                const chartsTab = document.getElementById('tab-charts');
                
                // 執行畫面截圖 (使用淡色背景確保截圖沒有透明區塊)
                const canvas = await html2canvas(chartsTab, { 
                    scale: 2, 
                    backgroundColor: '#f8fafc',
                    logging: false
                });
                
                const imgData = canvas.toDataURL('image/png');
                
                // 將圖表影像加入工作簿
                const imageId = wb.addImage({
                    base64: imgData,
                    extension: 'png',
                });
                
                // 設定圖片在 Excel 中的尺寸與位置
                ws2.addImage(imageId, {
                    tl: { col: 1, row: 1 }, // 放於 B2 位置 (索引為 1,1)
                    ext: { width: canvas.width / 2, height: canvas.height / 2 }
                });

                // 截圖完畢切回原本的分頁
                if (activeTabBtn !== 'btn-tab-charts') {
                    switchTab(activeTabBtn.replace('btn-tab-', ''));
                }

                // === Sheet 3: 詳細數據 ===
                const ws3 = wb.addWorksheet('詳細數據');
                ws3.getColumn(1).width = 30;
                ws3.getColumn(2).width = 12;
                ws3.getColumn(3).width = 12;

                const validData = tableData.filter(row => row.company.trim() !== '' || row.industry.trim() !== '' || row.position.trim() !== '' || row.date.trim() !== '');
                // 讓匯出的數據與畫面上目前的篩選狀態同步
                const filteredData = getFilteredPieData(validData);
                const total = filteredData.length;

                // 建立統計表格的輔助函式
                const addStatBlock = (title, key) => {
                    const titleRow = ws3.addRow([title]);
                    titleRow.font = { bold: true, size: 14 };
                    
                    const headerRow = ws3.addRow(['項目', '筆數', '佔比']);
                    headerRow.font = { bold: true };
                    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };

                    if (total === 0) {
                        ws3.addRow(['無資料', '', '']);
                    } else {
                        const counts = aggregateData(filteredData, key);
                        const sortedData = Object.entries(counts).sort((a, b) => b[1] - a[1]);
                        
                        sortedData.forEach(([k, v]) => {
                            const percentage = ((v / total) * 100).toFixed(1) + '%';
                            ws3.addRow([k, v, percentage]);
                        });
                    }
                    ws3.addRow([]); // 加入空行隔開表格
                };

                addStatBlock('【公司統計】', 'company');
                addStatBlock('【產業統計】', 'industry');
                addStatBlock('【職位統計】', 'position');
                addStatBlock('【職務統計】', 'c');

                // 輸出並下載 Excel 檔案，使用自訂的檔案名稱
                const buffer = await wb.xlsx.writeBuffer();
                saveAs(new Blob([buffer]), exportFileName);

            } catch (error) {
                console.error('Excel 匯出發生錯誤:', error);
                alert('匯出失敗，請確認您的瀏覽器是否阻擋了下載操作。');
            } finally {
                // 還原按鈕狀態
                btn.innerHTML = originalBtnContent;
                lucide.createIcons();
                btn.disabled = false;
            }
        }

        window.onload = () => {
            renderTable();
        };

    </script>
</body>
</html>
