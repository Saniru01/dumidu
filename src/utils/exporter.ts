import { AirFreightRecord } from '../types';

export function generateStandaloneReportHtml(appData: AirFreightRecord[]): string {
  const dataString = JSON.stringify(appData);
  return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Air Freight Failure Cost Analytics</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- SheetJS -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
</head>
<body class="bg-slate-100 text-slate-800 font-sans min-h-screen flex flex-col">

    <!-- Header Navigation -->
    <header class="bg-slate-900 text-white shadow-md sticky top-0 z-50">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 py-4 flex flex-col md:flex-row justify-between items-center gap-4">
            <div class="flex items-center space-x-3">
                <div class="bg-indigo-600 p-2.5 rounded-xl text-white shadow-inner">
                    <i class="fa-solid fa-plane-circle-exclamation text-xl"></i>
                </div>
                <div>
                    <h1 class="text-xl font-bold tracking-wide">Air Freight Cost of Failures</h1>
                    <p class="text-xs text-slate-400">Weekly Operations &amp; Delay Analytics</p>
                </div>
            </div>

            <!-- Upload, Export & Storage Action Controls -->
            <div class="flex items-center flex-wrap gap-2">
                <label class="cursor-pointer flex items-center gap-2 bg-indigo-600 hover:bg-indigo-500 text-white text-xs sm:text-sm font-semibold px-4 py-2 rounded-lg shadow transition duration-150">
                    <i class="fa-solid fa-cloud-arrow-up"></i>
                    <span id="uploadBtnText">Upload Excel</span>
                    <input type="file" id="fileInput" accept=".xlsx, .xls, .csv" class="hidden">
                </label>
                <button onclick="exportFilledHTML()" class="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs sm:text-sm font-semibold px-4 py-2 rounded-lg shadow transition duration-150" title="Export HTML file with data embedded for email sharing">
                    <i class="fa-solid fa-file-export"></i>
                    <span>Export for Email</span>
                </button>
                <button onclick="resetFilters()" class="text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 px-3 py-2 rounded-lg border border-slate-700 transition" title="Reset Filters">
                    <i class="fa-solid fa-rotate-right"></i> Reset Filters
                </button>
                <button onclick="clearSavedData()" class="text-xs bg-rose-950 hover:bg-rose-900 text-rose-300 px-3 py-2 rounded-lg border border-rose-800 transition" title="Clear Auto-Saved Data">
                    <i class="fa-solid fa-trash-can"></i> Clear Storage
                </button>
            </div>
        </div>
    </header>

    <div class="max-w-7xl mx-auto px-4 sm:px-6 py-6 space-y-6 flex-1 w-full">

        <!-- Notification / Status Banner -->
        <div id="uploadAlert" class="hidden p-3 bg-emerald-100 border border-emerald-300 text-emerald-800 text-xs font-semibold rounded-xl flex items-center justify-between transition-all">
            <span id="uploadAlertMsg"><i class="fa-solid fa-circle-check mr-2"></i> File loaded and saved in browser cache.</span>
            <button onclick="document.getElementById('uploadAlert').classList.add('hidden')" class="text-emerald-700 hover:text-emerald-900">&times;</button>
        </div>

        <!-- Filter Bar -->
        <section class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200">
            <div class="flex flex-col sm:flex-row items-start sm:items-center justify-between pb-3 border-b border-slate-100 mb-4 gap-2">
                <span class="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-2">
                    <i class="fa-solid fa-sliders text-indigo-600"></i> Interactive Filters
                </span>
                <div class="flex items-center gap-2">
                    <span id="storageStatusBadge" class="text-xs font-medium bg-emerald-50 text-emerald-700 px-2.5 py-0.5 rounded-full border border-emerald-200">
                        <i class="fa-solid fa-hard-drive mr-1"></i> Auto-Saved Active
                    </span>
                    <span id="recordCountBadge" class="text-xs font-medium bg-indigo-50 text-indigo-700 px-2.5 py-0.5 rounded-full border border-indigo-100">
                        0 records
                    </span>
                </div>
            </div>
            
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <!-- Year -->
                <div>
                    <label class="block text-xs font-semibold text-slate-600 mb-1">Year</label>
                    <select id="selYear" class="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
                        <option value="ALL">All Years</option>
                    </select>
                </div>

                <!-- Month -->
                <div>
                    <label class="block text-xs font-semibold text-slate-600 mb-1">Month</label>
                    <select id="selMonth" class="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
                        <option value="ALL">All Months</option>
                    </select>
                </div>

                <!-- Start Week -->
                <div>
                    <label class="block text-xs font-semibold text-slate-600 mb-1">From Week</label>
                    <select id="selWeekFrom" class="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
                        <option value="ALL">Start Week</option>
                    </select>
                </div>

                <!-- End Week -->
                <div>
                    <label class="block text-xs font-semibold text-slate-600 mb-1">To Week</label>
                    <select id="selWeekTo" class="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
                        <option value="ALL">End Week</option>
                    </select>
                </div>
            </div>
        </section>

        <!-- KPI Summary Cards -->
        <section class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                <div>
                    <p class="text-xs font-bold text-slate-400 uppercase tracking-wider">Total AF Cost</p>
                    <h3 id="kpiTotalCost" class="text-2xl font-black text-rose-600 mt-1">$0.00</h3>
                </div>
                <div class="bg-rose-50 text-rose-600 w-12 h-12 rounded-xl flex items-center justify-center text-xl">
                    <i class="fa-solid fa-dollar-sign"></i>
                </div>
            </div>

            <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                <div>
                    <p class="text-xs font-bold text-slate-400 uppercase tracking-wider">Total Incidents</p>
                    <h3 id="kpiTotalShipments" class="text-2xl font-black text-slate-800 mt-1">0</h3>
                </div>
                <div class="bg-indigo-50 text-indigo-600 w-12 h-12 rounded-xl flex items-center justify-center text-xl">
                    <i class="fa-solid fa-boxes-packing"></i>
                </div>
            </div>

            <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                <div>
                    <p class="text-xs font-bold text-slate-400 uppercase tracking-wider">Top Cost Dept</p>
                    <h3 id="kpiTopDept" class="text-base font-bold text-slate-800 mt-1 truncate max-w-[140px]">-</h3>
                </div>
                <div class="bg-emerald-50 text-emerald-600 w-12 h-12 rounded-xl flex items-center justify-center text-xl">
                    <i class="fa-solid fa-building-user"></i>
                </div>
            </div>

            <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                <div>
                    <p class="text-xs font-bold text-slate-400 uppercase tracking-wider">Top Delay Reason</p>
                    <h3 id="kpiTopReason" class="text-base font-bold text-slate-800 mt-1 truncate max-w-[140px]">-</h3>
                </div>
                <div class="bg-amber-50 text-amber-600 w-12 h-12 rounded-xl flex items-center justify-center text-xl">
                    <i class="fa-solid fa-triangle-exclamation"></i>
                </div>
            </div>
        </section>

        <!-- Charts Dashboard Grid -->
        <section class="grid grid-cols-1 lg:grid-cols-2 gap-6">

            <!-- Chart 1: Weekly Air Freight Cost Trend -->
            <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
                <div class="flex items-center justify-between mb-4">
                    <h2 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                        <i class="fa-solid fa-chart-line text-indigo-600"></i> 1. Weekly Air Freight Cost Trend ($ USD)
                    </h2>
                    <button onclick="copyChartImage('chartWeekly', this)" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold px-2.5 py-1 rounded-md border border-slate-300 transition flex items-center gap-1.5" title="Copy chart to clipboard">
                        <i class="fa-regular fa-copy"></i> <span>Copy</span>
                    </button>
                </div>
                <div class="relative h-96 w-full">
                    <canvas id="chartWeekly"></canvas>
                </div>
            </div>

            <!-- Chart 2: Responsible Department Wise (PIE CHART) -->
            <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
                <div class="flex items-center justify-between mb-4">
                    <h2 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                        <i class="fa-solid fa-chart-pie text-indigo-600"></i> 2. Responsible Department Cost Distribution
                    </h2>
                    <button onclick="copyChartImage('chartDept', this)" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold px-2.5 py-1 rounded-md border border-slate-300 transition flex items-center gap-1.5" title="Copy chart to clipboard">
                        <i class="fa-regular fa-copy"></i> <span>Copy</span>
                    </button>
                </div>
                <div class="relative h-96 w-full flex items-center justify-center">
                    <canvas id="chartDept"></canvas>
                </div>
            </div>

            <!-- Chart 3: Reason Wise Breakdown -->
            <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
                <div class="flex items-center justify-between mb-4">
                    <h2 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                        <i class="fa-solid fa-chart-bar text-indigo-600"></i> 3. Failure Reason Wise Cost (Top 10)
                    </h2>
                    <button onclick="copyChartImage('chartReason', this)" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold px-2.5 py-1 rounded-md border border-slate-300 transition flex items-center gap-1.5" title="Copy chart to clipboard">
                        <i class="fa-regular fa-copy"></i> <span>Copy</span>
                    </button>
                </div>
                <div class="relative h-96 w-full">
                    <canvas id="chartReason"></canvas>
                </div>
            </div>

            <!-- Chart 4: Plant Wise Breakdown -->
            <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
                <div class="flex items-center justify-between mb-4">
                    <h2 class="text-sm font-bold text-slate-800 flex items-center gap-2">
                        <i class="fa-solid fa-chart-pie text-indigo-600"></i> 4. Plant Wise Cost Distribution
                    </h2>
                    <button onclick="copyChartImage('chartPlant', this)" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold px-2.5 py-1 rounded-md border border-slate-300 transition flex items-center gap-1.5" title="Copy chart to clipboard">
                        <i class="fa-regular fa-copy"></i> <span>Copy</span>
                    </button>
                </div>
                <div class="relative h-96 w-full flex items-center justify-center">
                    <canvas id="chartPlant"></canvas>
                </div>
            </div>

        </section>

    </div>

    <!-- Application Engine -->
    <script>
        const STORAGE_KEY = 'AF_DASHBOARD_DATA';
        const STORAGE_META_KEY = 'AF_DASHBOARD_META';
        const STORAGE_FILTER_KEY = 'AF_DASHBOARD_FILTERS';

        /* EMBEDDED DATA */
        let appData = ${dataString};
        let chartInstances = {};

        const chartColors = [
            '#4f46e5', '#06b6d4', '#10b981', '#f59e0b', '#ef4444',
            '#8b5cf6', '#ec4899', '#3b82f6', '#14b8a6', '#f97316',
            '#6366f1', '#84cc16', '#a855f7', '#e11d48', '#0ea5e9',
            '#64748b', '#d97706', '#059669', '#2563eb', '#7c3aed'
        ];

        function findValue(row, possibleNames, defaultVal = '') {
            const keys = Object.keys(row);
            for (const name of possibleNames) {
                const target = name.toLowerCase().replace(/[^a-z0-9]/g, '');
                for (const k of keys) {
                    const cleanKey = k.toLowerCase().replace(/[^a-z0-9]/g, '');
                    if (cleanKey === target && row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') {
                        return row[k];
                    }
                }
            }
            return defaultVal;
        }

        function parseCleanRow(r) {
            let rawYear = findValue(r, ['No', 'Year', 'Shipment Year'], '2026');
            let year = String(parseInt(rawYear, 10) || rawYear).trim();

            let rawWeek = findValue(r, ['Shipment week', 'Week', 'Shipment Wk', 'Wk'], 0);
            let week = parseInt(rawWeek, 10) || 0;

            let plant = String(findValue(r, ['Column1', 'Plant', 'Plant Name', 'Unit'], 'Unspecified Plant')).trim();
            let dept = String(findValue(r, ['Responsible Department', 'Department', 'Resp Dept', 'Dept'], 'Unassigned')).trim();
            let reason = String(findValue(r, ['Reason for delay', 'Reason', 'Delay Reason'], 'Unspecified')).trim();

            let rawCost = findValue(r, ['Cost (USD)', 'Cost', 'Amount', 'Cost USD'], 0);
            let cost = typeof rawCost === 'number' ? rawCost : parseFloat(String(rawCost).replace(/[^0-9.-]+/g, '')) || 0;

            let month = String(findValue(r, ['Month', 'Shipment Month'], 'Unspecified')).trim();

            return { year, week, plant, dept, reason, cost, month };
        }

        // --- EXCEL UPLOAD & LOCAL STORAGE SYNC ---
        document.getElementById('fileInput').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const rawJson = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { defval: "" });

                    if (rawJson.length === 0) {
                        alert("The uploaded Excel sheet contains no data.");
                        return;
                    }

                    appData = rawJson.map(parseCleanRow).filter(item => item.cost > 0 || item.week > 0);
                    
                    localStorage.setItem(STORAGE_KEY, JSON.stringify(appData));
                    localStorage.setItem(STORAGE_META_KEY, JSON.stringify({
                        fileName: file.name,
                        uploadDate: new Date().toLocaleString()
                    }));

                    resetFilters();
                    populateFilterDropdowns();
                    refreshDashboard();
                    updateStorageIndicator();

                    const alertBox = document.getElementById('uploadAlert');
                    const alertMsg = document.getElementById('uploadAlertMsg');
                    alertMsg.innerHTML = '<i class="fa-solid fa-circle-check mr-2"></i> Successfully loaded and auto-saved <strong>' + file.name + '</strong> (' + appData.length + ' records).';
                    alertBox.classList.remove('hidden');
                    setTimeout(() => alertBox.classList.add('hidden'), 5000);
                } catch (err) {
                    console.error(err);
                    alert("Error processing Excel file. Please verify column structure.");
                }
            };

            reader.readAsArrayBuffer(file);
            this.value = '';
        });

        function exportFilledHTML() {
            if (!appData || appData.length === 0) {
                alert("Please upload an Excel file first before exporting.");
                return;
            }

            let htmlSource = document.documentElement.outerHTML;
            const cleanAlertRegex = /<div id="uploadAlert" class="[^"]*"/;
            htmlSource = htmlSource.replace(cleanAlertRegex, '<div id="uploadAlert" class="hidden');

            const dataString = JSON.stringify(appData);
            htmlSource = htmlSource.replace(/let appData = \\[.*?\\];/, 'let appData = ' + dataString + ';');

            const blob = new Blob(["<!DOCTYPE html>\\n<html lang=\\"en\\">\\n" + htmlSource + "\\n</html>"], { type: 'text/html;charset=utf-8' });
            const downloadLink = document.createElement('a');
            const dateStamp = new Date().toISOString().split('T')[0];
            downloadLink.download = "Air_Freight_Report_" + dateStamp + ".html";
            downloadLink.href = URL.createObjectURL(blob);
            
            document.body.appendChild(downloadLink);
            downloadLink.click();
            document.body.removeChild(downloadLink);
            URL.revokeObjectURL(downloadLink.href);
        }

        function clearSavedData() {
            if (confirm("Are you sure you want to clear the saved Excel data and active filters from your browser?")) {
                localStorage.removeItem(STORAGE_KEY);
                localStorage.removeItem(STORAGE_META_KEY);
                localStorage.removeItem(STORAGE_FILTER_KEY);
                appData = [];
                resetFilters();
                populateFilterDropdowns();
                refreshDashboard();
                updateStorageIndicator();
            }
        }

        function updateStorageIndicator() {
            const metaRaw = localStorage.getItem(STORAGE_META_KEY);
            const badge = document.getElementById('storageStatusBadge');
            if (appData.length > 0) {
                badge.className = "text-xs font-medium bg-emerald-50 text-emerald-700 px-2.5 py-0.5 rounded-full border border-emerald-200";
                badge.innerHTML = '<i class="fa-solid fa-hard-drive mr-1"></i> Data Ready (' + appData.length + ' records)';
            } else {
                badge.className = "text-xs font-medium bg-amber-50 text-amber-700 px-2.5 py-0.5 rounded-full border border-amber-200";
                badge.innerHTML = '<i class="fa-solid fa-triangle-exclamation mr-1"></i> No Data Loaded';
            }
        }

        async function copyChartImage(canvasId, btnElement) {
            const chartCanvas = document.getElementById(canvasId);
            if (!chartCanvas) return;

            const exportCanvas = document.createElement('canvas');
            exportCanvas.width = chartCanvas.width;
            exportCanvas.height = chartCanvas.height;
            const ctx = exportCanvas.getContext('2d');

            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, exportCanvas.width, exportCanvas.height);
            ctx.drawImage(chartCanvas, 0, 0);

            exportCanvas.toBlob(async (blob) => {
                try {
                    await navigator.clipboard.write([
                        new ClipboardItem({ 'image/png': blob })
                    ]);
                    
                    const originalHTML = btnElement.innerHTML;
                    btnElement.innerHTML = '<i class="fa-solid fa-check text-emerald-600"></i> <span class="text-emerald-700">Copied!</span>';
                    btnElement.classList.add('bg-emerald-50', 'border-emerald-300');
                    
                    setTimeout(() => {
                        btnElement.innerHTML = originalHTML;
                        btnElement.classList.remove('bg-emerald-50', 'border-emerald-300');
                    }, 2000);
                } catch (err) {
                    console.error("Clipboard copy failed:", err);
                    const link = document.createElement('a');
                    link.download = canvasId + "-export.png";
                    link.href = exportCanvas.toDataURL('image/png');
                    link.click();
                }
            }, 'image/png');
        }

        function saveCurrentFilters() {
            const filters = {
                year: document.getElementById('selYear').value,
                month: document.getElementById('selMonth').value,
                weekFrom: document.getElementById('selWeekFrom').value,
                weekTo: document.getElementById('selWeekTo').value
            };
            localStorage.setItem(STORAGE_FILTER_KEY, JSON.stringify(filters));
        }

        function populateFilterDropdowns() {
            const years = [...new Set(appData.map(d => d.year))].filter(Boolean).sort();
            const monthsOrder = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
            const rawMonths = [...new Set(appData.map(d => d.month))].filter(m => m && m !== 'Unspecified');
            const months = monthsOrder.filter(m => rawMonths.some(rm => rm.toLowerCase() === m.toLowerCase()));
            const weeks = [...new Set(appData.map(d => d.week))].filter(w => w > 0).sort((a,b) => a - b);

            const savedFilters = JSON.parse(localStorage.getItem(STORAGE_FILTER_KEY) || '{}');

            const fill = (elemId, items, placeholder, savedVal) => {
                const el = document.getElementById(elemId);
                el.innerHTML = '<option value="ALL">' + placeholder + '</option>';
                items.forEach(it => {
                    el.innerHTML += '<option value="' + it + '">' + it + '</option>';
                });
                if (savedVal && (items.includes(savedVal) || items.includes(parseInt(savedVal, 10)) || savedVal === 'ALL')) {
                    el.value = savedVal;
                }
            };

            fill('selYear', years, 'All Years', savedFilters.year);
            fill('selMonth', months, 'All Months', savedFilters.month);
            fill('selWeekFrom', weeks, 'Start Week (Min)', savedFilters.weekFrom);
            fill('selWeekTo', weeks, 'End Week (Max)', savedFilters.weekTo);

            ['selYear', 'selMonth', 'selWeekFrom', 'selWeekTo'].forEach(id => {
                document.getElementById(id).onchange = () => {
                    saveCurrentFilters();
                    refreshDashboard();
                };
            });
        }

        function resetFilters() {
            document.getElementById('selYear').value = 'ALL';
            document.getElementById('selMonth').value = 'ALL';
            document.getElementById('selWeekFrom').value = 'ALL';
            document.getElementById('selWeekTo').value = 'ALL';
            saveCurrentFilters();
            refreshDashboard();
        }

        function getFilteredDataset() {
            const y = document.getElementById('selYear').value;
            const m = document.getElementById('selMonth').value;
            const wf = document.getElementById('selWeekFrom').value;
            const wt = document.getElementById('selWeekTo').value;

            return appData.filter(d => {
                if (y !== 'ALL' && d.year !== y) return false;
                if (m !== 'ALL' && d.month.toLowerCase() !== m.toLowerCase()) return false;
                if (wf !== 'ALL' && d.week < parseInt(wf, 10)) return false;
                if (wt !== 'ALL' && d.week > parseInt(wt, 10)) return false;
                return true;
            });
        }

        function aggregateBy(records, key) {
            return records.reduce((acc, row) => {
                const val = row[key] || 'Other';
                acc[val] = (acc[val] || 0) + row.cost;
                return acc;
            }, {});
        }

        function createOrUpdateChart(canvasId, type, labels, data, labelName, isCurrency = true) {
            if (chartInstances[canvasId]) {
                chartInstances[canvasId].destroy();
            }

            const ctx = document.getElementById(canvasId).getContext('2d');
            const isCircular = type === 'doughnut' || type === 'pie';
            const totalSum = data.reduce((a, b) => a + b, 0);

            const legendConfig = isCircular ? {
                display: true,
                position: 'right',
                labels: {
                    boxWidth: 12,
                    font: { size: 11 },
                    generateLabels: function(chart) {
                        const original = Chart.overrides[type].plugins.legend.labels.generateLabels;
                        const defaultLabels = original(chart);
                        
                        return defaultLabels.map((item, index) => {
                            const val = data[index] || 0;
                            const pct = totalSum > 0 ? ((val / totalSum) * 100).toFixed(1) : 0;
                            const formattedVal = '$' + val.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 0});
                            item.text = labels[index] + ' - ' + formattedVal + ' (' + pct + '%)';
                            return item;
                        });
                    }
                }
            } : { display: false };

            chartInstances[canvasId] = new Chart(ctx, {
                type: type,
                data: {
                    labels: labels,
                    datasets: [{
                        label: labelName,
                        data: data,
                        backgroundColor: isCircular ? chartColors : 'rgba(79, 70, 229, 0.85)',
                        hoverBackgroundColor: isCircular ? chartColors : '#4338ca',
                        borderColor: isCircular ? '#ffffff' : (type === 'line' ? '#4f46e5' : '#4338ca'),
                        borderWidth: isCircular ? 2 : 1.5,
                        tension: 0.3,
                        fill: type === 'line' ? { target: 'origin', above: 'rgba(79, 70, 229, 0.1)' } : false,
                        borderRadius: type === 'bar' ? 4 : 0
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: legendConfig,
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    const val = context.raw || 0;
                                    if (isCircular) {
                                        const pct = totalSum > 0 ? ((val / totalSum) * 100).toFixed(1) : 0;
                                        return ' ' + context.label + ': $' + val.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2}) + ' (' + pct + '%)';
                                    }
                                    return ' ' + (context.dataset.label || '') + ': $' + val.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
                                }
                            }
                        }
                    },
                    scales: !isCircular ? {
                        y: {
                            beginAtZero: true,
                            grid: { color: '#f1f5f9' },
                            ticks: {
                                callback: val => isCurrency ? '$' + val.toLocaleString() : val,
                                font: { size: 11 }
                            }
                        },
                        x: {
                            grid: { display: false },
                            ticks: { font: { size: 11 }, maxRotation: 45, minRotation: 0 }
                        }
                    } : {}
                }
            });
        }

        function refreshDashboard() {
            const data = getFilteredDataset();

            document.getElementById('recordCountBadge').innerText = data.length + ' records';
            
            const totalCost = data.reduce((acc, r) => acc + r.cost, 0);
            document.getElementById('kpiTotalCost').innerText = '$' + totalCost.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            document.getElementById('kpiTotalShipments').innerText = data.length;

            const deptAgg = aggregateBy(data, 'dept');
            const sortedDepts = Object.entries(deptAgg).sort((a,b) => b[1] - a[1]);
            document.getElementById('kpiTopDept').innerText = sortedDepts[0] ? sortedDepts[0][0] : '-';

            const reasonAgg = aggregateBy(data, 'reason');
            const sortedReasons = Object.entries(reasonAgg).sort((a,b) => b[1] - a[1]);
            document.getElementById('kpiTopReason').innerText = sortedReasons[0] ? sortedReasons[0][0] : '-';

            const weeklyAgg = {};
            data.forEach(d => {
                if (d.week > 0) {
                    weeklyAgg[d.week] = (weeklyAgg[d.week] || 0) + d.cost;
                }
            });
            const sortedWeeks = Object.keys(weeklyAgg).sort((a,b) => parseInt(a) - parseInt(b));
            createOrUpdateChart('chartWeekly', 'line', sortedWeeks.map(w => 'WK ' + w), sortedWeeks.map(w => weeklyAgg[w]), 'Air Freight Cost');

            createOrUpdateChart('chartDept', 'pie', sortedDepts.map(d => d[0]), sortedDepts.map(d => d[1]), 'Cost by Department');

            const top10Reasons = sortedReasons.slice(0, 10);
            createOrUpdateChart('chartReason', 'bar', top10Reasons.map(r => r[0]), top10Reasons.map(r => r[1]), 'Cost by Reason');

            const plantAgg = aggregateBy(data, 'plant');
            const sortedPlants = Object.entries(plantAgg).sort((a,b) => b[1] - a[1]);
            createOrUpdateChart('chartPlant', 'doughnut', sortedPlants.map(p => p[0]), sortedPlants.map(p => p[1]), 'Plant Cost');
        }

        window.addEventListener('DOMContentLoaded', () => {
            populateFilterDropdowns();
            refreshDashboard();
            updateStorageIndicator();
        });
    </script>
</body>
</html>`;
}
