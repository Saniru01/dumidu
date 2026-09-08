import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react';
import { Chart, registerables } from 'chart.js';
import * as XLSX from 'xlsx';
import { AirFreightRecord, FilterState, StorageMeta } from './types';
import { chartColors, parseCleanRow, aggregateBy, monthsOrder } from './utils/dataParser';
import { generateStandaloneReportHtml } from './utils/exporter';

Chart.register(...registerables);

const STORAGE_KEY = 'AF_DASHBOARD_DATA';
const STORAGE_META_KEY = 'AF_DASHBOARD_META';
const STORAGE_FILTER_KEY = 'AF_DASHBOARD_FILTERS';

export default function App() {
  const [appData, setAppData] = useState<AirFreightRecord[]>([]);
  const [storageMeta, setStorageMeta] = useState<StorageMeta | null>(null);
  const [filters, setFilters] = useState<FilterState>({
    year: 'ALL',
    month: 'ALL',
    weekFrom: 'ALL',
    weekTo: 'ALL',
  });
  const [alert, setAlert] = useState<{ show: boolean; msg: string; fileName?: string }>({
    show: false,
    msg: '',
  });
  const [copiedChart, setCopiedChart] = useState<Record<string, boolean>>({});

  const fileInputRef = useRef<HTMLInputElement>(null);
  const chartWeeklyRef = useRef<HTMLCanvasElement>(null);
  const chartDeptRef = useRef<HTMLCanvasElement>(null);
  const chartReasonRef = useRef<HTMLCanvasElement>(null);
  const chartPlantRef = useRef<HTMLCanvasElement>(null);

  const chartInstances = useRef<{ [key: string]: Chart | null }>({
    chartWeekly: null,
    chartDept: null,
    chartReason: null,
    chartPlant: null,
  });

  // Load saved state on mount
  useEffect(() => {
    try {
      const savedDataRaw = localStorage.getItem(STORAGE_KEY);
      if (savedDataRaw) {
        const parsed = JSON.parse(savedDataRaw);
        if (Array.isArray(parsed) && parsed.length > 0) {
          setAppData(parsed);
        }
      }
      const metaRaw = localStorage.getItem(STORAGE_META_KEY);
      if (metaRaw) {
        setStorageMeta(JSON.parse(metaRaw));
      }
      const filterRaw = localStorage.getItem(STORAGE_FILTER_KEY);
      if (filterRaw) {
        const parsedFilters = JSON.parse(filterRaw);
        setFilters(prev => ({ ...prev, ...parsedFilters }));
      }
    } catch (err) {
      console.error('Failed to load localStorage data:', err);
    }
  }, []);

  // Filter options derived from dataset
  const years = useMemo(() => {
    const yearSet = new Set<string>();
    appData.forEach(d => {
      if (d.year) yearSet.add(d.year);
    });
    return Array.from(yearSet).sort();
  }, [appData]);

  const months = useMemo(() => {
    const monthSet = new Set<string>();
    appData.forEach(d => {
      if (d.month && d.month !== 'Unspecified') {
        monthSet.add(d.month.toLowerCase());
      }
    });
    return monthsOrder.filter(m => monthSet.has(m.toLowerCase()));
  }, [appData]);

  const weeks = useMemo(() => {
    const weekSet = new Set<number>();
    appData.forEach(d => {
      const w = Number(d.week);
      if (w > 0) weekSet.add(w);
    });
    return Array.from(weekSet).sort((a, b) => a - b);
  }, [appData]);

  // Handle filter changes and sync to localStorage
  const handleFilterChange = (key: keyof FilterState, val: string) => {
    const newFilters = { ...filters, [key]: val };
    setFilters(newFilters);
    localStorage.setItem(STORAGE_FILTER_KEY, JSON.stringify(newFilters));
  };

  const resetFilters = () => {
    const defaultFilters: FilterState = {
      year: 'ALL',
      month: 'ALL',
      weekFrom: 'ALL',
      weekTo: 'ALL',
    };
    setFilters(defaultFilters);
    localStorage.setItem(STORAGE_FILTER_KEY, JSON.stringify(defaultFilters));
  };

  const clearSavedData = () => {
    if (window.confirm('Are you sure you want to clear the saved Excel data and active filters from your browser?')) {
      localStorage.removeItem(STORAGE_KEY);
      localStorage.removeItem(STORAGE_META_KEY);
      localStorage.removeItem(STORAGE_FILTER_KEY);
      setAppData([]);
      setStorageMeta(null);
      resetFilters();
    }
  };

  // File Upload logic
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const rawJson = XLSX.utils.sheet_to_json<Record<string, any>>(workbook.Sheets[firstSheetName], { defval: '' });

        if (rawJson.length === 0) {
          window.alert('The uploaded Excel sheet contains no data.');
          return;
        }

        const cleaned = rawJson.map(parseCleanRow).filter(item => item.cost > 0 || item.week > 0);

        setAppData(cleaned);
        const meta: StorageMeta = {
          fileName: file.name,
          uploadDate: new Date().toLocaleString(),
        };
        setStorageMeta(meta);

        localStorage.setItem(STORAGE_KEY, JSON.stringify(cleaned));
        localStorage.setItem(STORAGE_META_KEY, JSON.stringify(meta));

        resetFilters();

        setAlert({
          show: true,
          msg: `Successfully loaded and auto-saved ${file.name} (${cleaned.length} records).`,
          fileName: file.name,
        });

        setTimeout(() => {
          setAlert(prev => ({ ...prev, show: false }));
        }, 5000);
      } catch (err) {
        console.error('Error processing Excel file:', err);
        window.alert('Error processing Excel file. Please verify column structure.');
      }
    };

    reader.readAsArrayBuffer(file);
    if (e.target) {
      e.target.value = '';
    }
  };

  // Export standalone filled HTML
  const exportFilledHTML = () => {
    if (!appData || appData.length === 0) {
      window.alert('Please upload an Excel file first before exporting.');
      return;
    }

    const htmlContent = generateStandaloneReportHtml(appData);
    const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
    const downloadLink = document.createElement('a');
    const dateStamp = new Date().toISOString().split('T')[0];
    downloadLink.download = `Air_Freight_Report_${dateStamp}.html`;
    downloadLink.href = URL.createObjectURL(blob);

    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(downloadLink.href);
  };

  // Copy Chart as PNG
  const copyChartImage = async (canvasId: string, canvasRef: React.RefObject<HTMLCanvasElement | null>) => {
    const chartCanvas = canvasRef.current;
    if (!chartCanvas) return;

    const exportCanvas = document.createElement('canvas');
    exportCanvas.width = chartCanvas.width;
    exportCanvas.height = chartCanvas.height;
    const ctx = exportCanvas.getContext('2d');
    if (!ctx) return;

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, exportCanvas.width, exportCanvas.height);
    ctx.drawImage(chartCanvas, 0, 0);

    exportCanvas.toBlob(async (blob) => {
      if (!blob) return;
      try {
        await navigator.clipboard.write([
          new ClipboardItem({ 'image/png': blob }),
        ]);

        setCopiedChart(prev => ({ ...prev, [canvasId]: true }));
        setTimeout(() => {
          setCopiedChart(prev => ({ ...prev, [canvasId]: false }));
        }, 2000);
      } catch (err) {
        console.error('Clipboard copy failed:', err);
        const link = document.createElement('a');
        link.download = `${canvasId}-export.png`;
        link.href = exportCanvas.toDataURL('image/png');
        link.click();
      }
    }, 'image/png');
  };

  // Filtered dataset
  const filteredData = useMemo(() => {
    return appData.filter(d => {
      if (filters.year !== 'ALL' && d.year !== filters.year) return false;
      if (filters.month !== 'ALL' && d.month.toLowerCase() !== filters.month.toLowerCase()) return false;
      if (filters.weekFrom !== 'ALL' && d.week < parseInt(filters.weekFrom, 10)) return false;
      if (filters.weekTo !== 'ALL' && d.week > parseInt(filters.weekTo, 10)) return false;
      return true;
    });
  }, [appData, filters]);

  // KPIs
  const totalCost = useMemo(() => {
    return filteredData.reduce((acc, r) => acc + r.cost, 0);
  }, [filteredData]);

  const sortedDepts = useMemo(() => {
    const deptAgg = aggregateBy(filteredData, 'dept');
    return Object.entries(deptAgg).sort((a, b) => b[1] - a[1]);
  }, [filteredData]);

  const sortedReasons = useMemo(() => {
    const reasonAgg = aggregateBy(filteredData, 'reason');
    return Object.entries(reasonAgg).sort((a, b) => b[1] - a[1]);
  }, [filteredData]);

  const sortedPlants = useMemo(() => {
    const plantAgg = aggregateBy(filteredData, 'plant');
    return Object.entries(plantAgg).sort((a, b) => b[1] - a[1]);
  }, [filteredData]);

  const topDept = sortedDepts[0] ? sortedDepts[0][0] : '-';
  const topReason = sortedReasons[0] ? sortedReasons[0][0] : '-';

  // Render or Update Chart instances
  const updateCharts = useCallback(() => {
    // 1. Chart 1: Weekly Trend (Line)
    if (chartWeeklyRef.current) {
      if (chartInstances.current.chartWeekly) {
        chartInstances.current.chartWeekly.destroy();
      }

      const weeklyAgg: Record<number, number> = {};
      filteredData.forEach(d => {
        if (d.week > 0) {
          weeklyAgg[d.week] = (weeklyAgg[d.week] || 0) + d.cost;
        }
      });
      const sortedWeeks = Object.keys(weeklyAgg).map(Number).sort((a, b) => a - b);
      const labels = sortedWeeks.map(w => 'WK ' + w);
      const data = sortedWeeks.map(w => weeklyAgg[w]);

      const ctx = chartWeeklyRef.current.getContext('2d');
      if (ctx) {
        chartInstances.current.chartWeekly = new Chart(ctx, {
          type: 'line',
          data: {
            labels,
            datasets: [{
              label: 'Air Freight Cost',
              data,
              backgroundColor: 'rgba(79, 70, 229, 0.85)',
              hoverBackgroundColor: '#4338ca',
              borderColor: '#4f46e5',
              borderWidth: 1.5,
              tension: 0.3,
              fill: { target: 'origin', above: 'rgba(79, 70, 229, 0.1)' },
            }],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { display: false },
              tooltip: {
                callbacks: {
                  label: (context) => {
                    const val = (context.raw as number) || 0;
                    return ` ${context.dataset.label || ''}: $${val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
                  },
                },
              },
            },
            scales: {
              y: {
                beginAtZero: true,
                grid: { color: '#f1f5f9' },
                ticks: {
                  callback: (val) => '$' + Number(val).toLocaleString(),
                  font: { size: 11 },
                },
              },
              x: {
                grid: { display: false },
                ticks: { font: { size: 11 }, maxRotation: 45, minRotation: 0 },
              },
            },
          },
        });
      }
    }

    // 2. Chart 2: Responsible Department Wise (Pie)
    if (chartDeptRef.current) {
      if (chartInstances.current.chartDept) {
        chartInstances.current.chartDept.destroy();
      }

      const labels = sortedDepts.map(d => d[0]);
      const data = sortedDepts.map(d => d[1]);
      const totalSum = data.reduce((a, b) => a + b, 0);

      const ctx = chartDeptRef.current.getContext('2d');
      if (ctx) {
        chartInstances.current.chartDept = new Chart(ctx, {
          type: 'pie',
          data: {
            labels,
            datasets: [{
              label: 'Cost by Department',
              data,
              backgroundColor: chartColors,
              hoverBackgroundColor: chartColors,
              borderColor: '#ffffff',
              borderWidth: 2,
            }],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: {
                display: true,
                position: 'right',
                labels: {
                  boxWidth: 12,
                  font: { size: 11 },
                  generateLabels: (chart) => {
                    const original = Chart.overrides.pie.plugins.legend.labels.generateLabels;
                    const defaultLabels = original(chart);
                    return defaultLabels.map((item, index) => {
                      const val = data[index] || 0;
                      const pct = totalSum > 0 ? ((val / totalSum) * 100).toFixed(1) : '0';
                      const formattedVal = '$' + val.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 });
                      item.text = `${labels[index]} - ${formattedVal} (${pct}%)`;
                      return item;
                    });
                  },
                },
              },
              tooltip: {
                callbacks: {
                  label: (context) => {
                    const val = (context.raw as number) || 0;
                    const pct = totalSum > 0 ? ((val / totalSum) * 100).toFixed(1) : '0';
                    return ` ${context.label}: $${val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} (${pct}%)`;
                  },
                },
              },
            },
          },
        });
      }
    }

    // 3. Chart 3: Reason Wise Breakdown (Bar - Top 10)
    if (chartReasonRef.current) {
      if (chartInstances.current.chartReason) {
        chartInstances.current.chartReason.destroy();
      }

      const top10Reasons = sortedReasons.slice(0, 10);
      const labels = top10Reasons.map(r => r[0]);
      const data = top10Reasons.map(r => r[1]);

      const ctx = chartReasonRef.current.getContext('2d');
      if (ctx) {
        chartInstances.current.chartReason = new Chart(ctx, {
          type: 'bar',
          data: {
            labels,
            datasets: [{
              label: 'Cost by Reason',
              data,
              backgroundColor: 'rgba(79, 70, 229, 0.85)',
              hoverBackgroundColor: '#4338ca',
              borderColor: '#4338ca',
              borderWidth: 1.5,
              borderRadius: 4,
            }],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: { display: false },
              tooltip: {
                callbacks: {
                  label: (context) => {
                    const val = (context.raw as number) || 0;
                    return ` ${context.dataset.label || ''}: $${val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
                  },
                },
              },
            },
            scales: {
              y: {
                beginAtZero: true,
                grid: { color: '#f1f5f9' },
                ticks: {
                  callback: (val) => '$' + Number(val).toLocaleString(),
                  font: { size: 11 },
                },
              },
              x: {
                grid: { display: false },
                ticks: { font: { size: 11 }, maxRotation: 45, minRotation: 0 },
              },
            },
          },
        });
      }
    }

    // 4. Chart 4: Plant Wise Breakdown (Doughnut)
    if (chartPlantRef.current) {
      if (chartInstances.current.chartPlant) {
        chartInstances.current.chartPlant.destroy();
      }

      const labels = sortedPlants.map(p => p[0]);
      const data = sortedPlants.map(p => p[1]);
      const totalSum = data.reduce((a, b) => a + b, 0);

      const ctx = chartPlantRef.current.getContext('2d');
      if (ctx) {
        chartInstances.current.chartPlant = new Chart(ctx, {
          type: 'doughnut',
          data: {
            labels,
            datasets: [{
              label: 'Plant Cost',
              data,
              backgroundColor: chartColors,
              hoverBackgroundColor: chartColors,
              borderColor: '#ffffff',
              borderWidth: 2,
            }],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
              legend: {
                display: true,
                position: 'right',
                labels: {
                  boxWidth: 12,
                  font: { size: 11 },
                  generateLabels: (chart) => {
                    const original = Chart.overrides.doughnut.plugins.legend.labels.generateLabels;
                    const defaultLabels = original(chart);
                    return defaultLabels.map((item, index) => {
                      const val = data[index] || 0;
                      const pct = totalSum > 0 ? ((val / totalSum) * 100).toFixed(1) : '0';
                      const formattedVal = '$' + val.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 });
                      item.text = `${labels[index]} - ${formattedVal} (${pct}%)`;
                      return item;
                    });
                  },
                },
              },
              tooltip: {
                callbacks: {
                  label: (context) => {
                    const val = (context.raw as number) || 0;
                    const pct = totalSum > 0 ? ((val / totalSum) * 100).toFixed(1) : '0';
                    return ` ${context.label}: $${val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })} (${pct}%)`;
                  },
                },
              },
            },
          },
        });
      }
    }
  }, [filteredData, sortedDepts, sortedReasons, sortedPlants]);

  // Re-draw charts whenever filteredData updates
  useEffect(() => {
    updateCharts();
    return () => {
      Object.keys(chartInstances.current).forEach((key) => {
        if (chartInstances.current[key]) {
          chartInstances.current[key]?.destroy();
          chartInstances.current[key] = null;
        }
      });
    };
  }, [updateCharts]);

  return (
    <div className="bg-slate-100 text-slate-800 font-sans min-h-screen flex flex-col">
      {/* Header Navigation */}
      <header className="bg-slate-900 text-white shadow-md sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-4 flex flex-col md:flex-row justify-between items-center gap-4">
          <div className="flex items-center space-x-3">
            <div className="bg-indigo-600 p-2.5 rounded-xl text-white shadow-inner">
              <i className="fa-solid fa-plane-circle-exclamation text-xl"></i>
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-wide">Air Freight Cost of Failures</h1>
              <p className="text-xs text-slate-400">Weekly Operations &amp; Delay Analytics</p>
            </div>
          </div>

          {/* Upload, Export & Storage Action Controls */}
          <div className="flex items-center flex-wrap gap-2">
            <label
              id="uploadBtnLabel"
              className="cursor-pointer flex items-center gap-2 bg-indigo-600 hover:bg-indigo-500 text-white text-xs sm:text-sm font-semibold px-4 py-2 rounded-lg shadow transition duration-150"
            >
              <i className="fa-solid fa-cloud-arrow-up"></i>
              <span id="uploadBtnText">Upload Excel</span>
              <input
                ref={fileInputRef}
                type="file"
                id="fileInput"
                accept=".xlsx, .xls, .csv"
                className="hidden"
                onChange={handleFileUpload}
              />
            </label>
            <button
              id="btnExportEmail"
              onClick={exportFilledHTML}
              className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs sm:text-sm font-semibold px-4 py-2 rounded-lg shadow transition duration-150"
              title="Export HTML file with data embedded for email sharing"
            >
              <i className="fa-solid fa-file-export"></i>
              <span>Export for Email</span>
            </button>
            <button
              id="btnResetFilters"
              onClick={resetFilters}
              className="text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 px-3 py-2 rounded-lg border border-slate-700 transition"
              title="Reset Filters"
            >
              <i className="fa-solid fa-rotate-right"></i> Reset Filters
            </button>
            <button
              id="btnClearStorage"
              onClick={clearSavedData}
              className="text-xs bg-rose-950 hover:bg-rose-900 text-rose-300 px-3 py-2 rounded-lg border border-rose-800 transition"
              title="Clear Auto-Saved Data"
            >
              <i className="fa-solid fa-trash-can"></i> Clear Storage
            </button>
          </div>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 py-6 space-y-6 flex-1 w-full">
        {/* Notification / Status Banner */}
        {alert.show && (
          <div
            id="uploadAlert"
            className="p-3 bg-emerald-100 border border-emerald-300 text-emerald-800 text-xs font-semibold rounded-xl flex items-center justify-between transition-all"
          >
            <span id="uploadAlertMsg">
              <i className="fa-solid fa-circle-check mr-2"></i> {alert.msg}
            </span>
            <button
              id="btnCloseAlert"
              onClick={() => setAlert(prev => ({ ...prev, show: false }))}
              className="text-emerald-700 hover:text-emerald-900 text-base"
            >
              &times;
            </button>
          </div>
        )}

        {/* Filter Bar */}
        <section className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200">
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between pb-3 border-b border-slate-100 mb-4 gap-2">
            <span className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-2">
              <i className="fa-solid fa-sliders text-indigo-600"></i> Interactive Filters
            </span>
            <div className="flex items-center gap-2">
              <span
                id="storageStatusBadge"
                className={`text-xs font-medium px-2.5 py-0.5 rounded-full border ${
                  appData.length > 0
                    ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
                    : 'bg-amber-50 text-amber-700 border-amber-200'
                }`}
              >
                {appData.length > 0 ? (
                  <>
                    <i className="fa-solid fa-hard-drive mr-1"></i> Data Ready ({appData.length} records)
                  </>
                ) : (
                  <>
                    <i className="fa-solid fa-triangle-exclamation mr-1"></i> No Data Loaded
                  </>
                )}
              </span>
              <span
                id="recordCountBadge"
                className="text-xs font-medium bg-indigo-50 text-indigo-700 px-2.5 py-0.5 rounded-full border border-indigo-100"
              >
                {filteredData.length} records
              </span>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            {/* Year */}
            <div>
              <label htmlFor="selYear" className="block text-xs font-semibold text-slate-600 mb-1">
                Year
              </label>
              <select
                id="selYear"
                value={filters.year}
                onChange={(e) => handleFilterChange('year', e.target.value)}
                className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
              >
                <option value="ALL">All Years</option>
                {years.map(y => (
                  <option key={y} value={y}>{y}</option>
                ))}
              </select>
            </div>

            {/* Month */}
            <div>
              <label htmlFor="selMonth" className="block text-xs font-semibold text-slate-600 mb-1">
                Month
              </label>
              <select
                id="selMonth"
                value={filters.month}
                onChange={(e) => handleFilterChange('month', e.target.value)}
                className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
              >
                <option value="ALL">All Months</option>
                {months.map(m => (
                  <option key={m} value={m}>{m}</option>
                ))}
              </select>
            </div>

            {/* Start Week */}
            <div>
              <label htmlFor="selWeekFrom" className="block text-xs font-semibold text-slate-600 mb-1">
                From Week
              </label>
              <select
                id="selWeekFrom"
                value={filters.weekFrom}
                onChange={(e) => handleFilterChange('weekFrom', e.target.value)}
                className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
              >
                <option value="ALL">Start Week (Min)</option>
                {weeks.map(w => (
                  <option key={w} value={String(w)}>{w}</option>
                ))}
              </select>
            </div>

            {/* End Week */}
            <div>
              <label htmlFor="selWeekTo" className="block text-xs font-semibold text-slate-600 mb-1">
                To Week
              </label>
              <select
                id="selWeekTo"
                value={filters.weekTo}
                onChange={(e) => handleFilterChange('weekTo', e.target.value)}
                className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
              >
                <option value="ALL">End Week (Max)</option>
                {weeks.map(w => (
                  <option key={w} value={String(w)}>{w}</option>
                ))}
              </select>
            </div>
          </div>
        </section>

        {/* KPI Summary Cards */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <div id="cardKpiTotalCost" className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
            <div>
              <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Total AF Cost</p>
              <h3 id="kpiTotalCost" className="text-2xl font-black text-rose-600 mt-1">
                ${totalCost.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </h3>
            </div>
            <div className="bg-rose-50 text-rose-600 w-12 h-12 rounded-xl flex items-center justify-center text-xl">
              <i className="fa-solid fa-dollar-sign"></i>
            </div>
          </div>

          <div id="cardKpiTotalShipments" className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
            <div>
              <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Total Incidents</p>
              <h3 id="kpiTotalShipments" className="text-2xl font-black text-slate-800 mt-1">
                {filteredData.length}
              </h3>
            </div>
            <div className="bg-indigo-50 text-indigo-600 w-12 h-12 rounded-xl flex items-center justify-center text-xl">
              <i className="fa-solid fa-boxes-packing"></i>
            </div>
          </div>

          <div id="cardKpiTopDept" className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
            <div>
              <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Top Cost Dept</p>
              <h3 id="kpiTopDept" className="text-base font-bold text-slate-800 mt-1 truncate max-w-[140px]">
                {topDept}
              </h3>
            </div>
            <div className="bg-emerald-50 text-emerald-600 w-12 h-12 rounded-xl flex items-center justify-center text-xl">
              <i className="fa-solid fa-building-user"></i>
            </div>
          </div>

          <div id="cardKpiTopReason" className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
            <div>
              <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">Top Delay Reason</p>
              <h3 id="kpiTopReason" className="text-base font-bold text-slate-800 mt-1 truncate max-w-[140px]">
                {topReason}
              </h3>
            </div>
            <div className="bg-amber-50 text-amber-600 w-12 h-12 rounded-xl flex items-center justify-center text-xl">
              <i className="fa-solid fa-triangle-exclamation"></i>
            </div>
          </div>
        </section>

        {/* Charts Dashboard Grid */}
        <section className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Chart 1: Weekly Air Freight Cost Trend */}
          <div id="containerChartWeekly" className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i className="fa-solid fa-chart-line text-indigo-600"></i> 1. Weekly Air Freight Cost Trend ($ USD)
              </h2>
              <button
                id="btnCopyChartWeekly"
                onClick={() => copyChartImage('chartWeekly', chartWeeklyRef)}
                className={`text-xs font-semibold px-2.5 py-1 rounded-md border transition flex items-center gap-1.5 ${
                  copiedChart.chartWeekly
                    ? 'bg-emerald-50 border-emerald-300 text-emerald-700'
                    : 'bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300'
                }`}
                title="Copy chart to clipboard"
              >
                {copiedChart.chartWeekly ? (
                  <>
                    <i className="fa-solid fa-check text-emerald-600"></i>
                    <span>Copied!</span>
                  </>
                ) : (
                  <>
                    <i className="fa-regular fa-copy"></i>
                    <span>Copy</span>
                  </>
                )}
              </button>
            </div>
            <div className="relative h-96 w-full">
              <canvas id="chartWeekly" ref={chartWeeklyRef}></canvas>
            </div>
          </div>

          {/* Chart 2: Responsible Department Wise (PIE CHART) */}
          <div id="containerChartDept" className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i className="fa-solid fa-chart-pie text-indigo-600"></i> 2. Responsible Department Cost Distribution
              </h2>
              <button
                id="btnCopyChartDept"
                onClick={() => copyChartImage('chartDept', chartDeptRef)}
                className={`text-xs font-semibold px-2.5 py-1 rounded-md border transition flex items-center gap-1.5 ${
                  copiedChart.chartDept
                    ? 'bg-emerald-50 border-emerald-300 text-emerald-700'
                    : 'bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300'
                }`}
                title="Copy chart to clipboard"
              >
                {copiedChart.chartDept ? (
                  <>
                    <i className="fa-solid fa-check text-emerald-600"></i>
                    <span>Copied!</span>
                  </>
                ) : (
                  <>
                    <i className="fa-regular fa-copy"></i>
                    <span>Copy</span>
                  </>
                )}
              </button>
            </div>
            <div className="relative h-96 w-full flex items-center justify-center">
              <canvas id="chartDept" ref={chartDeptRef}></canvas>
            </div>
          </div>

          {/* Chart 3: Reason Wise Breakdown */}
          <div id="containerChartReason" className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i className="fa-solid fa-chart-bar text-indigo-600"></i> 3. Failure Reason Wise Cost (Top 10)
              </h2>
              <button
                id="btnCopyChartReason"
                onClick={() => copyChartImage('chartReason', chartReasonRef)}
                className={`text-xs font-semibold px-2.5 py-1 rounded-md border transition flex items-center gap-1.5 ${
                  copiedChart.chartReason
                    ? 'bg-emerald-50 border-emerald-300 text-emerald-700'
                    : 'bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300'
                }`}
                title="Copy chart to clipboard"
              >
                {copiedChart.chartReason ? (
                  <>
                    <i className="fa-solid fa-check text-emerald-600"></i>
                    <span>Copied!</span>
                  </>
                ) : (
                  <>
                    <i className="fa-regular fa-copy"></i>
                    <span>Copy</span>
                  </>
                )}
              </button>
            </div>
            <div className="relative h-96 w-full">
              <canvas id="chartReason" ref={chartReasonRef}></canvas>
            </div>
          </div>

          {/* Chart 4: Plant Wise Breakdown */}
          <div id="containerChartPlant" className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <i className="fa-solid fa-chart-pie text-indigo-600"></i> 4. Plant Wise Cost Distribution
              </h2>
              <button
                id="btnCopyChartPlant"
                onClick={() => copyChartImage('chartPlant', chartPlantRef)}
                className={`text-xs font-semibold px-2.5 py-1 rounded-md border transition flex items-center gap-1.5 ${
                  copiedChart.chartPlant
                    ? 'bg-emerald-50 border-emerald-300 text-emerald-700'
                    : 'bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300'
                }`}
                title="Copy chart to clipboard"
              >
                {copiedChart.chartPlant ? (
                  <>
                    <i className="fa-solid fa-check text-emerald-600"></i>
                    <span>Copied!</span>
                  </>
                ) : (
                  <>
                    <i className="fa-regular fa-copy"></i>
                    <span>Copy</span>
                  </>
                )}
              </button>
            </div>
            <div className="relative h-96 w-full flex items-center justify-center">
              <canvas id="chartPlant" ref={chartPlantRef}></canvas>
            </div>
          </div>
        </section>
      </div>
    </div>
  );
}
