import React, { useState, useEffect, useRef } from "react";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import {
  Upload,
  Check,
  AlertTriangle,
  Trash2,
  Download,
  RefreshCw,
  FileSpreadsheet,
  Plus,
  X,
  Sparkles,
  ArrowRight,
  Info,
  Layers,
  ChevronRight,
  Database
} from "lucide-react";

// Normalization function to compare string headers
function normalizeHeader(s: string): string {
  return s.trim().toLowerCase().replace(/[\s-_()/]+/g, "");
}

// Client-side Levenshtein distance for fuzzy matching fallback
function getLevenshteinSimilarity(s1: string, s2: string): number {
  const len1 = s1.length;
  const len2 = s2.length;
  const matrix: number[][] = [];

  for (let i = 0; i <= len1; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= len2; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= len1; i++) {
    for (let j = 1; j <= len2; j++) {
      if (s1.charAt(i - 1) === s2.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }

  const distance = matrix[len1][len2];
  const maxLength = Math.max(len1, len2);
  return maxLength === 0 ? 1.0 : 1.0 - distance / maxLength;
}

export default function App() {
  // --- States ---
  // Golden Headers (Row A master template) - empty by default (no preset template)
  const [goldenHeaders, setGoldenHeaders] = useState<string[]>([]);

  const [newHeader, setNewHeader] = useState("");
  const [pastedText, setPastedText] = useState("");

  // Uploaded Excel File Info
  const [fileName, setFileName] = useState<string | null>(null);
  const [fileHeaders, setFileHeaders] = useState<string[]>([]);
  const [fileRecords, setFileRecords] = useState<any[][]>([]); // Column B, C, D...

  // Mapping state: [uploadedHeader] -> goldenHeader
  const [columnMappings, setColumnMappings] = useState<{ [key: string]: string | null }>({});
  const [mappingSource, setMappingSource] = useState<{ [key: string]: "exact" | "ai" | "unmatched" | "manual" }>({});

  const [isLoadingMapping, setIsLoadingMapping] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);

  const [dragActive, setDragActive] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // --- Handlers ---

  // Load Preset
  const handleLoadPreset = () => {
    setGoldenHeaders([
      "Finish",
      "Foot",
      "Front bezel",
      "Rear cover",
      "Regulatory",
      "Approvals",
      "Warranty",
      "Signal Input",
      "Sync Input",
      "Audio (In/Out)",
      "HDCP",
      "Built-in Speakers"
    ]);
    showSuccess("已載入預設 Golden Headers 欄位範本。");
  };

  // Add individual header manually
  const handleAddHeader = (e: React.FormEvent) => {
    e.preventDefault();
    const trimmed = newHeader.trim();
    if (!trimmed) return;
    if (goldenHeaders.includes(trimmed)) {
      setErrorMsg("該欄位名稱已存在於 Golden Headers 中。");
      return;
    }
    setGoldenHeaders([...goldenHeaders, trimmed]);
    setNewHeader("");
    setErrorMsg(null);
  };

  // Delete golden header
  const handleDeleteGoldenHeader = (index: number) => {
    setGoldenHeaders(goldenHeaders.filter((_, i) => i !== index));
  };

  // Handle Pasted Excel Row/Column
  const handleParsePasted = () => {
    if (!pastedText.trim()) return;

    // Split by tab, newline, or commas to handle both rows and columns
    const parsed = pastedText
      .split(/[\r\n\t,]+/)
      .map(s => s.trim())
      .filter(s => s.length > 0);

    if (parsed.length === 0) {
      setErrorMsg("貼上的內容無法解析成有效欄位。");
      return;
    }

    // Filter duplicates
    const uniqueParsed = parsed.filter(item => !goldenHeaders.includes(item));
    setGoldenHeaders([...goldenHeaders, ...uniqueParsed]);
    setPastedText("");
    setErrorMsg(null);
    showSuccess(`已成功解析並新增 ${uniqueParsed.length} 個欄位至 Golden Headers。`);
  };

  // Drag & Drop event handlers
  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFileSelection(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      handleFileSelection(e.target.files[0]);
    }
  };

  // Parse Excel via SheetJS
  const handleFileSelection = (file: File) => {
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (ext !== "xlsx" && ext !== "xls" && ext !== "csv") {
      setErrorMsg("不支援此檔案格式。請上傳 .xlsx, .xls 或 .csv 檔案。");
      return;
    }

    setErrorMsg(null);
    setSuccessMsg(null);

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Parse to raw array of arrays
        const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: "" });

        if (rows.length === 0) {
          setErrorMsg("上傳的 Excel 檔案為空，無任何資料。");
          return;
        }

        // Extract Column A (vertical specification headers) and dynamic column records (Product B, C, D...)
        let maxCols = 0;
        rows.forEach(r => {
          if (r.length > maxCols) maxCols = r.length;
        });
        const dataColCount = Math.max(0, maxCols - 1); // Column B, C, D...

        const headers: string[] = [];
        const records: any[][] = Array.from({ length: dataColCount }, () => []);

        rows.forEach((row) => {
          const header = String(row[0] || "").trim();
          if (header) {
            headers.push(header);
            for (let c = 0; c < dataColCount; c++) {
              records[c].push(row[c + 1] !== undefined ? row[c + 1] : "");
            }
          }
        });

        if (headers.length === 0) {
          setErrorMsg("無法在第一欄 (Column A) 找到任何有效的規格/欄位名稱。");
          return;
        }

        setFileHeaders(headers);
        setFileRecords(records);
        setFileName(file.name);

        // Auto trigger alignment
        await runAlignmentMapping(headers, goldenHeaders);

      } catch (err) {
        console.error(err);
        setErrorMsg("解析 Excel 檔案失敗，請確保檔案格式正確且無損毀。");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Run Column Alignment Matching via Server-side API or Client similarity fallback
  const runAlignmentMapping = async (uploaded: string[], golden: string[]) => {
    setIsLoadingMapping(true);
    setErrorMsg(null);

    try {
      const response = await fetch("/api/match-columns", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ goldenColumns: golden, uploadedColumns: uploaded }),
      });

      if (!response.ok) {
        throw new Error("API mapping failed, fallback to client-side");
      }

      const data = await response.json();
      const mappingsResult = data.mappings as Array<{
        target: string;
        golden: string | null;
        matchType: "exact" | "ai" | "unmatched";
      }>;

      const newMappings: { [key: string]: string | null } = {};
      const newSources: { [key: string]: "exact" | "ai" | "unmatched" | "manual" } = {};

      for (const item of mappingsResult) {
        newMappings[item.target] = item.golden;
        newSources[item.target] = item.matchType;
      }

      setColumnMappings(newMappings);
      setMappingSource(newSources);
      showSuccess(`欄位自動分析完成！已為 ${uploaded.length} 個欄位找到最佳對應。`);

    } catch (apiErr) {
      console.warn("Using client-side fuzzy similarity fallback match due to server error or offline mode.");
      
      const newMappings: { [key: string]: string | null } = {};
      const newSources: { [key: string]: "exact" | "ai" | "unmatched" | "manual" } = {};

      for (const target of uploaded) {
        const normTarget = normalizeHeader(target);

        // 1. Direct exact match
        const exactMatch = golden.find(g => normalizeHeader(g) === normTarget);
        if (exactMatch) {
          newMappings[target] = exactMatch;
          newSources[target] = "exact";
          continue;
        }

        // 2. Fuzzy match
        let bestMatch: string | null = null;
        let highestSim = 0;

        for (const g of golden) {
          const sim = getLevenshteinSimilarity(normTarget, normalizeHeader(g));
          if (sim > highestSim) {
            highestSim = sim;
            bestMatch = g;
          }
        }

        // Threshold of 0.4 similarity
        if (highestSim >= 0.4 && bestMatch) {
          newMappings[target] = bestMatch;
          newSources[target] = "ai";
        } else {
          newMappings[target] = null;
          newSources[target] = "unmatched";
        }
      }

      setColumnMappings(newMappings);
      setMappingSource(newSources);
      showSuccess("已使用本地模糊匹配算法完成欄位對應！");
    } finally {
      setIsLoadingMapping(false);
    }
  };

  // Re-run matching if golden headers update
  const handleReMatch = () => {
    if (fileHeaders.length > 0) {
      runAlignmentMapping(fileHeaders, goldenHeaders);
    }
  };

  // Modify manual mapping mapping dropdown
  const handleManualMappingChange = (uploadedHeader: string, value: string) => {
    const selectedGolden = value === "skip" ? null : value;
    setColumnMappings(prev => ({
      ...prev,
      [uploadedHeader]: selectedGolden
    }));
    setMappingSource(prev => ({
      ...prev,
      [uploadedHeader]: selectedGolden ? "manual" : "unmatched"
    }));
  };

  // Utility to trigger success toast
  const showSuccess = (msg: string) => {
    setSuccessMsg(msg);
    setTimeout(() => {
      setSuccessMsg(prev => (prev === msg ? null : prev));
    }, 4500);
  };

  // Clear golden template
  const handleClearGolden = () => {
    setGoldenHeaders([]);
    showSuccess("已清空 Golden Headers。");
  };

  // Export & Download XLSX aligned to Golden Headers
  const handleExportAndDownload = async () => {
    if (goldenHeaders.length === 0) {
      setErrorMsg("輸出失敗：您尚未設定任何規格項目。");
      return;
    }

    if (fileHeaders.length === 0 || fileRecords.length === 0) {
      setErrorMsg("輸出失敗：請先上傳使用者 Excel 檔案。");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Realigned Data");

      // Set columns
      const columnsDef = [
        { header: "Golden Column A (標準規格列表)", key: "golden", width: 32 },
        { header: "對應來源 / 狀態", key: "status", width: 38 }
      ];

      for (let c = 0; c < fileRecords.length; c++) {
        columnsDef.push({
          header: `資料 #${c + 1} (Column ${String.fromCharCode(66 + c)})`,
          key: `record_${c}`,
          width: 26
        });
      }
      worksheet.columns = columnsDef;

      // Style header row (Row 1)
      const headerRow = worksheet.getRow(1);
      headerRow.font = { name: "Segoe UI", size: 11, bold: true, color: { argb: "FFFFFF" } };
      headerRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "1F4E79" } // Classic Dark Blue header
      };
      headerRow.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      headerRow.height = 32;

      // Add data rows
      for (let gIdx = 0; gIdx < goldenHeaders.length; gIdx++) {
        const gHeader = goldenHeaders[gIdx];

        // Find which uploaded header maps to this goldenHeader
        const mappedTargetHeader = Object.keys(columnMappings).find(
          target => columnMappings[target] === gHeader
        );

        // Map status / source text
        const source = mappedTargetHeader ? mappingSource[mappedTargetHeader] : "unmatched";
        let statusText = "(未對應 - 將匯出為空值)";
        if (mappedTargetHeader) {
          const typeStr = source === "exact" ? "精確對應" : source === "ai" ? "AI 接近對應" : source === "manual" ? "手動校正" : "已對應";
          statusText = `← ${mappedTargetHeader} (${typeStr})`;
        }

        const rowData: any = {
          golden: gHeader,
          status: statusText
        };

        for (let c = 0; c < fileRecords.length; c++) {
          let cellValue = "";
          if (mappedTargetHeader) {
            const targetIdx = fileHeaders.indexOf(mappedTargetHeader);
            if (targetIdx !== -1) {
              const originalVal = fileRecords[c][targetIdx];
              cellValue = originalVal !== undefined ? originalVal : "";
            }
          }
          rowData[`record_${c}`] = cellValue;
        }

        const addedRow = worksheet.addRow(rowData);
        addedRow.height = 24;

        // Check if it's an AI match (non-exact matching) to make it red!
        const isAi = source === "ai";

        // Style cells in the row
        addedRow.eachCell((cell, colNumber) => {
          // Set cell borders & font family
          cell.border = {
            top: { style: "thin", color: { argb: "E2E8F0" } },
            left: { style: "thin", color: { argb: "E2E8F0" } },
            bottom: { style: "thin", color: { argb: "E2E8F0" } },
            right: { style: "thin", color: { argb: "E2E8F0" } }
          };

          if (colNumber === 1) {
            cell.alignment = { vertical: "middle", horizontal: "left" };
            if (isAi) {
              cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "DC2626" } }; // Darker Red for clean output
              cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FEF2F2" } // Soft red bg
              };
            } else {
              cell.font = { name: "Segoe UI", size: 10, bold: true, color: { argb: "1E293B" } };
            }
          } else if (colNumber === 2) {
            cell.alignment = { vertical: "middle", horizontal: "left" };
            if (isAi) {
              cell.font = { name: "Segoe UI", size: 10, bold: true, italic: true, color: { argb: "DC2626" } };
              cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FEF2F2" } // Soft red bg
              };
            } else if (source === "unmatched") {
              cell.font = { name: "Segoe UI", size: 10, italic: true, color: { argb: "94A3B8" } };
            } else {
              cell.font = { name: "Segoe UI", size: 10, color: { argb: "334155" } };
            }
          } else {
            cell.alignment = { vertical: "middle", horizontal: "left" };
            if (isAi) {
              cell.font = { name: "Segoe UI", size: 10, color: { argb: "DC2626" } };
              cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FEF2F2" } // Soft red bg
              };
            } else {
              cell.font = { name: "Segoe UI", size: 10, color: { argb: "475569" } };
            }
          }
        });
      }

      // Generate buffer and trigger download
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

      const originalBaseName = fileName ? fileName.substring(0, fileName.lastIndexOf(".")) : "UserFile";
      const downloadName = `Realigned_${originalBaseName}_aligned_to_Golden.xlsx`;

      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = downloadName;
      document.body.appendChild(anchor);
      anchor.click();
      document.body.removeChild(anchor);
      window.URL.revokeObjectURL(url);

      showSuccess(`成功匯出整合檔案：${downloadName}`);
    } catch (err) {
      console.error(err);
      setErrorMsg("匯出整合 Excel 失敗。");
    }
  };

  // --- Dynamic Stats for active upload ---
  const exactCount = Object.keys(columnMappings).filter(
    k => columnMappings[k] && mappingSource[k] === "exact"
  ).length;
  const aiCount = Object.keys(columnMappings).filter(
    k => columnMappings[k] && mappingSource[k] === "ai"
  ).length;
  const unmatchedCount = Object.keys(columnMappings).filter(
    k => !columnMappings[k]
  ).length;

  const confidenceScore = fileHeaders.length > 0 
    ? Math.round(((exactCount * 1.0 + aiCount * 0.75) / fileHeaders.length) * 100) 
    : 0;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans flex flex-col pb-16">
      {/* Header Section */}
      <header className="bg-white border-b border-slate-200 px-6 py-4 sticky top-0 z-40 shadow-sm">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center md:justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="bg-blue-600 p-2.5 rounded-xl shadow-md text-white flex items-center justify-center">
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight text-slate-900 flex items-center gap-2">
                MNT SPEC &lt;-&gt; PXM FORMAT <span className="text-blue-600 italic font-semibold">AI</span>
              </h1>
              <p className="text-xs text-slate-500 font-medium">Excel 欄位自動排序與對應整合系統 by terry</p>
            </div>
          </div>

          {/* Stepper Display */}
          <div className="hidden sm:flex items-center gap-4">
            <div className="flex items-center gap-2">
              <span className={`w-7 h-7 rounded-full flex items-center justify-center font-bold text-xs ${
                goldenHeaders.length > 0 ? "bg-blue-100 text-blue-700" : "bg-slate-100 text-slate-400"
              }`}>
                {goldenHeaders.length > 0 ? <Check className="w-3.5 h-3.5" /> : "1"}
              </span>
              <span className={`text-xs font-semibold ${goldenHeaders.length > 0 ? "text-blue-700" : "text-slate-400"}`}>Set Golden</span>
            </div>
            <div className="w-6 h-px bg-slate-300"></div>
            <div className="flex items-center gap-2">
              <span className={`w-7 h-7 rounded-full flex items-center justify-center font-bold text-xs ${
                fileName ? "bg-blue-600 text-white shadow-sm" : "bg-slate-100 text-slate-400"
              }`}>
                {fileName && fileRecords.length > 0 ? <Check className="w-3.5 h-3.5" /> : "2"}
              </span>
              <span className={`text-xs font-semibold ${fileName ? "text-slate-900" : "text-slate-400"}`}>Process & Preview</span>
            </div>
            <div className="w-6 h-px bg-slate-300"></div>
            <div className="flex items-center gap-2">
              <span className={`w-7 h-7 rounded-full border-2 flex items-center justify-center font-bold text-xs ${
                fileName && fileRecords.length > 0 ? "bg-emerald-50 text-emerald-600 border-emerald-200" : "border-slate-200 text-slate-400"
              }`}>
                3
              </span>
              <span className={`text-xs font-semibold ${fileName && fileRecords.length > 0 ? "text-emerald-600" : "text-slate-400"}`}>Export</span>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-6 flex flex-col gap-6">
        
        {/* Global Notifications */}
        <AnimatePresence>
          {errorMsg && (
            <motion.div
              initial={{ opacity: 0, y: -10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="p-4 bg-red-50 border border-red-200 text-red-700 rounded-xl flex items-start gap-3 shadow-sm"
              id="alert-error"
            >
              <AlertTriangle className="w-5 h-5 text-red-500 shrink-0 mt-0.5" />
              <div className="text-sm flex-1">
                <span className="font-semibold">錯誤：</span>
                {errorMsg}
              </div>
              <button onClick={() => setErrorMsg(null)} className="text-red-400 hover:text-red-600 transition-all">
                <X className="w-5 h-5" />
              </button>
            </motion.div>
          )}

          {successMsg && (
            <motion.div
              initial={{ opacity: 0, y: -10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="p-4 bg-blue-50 border border-blue-200 text-blue-700 rounded-xl flex items-start gap-3 shadow-sm"
              id="alert-success"
            >
              <Check className="w-5 h-5 text-blue-500 shrink-0 mt-0.5" />
              <div className="text-sm flex-1 font-medium">{successMsg}</div>
              <button onClick={() => setSuccessMsg(null)} className="text-blue-400 hover:text-blue-600 transition-all">
                <X className="w-5 h-5" />
              </button>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Top Information Row */}
        <div className="grid grid-cols-1 md:grid-cols-12 gap-6">
          {/* Active Upload summary card */}
          <div className="md:col-span-12 bg-white border border-slate-200 rounded-xl p-5 shadow-sm flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
            <div className="flex-1">
              <h2 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-1.5">Active Upload</h2>
              {fileName ? (
                <div>
                  <div className="flex items-center gap-2">
                    <span className="text-sm font-semibold text-slate-800 truncate max-w-[240px] block" title={fileName}>
                      {fileName}
                    </span>
                    <span className="bg-green-50 text-green-700 text-[10px] px-2 py-0.5 rounded-full border border-green-200 font-bold uppercase shrink-0">
                      Mapped Successfully
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500 mt-1">包含 {fileHeaders.length} 個規格項目 (Column A)，共 {fileRecords.length} 欄資料 (Column B+)</p>
                </div>
              ) : (
                <div>
                  <p className="text-sm font-semibold text-slate-400 italic">尚未上傳任何檔案</p>
                  <p className="text-xs text-slate-400 mt-1">請於下方步驟 2 拖曳或瀏覽上傳您的 Excel</p>
                </div>
              )}
            </div>
            {fileName && (
              <div className="flex items-center gap-4 shrink-0 border-t sm:border-t-0 pt-3 sm:pt-0 border-slate-100">
                <div className="text-right">
                  <p className="text-[10px] uppercase font-bold text-slate-400">AI Mapping Confidence</p>
                  <p className="text-xl font-bold text-blue-600">{confidenceScore}%</p>
                </div>
                <div className="w-px h-10 bg-slate-200"></div>
                <div className="text-right">
                  <p className="text-[10px] uppercase font-bold text-slate-400">Mismatched Columns</p>
                  <p className="text-xl font-bold text-red-500">{unmatchedCount}</p>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Bento Grid Step 1 & 2 */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          
          {/* STEP 1: Golden Excel Headers Template */}
          <section className="lg:col-span-7 bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col justify-between" id="section-step1">
            <div className="p-6">
              <div className="flex items-center justify-between border-b border-slate-100 pb-4 mb-4">
                <div className="flex items-center gap-2.5">
                  <span className="flex items-center justify-center w-6 h-6 rounded-full bg-blue-50 text-blue-600 font-bold text-xs">1</span>
                  <div>
                    <h2 className="font-semibold text-slate-900">設定 GOLDEN EXCEL Column A 規格範本</h2>
                    <p className="text-xs text-slate-500">此規格項目順序將作為最終輸出檔案 Column A 的規格排列標準</p>
                  </div>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={handleClearGolden}
                    className="text-slate-400 hover:text-red-500 p-1.5 rounded-lg hover:bg-slate-50 transition-all text-xs flex items-center gap-1"
                    title="清空範本"
                    id="btn-clear-golden"
                  >
                    <Trash2 className="w-4 h-4" />
                    <span>清空</span>
                  </button>
                </div>
              </div>

              {/* Paste Excel Area */}
              <div className="bg-slate-50/50 rounded-xl p-4 border border-slate-200/60 mb-5">
                <label className="block text-xs font-bold uppercase tracking-wider text-slate-600 mb-2 flex items-center gap-1.5">
                  <Layers className="w-3.5 h-3.5 text-blue-600" />
                  貼上從 EXCEL 複製過來的規格項目列表
                </label>
                <div className="flex flex-col sm:flex-row gap-2">
                  <textarea
                    className="flex-1 text-sm bg-white border border-slate-200 rounded-lg p-2.5 focus:outline-none focus:border-blue-500 transition-all font-mono min-h-[50px] resize-y"
                    placeholder="可在此貼上整欄(Column)或整列(Row)複製過來的規格儲存格... (例如：Finish \n Foot \n Front bezel)"
                    value={pastedText}
                    onChange={(e) => setPastedText(e.target.value)}
                    id="textarea-pasted-headers"
                  />
                  <button
                    onClick={handleParsePasted}
                    disabled={!pastedText.trim()}
                    className="bg-blue-600 hover:bg-blue-700 disabled:bg-slate-100 disabled:text-slate-400 text-white font-semibold text-sm px-4 py-2.5 rounded-lg transition-all flex items-center justify-center gap-1.5 shrink-0 shadow-sm active:scale-95"
                    id="btn-parse-pasted"
                  >
                    <Plus className="w-4 h-4" />
                    解析並加入
                  </button>
                </div>
                <p className="text-[10px] text-slate-400 mt-1.5 leading-normal">
                  💡 系統會自動以換行、Tab 或逗號切分並解析為一筆筆垂直規格項目名稱。
                </p>
              </div>

              {/* Add Single Header Manually */}
              <form onSubmit={handleAddHeader} className="flex gap-2 mb-5">
                <input
                  type="text"
                  placeholder="手動新增單一規格項目..."
                  className="flex-1 text-sm bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 focus:outline-none focus:border-blue-500 transition-all text-slate-800 font-medium"
                  value={newHeader}
                  onChange={(e) => setNewHeader(e.target.value)}
                  id="input-new-header"
                />
                <button
                  type="submit"
                  className="bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold px-4 py-2 rounded-lg transition-all shadow-sm active:scale-95"
                  id="btn-add-header-manual"
                >
                  新增
                </button>
              </form>

              {/* Render Golden Header Badges */}
              <div>
                <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3 flex items-center justify-between">
                  <span>目前設定的規格項目 (共 {goldenHeaders.length} 個)</span>
                  {goldenHeaders.length > 0 && <span className="text-blue-600 lowercase text-[10px] font-semibold bg-blue-50 px-2 py-0.5 rounded-full border border-blue-100">主範本順序</span>}
                </h3>
                {goldenHeaders.length === 0 ? (
                  <div className="border border-dashed border-slate-200 rounded-xl p-8 text-center bg-slate-50/30">
                    <p className="text-sm text-slate-400 font-medium">目前尚無任何項目，請在上方貼上規格項目或手動新增單一規格項目。</p>
                  </div>
                ) : (
                  <div className="flex flex-wrap gap-2 max-h-[220px] overflow-y-auto p-3 bg-slate-50/30 border border-slate-200 rounded-xl">
                    {goldenHeaders.map((header, idx) => (
                      <motion.div
                        key={`${header}-${idx}`}
                        initial={{ scale: 0.95, opacity: 0 }}
                        animate={{ scale: 1, opacity: 1 }}
                        className="px-2.5 py-1 bg-slate-100 border border-slate-200 rounded-lg text-xs font-mono text-slate-700 flex items-center gap-1.5 shadow-sm"
                      >
                        <span className="text-slate-400 font-mono text-[9px] font-bold">{idx + 1}</span>
                        <span className="truncate max-w-[120px] font-medium" title={header}>{header}</span>
                        <button
                          type="button"
                          onClick={() => handleDeleteGoldenHeader(idx)}
                          className="text-slate-400 hover:text-red-500 hover:bg-slate-200 p-0.5 rounded transition-all"
                        >
                          <X className="w-3 h-3" />
                        </button>
                      </motion.div>
                    ))}
                  </div>
                )}
              </div>
            </div>

            {/* Step 1 Footer advice */}
            <div className="bg-slate-50/50 p-4 border-t border-slate-100 flex items-center gap-2 text-xs text-slate-500">
              <Info className="w-4 h-4 text-slate-400 shrink-0" />
              <span>當您上傳使用者檔案後，系統將自動解析其規格，並將對應資料重新排列進此規格範本的順序中。</span>
            </div>
          </section>

          {/* STEP 2: Upload User Excel File */}
          <section className="lg:col-span-5 bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col justify-between" id="section-step2">
            <div className="p-6 flex-1 flex flex-col justify-between">
              <div>
                <div className="flex items-center gap-2.5 border-b border-slate-100 pb-4 mb-4">
                  <span className="flex items-center justify-center w-6 h-6 rounded-full bg-blue-50 text-blue-600 font-bold text-xs">2</span>
                  <div>
                    <h2 className="font-semibold text-slate-900">上傳使用者 EXCEL 檔案</h2>
                    <p className="text-xs text-slate-500">系統將重排其規格順序與所有對應的產品欄位資料</p>
                  </div>
                </div>

                {/* Upload Dropzone */}
                <div
                  onDragEnter={handleDrag}
                  onDragOver={handleDrag}
                  onDragLeave={handleDrag}
                  onDrop={handleDrop}
                  onClick={() => fileInputRef.current?.click()}
                  className={`border-2 border-dashed rounded-xl p-6 text-center cursor-pointer transition-all flex flex-col items-center justify-center min-h-[170px] ${
                    dragActive
                      ? "border-blue-500 bg-blue-50/40"
                      : "border-slate-200 hover:border-blue-400 hover:bg-slate-50/50"
                  }`}
                  id="div-dropzone"
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    className="hidden"
                    accept=".xlsx, .xls, .csv"
                    onChange={handleFileChange}
                  />
                  
                  <div className="w-11 h-11 bg-blue-50 text-blue-600 rounded-full flex items-center justify-center mb-3 shadow-sm border border-blue-100">
                     <Upload className="w-5 h-5" />
                  </div>
                  
                  <p className="text-sm font-bold text-slate-900">
                    拖曳檔案至此，或點擊瀏覽檔案
                  </p>
                  <p className="text-xs text-slate-400 mt-1">
                    支援 .xlsx, .xls, .csv (最大 50MB)
                  </p>
                </div>
              </div>

              {/* File details */}
              {fileName && (
                <motion.div
                  initial={{ opacity: 0, scale: 0.95 }}
                  animate={{ opacity: 1, scale: 1 }}
                  className="mt-4 p-4 bg-slate-50 border border-slate-200/80 rounded-xl"
                  id="div-uploaded-file-details"
                >
                  <div className="flex items-center justify-between border-b border-slate-200/50 pb-2 mb-2">
                    <div className="flex items-center gap-2 text-xs font-semibold text-slate-900">
                      <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
                      <span className="truncate max-w-[140px]" title={fileName}>{fileName}</span>
                    </div>
                    <span className="bg-emerald-50 text-emerald-700 text-[9px] font-bold px-2 py-0.5 rounded-full border border-emerald-200 flex items-center gap-0.5">
                      <Check className="w-2.5 h-2.5" /> 已載入
                    </span>
                  </div>

                  <div className="grid grid-cols-2 gap-2 text-xs text-slate-600">
                    <div>
                      <p className="text-slate-400 font-medium text-[10px] uppercase">規格/欄位數量 (Column A)</p>
                      <p className="text-sm font-bold text-slate-800">{fileHeaders.length}</p>
                    </div>
                    <div>
                      <p className="text-slate-400 font-medium text-[10px] uppercase">資料筆數 (Column B+)</p>
                      <p className="text-sm font-bold text-slate-800">{fileRecords.length} 欄</p>
                    </div>
                  </div>
                </motion.div>
              )}
            </div>

            {/* Step 2 Footer info */}
            <div className="bg-slate-50/50 p-4 border-t border-slate-100 flex items-center justify-between text-xs text-slate-500">
              <span className="flex items-center gap-1">
                <Info className="w-4 h-4 text-slate-400 shrink-0" />
                自動與 Golden Column A 規格進行語意比對
              </span>
              {fileName && (
                <button
                  onClick={handleReMatch}
                  disabled={isLoadingMapping}
                  className="text-blue-600 hover:text-blue-800 flex items-center gap-1 font-semibold disabled:text-slate-400 transition-all cursor-pointer"
                  id="btn-rematch"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isLoadingMapping ? "animate-spin" : ""}`} />
                  重新比對
                </button>
              )}
            </div>
          </section>

        </div>

        {/* STEP 3: Preview and Match Verification - Shown only when file is selected */}
        <AnimatePresence>
          {fileName && (
            <motion.div
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 15 }}
              transition={{ duration: 0.25 }}
              className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col"
              id="section-preview"
            >
              {/* Preview Header */}
              <div className="px-6 py-4 border-b border-slate-200 bg-slate-900 text-white flex flex-col md:flex-row md:items-center md:justify-between gap-4">
                <div className="flex items-center gap-2.5">
                  <span className="flex items-center justify-center w-6 h-6 rounded-full bg-blue-500/20 text-blue-300 font-bold text-xs border border-blue-500/30">3</span>
                  <div>
                    <h3 className="font-semibold text-white">驗證規格對應與直式資料預覽</h3>
                    <p className="text-xs text-slate-400 mt-0.5">請確認 Column A 規格對應。非精確對應已藉由 AI 判斷，其來源項目與不吻合儲存格以<span className="text-red-400 font-semibold px-1">紅色/紅色文字</span>標記。</p>
                  </div>
                </div>
                <div className="flex gap-2 text-[11px] text-slate-400 shrink-0">
                  <span className="flex items-center gap-1.5"><div className="w-2 h-2 rounded-full bg-slate-500"></div> 精確比對</span>
                  <span className="flex items-center gap-1.5"><div className="w-2 h-2 rounded-full bg-red-500"></div> AI 自動語意關聯 (可能存在差異)</span>
                </div>
              </div>

              <div className="grid grid-cols-1 xl:grid-cols-12">
                
                {/* Column Mappings list (xl:col-span-5) */}
                <div className="xl:col-span-5 border-b xl:border-b-0 xl:border-r border-slate-200 p-6 bg-slate-50/40">
                  <div className="flex items-center justify-between mb-3">
                    <h3 className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
                      <Sparkles className="w-3.5 h-3.5 text-blue-600 animate-pulse" />
                      規格自動對應設定 (共 {fileHeaders.length} 個規格來源)
                    </h3>
                    {isLoadingMapping && (
                      <span className="text-xs text-blue-600 font-semibold flex items-center gap-1">
                        <RefreshCw className="w-3 h-3 animate-spin" /> AI 分析中...
                      </span>
                    )}
                  </div>

                  <p className="text-xs text-slate-500 mb-4 bg-slate-100/80 p-3 rounded-lg border border-slate-200/60 leading-relaxed">
                    💡 您可以自由調整下拉選單，校正自動比對的結果。若上傳檔案包含多餘規格列，請將其設為<b>「不對應 (略過此規格)」</b>。
                  </p>

                  <div className="space-y-2.5 max-h-[420px] overflow-y-auto pr-1">
                    {fileHeaders.map((uploadedHeader, index) => {
                      const matchedGolden = columnMappings[uploadedHeader];
                      const source = mappingSource[uploadedHeader];

                      const isAiMatch = source === "ai";
                      const isUnmatched = source === "unmatched";
                      const isExact = source === "exact";

                      return (
                        <div
                          key={`${uploadedHeader}_${index}`}
                          className={`p-3 rounded-xl border transition-all ${
                            isAiMatch
                              ? "bg-red-50/50 border-red-200"
                              : isUnmatched
                              ? "bg-slate-50 border-slate-200"
                              : "bg-white border-slate-200 shadow-sm hover:border-slate-300"
                          }`}
                        >
                          <div className="flex items-center justify-between gap-3">
                            <div className="flex-1 min-w-0">
                              <span className="text-[9px] font-mono block text-slate-400 uppercase tracking-wider mb-0.5">
                                上傳檔案 Column A 原始規格
                              </span>
                              <span
                                className={`text-sm font-bold truncate block ${
                                  isAiMatch ? "text-red-600 font-medium" : "text-slate-800"
                                }`}
                                title={uploadedHeader}
                              >
                                {uploadedHeader}
                              </span>
                            </div>

                            <div className="shrink-0 text-slate-300">
                              <ChevronRight className="w-4 h-4" />
                            </div>

                            <div className="flex-1 min-w-0">
                              <span className="text-[9px] font-mono block text-slate-400 uppercase tracking-wider mb-0.5 flex justify-between items-center">
                                <span>對應至 Golden Header</span>
                                {isExact && (
                                  <span className="text-emerald-700 font-bold bg-emerald-50 px-1.5 py-0.2 rounded border border-emerald-100 scale-90">
                                    精確
                                  </span>
                                )}
                                {isAiMatch && (
                                  <span className="text-red-700 font-bold bg-red-50 px-1.5 py-0.2 rounded border border-red-100 scale-90">
                                    AI 接近
                                  </span>
                                )}
                                {isUnmatched && (
                                  <span className="text-slate-500 font-medium bg-slate-100 px-1.5 py-0.2 rounded border scale-90">
                                    無
                                  </span>
                                )}
                              </span>

                              <select
                                className={`w-full text-xs font-semibold rounded-lg border p-1.5 bg-white focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 transition-all ${
                                  isAiMatch
                                    ? "border-red-300 text-red-700"
                                    : "border-slate-200 text-slate-700"
                                }`}
                                value={matchedGolden || "skip"}
                                onChange={(e) => handleManualMappingChange(uploadedHeader, e.target.value)}
                              >
                                <option value="skip" className="text-slate-400 italic">-- 不對應 (略過此項目) --</option>
                                {goldenHeaders.map((gh, ghIdx) => (
                                  <option key={`${gh}_${ghIdx}`} value={gh} className="text-slate-700 font-medium">
                                    {gh}
                                  </option>
                                ))}
                              </select>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                {/* Rearranged Data Grid Preview (xl:col-span-7) */}
                <div className="xl:col-span-7 p-6 flex flex-col justify-between bg-white">
                  <div>
                    <div className="flex items-center gap-2 mb-3">
                      <div className="w-2 h-2 bg-emerald-500 rounded-full animate-ping" />
                      <h3 className="text-sm font-bold text-slate-900">
                        重排後 Golden Excel 直式預覽畫面 (共 {fileRecords.length} 欄資料)
                      </h3>
                    </div>

                    <p className="text-xs text-slate-500 mb-4 leading-relaxed">
                      這是在匯出前，依據您的 <b>Golden Column A 規格順序</b> 自動重組與填入後的即時直向預覽 (僅預覽前 5 欄)：
                    </p>

                    {/* Aligned Preview Table Container */}
                    <div className="border border-slate-200 rounded-xl overflow-x-auto shadow-inner bg-slate-50/30">
                      <table className="w-full text-left border-collapse min-w-[700px]">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-200">
                            <th className="px-4 py-3 text-xs font-bold text-slate-700 border-r border-slate-200 w-[200px]">Golden Column A (標準規格列表)</th>
                            <th className="px-4 py-3 text-xs font-bold text-slate-700 border-r border-slate-200 w-[140px]">對應來源 / 狀態</th>
                            {/* Columns for the records */}
                            {Array.from({ length: Math.min(5, fileRecords.length) }).map((_, cIdx) => (
                              <th key={cIdx} className="px-4 py-3 text-xs font-bold text-slate-700 border-r border-slate-200 last:border-r-0">
                                資料 #{cIdx + 1} (Column {String.fromCharCode(66 + cIdx)})
                              </th>
                            ))}
                            {fileRecords.length > 5 && (
                              <th className="px-4 py-3 text-xs font-bold text-slate-400">...還有 {fileRecords.length - 5} 欄</th>
                            )}
                          </tr>
                        </thead>
                        <tbody>
                          {goldenHeaders.length === 0 ? (
                            <tr>
                              <td colSpan={2 + Math.min(5, fileRecords.length) + (fileRecords.length > 5 ? 1 : 0)} className="text-center py-8 text-slate-400 text-xs">
                                請先設定 Golden Headers 規格範本。
                              </td>
                            </tr>
                          ) : (
                            goldenHeaders.map((gh, ghIdx) => {
                              // Is there any uploaded column mapping to this golden header?
                              const mappedTargetHeader = Object.keys(columnMappings).find(
                                target => columnMappings[target] === gh
                              );
                              const source = mappedTargetHeader ? mappingSource[mappedTargetHeader] : "unmatched";
                              const isAi = source === "ai";
                              const isExact = source === "exact";
                              const isManual = source === "manual";

                              return (
                                <tr key={`${gh}_${ghIdx}`} className="border-b border-slate-100 last:border-0 bg-white hover:bg-blue-50/30 transition-colors text-xs text-slate-600">
                                  {/* Column 1: Golden Header Name */}
                                  <td className="px-4 py-2.5 font-bold text-slate-800 border-r border-slate-200">
                                    <span className={isAi ? "text-red-600 font-semibold" : ""}>{gh}</span>
                                  </td>
                                  {/* Column 2: Status / Source label */}
                                  <td className="px-4 py-2.5 border-r border-slate-200 text-[10px]">
                                    {mappedTargetHeader ? (
                                      <div className="flex flex-col gap-0.5">
                                        <span className="font-mono text-[9px] text-slate-400 block truncate max-w-[120px]" title={mappedTargetHeader}>
                                          ← {mappedTargetHeader}
                                        </span>
                                        {isExact && <span className="text-emerald-700 font-bold bg-emerald-50 px-1 py-0.2 rounded border border-emerald-100 w-max">精確對應</span>}
                                        {isAi && <span className="text-red-700 font-bold bg-red-50 px-1 py-0.2 rounded border border-red-100 w-max animate-pulse">AI 接近對應</span>}
                                        {isManual && <span className="text-blue-700 font-bold bg-blue-50 px-1 py-0.2 rounded border border-blue-100 w-max">手動校正</span>}
                                      </div>
                                    ) : (
                                      <span className="text-slate-400 italic">(未對應 - 將匯出為空值)</span>
                                    )}
                                  </td>
                                  {/* Column 3+: Product Values */}
                                  {Array.from({ length: Math.min(5, fileRecords.length) }).map((_, cIdx) => {
                                    let val = "";
                                    if (mappedTargetHeader) {
                                      const targetIdx = fileHeaders.indexOf(mappedTargetHeader);
                                      if (targetIdx !== -1) {
                                        val = fileRecords[cIdx][targetIdx];
                                      }
                                    }
                                    return (
                                      <td key={cIdx} className="px-4 py-2.5 border-r border-slate-200 last:border-r-0 truncate max-w-[180px]" title={String(val)}>
                                        {val === "" ? (
                                          <span className="text-slate-300 italic font-light">-</span>
                                        ) : (
                                          <span className={isAi ? "text-red-600 font-medium" : "text-slate-700"}>{String(val)}</span>
                                        )}
                                      </td>
                                    );
                                  })}
                                  {fileRecords.length > 5 && (
                                    <td className="px-4 py-2.5 text-slate-400 italic text-[10px]">...</td>
                                  )}
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                    {fileRecords.length > 5 && (
                      <p className="text-[10px] text-slate-400 mt-2 text-right italic">
                        * 僅預覽前 5 欄，匯出下載檔案將包含完整所有的 {fileRecords.length} 欄資料。
                      </p>
                    )}
                  </div>

                  {/* Actions Footer */}
                  <div className="mt-6 pt-4 border-t border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-4">
                    <button
                      onClick={() => {
                        setFileName(null);
                        setFileHeaders([]);
                        setFileRecords([]);
                        setColumnMappings({});
                        setMappingSource({});
                        showSuccess("已取消並清空目前上傳的檔案。");
                      }}
                      className="w-full sm:w-auto px-4 py-2 border border-slate-300 rounded-lg text-xs font-semibold text-slate-600 hover:bg-slate-100 transition-all cursor-pointer"
                    >
                      取消並清空檔案
                    </button>
                    <div className="flex gap-3 w-full sm:w-auto">
                      <button
                        onClick={handleReMatch}
                        className="flex-1 sm:flex-none px-4 py-2 border border-blue-600 rounded-lg text-xs font-semibold text-blue-600 hover:bg-blue-50 transition-all cursor-pointer"
                      >
                        重新比對
                      </button>
                      <button
                        onClick={handleExportAndDownload}
                        className="flex-1 sm:flex-none px-6 py-2 bg-blue-600 rounded-lg text-xs font-bold text-white shadow-lg shadow-blue-200 hover:bg-blue-700 transition-all flex items-center justify-center gap-1.5 cursor-pointer active:scale-95"
                        id="btn-export-download-mini"
                      >
                        <Download className="w-3.5 h-3.5" />
                        整合並下載
                      </button>
                    </div>
                  </div>
                </div>

              </div>
            </motion.div>
          )}
        </AnimatePresence>

      </main>
    </div>
  );
}
