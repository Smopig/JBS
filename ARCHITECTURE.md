# JBS Wafer Analysis — 架構說明（ARCHITECTURE）

> 本文件說明 `JBS Wafer Analysis.html` 的內部結構、資料流與模組職責，供維護與後續重構參考。
> 目前應用為**單一 HTML 檔**（~9,700 行），所有 CSS / JS 內嵌於同一檔案，可完全離線運作。

## 1. 總覽

`JBS Wafer Analysis.html` 是半導體晶圓 CP（Circuit Probe）測試資料的互動分析工具。
使用者上傳多個 CP 測試檔（每檔 = 一個 Lot/Wafer），工具自動辨識格式、解析量測資料，
並提供晶圓地圖、分佈圖、統計、良率、Bin Pareto、Split 關聯、SPC、QC 等多種分析分頁。

### 相依函式庫（CDN，另支援離線 `libs/`）
| 函式庫 | 版本 | 用途 |
|--------|------|------|
| `xlsx-js-style` | 1.2.0 | Excel / CSV 讀寫與樣式 |
| `plotly.js-dist-min` | 2.35.2 | 所有互動圖表與晶圓熱圖 |
| `html2canvas` | 1.4.1 | 圖表截圖 / 批次圖片匯出 |

## 2. 檔案內部分層

單檔由上而下大致分為：
1. `<head>`：Google Fonts、`<style>`（CSS 變數設計系統、主題、密度、RWD、無障礙焦點環）。
2. `<body>`：Header、分頁導覽（`role="tablist"`）、各 `tab-panel` 內容、地圖匯出 Modal、Toast 容器。
3. `<script>`（主邏輯，~7,800 行）：常數 → 全域狀態 `S` → 解析器 → 資料/渲染/工具函式 → 事件處理。

## 3. 全域狀態 `S`（依關注點分群）

`S` 為單一全域物件，集中所有應用狀態。原始碼中已加上分區註解，分群如下：

| 分群 | 代表屬性 | 職責 |
|------|----------|------|
| 資料模型 Data Model | `data, lots, items, globalBins, activeLots, lotColors` | 解析後的量測列、Lot/項目清單、顯示篩選 |
| 載入 / 解析 Parsing | `_parsed, _activeTemplates, _autoDetect, _colSpecsRaw/User, _excludedCols, _colConvOverrides, _headerIdx, _dataStartIdx` | 檔案解析快取、模板覆寫、欄位換算 / 排除 / SPEC 抽取 |
| 統計 Statistics | `activeStats, customStats, statSpecs, statFormats, statItems, *OnlyInSpec` | 統計分頁的統計函式選擇、自訂統計、SPEC |
| 圖表分析 Chart | `chartItems, chartType, chartSpecs, chartSplit*` | 圖表分析分頁的項目 / 類型 / 分組設定 |
| 分布 Distribution | `cntSplit*, pctSplit*` | Count / % 分布的 Split 設定 |
| Split 關聯 | `splitProfiles, splitLotMaps, activeProfile, chartSplitProfile, splitItems, splitEditMode` | Split 關聯分析的 profile 與對應 |
| 模板設計器 Template | `splitTemplates, activeTemplate, splitBatchTemplates, splitBatchLocked` | Split 模板設計與批次連結 |
| 遮罩 Mask | `maskSet, maskModeOn` | 晶圓座標 die 遮罩 |

> **Proxy 注意**：`S.splitData` / `S.splitLotMap` 透過 `Object.defineProperty` 代理至
> `S.splitProfiles[S.activeProfile]`，讓舊程式碼無需修改即可切換 profile。

### 常數
`STD_RENAME`（標準命名對照）、`META_COLS` / `CANONICAL_META`（非量測欄位）、
`N_BINS`（直方圖分箱數）、`DEFAULT_STATS`、`BUILTIN_TEMPLATES`（內建 CP 格式）、
`COLORSCALES`（色階預設）。

## 4. 模組職責（函式分類）

雖在同一作用域，函式依職責可歸為以下「邏輯模組」：

- **Parser（解析器）**：`detectTemplate`、`parseWithTemplate`、`_extractMetadata`、
  `resolveLotId`、`buildRecordsFromEntry`、`parseUploadedFiles`。
  依 `BUILTIN_TEMPLATES` 評分辨識格式，定位 header / data 列，抽取 metadata 與 SPEC。
- **DataModel（資料模型）**：`applyAndAnalyze`（解析→合併→建索引）、`getActiveLots`、
  `getLotRows`（**記憶化 per-Lot 索引**，取代重複的 `S.data.filter`）、`validData`、
  `_isMasked`、`makeHist`。
- **Renderer（渲染器）**：
  - 晶圓地圖：`renderMap`、`renderDeltaMap`（Wafer-to-Wafer 差異圖）。
  - 分布：`renderCount`、`renderPct`。
  - 統計：`renderStats`、`buildItemStatsBlock`。
  - 良率 / Bin：`renderYield`、`renderBinPareto`。
  - 圖表分析：`renderChart` 分派至 `caBox / caViolin / caScatter / caTrend / caBar /
    caSPC / caCDF / caQQ / caCorr / caDual`。
  - SPC run rules：`spcRunRules`（Western Electric / Nelson 規則違反偵測）。
- **Utils（工具）**：數學 `mean / std / median / percentile / linRegress / rd`、
  色彩 `getColor / getLotColor / _scaleColorHex`、`escHtml / uid / tick`、`Toast`。
- **State / Persistence（持久化）**：`saveSettings / loadSettings`（localStorage，
  `SETTINGS_KEY = 'jbs_settings_v2'`）、`debounceSave`、`exportSession / importSession`、
  `copyDataCsv / resetSettings`。
- **UI / 事件**：`switchTab`（含 `aria-selected` 同步）、`_refreshAllPanels`、
  拖放與檔案處理、各 `on*` 變更處理器。

## 5. 資料流

```
使用者上傳檔案
  → handleFiles            （存入 S.files）
  → parseUploadedFiles     （XLSX 讀取 → detectTemplate 評分辨識）
  → parseWithTemplate      （定位 header/data 列、抽 metadata 與 SPEC）
  → buildRecordsFromEntry  （轉物件、套用 rename/換算/排除）
  → applyAndAnalyze        （合併進 S.data，建立 S.lots/S.items/globalBins）
  → populateSels / 各 render*（依目前分頁渲染 Plotly 圖表）
```

`getLotRows(lot)` 在 `S.data` 重新指派或長度變動時自動重建索引，供所有渲染函式查表。

## 6. 演進記錄（近期優化批次）

- **P0**：`getLotRows` 記憶化索引；複製資料 CSV / 重置設定。
- **P1**：`:focus-visible` 焦點環、分頁 / Modal ARIA、ESC 關閉、小螢幕 RWD。
- **P2**：Bin Pareto、Wafer Delta Map、SPC run rules。
- **P3**：狀態分群註解 + 本架構文件（路線 B，零行為變更）。

## 7. 後續重構建議（尚未執行）

- **路線 A — Build 拆檔**：將單檔拆為 `src/` 內 CSS / JS 模組，以 build 腳本 concat 回單檔，
  保留離線特性同時提升可維護性（產物可用「與原檔差異」客觀驗證）。
- **路線 C — 三工具合併**：將 `wafer_tool/index.html` 視為 JBS 子集、以 feature flag 共用核心，
  並與 `wafer_analysis.py` 抽出共同 JSON schema，消除 ~90% 重複。風險最高，建議於可瀏覽器驗證的環境進行。
