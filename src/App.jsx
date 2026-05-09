import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart3,
  Brush,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  Filter,
  LineChart,
  PaintBucket,
  PieChart,
  Plus,
  RotateCcw,
  Sparkles,
  Table2,
  Upload,
  Wand2,
} from 'lucide-react';

const defaultTheme = {
  title: 'Enterprise Data Cleaning and Analytics',
  subtitle: 'Upload, profile, clean, compare, analyze, and visualize tabular datasets.',
  primary: '#6d5dfc',
  secondary: '#25c4c2',
  accent: '#94bfff',
  ink: '#252636',
  surface: '#ffffff',
};

const DB_NAME = 'fedm-data-cleaning-workspace';
const DB_STORE = 'snapshots';
const SNAPSHOT_KEY = 'current';

function openWorkspaceDb() {
  return new Promise((resolve, reject) => {
    if (!('indexedDB' in window)) {
      reject(new Error('IndexedDB is not supported'));
      return;
    }
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      request.result.createObjectStore(DB_STORE);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function readWorkspaceSnapshot() {
  const db = await openWorkspaceDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(DB_STORE, 'readonly');
    const request = transaction.objectStore(DB_STORE).get(SNAPSHOT_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
    transaction.oncomplete = () => db.close();
  });
}

async function saveWorkspaceSnapshot(snapshot) {
  const db = await openWorkspaceDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(DB_STORE, 'readwrite');
    const request = transaction.objectStore(DB_STORE).put(snapshot, SNAPSHOT_KEY);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
    transaction.oncomplete = () => db.close();
  });
}

async function clearWorkspaceSnapshot() {
  const db = await openWorkspaceDb();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(DB_STORE, 'readwrite');
    const request = transaction.objectStore(DB_STORE).delete(SNAPSHOT_KEY);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
    transaction.oncomplete = () => db.close();
  });
}

function isMissing(value) {
  return value === null || value === undefined || String(value).trim() === '' || String(value).trim().toLowerCase() === 'na';
}

function toNumber(value) {
  if (isMissing(value)) return null;
  const number = Number(String(value).replace(/,/g, ''));
  return Number.isFinite(number) ? number : null;
}

function median(values) {
  const sorted = values.slice().sort((a, b) => a - b);
  if (!sorted.length) return 0;
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}

function mode(values) {
  const counts = new Map();
  values.forEach((value) => {
    const key = String(value);
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? '';
}

function titleCase(value) {
  return String(value)
    .toLowerCase()
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function formatNumber(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return '0';
  return new Intl.NumberFormat('en-US', { maximumFractionDigits: 1 }).format(number);
}

function normalizeRows(rows) {
  return rows.map((row) =>
    Object.fromEntries(
      Object.entries(row).map(([key, value]) => [
        String(key).trim() || 'Column',
        value instanceof Date ? value.toISOString().slice(0, 10) : value,
      ]),
    ),
  );
}

function inferType(values) {
  const clean = values.filter((value) => !isMissing(value));
  if (!clean.length) return 'empty';
  const numeric = clean.filter((value) => toNumber(value) !== null).length;
  const dates = clean.filter((value) => !Number.isNaN(Date.parse(value))).length;
  const booleans = clean.filter((value) => ['true', 'false', 'yes', 'no', '1', '0'].includes(String(value).toLowerCase())).length;
  if (numeric / clean.length >= 0.8) return 'number';
  if (dates / clean.length >= 0.8) return 'date';
  if (booleans / clean.length >= 0.8) return 'boolean';
  return 'text';
}

function getColumns(rows) {
  return [...new Set(rows.flatMap((row) => Object.keys(row)))];
}

function buildProfile(rows) {
  const columns = getColumns(rows);
  const rowFingerprints = rows.map((row) => JSON.stringify(columns.map((column) => row[column] ?? '')));
  const duplicateRows = rowFingerprints.length - new Set(rowFingerprints).size;

  const columnProfiles = columns.map((column) => {
    const values = rows.map((row) => row[column]);
    const clean = values.filter((value) => !isMissing(value));
    const type = inferType(values);
    const numbers = clean.map(toNumber).filter((value) => value !== null);
    const frequent = mode(clean);
    return {
      column,
      type,
      missing: values.length - clean.length,
      unique: new Set(clean.map(String)).size,
      frequent,
      min: numbers.length ? Math.min(...numbers) : null,
      max: numbers.length ? Math.max(...numbers) : null,
      mean: numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : null,
      median: numbers.length ? median(numbers) : null,
    };
  });

  return {
    rows: rows.length,
    columns: columns.length,
    missing: columnProfiles.reduce((sum, profile) => sum + profile.missing, 0),
    duplicateRows,
    columnsList: columns,
    columnProfiles,
  };
}

function groupRows(rows, categoryColumn, valueColumn, limit = 8, sortMode = 'value') {
  const grouped = new Map();
  rows.forEach((row) => {
    const key = String(row[categoryColumn] ?? 'Blank').trim() || 'Blank';
    const numeric = toNumber(row[valueColumn]);
    grouped.set(key, (grouped.get(key) || 0) + (numeric ?? 1));
  });
  const result = [...grouped.entries()].map(([label, value]) => ({ label, value }));
  result.sort((a, b) => (sortMode === 'label' ? String(a.label).localeCompare(String(b.label)) : b.value - a.value));
  return result.slice(0, limit);
}

function getNumericColumns(profile) {
  return profile.columnProfiles.filter((column) => column.type === 'number').map((column) => column.column);
}

function getTextColumns(profile) {
  return profile.columnProfiles.filter((column) => column.type !== 'number').map((column) => column.column);
}

function BarChartCard({ data, theme }) {
  if (!data.length) return <EmptyChart message="Upload data to build a bar chart." />;
  const max = Math.max(...data.map((item) => item.value), 1);
  return (
    <div className="chart-shell bar-shell">
      {data.map((item, index) => (
        <div className="bar-item" key={item.label}>
          <div className="bar-stack">
            <span style={{ height: `${Math.max((item.value / max) * 100, 4)}%`, background: index % 2 ? theme.primary : theme.accent }} />
            <span style={{ height: `${Math.max((item.value / max) * 42, 3)}%`, background: theme.secondary }} />
          </div>
          <small>{item.label.slice(0, 9)}</small>
        </div>
      ))}
    </div>
  );
}

function DonutChart({ data, theme }) {
  if (!data.length) return <EmptyChart message="Upload data to build a donut chart." />;
  const total = data.reduce((sum, item) => sum + item.value, 0) || 1;
  let offset = 25;
  const colors = [theme.primary, '#8477f6', theme.secondary, theme.accent, '#d8f7f6'];
  return (
    <div className="donut-wrap">
      <svg viewBox="0 0 44 44" className="donut">
        <circle cx="22" cy="22" r="15.915" fill="transparent" stroke="#eef0f7" strokeWidth="8" />
        {data.map((item, index) => {
          const dash = (item.value / total) * 100;
          const element = (
            <circle
              key={item.label}
              cx="22"
              cy="22"
              r="15.915"
              fill="transparent"
              stroke={colors[index % colors.length]}
              strokeWidth="8"
              strokeDasharray={`${dash} ${100 - dash}`}
              strokeDashoffset={offset}
            />
          );
          offset -= dash;
          return element;
        })}
      </svg>
      <div className="donut-center">
        <small>Total</small>
        <strong>{formatNumber(total)}</strong>
      </div>
      <div className="legend-stack">
        {data.slice(0, 3).map((item, index) => (
          <span key={item.label}>
            <i style={{ background: colors[index] }} />
            {item.label} <b>{Math.round((item.value / total) * 100)}%</b>
          </span>
        ))}
      </div>
    </div>
  );
}

function GaugeCard({ ratio, theme, numeratorLabel, denominatorLabel }) {
  if (ratio === null) return <EmptyChart message="Upload data and choose two numeric columns to build a ratio gauge." />;
  const value = Math.max(0, Math.min(ratio, 1.25));
  const percent = Math.min(value / 1.25, 1);
  const angle = -180 + percent * 180;
  const markerX = 90 + Math.cos((angle * Math.PI) / 180) * 64;
  const markerY = 90 + Math.sin((angle * Math.PI) / 180) * 64;
  return (
    <div className="gauge-wrap">
      <div className="mini-legend">
        <span><i style={{ background: theme.primary }} />{denominatorLabel || 'Denominator'}</span>
        <span><i style={{ background: '#8878f5' }} />{numeratorLabel || 'Numerator'}</span>
      </div>
      <svg viewBox="0 0 180 100" className="gauge">
        <path d="M25 90 A65 65 0 0 1 155 90" fill="none" stroke="#eef0f7" strokeWidth="36" />
        <path
          d="M25 90 A65 65 0 0 1 155 90"
          fill="none"
          stroke={theme.primary}
          strokeWidth="36"
          strokeDasharray={`${percent * 205} 205`}
          strokeLinecap="butt"
        />
        <circle cx={markerX} cy={markerY} r="7" fill="#fff" stroke={theme.primary} strokeWidth="4" />
      </svg>
      <div className="gauge-label">
        <strong>{Math.round(ratio * 100)}%</strong>
        <span>{numeratorLabel || 'Value'} to {denominatorLabel || 'total'}</span>
      </div>
    </div>
  );
}

function PieChartCard({ data, theme }) {
  if (!data.length) return <EmptyChart message="Upload data to build a pie chart." />;
  const total = data.reduce((sum, item) => sum + item.value, 0) || 1;
  let current = 0;
  const colors = [theme.primary, theme.secondary, theme.accent, '#ff8d7d', '#a6e7d8', '#d7d2ff'];
  const segments = data.map((item, index) => {
    const start = current / total;
    const end = (current + item.value) / total;
    current += item.value;
    const large = end - start > 0.5 ? 1 : 0;
    const startAngle = start * Math.PI * 2 - Math.PI / 2;
    const endAngle = end * Math.PI * 2 - Math.PI / 2;
    const x1 = 50 + Math.cos(startAngle) * 42;
    const y1 = 50 + Math.sin(startAngle) * 42;
    const x2 = 50 + Math.cos(endAngle) * 42;
    const y2 = 50 + Math.sin(endAngle) * 42;
    return (
      <path
        key={item.label}
        d={`M 50 50 L ${x1} ${y1} A 42 42 0 ${large} 1 ${x2} ${y2} Z`}
        fill={colors[index % colors.length]}
      />
    );
  });

  return (
    <div className="pie-wrap">
      <svg viewBox="0 0 100 100" className="pie-chart">{segments}</svg>
      <div className="legend-stack">
        {data.slice(0, 5).map((item, index) => (
          <span key={item.label}>
            <i style={{ background: colors[index % colors.length] }} />
            {item.label} <b>{Math.round((item.value / total) * 100)}%</b>
          </span>
        ))}
      </div>
    </div>
  );
}

function LineChartCard({ data, theme }) {
  if (!data.length) return <EmptyChart message="Upload data to build a line chart." />;
  const width = 620;
  const height = 230;
  const max = Math.max(...data.map((item) => item.value), 1);
  const points = data.map((item, index) => {
    const x = data.length === 1 ? width / 2 : (index / (data.length - 1)) * width;
    const y = height - (item.value / max) * (height - 28) - 10;
    return `${x},${y}`;
  });
  const comparisonPoints = data.map((item, index) => {
    const x = data.length === 1 ? width / 2 : (index / (data.length - 1)) * width;
    const wave = 0.62 + Math.sin(index * 0.9) * 0.18;
    const y = height - ((item.value * wave) / max) * (height - 42) - 18;
    return `${x},${Math.max(12, Math.min(height - 12, y))}`;
  });

  return (
    <div className="line-shell">
      <svg viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none">
        {[0, 1, 2, 3, 4].map((line) => (
          <line key={line} x1="0" x2={width} y1={(height / 5) * line + 10} y2={(height / 5) * line + 10} stroke="#e6e8f0" strokeDasharray="4 6" />
        ))}
        <polyline fill="none" stroke={theme.primary} strokeWidth="5" strokeLinecap="round" strokeLinejoin="round" points={points.join(' ')} />
        <polyline fill="none" stroke="#ff8d7d" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" opacity="0.48" points={comparisonPoints.join(' ')} />
        {data.map((item, index) => {
          const [x, y] = points[index].split(',').map(Number);
          return <circle key={item.label} cx={x} cy={y} r="5" fill={theme.primary} />;
        })}
      </svg>
      <div className="line-axis">
        {data.map((item) => <small key={item.label}>{item.label.slice(0, 4)}</small>)}
      </div>
    </div>
  );
}

function ScatterPlotCard({ rows, xColumn, yColumn, colorColumn, theme }) {
  if (!rows.length || !xColumn || !yColumn) return <EmptyChart message="Upload data and choose numeric X/Y columns." />;
  const width = 620;
  const height = 260;
  const points = rows
    .map((row) => ({
      x: toNumber(row[xColumn]),
      y: toNumber(row[yColumn]),
      color: String(row[colorColumn] ?? 'Blank') || 'Blank',
    }))
    .filter((point) => point.x !== null && point.y !== null);
  if (!points.length) return <EmptyChart message="No numeric points available for this scatter plot." />;
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const minX = Math.min(...xs, 0);
  const maxX = Math.max(...xs, 1);
  const minY = Math.min(...ys, 0);
  const maxY = Math.max(...ys, 1);
  const colors = [theme.primary, theme.secondary, theme.accent, '#ff8d7d', '#a6e7d8', '#2d3044'];
  const colorValues = [...new Set(points.map((point) => point.color))];

  return (
    <div className="scatter-shell">
      <svg viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none">
        {[0, 1, 2, 3, 4].map((line) => (
          <React.Fragment key={line}>
            <line x1="0" x2={width} y1={(height / 4) * line} y2={(height / 4) * line} stroke="#e6e8f0" strokeDasharray="4 6" />
            <line x1={(width / 4) * line} x2={(width / 4) * line} y1="0" y2={height} stroke="#eef0f7" strokeDasharray="4 6" />
          </React.Fragment>
        ))}
        {points.map((point, index) => {
          const x = ((point.x - minX) / Math.max(maxX - minX, 1)) * (width - 32) + 16;
          const y = height - (((point.y - minY) / Math.max(maxY - minY, 1)) * (height - 32) + 16);
          return (
            <circle
              key={`${point.x}-${point.y}-${index}`}
              cx={x}
              cy={y}
              r="7"
              fill={colors[colorValues.indexOf(point.color) % colors.length]}
              opacity="0.82"
            />
          );
        })}
      </svg>
      <div className="scatter-legend">
        {colorValues.slice(0, 6).map((value, index) => (
          <span key={value}><i style={{ background: colors[index % colors.length] }} />{value}</span>
        ))}
      </div>
    </div>
  );
}

function EmptyChart({ message }) {
  return (
    <div className="empty-chart">
      <FileSpreadsheet size={28} />
      <span>{message}</span>
    </div>
  );
}

function DiagramRenderer({ type, data, rows, xColumn, yColumn, colorColumn, theme }) {
  if (type === 'line') return <LineChartCard data={data} theme={theme} />;
  if (type === 'pie') return <PieChartCard data={data} theme={theme} />;
  if (type === 'donut') return <DonutChart data={data} theme={theme} />;
  if (type === 'scatter') return <ScatterPlotCard rows={rows} xColumn={xColumn} yColumn={yColumn} colorColumn={colorColumn} theme={theme} />;
  return <BarChartCard data={data} theme={theme} />;
}

function App() {
  const [originalRows, setOriginalRows] = useState([]);
  const [cleanRows, setCleanRows] = useState([]);
  const [fileName, setFileName] = useState('No file uploaded');
  const [activeView, setActiveView] = useState('cleaned');
  const [activePage, setActivePage] = useState('upload');
  const [isDragging, setIsDragging] = useState(false);
  const [notice, setNotice] = useState('');
  const [snapshotReady, setSnapshotReady] = useState(false);
  const [theme, setTheme] = useState(defaultTheme);
  const [notes, setNotes] = useState(['Defense note: explain why each cleaning method was applied and compare the before/after metrics.']);
  const [newNote, setNewNote] = useState('');
  const [history, setHistory] = useState(['Waiting for uploaded dataset']);
  const [changedCells, setChangedCells] = useState(new Set());
  const [cleaningApplied, setCleaningApplied] = useState(false);
  const [selectedColumn, setSelectedColumn] = useState('');
  const [missingMethod, setMissingMethod] = useState('mean');
  const [customFill, setCustomFill] = useState('Unknown');
  const [convertType, setConvertType] = useState('number');
  const [standardizeMethod, setStandardizeMethod] = useState('trim');
  const [minValue, setMinValue] = useState('0');
  const [maxValue, setMaxValue] = useState('');
  const [gaugeNumerator, setGaugeNumerator] = useState('');
  const [gaugeDenominator, setGaugeDenominator] = useState('');
  const [barCategory, setBarCategory] = useState('');
  const [barValue, setBarValue] = useState('');
  const [pieCategory, setPieCategory] = useState('');
  const [pieValue, setPieValue] = useState('');
  const [lineX, setLineX] = useState('');
  const [lineY, setLineY] = useState('');
  const [diagramType, setDiagramType] = useState('scatter');
  const [diagramCategory, setDiagramCategory] = useState('');
  const [diagramValue, setDiagramValue] = useState('');
  const [diagramX, setDiagramX] = useState('');
  const [diagramY, setDiagramY] = useState('');
  const [diagramColor, setDiagramColor] = useState('');
  const fileRef = useRef(null);

  const originalProfile = useMemo(() => buildProfile(originalRows), [originalRows]);
  const cleanProfile = useMemo(() => buildProfile(cleanRows), [cleanRows]);
  const columns = cleanProfile.columnsList;
  const numericColumns = getNumericColumns(cleanProfile);
  const textColumns = getTextColumns(cleanProfile);
  const hasData = cleanRows.length > 0;

  const dashboard = useMemo(() => {
    const numeratorTotal = cleanRows.reduce((sum, row) => sum + (toNumber(row[gaugeNumerator]) ?? 0), 0);
    const denominatorTotal = cleanRows.reduce((sum, row) => sum + (toNumber(row[gaugeDenominator]) ?? 0), 0);
    return {
      ratioGauge: hasData && gaugeNumerator && gaugeDenominator && denominatorTotal ? numeratorTotal / denominatorTotal : null,
      bar: groupRows(cleanRows, barCategory, barValue),
      pie: groupRows(cleanRows, pieCategory, pieValue),
      line: groupRows(cleanRows, lineX, lineY, 12, 'label'),
      diagram: groupRows(cleanRows, diagramCategory, diagramValue, 12, diagramType === 'line' ? 'label' : 'value'),
    };
  }, [barCategory, barValue, cleanRows, diagramCategory, diagramType, diagramValue, gaugeDenominator, gaugeNumerator, hasData, lineX, lineY, pieCategory, pieValue]);

  const insights = useMemo(() => {
    const missingReduced = Math.max(originalProfile.missing - cleanProfile.missing, 0);
    const duplicateReduced = Math.max(originalProfile.duplicateRows - cleanProfile.duplicateRows, 0);
    const topFrequent = cleanProfile.columnProfiles
      .filter((column) => column.frequent)
      .slice(0, 4)
      .map((column) => `${column.column}: ${column.frequent}`);
    const bestNumeric = cleanProfile.columnProfiles
      .filter((column) => column.mean !== null)
      .sort((a, b) => Math.abs(b.mean) - Math.abs(a.mean))[0];

    return [
      `${missingReduced} missing value${missingReduced === 1 ? '' : 's'} cleaned from the uploaded data.`,
      `${duplicateReduced} duplicate row${duplicateReduced === 1 ? '' : 's'} removed from the dataset.`,
      bestNumeric ? `${bestNumeric.column} has an average of ${formatNumber(bestNumeric.mean)} and ranges from ${formatNumber(bestNumeric.min)} to ${formatNumber(bestNumeric.max)}.` : 'No numeric column is available yet for statistical interpretation.',
      topFrequent.length ? `Most frequent values include ${topFrequent.join(', ')}.` : 'Categorical patterns will appear after a dataset is uploaded.',
    ];
  }, [cleanProfile, originalProfile]);

  useEffect(() => {
    let cancelled = false;
    readWorkspaceSnapshot()
      .then((snapshot) => {
        if (cancelled || !snapshot) return;
        setOriginalRows(snapshot.originalRows || []);
        setCleanRows(snapshot.cleanRows || []);
        setFileName(snapshot.fileName || 'No file uploaded');
        setActiveView(snapshot.activeView || 'cleaned');
        setActivePage(snapshot.activePage || 'upload');
        setTheme({ ...defaultTheme, ...(snapshot.theme || {}) });
        setNotes(snapshot.notes?.length ? snapshot.notes : ['Defense note: explain why each cleaning method was applied and compare the before/after metrics.']);
        setHistory(snapshot.history?.length ? snapshot.history : ['Waiting for uploaded dataset']);
        setChangedCells(new Set(snapshot.changedCells || []));
        setCleaningApplied(Boolean(snapshot.cleaningApplied));
        setSelectedColumn(snapshot.selectedColumn || '');
        setMissingMethod(snapshot.missingMethod || 'mean');
        setCustomFill(snapshot.customFill || 'Unknown');
        setConvertType(snapshot.convertType || 'number');
        setStandardizeMethod(snapshot.standardizeMethod || 'trim');
        setMinValue(snapshot.minValue ?? '0');
        setMaxValue(snapshot.maxValue ?? '');
        setGaugeNumerator(snapshot.gaugeNumerator || '');
        setGaugeDenominator(snapshot.gaugeDenominator || '');
        setBarCategory(snapshot.barCategory || '');
        setBarValue(snapshot.barValue || '');
        setPieCategory(snapshot.pieCategory || '');
        setPieValue(snapshot.pieValue || '');
        setLineX(snapshot.lineX || '');
        setLineY(snapshot.lineY || '');
        setDiagramType(snapshot.diagramType || 'scatter');
        setDiagramCategory(snapshot.diagramCategory || '');
        setDiagramValue(snapshot.diagramValue || '');
        setDiagramX(snapshot.diagramX || '');
        setDiagramY(snapshot.diagramY || '');
        setDiagramColor(snapshot.diagramColor || '');
        showNotice('Restored saved workspace');
      })
      .catch(() => {
        showNotice('Offline saving is not available in this browser.');
      })
      .finally(() => {
        if (!cancelled) setSnapshotReady(true);
      });

    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!snapshotReady) return undefined;
    const snapshot = {
      version: 1,
      savedAt: new Date().toISOString(),
      originalRows,
      cleanRows,
      fileName,
      activeView,
      activePage,
      theme,
      notes,
      history,
      changedCells: [...changedCells],
      cleaningApplied,
      selectedColumn,
      missingMethod,
      customFill,
      convertType,
      standardizeMethod,
      minValue,
      maxValue,
      gaugeNumerator,
      gaugeDenominator,
      barCategory,
      barValue,
      pieCategory,
      pieValue,
      lineX,
      lineY,
      diagramType,
      diagramCategory,
      diagramValue,
      diagramX,
      diagramY,
      diagramColor,
    };
    const timer = window.setTimeout(() => {
      saveWorkspaceSnapshot(snapshot).catch(() => {
        console.warn('Workspace snapshot could not be saved.');
      });
    }, 450);

    return () => window.clearTimeout(timer);
  }, [
    activePage,
    activeView,
    barCategory,
    barValue,
    changedCells,
    cleanRows,
    cleaningApplied,
    convertType,
    customFill,
    diagramCategory,
    diagramColor,
    diagramType,
    diagramValue,
    diagramX,
    diagramY,
    fileName,
    gaugeDenominator,
    gaugeNumerator,
    history,
    lineX,
    lineY,
    maxValue,
    minValue,
    missingMethod,
    notes,
    originalRows,
    pieCategory,
    pieValue,
    selectedColumn,
    snapshotReady,
    standardizeMethod,
    theme,
  ]);

  function updateDataset(nextRows, message, cellChanges = new Set()) {
    setCleanRows(nextRows);
    setChangedCells(cellChanges);
    setCleaningApplied(true);
    setHistory((items) => [message, ...items].slice(0, 8));
    showNotice(message);
  }

  function showNotice(message) {
    setNotice(message);
    window.setTimeout(() => setNotice(''), 2600);
  }

  async function parseUploadedFile(file) {
    if (!file) return;
    let rows = [];
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      rows = normalizeRows(XLSX.utils.sheet_to_json(sheet, { defval: '' }));
    } catch {
      showNotice('Could not read that file. Please upload a valid CSV or Excel file.');
      return;
    }
    if (!rows.length) {
      showNotice('The uploaded file has no readable rows.');
      return;
    }
    setFileName(file.name);
    setOriginalRows(rows);
    setCleanRows(rows);
    setChangedCells(new Set());
    setCleaningApplied(false);
    setHistory([`Uploaded ${file.name}`]);
    showNotice(`Uploaded ${file.name}`);
    const nextProfile = buildProfile(rows);
    const nextTextColumns = getTextColumns(nextProfile);
    const nextNumericColumns = getNumericColumns(nextProfile);
    setSelectedColumn(nextProfile.columnsList[0] || '');
    setGaugeDenominator(nextNumericColumns[0] || '');
    setGaugeNumerator(nextNumericColumns[1] || nextNumericColumns[0] || '');
    setBarCategory(nextTextColumns[0] || nextProfile.columnsList[0] || '');
    setPieCategory(nextTextColumns[1] || nextTextColumns[0] || nextProfile.columnsList[0] || '');
    setBarValue(nextNumericColumns[0] || nextProfile.columnsList[0] || '');
    setPieValue(nextNumericColumns[1] || nextNumericColumns[0] || nextProfile.columnsList[0] || '');
    setLineX(nextProfile.columnProfiles.find((column) => column.type === 'date')?.column || nextProfile.columnsList[0] || '');
    setLineY(nextNumericColumns[0] || nextProfile.columnsList[0] || '');
    setDiagramCategory(nextTextColumns[0] || nextProfile.columnsList[0] || '');
    setDiagramValue(nextNumericColumns[0] || nextProfile.columnsList[0] || '');
    setDiagramX(nextNumericColumns[1] || nextNumericColumns[0] || nextProfile.columnsList[0] || '');
    setDiagramY(nextNumericColumns[0] || nextProfile.columnsList[0] || '');
    setDiagramColor(nextTextColumns[0] || nextProfile.columnsList[0] || '');
  }

  async function handleFile(event) {
    await parseUploadedFile(event.target.files?.[0]);
    event.target.value = '';
  }

  async function handleDrop(event) {
    event.preventDefault();
    setIsDragging(false);
    await parseUploadedFile(event.dataTransfer.files?.[0]);
  }

  async function clearSavedWorkspace() {
    setOriginalRows([]);
    setCleanRows([]);
    setFileName('No file uploaded');
    setActiveView('cleaned');
    setActivePage('upload');
    setTheme(defaultTheme);
    setNotes(['Defense note: explain why each cleaning method was applied and compare the before/after metrics.']);
    setHistory(['Waiting for uploaded dataset']);
    setChangedCells(new Set());
    setCleaningApplied(false);
    setSelectedColumn('');
    setMissingMethod('mean');
    setCustomFill('Unknown');
    setConvertType('number');
    setStandardizeMethod('trim');
    setMinValue('0');
    setMaxValue('');
    setGaugeNumerator('');
    setGaugeDenominator('');
    setBarCategory('');
    setBarValue('');
    setPieCategory('');
    setPieValue('');
    setLineX('');
    setLineY('');
    setDiagramType('scatter');
    setDiagramCategory('');
    setDiagramValue('');
    setDiagramX('');
    setDiagramY('');
    setDiagramColor('');
    await clearWorkspaceSnapshot().catch(() => undefined);
    showNotice('Saved workspace cleared');
  }

  function resetCleaning() {
    setCleanRows(originalRows);
    setChangedCells(new Set());
    setCleaningApplied(false);
    setHistory((items) => ['Reset cleaned data to original upload', ...items].slice(0, 8));
  }

  function autoCleanDataset() {
    if (!hasData) return;
    const profile = buildProfile(cleanRows);
    const numericColumnsMap = new Set(getNumericColumns(profile));
    const changes = new Set();
    const seen = new Set();
    let removedInvalid = 0;
    let removedDuplicates = 0;

    const replacements = Object.fromEntries(
      profile.columnProfiles.map((columnProfile) => {
        const values = cleanRows.map((row) => row[columnProfile.column]).filter((value) => !isMissing(value));
        const numericValues = values.map(toNumber).filter((value) => value !== null);
        if (columnProfile.type === 'number') {
          const average = numericValues.length ? numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length : 0;
          return [columnProfile.column, Number(average.toFixed(2))];
        }
        return [columnProfile.column, mode(values) || 'Unknown'];
      }),
    );

    const rows = [];
    cleanRows.forEach((row, rowIndex) => {
      const nextRow = { ...row };
      let invalid = false;

      columns.forEach((column) => {
        const current = nextRow[column];
        let next = current;
        if (isMissing(next)) next = replacements[column];
        if (typeof next === 'string') next = titleCase(next.trim());
        if (numericColumnsMap.has(column)) {
          const numeric = toNumber(next);
          next = numeric === null ? next : Number(numeric.toFixed(2));
          if (numeric !== null && numeric < 0) invalid = true;
        }
        if (String(next ?? '') !== String(current ?? '')) changes.add(`${rows.length}-${column}`);
        nextRow[column] = next;
      });

      if (invalid) {
        removedInvalid += 1;
        return;
      }

      const fingerprint = JSON.stringify(columns.map((column) => nextRow[column] ?? ''));
      if (seen.has(fingerprint)) {
        removedDuplicates += 1;
        return;
      }
      seen.add(fingerprint);
      rows.push(nextRow);
    });

    updateDataset(
      rows,
      `Auto cleaned dataset: ${changes.size} cell change(s), ${removedDuplicates} duplicate row(s), ${removedInvalid} invalid row(s)`,
      changes,
    );
  }

  function fillMissing() {
    if (!hasData || !selectedColumn) return;
    const profile = cleanProfile.columnProfiles.find((column) => column.column === selectedColumn);
    const numericValues = cleanRows.map((row) => toNumber(row[selectedColumn])).filter((value) => value !== null);
    const replacement = {
      mean: numericValues.length ? numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length : 0,
      median: median(numericValues),
      mode: mode(cleanRows.map((row) => row[selectedColumn]).filter((value) => !isMissing(value))),
      zero: 0,
      blank: '',
      custom: customFill,
    }[missingMethod];

    if (missingMethod === 'remove-row') {
      updateDataset(cleanRows.filter((row) => !isMissing(row[selectedColumn])), `Removed rows missing ${selectedColumn}`);
      return;
    }

    const changes = new Set();
    const rows = cleanRows.map((row, rowIndex) => {
      if (!isMissing(row[selectedColumn])) return row;
      changes.add(`${rowIndex}-${selectedColumn}`);
      return { ...row, [selectedColumn]: profile?.type === 'number' ? Number(Number(replacement).toFixed(2)) : replacement };
    });
    updateDataset(rows, `Filled missing ${selectedColumn} values with ${missingMethod}`, changes);
  }

  function removeDuplicates() {
    if (!hasData) return;
    const seen = new Set();
    const rows = cleanRows.filter((row) => {
      const fingerprint = JSON.stringify(columns.map((column) => row[column] ?? ''));
      if (seen.has(fingerprint)) return false;
      seen.add(fingerprint);
      return true;
    });
    updateDataset(rows, `Removed ${cleanRows.length - rows.length} duplicate row(s)`);
  }

  function convertColumn() {
    if (!hasData || !selectedColumn) return;
    const changes = new Set();
    const rows = cleanRows.map((row, rowIndex) => {
      const current = row[selectedColumn];
      if (isMissing(current)) return row;
      let next = current;
      if (convertType === 'number') next = toNumber(current) ?? current;
      if (convertType === 'text') next = String(current);
      if (convertType === 'date') {
        const parsed = new Date(current);
        next = Number.isNaN(parsed.getTime()) ? current : parsed.toISOString().slice(0, 10);
      }
      if (convertType === 'boolean') next = ['true', 'yes', '1', 'active'].includes(String(current).toLowerCase());
      if (String(next) !== String(current)) changes.add(`${rowIndex}-${selectedColumn}`);
      return { ...row, [selectedColumn]: next };
    });
    updateDataset(rows, `Converted ${selectedColumn} to ${convertType}`, changes);
  }

  function standardizeColumn() {
    if (!hasData || !selectedColumn) return;
    const changes = new Set();
    const rows = cleanRows.map((row, rowIndex) => {
      const current = row[selectedColumn];
      if (isMissing(current)) return row;
      let next = String(current);
      if (standardizeMethod === 'trim') next = next.trim();
      if (standardizeMethod === 'upper') next = next.toUpperCase();
      if (standardizeMethod === 'lower') next = next.toLowerCase();
      if (standardizeMethod === 'title') next = titleCase(next.trim());
      if (next !== String(current)) changes.add(`${rowIndex}-${selectedColumn}`);
      return { ...row, [selectedColumn]: next };
    });
    updateDataset(rows, `Standardized ${selectedColumn} using ${standardizeMethod}`, changes);
  }

  function filterInvalid() {
    if (!hasData || !selectedColumn) return;
    const min = minValue === '' ? -Infinity : Number(minValue);
    const max = maxValue === '' ? Infinity : Number(maxValue);
    const rows = cleanRows.filter((row) => {
      const value = toNumber(row[selectedColumn]);
      return value === null || (value >= min && value <= max);
    });
    updateDataset(rows, `Filtered invalid ${selectedColumn} values`);
  }

  function downloadCleanCsv() {
    if (!hasData) return;
    const sheet = XLSX.utils.json_to_sheet(cleanRows);
    const csv = XLSX.utils.sheet_to_csv(sheet);
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = `cleaned-${fileName.replace(/\.[^.]+$/, '')}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }

  const tableRows = activePage === 'upload' ? originalRows : activeView === 'original' ? originalRows : cleanRows;
  const comparisonSummary = cleaningApplied
    ? `${changedCells.size} changed cell(s), ${Math.max(originalRows.length - cleanRows.length, 0)} removed row(s).`
    : hasData
      ? 'Original and Cleaned are the same right after import. Apply a cleaning method or Auto Clean to create differences.'
      : 'Upload a dataset to compare original and cleaned versions.';
  const pageMeta = {
    upload: ['Upload Dataset', 'Import a CSV or Excel file, then preview the raw table before cleaning.'],
    profile: ['Data Profile', 'Review row counts, column counts, detected data types, missing values, and statistics.'],
    clean: ['Cleaning Workbench', 'Apply cleaning methods and compare the original dataset with the cleaned version.'],
    analyze: ['Insights', 'Read generated interpretations, frequent values, and cleaning history.'],
    visualize: ['Dashboard Visualization', 'Choose variables and chart types to explore patterns in your cleaned data.'],
  };
  const navigationTabs = [
    ['upload', 'Upload', Upload],
    ['profile', 'Profile', Table2],
    ['clean', 'Clean', Wand2],
    ['analyze', 'Analyze', Sparkles],
    ['visualize', 'Visualize', BarChart3],
  ];

  return (
    <main className="app" style={{ '--primary': theme.primary, '--secondary': theme.secondary, '--accent': theme.accent, '--ink': theme.ink, '--surface': theme.surface }}>
      {notice && <div className="toast">{notice}</div>}
      <header className="topbar">
        <div className="brand">
          <span className="brand-mark"><FileSpreadsheet size={22} /></span>
          <div>
            <strong>BAT403</strong>
            <small>Data Cleaning Explorer</small>
          </div>
        </div>

        <nav className="top-tabs">
          {navigationTabs.map(([page, label, Icon]) => (
            <button type="button" className={activePage === page ? 'active' : ''} onClick={() => setActivePage(page)} key={page}>
              <Icon size={17} />
              {label}
            </button>
          ))}
        </nav>
      </header>

      <section className="workspace">
        {activePage !== 'upload' && (
        <div className="page-heading">
          <div>
            <p className="eyebrow">{'Upload -> Profile -> Clean -> Analyze -> Visualize'}</p>
            <h1>{pageMeta[activePage][0]}</h1>
            <p>{pageMeta[activePage][1]}</p>
          </div>
        </div>
        )}

        {activePage === 'upload' && (
        <header className="hero">
          <div>
            <p className="eyebrow">{'Upload -> Profile -> Clean -> Analyze -> Visualize'}</p>
            <h1>{theme.title}</h1>
            <p>{theme.subtitle}</p>
          </div>
          <div className="hero-actions" id="upload">
            <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} hidden />
            <div
              className={`upload-dropzone ${isDragging ? 'dragging' : ''}`}
              onDragOver={(event) => {
                event.preventDefault();
                setIsDragging(true);
              }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={handleDrop}
            >
              <Upload size={26} />
              <strong>Drop CSV or Excel here</strong>
              <span>or browse from your device</span>
              <button className="primary-btn" onClick={() => fileRef.current?.click()}><Upload size={18} /> Choose File</button>
            </div>
            <button onClick={downloadCleanCsv}><Download size={18} /> Export Clean CSV</button>
            <button onClick={clearSavedWorkspace}>Clear Saved Workspace</button>
            {hasData && <button onClick={() => setActivePage('profile')}>Continue to Profile</button>}
          </div>
        </header>
        )}

        {activePage === 'profile' && (
        <>
        <section className="metrics" id="profile">
          <Metric label="Rows" value={cleanProfile.rows} detail={`Original: ${originalProfile.rows}`} />
          <Metric label="Columns" value={cleanProfile.columns} detail="Detected fields" />
          <Metric label="Missing" value={cleanProfile.missing} detail={`Before: ${originalProfile.missing}`} />
          <Metric label="Duplicates" value={cleanProfile.duplicateRows} detail={`Before: ${originalProfile.duplicateRows}`} />
        </section>
        <div className="page-actions">
          <button onClick={() => setActivePage('upload')}>Back to Upload</button>
          <button className="primary-action" onClick={() => setActivePage('clean')} disabled={!hasData}>Continue to Clean</button>
        </div>
        </>
        )}

        {activePage === 'visualize' && (
        <section className="visualize-layout">
        <section className="tool-panel">
          <h3><PaintBucket size={17} /> Design Studio</h3>
          <label>
            Dashboard title
            <input value={theme.title} onChange={(event) => setTheme({ ...theme, title: event.target.value })} />
          </label>
          <label>
            Subtitle
            <textarea value={theme.subtitle} rows="3" onChange={(event) => setTheme({ ...theme, subtitle: event.target.value })} />
          </label>
          <div className="color-grid">
            {['primary', 'secondary', 'accent', 'ink'].map((key) => (
              <label key={key}>
                {key}
                <input type="color" value={theme[key]} onChange={(event) => setTheme({ ...theme, [key]: event.target.value })} />
              </label>
            ))}
          </div>
          <label>
            Add text note
            <div className="inline-action">
              <input value={newNote} onChange={(event) => setNewNote(event.target.value)} placeholder="Type dashboard note" />
              <button type="button" onClick={() => {
                if (!newNote.trim()) return;
                setNotes((items) => [newNote.trim(), ...items]);
                setNewNote('');
              }} aria-label="Add note"><Plus size={17} /></button>
            </div>
          </label>
        </section>

        <section className="diagram-builder">
          <div className="section-title">
            <div>
              <h2>Diagram Builder</h2>
              <p>{diagramType === 'scatter' ? `${diagramX} vs ${diagramY}` : `${diagramValue} by ${diagramCategory}`}</p>
            </div>
            <div className="diagram-mode" aria-label="Choose diagram type">
              {[
                ['bar', BarChart3, 'Bar'],
                ['line', LineChart, 'Line'],
                ['pie', PieChart, 'Pie'],
                ['donut', PieChart, 'Donut'],
                ['scatter', Sparkles, 'Scatter'],
              ].map(([type, Icon, label]) => (
                <button key={type} className={diagramType === type ? 'active' : ''} onClick={() => setDiagramType(type)}>
                  <Icon size={16} />
                  {label}
                </button>
              ))}
            </div>
          </div>
          <div className="diagram-layout">
            <div className="diagram-controls">
              {diagramType === 'scatter' ? (
                <>
                  <label>
                    X axis
                    <select value={diagramX} onChange={(event) => setDiagramX(event.target.value)}>
                      {(numericColumns.length ? numericColumns : columns).map((column) => <option key={column}>{column}</option>)}
                    </select>
                  </label>
                  <label>
                    Y axis
                    <select value={diagramY} onChange={(event) => setDiagramY(event.target.value)}>
                      {(numericColumns.length ? numericColumns : columns).map((column) => <option key={column}>{column}</option>)}
                    </select>
                  </label>
                  <label>
                    Color / class
                    <select value={diagramColor} onChange={(event) => setDiagramColor(event.target.value)}>
                      {(textColumns.length ? textColumns : columns).map((column) => <option key={column}>{column}</option>)}
                    </select>
                  </label>
                </>
              ) : (
                <>
                  <label>
                    Category
                    <select value={diagramCategory} onChange={(event) => setDiagramCategory(event.target.value)}>
                      {columns.map((column) => <option key={column}>{column}</option>)}
                    </select>
                  </label>
                  <label>
                    Value
                    <select value={diagramValue} onChange={(event) => setDiagramValue(event.target.value)}>
                      {(numericColumns.length ? numericColumns : columns).map((column) => <option key={column}>{column}</option>)}
                    </select>
                  </label>
                </>
              )}
            </div>
            <div className="diagram-stage">
              <DiagramRenderer
                type={diagramType}
                data={dashboard.diagram}
                rows={cleanRows}
                xColumn={diagramX}
                yColumn={diagramY}
                colorColumn={diagramColor}
                theme={theme}
              />
            </div>
          </div>
        </section>
        </section>
        )}

        {activePage === 'visualize' && (
        <section className="dashboard-grid" id="visualize">
          <Card title="Ratio Gauge" menu="Gauge">
            <ChartControls
              firstLabel="Numerator"
              firstValue={gaugeNumerator}
              firstOptions={numericColumns}
              onFirst={setGaugeNumerator}
              secondLabel="Denominator"
              secondValue={gaugeDenominator}
              secondOptions={numericColumns}
              onSecond={setGaugeDenominator}
            />
            <GaugeCard
              ratio={dashboard.ratioGauge}
              theme={theme}
              numeratorLabel={gaugeNumerator}
              denominatorLabel={gaugeDenominator}
            />
          </Card>
          <Card title="Engagement Analytics" menu="Platform">
            <DonutChart data={dashboard.pie} theme={theme} />
          </Card>
          <Card title="Salary Distribution" wide={false}>
            <ChartControls
              firstLabel="Category"
              firstValue={barCategory}
              firstOptions={columns}
              onFirst={setBarCategory}
              secondLabel="Value"
              secondValue={barValue}
              secondOptions={numericColumns.length ? numericColumns : columns}
              onSecond={setBarValue}
            />
            <BarChartCard data={dashboard.bar} theme={theme} />
          </Card>
          <Card title="Revenue Analytics" menu="Trend">
            <ChartControls
              firstLabel="X axis"
              firstValue={lineX}
              firstOptions={columns}
              onFirst={setLineX}
              secondLabel="Y axis"
              secondValue={lineY}
              secondOptions={numericColumns.length ? numericColumns : columns}
              onSecond={setLineY}
            />
            <LineChartCard data={dashboard.line} theme={theme} />
          </Card>
        </section>
        )}

        {activePage === 'clean' && (
        <>
        <section className="workbench" id="clean">
          <div className="cleaner">
            <div className="section-title">
              <h2>Cleaning Workbench</h2>
              <div className="header-actions">
                <button onClick={autoCleanDataset} disabled={!hasData}><Sparkles size={17} /> Auto Clean</button>
                <button onClick={resetCleaning}><RotateCcw size={17} /> Reset</button>
              </div>
            </div>
            <div className="clean-grid">
              <label>
                Target column
                <select value={selectedColumn} onChange={(event) => setSelectedColumn(event.target.value)} disabled={!hasData}>
                  {!hasData && <option>Upload data first</option>}
                  {columns.map((column) => <option key={column}>{column}</option>)}
                </select>
              </label>
              <label>
                Missing values
                <select value={missingMethod} onChange={(event) => setMissingMethod(event.target.value)}>
                  <option value="mean">Fill mean</option>
                  <option value="median">Fill median</option>
                  <option value="mode">Fill mode</option>
                  <option value="zero">Fill zero</option>
                  <option value="blank">Fill blank</option>
                  <option value="custom">Fill custom</option>
                  <option value="remove-row">Remove rows</option>
                </select>
              </label>
              <label>
                Custom fill
                <input value={customFill} onChange={(event) => setCustomFill(event.target.value)} />
              </label>
              <button onClick={fillMissing} disabled={!hasData}><Wand2 size={17} /> Apply Missing Fix</button>

              <label>
                Convert type
                <select value={convertType} onChange={(event) => setConvertType(event.target.value)}>
                  <option value="number">Number</option>
                  <option value="text">Text</option>
                  <option value="date">Date</option>
                  <option value="boolean">Boolean</option>
                </select>
              </label>
              <button onClick={convertColumn} disabled={!hasData}><CheckCircle2 size={17} /> Convert</button>

              <label>
                Standardize
                <select value={standardizeMethod} onChange={(event) => setStandardizeMethod(event.target.value)}>
                  <option value="trim">Trim spaces</option>
                  <option value="upper">Uppercase</option>
                  <option value="lower">Lowercase</option>
                  <option value="title">Title case</option>
                </select>
              </label>
              <button onClick={standardizeColumn} disabled={!hasData}><Brush size={17} /> Standardize</button>

              <label>
                Min
                <input value={minValue} onChange={(event) => setMinValue(event.target.value)} />
              </label>
              <label>
                Max
                <input value={maxValue} onChange={(event) => setMaxValue(event.target.value)} placeholder="Optional" />
              </label>
              <button onClick={filterInvalid} disabled={!hasData}><Filter size={17} /> Filter Invalid</button>
              <button onClick={removeDuplicates} disabled={!hasData}><Table2 size={17} /> Remove Duplicates</button>
            </div>
          </div>

          <div className="insights" id="analyze">
            <h2>Insights</h2>
            {insights.map((insight) => <p key={insight}>{insight}</p>)}
            <h3>Cleaning Log</h3>
            <ul>
              {history.map((item, index) => <li key={`${item}-${index}`}>{item}</li>)}
            </ul>
          </div>
        </section>
        <div className="page-actions">
          <button onClick={() => setActivePage('profile')}>Back to Profile</button>
          <button className="primary-action" onClick={() => setActivePage('analyze')} disabled={!hasData}>Continue to Analyze</button>
        </div>
        </>
        )}

        {(activePage === 'upload' || activePage === 'clean') && (
        <section className="data-area">
          <div className="section-title">
            <div>
              <h2>{activePage === 'upload' ? 'Uploaded Dataset Preview' : 'Original vs Cleaned Dataset'}</h2>
              <p>
                {activePage === 'upload'
                  ? `${fileName} - raw imported data only`
                  : `${fileName} - highlighted cells show changed values`}
              </p>
            </div>
            {activePage === 'clean' && (
            <div className="segmented">
              <button className={activeView === 'original' ? 'active' : ''} onClick={() => setActiveView('original')}>Original</button>
              <button className={activeView === 'cleaned' ? 'active' : ''} onClick={() => setActiveView('cleaned')}>Cleaned</button>
            </div>
            )}
          </div>
          {activePage === 'clean' && (
          <div className={`comparison-summary ${cleaningApplied ? 'active' : ''}`}>
            {comparisonSummary}
          </div>
          )}
          <div className="table-wrap">
            {hasData ? (
              <table>
                <thead>
                  <tr>
                    {columns.map((column) => <th key={column}>{column}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {tableRows.slice(0, 80).map((row, rowIndex) => (
                    <tr key={`${rowIndex}-${JSON.stringify(row).slice(0, 15)}`}>
                      {columns.map((column) => (
                        <td key={column} className={activePage === 'clean' && activeView === 'cleaned' && changedCells.has(`${rowIndex}-${column}`) ? 'changed' : ''}>
                          {String(row[column] ?? '')}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <EmptyChart message="Upload a CSV or Excel file to preview your dataset." />
            )}
          </div>
        </section>
        )}

        {activePage === 'profile' && (
        <section className="profile-table">
          <div className="section-title">
            <h2>Data Types and Statistics</h2>
          </div>
          <div className="table-wrap compact">
            {hasData ? (
              <table>
                <thead>
                  <tr>
                    <th>Column</th>
                    <th>Type</th>
                    <th>Missing</th>
                    <th>Unique</th>
                    <th>Most Frequent</th>
                    <th>Mean</th>
                    <th>Min</th>
                    <th>Max</th>
                  </tr>
                </thead>
                <tbody>
                  {cleanProfile.columnProfiles.map((profile) => (
                    <tr key={profile.column}>
                      <td>{profile.column}</td>
                      <td><span className="type-pill">{profile.type}</span></td>
                      <td>{profile.missing}</td>
                      <td>{profile.unique}</td>
                      <td>{String(profile.frequent)}</td>
                      <td>{profile.mean === null ? '-' : formatNumber(profile.mean)}</td>
                      <td>{profile.min === null ? '-' : formatNumber(profile.min)}</td>
                      <td>{profile.max === null ? '-' : formatNumber(profile.max)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <EmptyChart message="Profiling results will appear after upload." />
            )}
          </div>
        </section>
        )}

        {activePage === 'analyze' && (
        <section className="analyze-page">
          <div className="insights" id="analyze">
            <h2>Insights</h2>
            {insights.map((insight) => <p key={insight}>{insight}</p>)}
            <h3>Cleaning Log</h3>
            <ul>
              {history.map((item, index) => <li key={`${item}-${index}`}>{item}</li>)}
            </ul>
          </div>
          <section className="notes">
          {notes.map((note, index) => (
            <article key={`${note}-${index}`}>
              <strong>Text Block {notes.length - index}</strong>
              <p>{note}</p>
            </article>
          ))}
          </section>
          <div className="page-actions">
            <button onClick={() => setActivePage('clean')}>Back to Clean</button>
            <button className="primary-action" onClick={() => setActivePage('visualize')} disabled={!hasData}>Continue to Visualize</button>
          </div>
        </section>
        )}
      </section>
    </main>
  );
}

function Metric({ label, value, detail }) {
  return (
    <article className="metric-card">
      <span>{label}</span>
      <strong>{formatNumber(value)}</strong>
      <small>{detail}</small>
    </article>
  );
}

function WorkflowStrip({ pages, activePage, hasData, onNavigate }) {
  const activeIndex = pages.findIndex(([page]) => page === activePage);
  return (
    <div className="workflow-strip">
      {pages.map(([page, label], index) => {
        const isActive = page === activePage;
        const isDone = hasData && index < activeIndex;
        return (
          <button
            type="button"
            key={page}
            className={`${isActive ? 'active' : ''} ${isDone ? 'done' : ''}`}
            onClick={() => onNavigate(page)}
          >
            <span>{index + 1}</span>
            {label}
          </button>
        );
      })}
    </div>
  );
}

function Card({ title, menu, children }) {
  return (
    <article className="dash-card">
      <header>
        <h2>{title}</h2>
        {menu && <span className="mini-menu">{menu}</span>}
      </header>
      {children}
    </article>
  );
}

function ChartControls({ firstLabel, firstValue, firstOptions, onFirst, secondLabel, secondValue, secondOptions, onSecond }) {
  return (
    <div className="chart-controls">
      <label>
        {firstLabel}
        <select value={firstValue} onChange={(event) => onFirst(event.target.value)} disabled={!firstOptions.length}>
          {!firstOptions.length && <option value="">Upload data first</option>}
          {firstOptions.map((option) => <option key={option} value={option}>{option}</option>)}
        </select>
      </label>
      <label>
        {secondLabel}
        <select value={secondValue} onChange={(event) => onSecond(event.target.value)} disabled={!secondOptions.length}>
          {!secondOptions.length && <option value="">Upload data first</option>}
          {secondOptions.map((option) => <option key={option} value={option}>{option}</option>)}
        </select>
      </label>
    </div>
  );
}

export default App;
