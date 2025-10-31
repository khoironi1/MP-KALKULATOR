import { useState } from 'react';
import * as XLSX from 'xlsx';
import InputField from './components/InputField.jsx';
import Card from './components/Card.jsx';
import ResultRow from './components/ResultRow.jsx';

// Helper: safe numeric parse
const n = (v) => {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/[^0-9.-]+/g, '');
  const num = Number(s);
  return Number.isFinite(num) ? num : 0;
};

const round2 = (x) => (x === null || x === undefined ? null : Math.round((x + Number.EPSILON) * 100) / 100);

export default function App() {
  const [rows, setRows] = useState([]);
  const [useExcelValues, setUseExcelValues] = useState(true);

  const format = (v) => {
    if (v === null || v === undefined || v === '') return '-';
    if (typeof v === 'number') return new Intl.NumberFormat('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v);
    return v;
  };

  // Core JS calc matching canonical formulas
  const calcRow = (r) => {
    const HPP = n(r.HPP || r.hpp || r['Harga Pokok'] || r['HPP']);
    const QTY = Math.max(0, Math.round(n(r.QTY || r.qty || r['QTY'] || r['Quantity'])));
    const MARKUP = n(r.MARKUP || r.markup || r['MARKUP'] || r['Markup']);
    const DISKON = n(r.DISKON || r.diskon || r['DISKON'] || r['Diskon']);
    const ADS = n(r.ADS || r.ads || r['ADS'] || r['Ads']);
    const BIAYA_PROSES = n(r.BIAYA_PROSES || r.biaya_proses || r['BIAYA_PROSES'] || r['Biaya Proses'] || r['biaya proses']);
    const TARGET_ROAS = n(r.TARGET_ROAS || r.target_roas || r['TARGET ROAS'] || r['Target ROAS']);

    const HARGA_TAYANG = HPP * (1 + MARKUP);
    const HARGA_NET = HARGA_TAYANG * (1 - DISKON);
    const REVENUE = HARGA_NET * QTY;
    const GROSS_PROFIT = (HARGA_NET - HPP) * QTY;
    const NET_PROFIT = GROSS_PROFIT - ADS - BIAYA_PROSES * QTY;
    const ROAS = ADS && ADS !== 0 ? REVENUE / ADS : null;
    const denom = ADS + BIAYA_PROSES * QTY;
    const BATAS_ROAS_MAX = denom && denom !== 0 ? REVENUE / denom : null;

    return {
      ...r,
      HPP, QTY, MARKUP, DISKON, ADS, BIAYA_PROSES, TARGET_ROAS,
      HARGA_TAYANG: round2(HARGA_TAYANG),
      HARGA_NET: round2(HARGA_NET),
      REVENUE: round2(REVENUE),
      GROSS_PROFIT: round2(GROSS_PROFIT),
      NET_PROFIT: round2(NET_PROFIT),
      ROAS: ROAS === null ? null : round2(ROAS),
      BATAS_ROAS_MAX: BATAS_ROAS_MAX === null ? null : round2(BATAS_ROAS_MAX),
    };
  };

  const importXLSX = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const first = wb.SheetNames[0];
      const ws = wb.Sheets[first];
      // Get values (sheet_to_json reads computed values by default)
      const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
      const mapped = json.map((row) => {
        const r = {
          SKU: row.SKU || row.sku || row['SKU'] || '',
          HPP: row.HPP || row.hpp || row['HPP'] || row['Harga Pokok'] || row['HARGA POKOK'] || row['COST'] || '',
          QTY: row.QTY || row.qty || row['QTY'] || row['Quantity'] || 1,
          MARKUP: row.MARKUP || row.markup || row['MARKUP'] || row['Markup'] || 0,
          DISKON: row.DISKON || row.diskon || row['DISKON'] || row['Diskon'] || 0,
          ADS: row.ADS || row.ads || row['ADS'] || 0,
          BIAYA_PROSES: row.BIAYA_PROSES || row.biaya_proses || row['BIAYA_PROSES'] || row['Biaya Proses'] || 0,
          TARGET_ROAS: row.TARGET_ROAS || row.target_roas || row['TARGET ROAS'] || 0,
          HARGA_TAYANG_SHEET: row['HARGA TAYANG'] || row['Harga Tayang'] || row.HARGA_TAYANG || row.harga_tayang || row['HARGA_TAYANG'] || '',
          HARGA_NET_SHEET: row['HARGA NET'] || row['Harga Net'] || row.HARGA_NET || row.harga_net || '',
          GROSS_PROFIT_SHEET: row['GROSS PROFIT'] || row['Gross Profit'] || row.GROSS_PROFIT || row.gross_profit || '',
          NET_PROFIT_SHEET: row['NET PROFIT'] || row['Net Profit'] || row.NET_PROFIT || row.net_profit || '',
          ROAS_SHEET: row['ROAS'] || row.Roas || row['ROAS'] || '',
          BATAS_ROAS_MAX_SHEET: row['BATAS ROAS MAX'] || row['Batas ROAS Max'] || row.BATAS_ROAS_MAX || '',
        };
        const calc = calcRow(r);
        if (useExcelValues) {
          const final = { ...calc };
          if (r.HARGA_TAYANG_SHEET !== '' && r.HARGA_TAYANG_SHEET !== null) final.HARGA_TAYANG = round2(n(r.HARGA_TAYANG_SHEET));
          if (r.HARGA_NET_SHEET !== '' && r.HARGA_NET_SHEET !== null) final.HARGA_NET = round2(n(r.HARGA_NET_SHEET));
          if (r.GROSS_PROFIT_SHEET !== '' && r.GROSS_PROFIT_SHEET !== null) final.GROSS_PROFIT = round2(n(r.GROSS_PROFIT_SHEET));
          if (r.NET_PROFIT_SHEET !== '' && r.NET_PROFIT_SHEET !== null) final.NET_PROFIT = round2(n(r.NET_PROFIT_SHEET));
          if (r.ROAS_SHEET !== '' && r.ROAS_SHEET !== null) final.ROAS = round2(n(r.ROAS_SHEET));
          if (r.BATAS_ROAS_MAX_SHEET !== '' && r.BATAS_ROAS_MAX_SHEET !== null) final.BATAS_ROAS_MAX = round2(n(r.BATAS_ROAS_MAX_SHEET));
          return final;
        } else {
          return calc;
        }
      });
      setRows(mapped);
    };
    reader.readAsArrayBuffer(file);
  };

  const exportXLSX = () => {
    const out = rows.map(r => ({
      SKU: r.SKU || '',
      HPP: r.HPP,
      QTY: r.QTY,
      MARKUP: r.MARKUP,
      HARGA_TAYANG: r.HARGA_TAYANG,
      DISKON: r.DISKON,
      HARGA_NET: r.HARGA_NET,
      REVENUE: r.REVENUE,
      GROSS_PROFIT: r.GROSS_PROFIT,
      ADS: r.ADS,
      BIAYA_PROSES: r.BIAYA_PROSES,
      NET_PROFIT: r.NET_PROFIT,
      ROAS: r.ROAS,
      BATAS_ROAS_MAX: r.BATAS_ROAS_MAX,
      TARGET_ROAS: r.TARGET_ROAS,
    }));
    const ws = XLSX.utils.json_to_sheet(out);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'ProfitData');
    XLSX.writeFile(wb, 'astrada_profit_export.xlsx');
  };

  const totals = rows.reduce((acc, r) => {
    acc.REVENUE += n(r.REVENUE);
    acc.GROSS += n(r.GROSS_PROFIT);
    acc.NET += n(r.NET_PROFIT);
    acc.ADS += n(r.ADS);
    return acc;
  }, { REVENUE:0, GROSS:0, NET:0, ADS:0 });

  const avgRoas = rows.length ? (rows.reduce((a,b)=> a + (b.ROAS||0),0) / rows.length) : 0;

  return (
    <div className="p-6 max-w-7xl mx-auto font-sans">
      <header className="flex items-center justify-between mb-6">
        <h1 className="text-2xl font-bold text-slate-800">Astrada — Profit Calculator (FINAL)</h1>
        <div className="flex gap-2 items-center">
          <label className="inline-flex items-center gap-2 bg-blue-50 border border-blue-200 px-3 py-2 rounded cursor-pointer">
            <input type="file" accept=".xlsx, .xls" onChange={importXLSX} className="hidden" />
            <span className="text-sm text-blue-700">Import XLSX</span>
          </label>
          <button onClick={exportXLSX} className="px-3 py-2 rounded bg-blue-600 text-white">Export XLSX</button>
          <label className="inline-flex items-center gap-2 ml-2 text-sm">
            <input type="checkbox" checked={useExcelValues} onChange={(e)=>setUseExcelValues(e.target.checked)} />
            <span>Use Excel computed values when present</span>
          </label>
        </div>
      </header>

      <div className="overflow-x-auto border rounded shadow-sm">
        <table className="min-w-full table-auto border-collapse">
          <thead>
            <tr className="bg-blue-600 text-white">
              <th className="p-2 border">SKU</th>
              <th className="p-2 border">HPP</th>
              <th className="p-2 border">QTY</th>
              <th className="p-2 border">Markup</th>
              <th className="p-2 border">Harga Tayang</th>
              <th className="p-2 border">Diskon</th>
              <th className="p-2 border">Harga Net</th>
              <th className="p-2 border">Revenue</th>
              <th className="p-2 border">Gross Profit</th>
              <th className="p-2 border">Ads</th>
              <th className="p-2 border">Biaya Proses</th>
              <th className="p-2 border">Net Profit</th>
              <th className="p-2 border">ROAS</th>
              <th className="p-2 border">Batas ROAS Max</th>
              <th className="p-2 border">Target ROAS</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r, idx) => (
              <tr key={idx} className="odd:bg-white even:bg-blue-50">
                <td className="p-2 border">{r.SKU || '-'}</td>
                <td className="p-2 border">{r.HPP}</td>
                <td className="p-2 border">{r.QTY}</td>
                <td className="p-2 border">{r.MARKUP}</td>
                <td className="p-2 border">{r.HARGA_TAYANG !== null ? r.HARGA_TAYANG.toFixed(2) : '-'}</td>
                <td className="p-2 border">{r.DISKON}</td>
                <td className="p-2 border">{r.HARGA_NET !== null ? r.HARGA_NET.toFixed(2) : '-'}</td>
                <td className="p-2 border">{r.REVENUE !== null ? r.REVENUE.toFixed(2) : '-'}</td>
                <td className="p-2 border">{r.GROSS_PROFIT !== null ? r.GROSS_PROFIT.toFixed(2) : '-'}</td>
                <td className="p-2 border">{r.ADS}</td>
                <td className="p-2 border">{r.BIAYA_PROSES}</td>
                <td className="p-2 border">{r.NET_PROFIT !== null ? r.NET_PROFIT.toFixed(2) : '-'}</td>
                <td className="p-2 border">{r.ROAS === null ? '-' : r.ROAS.toFixed(2)}</td>
                <td className="p-2 border">{r.BATAS_ROAS_MAX === null ? '-' : r.BATAS_ROAS_MAX.toFixed(2)}</td>
                <td className="p-2 border">{r.TARGET_ROAS}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <footer className="mt-6 grid grid-cols-3 gap-4">
        <div className="p-4 border rounded bg-blue-50">
          <div className="text-sm text-blue-700">Total Revenue</div>
          <div className="text-2xl font-bold text-slate-800">{format(totals.REVENUE)}</div>
        </div>
        <div className="p-4 border rounded bg-blue-50">
          <div className="text-sm text-blue-700">Total Net Profit</div>
          <div className="text-2xl font-bold text-slate-800">{format(totals.NET)}</div>
        </div>
        <div className="p-4 border rounded bg-blue-50">
          <div className="text-sm text-blue-700">Average ROAS</div>
          <div className="text-2xl font-bold text-slate-800">{avgRoas ? avgRoas.toFixed(2) : '-'}</div>
        </div>
      </footer>

      <div className="mt-4 text-sm text-slate-600">
        <p>Note: toggle "Use Excel computed values" to prefer values exactly as in your spreadsheet. If off, app recalculates using canonical formulas.</p>
      </div>
    </div>
  );
}
