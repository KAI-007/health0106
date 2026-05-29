import { useMemo, useState } from 'react';

const API_BASE = '';
const approvalLevels = ['行办会', '行领导', '总经理室', '部门科经理', '合计'];
const yearlyRows = ['2026年度', '2025年度'];
const chartCategories = ['本部直属', '公金应用研发中心', '管理支持中心', '集团基础应用研发中心', '零售应用研发中心', '同业应用研发中心', '数据中心'];
const chartGeometry = { width: 920, height: 510, marginTop: 58, marginRight: 24, marginBottom: 198, marginLeft: 92 };
const defaultTitles = ['行办会审议立项项目-健康度', '行领导签报立项项目-健康度', '总经理室签报立项项目-健康度', '部门科经理签报立项项目'];

function createDefaultEmailData() {
  return {
    header: {
      greeting: '各位领导、同事，您好！',
      summary: '2026年4月第一周的金融科技部尚未上线开发建设类项目情况如下。',
      meta: '一、健康度情况（数据来源：科管平台，取数时间：2026/04/13 18:00:00）',
    },
    healthSummary: Object.fromEntries(approvalLevels.map((x) => [x, { 红: 0, 黄: 0, 绿: 0 }])),
    strategySummary: Object.fromEntries(yearlyRows.map((x) => [x, { 战略数: 0, 实施中需求数: 0, 红: 0, 黄: 0, 绿: 0 }])),
    chartCards: defaultTitles.map((title) => ({ title, series: { 红: [0, 0, 0, 0, 0, 0, 0], 黄: [0, 0, 0, 0, 0, 0, 0], 绿: [0, 0, 0, 3, 0, 0, 0] }, max: 4 })),
    projectSummary: Object.fromEntries(approvalLevels.map((x) => [x, { 实施中: 0, 待投产: 0, 已投产: 0, 小计: 0, '其中，超期投产': 0 }])),
    launchSummary: {
      title: '按时上线率（26年四月第1、2周）',
      rows: Object.fromEntries(approvalLevels.map((x) => [x, { 按计划上线数: 0, 实际上线数: 0, 按时上线率: '0%' }])),
    },
  };
}

function normalizeEmailData(rawData) {
  const defaults = createDefaultEmailData();
  const rawLaunch = rawData?.launchSummary;
  const normalizedLaunchRows = Object.fromEntries(
    approvalLevels.map((item) => {
      const row = rawLaunch?.rows?.[item] || rawLaunch?.[item] || {};
      return [item, {
        按计划上线数: row['按计划上线数'] ?? row['按时上线（含按计划延期、4周之内）'] ?? 0,
        实际上线数: row.实际上线数 ?? 0,
        按时上线率: row.按时上线率 ?? '0%',
      }];
    }),
  );

  return {
    ...defaults,
    ...rawData,
    launchSummary: {
      title: rawLaunch?.title || defaults.launchSummary.title,
      rows: normalizedLaunchRows,
    },
  };
}

function ChartCard({ title, series, max }) {
  const { width, height, marginTop, marginRight, marginBottom, marginLeft } = chartGeometry;
  const plotWidth = width - marginLeft - marginRight;
  const plotHeight = height - marginTop - marginBottom;
  const totals = chartCategories.map((_, index) => (series?.红?.[index] ?? 0) + (series?.黄?.[index] ?? 0) + (series?.绿?.[index] ?? 0));
  const dataMax = Math.max(max || 0, ...totals, 0);

  const getNiceStep = (value) => {
    if (value <= 5) return 1;
    const rough = value / 5;
    const magnitude = 10 ** Math.floor(Math.log10(rough));
    const normalized = rough / magnitude;
    if (normalized <= 1) return magnitude;
    if (normalized <= 2) return 2 * magnitude;
    if (normalized <= 5) return 5 * magnitude;
    return 10 * magnitude;
  };

  const tickStep = getNiceStep(dataMax || 1);
  const safeMax = Math.max(tickStep, Math.ceil((dataMax || 1) / tickStep) * tickStep);
  const ticks = Array.from({ length: Math.floor(safeMax / tickStep) + 1 }, (_, i) => safeMax - i * tickStep);
  const stepX = plotWidth / chartCategories.length;
  const barWidth = 32;
  const getY = (v) => marginTop + plotHeight - (v / safeMax) * plotHeight;

  return <div className="chart-card svg-chart-card"><svg viewBox={`0 0 ${width} ${height}`} className="chart-svg" role="img" aria-label={title}><text x={width / 2} y={34} textAnchor="middle" className="svg-chart-title">{title}</text>{ticks.map((tick) => { const y = getY(tick); return <g key={tick}><text x={marginLeft - 14} y={y + 4} textAnchor="end" className="svg-axis-text">{tick}</text><line x1={marginLeft} y1={y} x2={width - marginRight} y2={y} className="svg-grid-line" /></g>; })}<line x1={marginLeft} y1={marginTop + plotHeight} x2={width - marginRight} y2={marginTop + plotHeight} className="svg-axis-line" />{chartCategories.map((category, index) => { const red = series?.红?.[index] ?? 0; const yellow = series?.黄?.[index] ?? 0; const green = series?.绿?.[index] ?? 0; const total = red + yellow + green; const xCenter = marginLeft + stepX * index + stepX / 2; const barX = xCenter - barWidth / 2; const redH = (red / safeMax) * plotHeight; const yellowH = (yellow / safeMax) * plotHeight; const greenH = (green / safeMax) * plotHeight; const greenY = marginTop + plotHeight - greenH; const yellowY = greenY - yellowH; const redY = yellowY - redH; const topY = total > 0 ? redY : marginTop + plotHeight; const labelY = marginTop + plotHeight + 18; return <g key={category}>{red > 0 ? <rect x={barX} y={redY} width={barWidth} height={redH} className="svg-bar-red" /> : null}{yellow > 0 ? <rect x={barX} y={yellowY} width={barWidth} height={yellowH} className="svg-bar-yellow" /> : null}{green > 0 ? <rect x={barX} y={greenY} width={barWidth} height={greenH} className="svg-bar-green" /> : null}{total > 0 ? <text x={xCenter} y={Math.max(marginTop + 14, topY - 6)} textAnchor="middle" className="svg-bar-value">{total}</text> : null}<text x={xCenter} y={labelY} textAnchor="end" transform={`rotate(-48 ${xCenter} ${labelY})`} className="svg-x-label">{category}</text></g>; })}<g transform={`translate(${width / 2 - 72}, ${height - 18})`}><rect x="0" y="-9" width="12" height="12" className="svg-legend-red" /><text x="18" y="1" className="svg-legend-text">红</text><rect x="44" y="-9" width="12" height="12" className="svg-legend-yellow" /><text x="62" y="1" className="svg-legend-text">黄</text><rect x="88" y="-9" width="12" height="12" className="svg-legend-green" /><text x="106" y="1" className="svg-legend-text">绿</text></g></svg></div>;
}

function getErrorMessage(error, fallback) {
  if (!error) return fallback;
  if (typeof error === 'string') return error;
  if (error instanceof Error) return error.message || fallback;
  if (typeof error === 'object') {
    try {
      return JSON.stringify(error);
    } catch {
      return fallback;
    }
  }
  return fallback;
}

function EmptyCell({ children = '0' }) { return <span className="cell-placeholder">{children}</span>; }
function EmailSection({ title, intro, children }) { return <section className="email-section"><p className="section-title">{title}</p>{intro ? <p className="email-paragraph">{intro}</p> : null}{children}</section>; }

export default function App() {
  const [healthFile, setHealthFile] = useState(null), [strategyFile, setStrategyFile] = useState(null), [emailFile, setEmailFile] = useState(null), [loading, setLoading] = useState(false), [emailLoading, setEmailLoading] = useState(false), [emailExporting, setEmailExporting] = useState(false), [message, setMessage] = useState('请上传健康度源数据和战略文件。'), [emailMessage, setEmailMessage] = useState('请上传邮件表格文件以刷新页面内容。'), [executionLogs, setExecutionLogs] = useState(''), [emailData, setEmailData] = useState(createDefaultEmailData);
  const canSubmit = useMemo(() => healthFile && strategyFile && !loading, [healthFile, strategyFile, loading]);
  const canUploadEmail = useMemo(() => emailFile && !emailLoading, [emailFile, emailLoading]);
  const canExportEmail = useMemo(() => emailFile && !emailExporting, [emailFile, emailExporting]);

  const handleSubmit = async (event) => {
    event.preventDefault();
    if (!healthFile || !strategyFile) return setMessage('请先选择两个 Excel 文件。');
    const formData = new FormData();
    formData.append('health_file', healthFile); formData.append('strategy_file', strategyFile);
    setLoading(true); setMessage('正在生成报表，请稍候...'); setExecutionLogs('');
    try {
      const response = await fetch(`${API_BASE}/api/generate`, { method: 'POST', body: formData });
      if (!response.ok) {
        let detail = '生成失败';
        try { const errorData = await response.json(); detail = errorData.detail || detail; } catch { detail = await response.text(); }
        throw new Error(detail);
      }
      const result = await response.json();
      setExecutionLogs(result.logs || '未输出日志。');
      if (result.downloadName) {
        const downloadResponse = await fetch(`${API_BASE}/api/download/${encodeURIComponent(result.downloadName)}`);
        if (!downloadResponse.ok) throw new Error('报表已生成，但下载失败。');
        const blob = await downloadResponse.blob();
        const url = window.URL.createObjectURL(blob), link = document.createElement('a');
        link.href = url; link.download = result.downloadName; document.body.appendChild(link); link.click(); link.remove(); window.URL.revokeObjectURL(url);
      }
      setMessage('报表已生成并开始下载，下方可查看执行日志。');
    } catch (error) { setMessage(getErrorMessage(error, '生成失败，请稍后重试。')); } finally { setLoading(false); }
  };

  const handleEmailUpload = async () => {
    if (!emailFile) return setEmailMessage('请先选择邮件表格文件。');
    const formData = new FormData();
    formData.append('email_file', emailFile);
    setEmailLoading(true); setEmailMessage('正在解析邮件表格并刷新页面...');
    try {
      const response = await fetch(`${API_BASE}/api/email-data`, { method: 'POST', body: formData });
      if (!response.ok) {
        let detail = '邮件表格解析失败';
        try { const errorData = await response.json(); detail = errorData.detail || detail; } catch { detail = await response.text(); }
        throw new Error(detail);
      }
      setEmailData(normalizeEmailData(await response.json()));
      setEmailMessage('邮件表格已解析，页面内容已更新。');
    } catch (error) { setEmailMessage(getErrorMessage(error, '邮件表格解析失败，请稍后重试。')); } finally { setEmailLoading(false); }
  };

  const handleEmailExport = async () => {
    if (!emailFile) return setEmailMessage('请先选择邮件表格文件。');
    setEmailExporting(true);
    setEmailMessage('正在生成邮件工作表并下载...');
    try {
      const formData = new FormData();
      formData.append('email_file', emailFile);

      const response = await fetch(`${API_BASE}/api/email-export`, { method: 'POST', body: formData });
      if (!response.ok) {
        let detail = '邮件工作表导出失败';
        try { const errorData = await response.json(); detail = errorData.detail || detail; } catch { detail = await response.text(); }
        throw new Error(detail);
      }

      const result = await response.json();
      if (result.downloadName) {
        const downloadResponse = await fetch(`${API_BASE}/api/download/${encodeURIComponent(result.downloadName)}`);
        if (!downloadResponse.ok) throw new Error('邮件工作表已生成，但下载失败。');
        const blob = await downloadResponse.blob();
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = result.downloadName;
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.URL.revokeObjectURL(url);
      }
      setEmailMessage('邮件工作表已生成并开始下载。');
    } catch (error) {
      setEmailMessage(getErrorMessage(error, '邮件工作表导出失败，请稍后重试。'));
    } finally {
      setEmailExporting(false);
    }
  };

  return <div className="page email-page"><div className="email-shell"><section className="upload-card"><div className="upload-header"><div><p className="eyebrow">Health Report Mail</p><h1>健康度邮件展示页</h1><p className="subtitle">页面按邮件正文形式展示，表格样式参考 Excel 周报截图，便于直接浏览、截图或后续复制到邮件正文。</p></div></div><form className="upload-panel" onSubmit={handleSubmit}><div className="upload-grid upload-grid-three"><label className="field"><span>健康度源数据文件</span><input type="file" accept=".xlsx,.xls" onChange={(e) => setHealthFile(e.target.files?.[0] || null)} /><strong>{healthFile ? healthFile.name : '未选择文件'}</strong></label><label className="field"><span>战略文件</span><input type="file" accept=".xlsx,.xls" onChange={(e) => setStrategyFile(e.target.files?.[0] || null)} /><strong>{strategyFile ? strategyFile.name : '未选择文件'}</strong></label><label className="field"><span>邮件表格文件</span><input type="file" accept=".xlsx,.xls" onChange={(e) => setEmailFile(e.target.files?.[0] || null)} /><strong>{emailFile ? emailFile.name : '未选择文件'}</strong></label></div><div className="action-row action-row-wrap"><button className="submit" type="submit" disabled={!canSubmit}>{loading ? '生成中...' : '生成并下载报表'}</button><div className="message" role="status">{message}</div></div><div className="action-row action-row-wrap"><button className="submit submit-secondary" type="button" disabled={!canUploadEmail} onClick={handleEmailUpload}>{emailLoading ? '更新中...' : '上传邮件表格并刷新页面'}</button><button className="submit submit-secondary" type="button" disabled={!canExportEmail} onClick={handleEmailExport}>{emailExporting ? '导出中...' : '导出邮件工作表并下载'}</button><div className="message" role="status">{emailMessage}</div></div>{executionLogs ? <pre className="log-panel">{executionLogs}</pre> : null}</form></section><article className="email-paper"><header className="email-header"><p>{emailData.header?.greeting}</p><p>{emailData.header?.summary}</p><p>{emailData.header?.meta}</p></header><EmailSection title="（一）立项项目健康度情况" intro="按审批级别统计本期立项项目健康度情况，分别展示红、黄、绿三类项目数量。"><table className="report-table excel-table health-table"><thead><tr><th rowSpan="2" className="side-header">审批级别</th><th colSpan="3" className="group-header">本期健康度</th></tr><tr><th className="sub-header health-red">红</th><th className="sub-header health-yellow">黄</th><th className="sub-header health-green">绿</th></tr></thead><tbody>{approvalLevels.map((item) => <tr key={item}><th>{item}</th><td><EmptyCell>{emailData.healthSummary?.[item]?.红 ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.healthSummary?.[item]?.黄 ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.healthSummary?.[item]?.绿 ?? 0}</EmptyCell></td></tr>)}</tbody></table></EmailSection><EmailSection title="战略项目计划立项健康度情况" intro="按年度统计战略数量、实施中需求数及对应健康度，用于反映战略项目计划立项整体推进情况。"><table className="report-table excel-table strategy-table"><thead><tr><th rowSpan="2" className="side-header">年度</th><th rowSpan="2" className="side-header">战略数</th><th rowSpan="2" className="side-header">实施中需求数</th><th colSpan="3" className="group-header">健康度</th></tr><tr><th className="sub-header health-red">红</th><th className="sub-header health-yellow">黄</th><th className="sub-header health-green">绿</th></tr></thead><tbody>{yearlyRows.map((item) => <tr key={item}><th>{item}</th><td><EmptyCell>{emailData.strategySummary?.[item]?.战略数 ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.strategySummary?.[item]?.实施中需求数 ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.strategySummary?.[item]?.红 ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.strategySummary?.[item]?.黄 ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.strategySummary?.[item]?.绿 ?? 0}</EmptyCell></td></tr>)}</tbody></table></EmailSection><EmailSection title="（二）各中心项目健康度分布" intro="按审批级别分别展示各中心立项项目健康度分布情况，图表按两行两列进行排布。"><div className="charts-grid">{(emailData.chartCards || []).map((chart) => <ChartCard key={chart.title} title={chart.title} series={chart.series} max={chart.max} />)}</div></EmailSection><EmailSection title="（三）项目阶段情况" intro="按审批级别统计项目阶段分布情况，分别展示实施中、待投产、已投产、小计及其中超期投产数量。"><table className="report-table excel-table stage-table"><thead><tr><th rowSpan="2" className="side-header">审批级别</th><th colSpan="5" className="group-header">项目阶段</th></tr><tr><th className="sub-header">实施中</th><th className="sub-header">待投产</th><th className="sub-header">已投产</th><th className="sub-header">小计</th><th className="sub-header">其中，超期投产</th></tr></thead><tbody>{approvalLevels.map((item) => <tr key={item}><th>{item}</th><td><EmptyCell>{emailData.projectSummary?.[item]?.实施中 ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.projectSummary?.[item]?.待投产 ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.projectSummary?.[item]?.已投产 ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.projectSummary?.[item]?.小计 ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.projectSummary?.[item]?.['其中，超期投产'] ?? 0}</EmptyCell></td></tr>)}</tbody></table></EmailSection><EmailSection title="（四）按时上线率情况" intro="按审批级别统计计划上线数、实际上线数及按时上线率，用于衡量各审批级别项目投产达成情况。"><table className="report-table excel-table launch-table"><thead><tr><th rowSpan="2" className="side-header">审批级别</th><th colSpan="3" className="group-header">{emailData.launchSummary?.title || '按时上线率'}</th></tr><tr><th className="sub-header">按计划上线数</th><th className="sub-header">实际上线数</th><th className="sub-header">按时上线率</th></tr></thead><tbody>{approvalLevels.map((item) => <tr key={item}><th>{item}</th><td><EmptyCell>{emailData.launchSummary?.rows?.[item]?.['按计划上线数'] ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.launchSummary?.rows?.[item]?.实际上线数 ?? 0}</EmptyCell></td><td><EmptyCell>{emailData.launchSummary?.rows?.[item]?.按时上线率 ?? '0%'}</EmptyCell></td></tr>)}</tbody></table></EmailSection><footer className="email-footer"><p>以上为本期健康度情况，请审阅。</p></footer></article></div></div>;
}
