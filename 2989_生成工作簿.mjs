// SPDX-License-Identifier: MIT
// Authoring uses the separately installed Codex bundled SDK; no dependency is vendored.
import fs from 'node:fs/promises';
import { constants } from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
import { spawnSync } from 'node:child_process';

const BLUE = '#244A6A', INK = '#243649', MUTED = '#64758A', LINE = '#DFE6ED';
const CLEAN = '2989_有效预订', SUMMARY = '2989_汇总';
const START = 6;

function options(argv) {
  const mode = argv.shift();
  if (!['source', 'result'].includes(mode)) throw new Error('首参数须为 source 或 result');
  const args = { mode };
  while (argv.length) {
    const key = argv.shift();
    if (!['--model', '--output', '--previews', '--modules'].includes(key) || !argv.length) throw new Error('参数：source|result --model JSON --output 新.xlsx --previews 新目录 [--modules node_modules]');
    if (args[key.slice(2)]) throw new Error(`参数重复：${key}`);
    args[key.slice(2)] = path.resolve(argv.shift());
  }
  if (!args.model || !args.output || !args.previews) throw new Error('model、output、previews 必填');
  if (!args.output.toLowerCase().endsWith('.xlsx')) throw new Error('输出必须为 .xlsx');
  if (new Set([args.model, args.output, args.previews]).size !== 3) throw new Error('输入与输出路径须不同');
  return args;
}

async function absent(filename) {
  try { await fs.lstat(filename); }
  catch (error) { if (error.code === 'ENOENT') return; throw error; }
  throw new Error(`已有目标，拒绝覆盖：${filename}`);
}

function text(value) {
  if (value == null) return '';
  const string = String(value);
  if (string.length > 32767) throw new Error('文本超过 Excel 单元格32767字符限制；不截断原始内容');
  // Public range.values interprets only an initial '=' as a formula, and "'="
  // as its escape. Other text must remain byte-for-byte unchanged.
  if (string.startsWith("'=")) throw new Error("原文以单引号和等号开头（'=）；当前SDK公开写值接口无法保真，拒绝导出而不改字");
  return string.startsWith('=') ? `'${string}` : string;
}

function isoDate(value) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error(`日期不是 YYYY-MM-DD：${value}`);
  const date = new Date(`${value}T00:00:00.000Z`);
  if (!Number.isFinite(date.valueOf()) || date.toISOString().slice(0, 10) !== value) throw new Error(`无效日期：${value}`);
  return date;
}

function primitive(value) {
  if (value == null) return null;
  if (typeof value === 'string') return text(value);
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'boolean') return value;
  if (typeof value === 'object' && !Array.isArray(value) && Object.keys(value).length === 1 && '$date' in value) return isoDate(value.$date);
  throw new Error('源模型仅允许普通值及单键 $date 对象');
}

function column(index) {
  let output = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) output = String.fromCharCode(65 + (n - 1) % 26) + output;
  return output;
}

function base(sheet, range) {
  sheet.showGridLines = false;
  sheet.tabColor = BLUE;
  range.format.font = { name: 'Arial', size: 10, color: INK };
  range.format.verticalAlignment = 'center';
  range.format.rowHeight = 24;
  range.format.horizontalAlignment = 'left';
}

function heading(range) {
  range.format = { fill: BLUE, font: { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' }, rowHeight: 30, horizontalAlignment: 'center', verticalAlignment: 'center', wrapText: true };
  range.format.borders = { insideVertical: { style: 'thin', color: '#FFFFFF' }, bottom: { style: 'thin', color: BLUE } };
}

function title(sheet, label, note, last = 'J', rows = 20) {
  base(sheet, sheet.getRange(`A1:${last}${Math.max(rows, START)}`));
  sheet.getRange(`A1:${last}1`).format.rowHeight = 9;
  sheet.getRange('A2').values = [[text(label)]];
  sheet.getRange('A2').format.font = { name: 'Arial', size: 14, bold: true, color: BLUE };
  sheet.getRange(`A2:${last}2`).format.rowHeight = 30;
  sheet.getRange('A3').values = [[text(note)]];
  sheet.getRange(`A3:${last}3`).format.font = { name: 'Arial', size: 10, italic: true, color: MUTED };
  sheet.getRange(`A4:${last}4`).format.rowHeight = 10;
}

function textMask(values, index) {
  return values.map(row => {
    const raw = row[index];
    return [typeof raw === 'string' && /^0\d+$/.test(raw) ? '0'.repeat(raw.length) : '@'];
  });
}

function table(sheet, headers, rows, name, widths, textColumns = []) {
  const end = START + rows.length - 1;
  const last = column(headers.length - 1);
  sheet.getRange(`A5:${last}5`).values = [headers];
  if (rows.length) {
    sheet.getRange(`A${START}:${last}${end}`).setNumberFormat('@');
    sheet.getRange(`A${START}:${last}${end}`).values = rows;
    for (const index of textColumns) sheet.getRange(`${column(index)}${START}:${column(index)}${end}`).format.numberFormat = textMask(rows, index);
    const item = sheet.tables.add(`A5:${last}${end}`, true, name);
    item.style = 'TableStyleMedium2'; item.showFilterButton = true; item.showBandedColumns = false;
    sheet.getRange(`A${START}:${last}${end}`).format.borders = { insideHorizontal: { style: 'thin', color: LINE } };
  }
  heading(sheet.getRange(`A5:${last}5`));
  widths.forEach((width, index) => { sheet.getRange(`${column(index)}1:${column(index)}${Math.max(end, 5)}`).format.columnWidth = width; });
  sheet.freezePanes.freezeRows(5);
  return end;
}

function normalize(value) { return String(value ?? '').replace(/[İIı]/g, 'i').normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
const ALIASES = /^(rez|rezerv|booking|ad$|adi$|isim|first|soyad|surname|last|giris|cikis|check|fiyat|ucret|tutar|price|not$|notlar|notes|aciklama)/;

function sourceBook(book, model) {
  if (!Array.isArray(model.sheets) || !model.sheets.length) throw new Error('source 模型须有非空 sheets');
  const names = new Set(), plan = [];
  for (const item of model.sheets) {
    if (typeof item.name !== 'string' || !item.name.startsWith('2989_') || item.name.length > 31 || /[\\/?*\[\]:]/.test(item.name)) throw new Error(`源工作表名称不符合约定：${item.name}`);
    if (names.has(item.name.toLowerCase())) throw new Error('工作表名重复');
    names.add(item.name.toLowerCase());
    if (!Array.isArray(item.rows) || item.rows.some(row => !Array.isArray(row))) throw new Error(`rows 结构错误：${item.name}`);
    const cols = Math.max(1, ...item.rows.map(row => row.length));
    if (cols > 16384 || item.rows.length > 1048576) throw new Error('源模型超出 Excel 行列限制');
    const data = item.rows.map(row => Array.from({ length: cols }, (_, index) => primitive(row[index])));
    let headerIndex = 0, max = 0;
    item.rows.slice(0, 8).forEach((row, index) => {
      const score = row.filter(value => typeof value === 'string' && ALIASES.test(normalize(value))).length;
      if (score > max) { max = score; headerIndex = index; }
    });
    const height = Math.max(8, item.rows.length), last = column(cols - 1), sheet = book.worksheets.add(item.name);
    base(sheet, sheet.getRange(`A1:${last}${height}`));
    sheet.getRange(`A1:${last}${height}`).setNumberFormat('@');
    if (data.length) sheet.getRange(`A1:${last}${data.length}`).values = data;
    for (let c = 0; c < cols; c++) {
      const token = normalize(item.rows[headerIndex]?.[c]);
      const width = /not|aciklama/.test(token) ? 37 : /fiyat|ucret|price|tutar|giris|cikis|check/.test(token) ? 19 : /booking|rez/.test(token) ? 21 : 20;
      sheet.getRange(`${column(c)}1:${column(c)}${height}`).format.columnWidth = width;
      sheet.getRange(`${column(c)}1:${column(c)}${Math.max(1, data.length)}`).format.numberFormat = data.length ? textMask(data, c) : [['@']];
    }
    item.rows.forEach((row, r) => row.forEach((value, c) => {
      if (value && typeof value === 'object' && '$date' in value) sheet.getCell(r, c).setNumberFormat('yyyy-mm-dd');
      else if (typeof value === 'number') sheet.getCell(r, c).setNumberFormat(Number.isInteger(value) ? '#,##0' : '0.00');
    }));
    if (max >= 2) heading(sheet.getRange(`A${headerIndex + 1}:${last}${headerIndex + 1}`));
    sheet.freezePanes.freezeRows(max >= 2 ? headerIndex + 1 : 1);
    sheet.freezePanes.freezeColumns(1);
    plan.push({ name: item.name, range: `A1:${last}${Math.min(height, Math.max(8, headerIndex + 6))}` });
  }
  return { plan, stats: { source_sheets: model.sheets.length, source_raw_rows: model.sheets.reduce((total, sheet) => total + sheet.rows.length, 0) } };
}

function acceptedRows(model) {
  if (!Array.isArray(model.accepted) || !Array.isArray(model.ledger) || !Array.isArray(model.summary)) throw new Error('result 模型须有 accepted、ledger、summary 数组');
  return model.accepted.map(row => {
    if (typeof row.price_usd !== 'number' || !Number.isFinite(row.price_usd) || !Number.isInteger(row.nights) || row.nights < 1) throw new Error('有效预订的价格或晚数类型无效');
    if (!Number.isInteger(row.source_row) || row.source_row < 1) throw new Error('来源行号无效');
    return [text(row.booking_id), text(row.first_name), text(row.last_name), isoDate(row.check_in), isoDate(row.check_out), row.nights, row.price_usd, text(row.notes), text(row.source_sheet), row.source_row];
  });
}

function summaryExpected(accepted, period) {
  const included = accepted.filter(row => row.check_in.startsWith(period));
  return { count: included.length, amount: included.reduce((sum, row) => sum + row.price_usd, 0) };
}

function resultBook(book, model) {
  const data = acceptedRows(model), summary = book.worksheets.add(SUMMARY), clean = book.worksheets.add(CLEAN);
  title(clean, '已核实酒店预订', '本次清洗静态快照。日期和美元价格为真实类型；修改此表不会重新执行源表识别。', 'J', START + data.length - 1);
  const cleanEnd = table(clean, ['预订编号', '名字', '姓氏', '入住日期', '退房日期', '入住晚数', '美元价格', '备注', '来源工作表', '来源行号'], data, 'HotelAccepted2989', [22, 19, 19, 17, 17, 13, 17, 49, 27, 13], [0, 1, 2, 7, 8]);
  if (data.length) {
  clean.getRange(`D${START}:E${cleanEnd}`).setNumberFormat('yyyy-mm-dd');
  clean.getRange(`F${START}:F${cleanEnd}`).setNumberFormat('#,##0');
  clean.getRange(`G${START}:G${cleanEnd}`).setNumberFormat('"$"#,##0.00');
  clean.getRange(`J${START}:J${cleanEnd}`).setNumberFormat('#,##0');
  for (const col of ['D', 'E', 'F', 'G', 'J']) clean.getRange(`${col}${START}:${col}${cleanEnd}`).format.horizontalAlignment = 'right';
  }
  clean.freezePanes.freezeColumns(3);

  const ledgerRows = model.ledger.map(row => {
    if (!Number.isInteger(row.source_row) || row.source_row < 1) throw new Error('来源账本行号无效');
    return [text(row.source_sheet), row.source_row, text(row.status), text(row.reason), text(row.raw_json)];
  });
  const review = ledgerRows.filter((row, index) => model.ledger[index].status === 'review');
  const reviewSheet = book.worksheets.add('2989_待复核');
  const ledgerSheet = book.worksheets.add('2989_来源记录');
  for (const [sheet, rows, name, label, note] of [
    [reviewSheet, review, 'HotelReview2989', '需要人工复核的来源行', '原始JSON完整保留。长文本可在公式栏或展开单元格查看，不截断底层内容。'],
    [ledgerSheet, ledgerRows, 'HotelLedger2989', '来源记录与处置', '每个来源行保留工作表、行号、状态、原因与原始JSON，包含未纳入有效预订的记录。'],
  ]) {
    title(sheet, label, note, 'E', START + rows.length - 1);
    const end = table(sheet, ['来源工作表', '来源行号', '处置状态', '原因', '原始记录JSON'], rows, name, [27, 13, 22, 63, 100], [0, 2, 3, 4]);
    if (rows.length) {
      sheet.getRange(`B${START}:B${end}`).setNumberFormat('#,##0');
      sheet.getRange(`B${START}:B${end}`).format.horizontalAlignment = 'right';
      sheet.getRange(`C${START}:C${end}`).format.horizontalAlignment = 'center';
    }
    sheet.freezePanes.freezeColumns(2);
  }

  const years = model.summary.filter(row => /^\d{4}$/.test(row.period)).sort((a, b) => a.period.localeCompare(b.period));
  const months = model.summary.filter(row => /^\d{4}-\d{2}$/.test(row.period)).sort((a, b) => a.period.localeCompare(b.period));
  if (years.length + months.length !== model.summary.length || new Set(model.summary.map(row => row.period)).size !== model.summary.length) throw new Error('summary 期间必须唯一且采用 YYYY 或 YYYY-MM');
  title(summary, '酒店预订月表合并汇总', '自造数据，按入住月归集整笔已核实预订金额，非逐夜收入确认', 'F', 14 + months.length);
  summary.getRange('A1:A70').format.columnWidth = 18;
  summary.getRange('B1:B70').format.columnWidth = 16;
  summary.getRange('C1:C70').format.columnWidth = 22;
  summary.getRange('D1:D70').format.columnWidth = 4;
  summary.getRange('E1:E70').format.columnWidth = 84;
  summary.getRange('F1:F70').format.columnWidth = 4;
  const calculations = [];
  const dates = `'${CLEAN}'!$D$${START}:$D$${Math.max(START, cleanEnd)}`, amounts = `'${CLEAN}'!$G$${START}:$G$${Math.max(START, cleanEnd)}`;
  function section(items, header, monthly, name) {
    summary.getRange(`A${header}:C${header}`).values = [[monthly ? '入住月份' : '入住年度', '已核实预订数', '预订金额(USD)']];
    for (let index = 0; index < items.length; index++) {
      const item = items[index], row = header + index + 1;
      const begin = monthly ? `${item.period}-01` : `${item.period}-01-01`;
      summary.getRange(`A${row}`).values = [[isoDate(begin)]];
      summary.getRange(`A${row}`).setNumberFormat(monthly ? 'yyyy-mm' : 'yyyy');
      const end = monthly ? `EDATE(A${row},1)` : `DATE(YEAR(A${row})+1,1,1)`;
      summary.getRange(`B${row}:C${row}`).formulas = [[`=COUNTIFS(${dates},">="&A${row},${dates},"<"&${end})`, `=SUMIFS(${amounts},${dates},">="&A${row},${dates},"<"&${end})`]];
      calculations.push({ period: item.period, row, expected_count: item.bookings, expected_amount: item.revenue_usd });
    }
    if (items.length) {
      const tab = summary.tables.add(`A${header}:C${header + items.length}`, true, name);
      tab.style = 'TableStyleMedium2'; tab.showFilterButton = true;
      summary.getRange(`B${header + 1}:B${header + items.length}`).setNumberFormat('#,##0');
      summary.getRange(`C${header + 1}:C${header + items.length}`).setNumberFormat('"$"#,##0.00');
      summary.getRange(`B${header + 1}:C${header + items.length}`).format.horizontalAlignment = 'right';
    }
    heading(summary.getRange(`A${header}:C${header}`));
  }
  section(years, 5, false, 'HotelAnnual2989');
  const monthlyHeader = Math.max(13, 5 + years.length + 4);
  section(months, monthlyHeader, true, 'HotelMonthly2989');
  summary.getRange('E5').values = [['范围与需要复核的情况']];
  summary.getRange('E5').format.font = { name: 'Arial', size: 10, bold: true, color: BLUE };
  const notes = [
    [6, '不猜测歧义日期，不根据金额反推缺失日期。'],
    [7, '不换算TRY，不补姓名，不猜缺失的土耳其字符。'],
    [8, '每晚40美元；日期与金额不一致的来源行进入复核。'],
    [12, '建议今后固定记录'],
    [13, '固定预订号、名字与姓氏分列、ISO日期、明确币种。'],
    [14, '单独记录状态、房间、实收/退款和最后更新时间。'],
    [18, '有效预订是清洗静态快照，Excel公式仅重算汇总。'],
    [19, 'Excel编辑不会重新识别源工作表或自动清洗数据。'],
    [23, '来源工作簿SHA-256'],
    [24, String(model.report?.source_sha256 ?? '')],
  ];
  notes.forEach(([row, value]) => { summary.getRange(`E${row}`).values = [[text(value)]]; });
  summary.getRange('E6:E24').format.font = { name: 'Arial', size: 10, color: MUTED };
  summary.freezePanes.freezeRows(monthlyHeader);
  return { clean, summary, calculations, months, plan: [
    { name: SUMMARY, range: `A1:E${Math.min(monthlyHeader + 12, monthlyHeader + months.length)}` },
    { name: CLEAN, range: `A1:J${Math.min(Math.max(5, cleanEnd), 14)}` },
    { name: '2989_待复核', range: `A1:D${Math.min(Math.max(5, START + review.length - 1), 12)}` },
    { name: '2989_来源记录', range: `A1:D${Math.min(Math.max(5, START + ledgerRows.length - 1), 12)}` },
  ], stats: { accepted: data.length, review: review.length, ledger: ledgerRows.length, annual_periods: years.length, monthly_periods: months.length } };
}

function verifySummary(state, accepted, useModel = false) {
  for (const item of state.calculations) {
    const expected = summaryExpected(accepted, item.period);
    if (useModel && (item.expected_count !== expected.count || Math.abs(item.expected_amount - expected.amount) > 1e-7)) throw new Error(`模型汇总不等于有效预订：${item.period}`);
    const actual = state.summary.getRange(`B${item.row}:C${item.row}`).values[0];
    if (actual[0] !== expected.count || typeof actual[1] !== 'number' || Math.abs(actual[1] - expected.amount) > 1e-7) throw new Error(`汇总重算失败：${item.period}，实际 ${JSON.stringify(actual)}`);
  }
}

const PYTHON_CHECK = String.raw`
import sys,json,datetime,math,hashlib
from pathlib import Path
from openpyxl import load_workbook
p,modelpath,mode=map(str,sys.argv[1:4]); before=hashlib.sha256(Path(p).read_bytes()).hexdigest()
model=json.loads(Path(modelpath).read_text(encoding='utf-8-sig'))
f=load_workbook(p,read_only=True,data_only=False); v=load_workbook(p,read_only=True,data_only=True)
checks=0
def same(a,b):
 if b is None or b=='':return a is None or a==''
 if isinstance(b,dict) and '$date' in b:return isinstance(a,(datetime.datetime,datetime.date)) and a.isoformat()[:10]==b['$date']
 if isinstance(b,(int,float)) and not isinstance(b,bool):return isinstance(a,(int,float)) and not isinstance(a,bool) and math.isclose(a,b,rel_tol=0,abs_tol=1e-8)
 return a==b
if mode=='source':
 assert f.sheetnames==[s['name'] for s in model['sheets']]
 for item in model['sheets']:
  ws=list(f[item['name']].iter_rows())
  for r,row in enumerate(item['rows'],1):
   for c,expected in enumerate(row,1):
    cell=ws[r-1][c-1]; assert cell.data_type!='f',(item['name'],r,c,'unexpected formula')
    assert same(cell.value,expected),(item['name'],r,c,repr(cell.value),repr(expected)); checks+=1
else:
 assert f.sheetnames==['2989_汇总','2989_有效预订','2989_待复核','2989_来源记录']
 ws=list(f['2989_有效预订'].iter_rows()); keys=['booking_id','first_name','last_name','check_in','check_out','nights','price_usd','notes','source_sheet','source_row']
 for r,row in enumerate(model['accepted'],6):
  for c,key in enumerate(keys,1):
   cell=ws[r-1][c-1]; expected=row.get(key)
   if key in ('check_in','check_out'):expected={'$date':expected}
   assert cell.data_type!='f' and same(cell.value,expected),(r,key,repr(cell.value),repr(expected)); checks+=1
  assert ws[r-1][6].data_type=='n'
  assert ws[r-1][3].data_type=='d' and ws[r-1][4].data_type=='d'
 for sheetname,rows in [('2989_待复核',[r for r in model['ledger'] if r['status']=='review']),('2989_来源记录',model['ledger'])]:
  ws=list(f[sheetname].iter_rows())
  for r,row in enumerate(rows,6):
   for c,key in enumerate(['source_sheet','source_row','status','reason','raw_json'],1):
    cell=ws[r-1][c-1];assert cell.data_type!='f' and same(cell.value,row.get(key)),(sheetname,r,key);checks+=1
 summary=v['2989_汇总']; formulas=list(f['2989_汇总'].iter_rows()); lookup={row['period']:row for row in model['summary']}
 for row in summary.iter_rows(min_row=6,max_col=3):
  date=row[0].value
  if not isinstance(date,(datetime.date,datetime.datetime)):continue
  fmt=row[0].number_format; period=date.strftime('%Y-%m') if fmt=='yyyy-mm' else date.strftime('%Y')
  expected=lookup[period];assert row[1].value==expected['bookings'] and math.isclose(row[2].value,expected['revenue_usd'],rel_tol=0,abs_tol=1e-7),(period,'cache')
  assert formulas[row[0].row-1][1].data_type=='f' and formulas[row[0].row-1][2].data_type=='f';checks+=2
 errors=[(ws.title,cell.coordinate) for ws in v for row in ws for cell in row if cell.data_type=='e'];assert not errors,errors
f.close();v.close();assert hashlib.sha256(Path(p).read_bytes()).hexdigest()==before
print(json.dumps({'status':'passed','cells_checked':checks,'xlsx_sha256':before,'read_only':True,'native_excel_opened':False}))
`;

async function main() {
  const began = performance.now(), args = options(process.argv.slice(2));
  await absent(args.output); await absent(args.previews);
  const model = JSON.parse((await fs.readFile(args.model, 'utf8')).replace(/^\uFEFF/, ''));
  const modules = args.modules || process.env.ARTIFACT_TOOL_NODE_MODULES || path.join(os.homedir(), '.cache', 'codex-runtimes', 'codex-primary-runtime', 'dependencies', 'node', 'node_modules');
  const resolver = createRequire(path.join(modules, '2989_模块入口.cjs'));
  let api;
  try { api = await import(pathToFileURL(resolver.resolve('@oai/artifact-tool')).href); }
  catch (error) { throw new Error(`需要另行具备的 Codex bundled @oai/artifact-tool；依赖不随仓库提供，本工具不会购买或下载。${error.message}`); }
  const realModules = await fs.realpath(modules), python = path.join(path.dirname(path.dirname(realModules)), 'python', 'python.exe');
  await fs.access(python);
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), '2989_表格暂存_'));
  try {
    const { Workbook, SpreadsheetFile } = api, book = Workbook.create();
    const state = args.mode === 'source' ? sourceBook(book, model) : resultBook(book, model);
    let mutation = null;
    if (args.mode === 'result' && model.accepted.length && model.summary.length) {
      book.recalculate(); verifySummary(state, model.accepted, true);
      const first = model.accepted[0], target = state.months.find(item => item.period.slice(0, 4) !== first.check_in.slice(0, 4) && item.period !== first.check_in.slice(0, 7));
      if (!target) throw new Error('汇总模型需至少两个年度，以完成跨月跨年重算验收');
      const changed = model.accepted.map((row, index) => index ? row : { ...row, check_in: `${target.period}-01`, price_usd: row.price_usd + 40 });
      state.clean.getRange('D6').values = [[isoDate(changed[0].check_in)]];
      state.clean.getRange('G6').values = [[changed[0].price_usd]];
      book.recalculate(); verifySummary(state, changed);
      state.clean.getRange('D6').values = [[isoDate(first.check_in)]];
      state.clean.getRange('G6').values = [[first.price_usd]];
      mutation = { price_delta_usd: 40, original_check_in: first.check_in, tested_check_in: changed[0].check_in, all_period_formulas_updated: true, original_values_restored: true };
    }
    book.recalculate();
    if (args.mode === 'result') verifySummary(state, model.accepted, true);
    const inspected = await book.inspect({ kind: 'table', range: args.mode === 'result' ? `${SUMMARY}!A5:C9` : `${model.sheets[0].name}!A1:G6`, include: 'values,formulas', tableMaxRows: 6, tableMaxCols: 7, maxChars: 2600 });
    const errors = await book.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!', options: { useRegex: true, maxResults: 30 }, maxChars: 1500 });
    const previewStage = path.join(temporary, '2989_预览'); await fs.mkdir(previewStage);
    const images = [];
    for (const [index, item] of state.plan.entries()) {
      const blob = await book.render({ sheetName: item.name, range: item.range, scale: 1.3, format: 'png' });
      const filename = `2989_${String(index + 1).padStart(2, '0')}_${item.name.replace(/^2989_/, '')}预览.png`;
      await fs.writeFile(path.join(previewStage, filename), new Uint8Array(await blob.arrayBuffer()), { flag: 'wx' });
      images.push({ sheet: item.name, range: item.range, filename });
      if ((index + 1) % 8 === 0 || index + 1 === state.plan.length) process.stderr.write(`2989：已渲染 ${index + 1}/${state.plan.length} 张工作表\n`);
    }
    const stage = path.join(temporary, '2989_工作簿.xlsx'), logger = console.log;
    try { console.log = (...values) => process.stderr.write(values.join(' ') + '\n'); const output = await SpreadsheetFile.exportXlsx(book); await output.save(stage); }
    finally { console.log = logger; }
    const verification = spawnSync(python, ['-X', 'utf8', '-B', '-c', PYTHON_CHECK, stage, args.model, args.mode], { encoding: 'utf8', timeout: 180000, maxBuffer: 500000 });
    if (verification.status !== 0) throw new Error(`openpyxl只读核验失败：${verification.stderr || verification.stdout || verification.error}`);
    const native = JSON.parse(verification.stdout.trim());
    const record = { status: 'passed', mode: args.mode, model: args.model, output: args.output, ...state.stats, previews: images, live_recalculation_test: mutation, readonly_python_check: native, formula_error_scan: errors.ndjson, inspection: inspected.ndjson, render_completed: true, contact_sheet_visual_review_by_root: 'pending', native_excel_opened: false, dependencies_provided_with_repository: false, seconds: Number(((performance.now() - began) / 1000).toFixed(3)) };
    await fs.writeFile(path.join(previewStage, '2989_工作簿核验.json'), JSON.stringify(record, null, 2) + '\n', { flag: 'wx' });
    await fs.mkdir(path.dirname(args.output), { recursive: true }); await fs.mkdir(path.dirname(args.previews), { recursive: true });
    await fs.copyFile(stage, args.output, constants.COPYFILE_EXCL);
    await absent(args.previews); await fs.rename(previewStage, args.previews);
    console.log(JSON.stringify({ status: record.status, mode: args.mode, output: args.output, previews: args.previews, statistics: state.stats, readonly_python_check: native, record: path.join(args.previews, '2989_工作簿核验.json'), seconds: record.seconds }, null, 2));
  } finally {
    const resolved = path.resolve(temporary);
    if (path.dirname(resolved) !== path.resolve(os.tmpdir()) || !path.basename(resolved).startsWith('2989_表格暂存_')) throw new Error('临时清理边界校验失败');
    await fs.rm(resolved, { recursive: true, force: true });
  }
}

main().catch(error => { console.error(JSON.stringify({ status: 'failed', error: error.message })); process.exitCode = 1; });
