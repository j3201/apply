import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { saveAs } from 'file-saver'

export interface ImportedTableARow {
  id: string
  testDate: string
  location: string
  region?: string
  workTime: string
  allowance: number
  name?: string
  payMonth?: string
  portFee?: number
  remark?: string
}

interface ExportTrafficRow {
  outdoorDate: string
  payMonth: string
  employeeId: string
  name: string
  region: string
  dock: string
  address: string
  trafficAllowance: number
}
const headerMap: Record<string, keyof ImportedTableARow> = {
  // 工号相关
  id: 'id',
  单号: 'id',
  编号: 'id',
  工号: 'id',
  工单号: 'id',
  员工编号: 'id',
  jobnumber: 'id',
  'jobno.': 'id',

  // 日期相关
  testdate: 'testDate',
  检验日期: 'testDate',
  检测日期: 'testDate',
  外勤日期: 'testDate',
  工作日期: 'testDate',
  日期: 'testDate',

  // 地点相关 - 检验地点用于交通明细表检索
  location: 'location',
  地点: 'location',
  检验地点: 'location',
  工作地点: 'location',
  所在地: 'location',

  // 区域相关 - 区域用于津贴明细表的 location 字段
  region: 'region',
  区域: 'region',
  地区: 'region',

  // 工作时间相关
  worktime: 'workTime',
  工作时间段: 'workTime',
  工作时间: 'workTime',
  时间段: 'workTime',
  工时: 'workTime',

  // 检验员津贴（不与港口交通费冲突的纯津贴类）
  allowance: 'allowance',
  检验员津贴: 'allowance',
  津贴: 'allowance',
  检验补贴: 'allowance',
  出差补贴: 'allowance',

  // 姓名相关
  name: 'name',
  姓名: 'name',
  员工姓名: 'name',
  申请人: 'name',
  检验员: 'name',
  员工: 'name',
  人员: 'name',

  // 支付月份相关
  paymonth: 'payMonth',
  支付月份: 'payMonth',
  报销月份: 'payMonth',
  结算月份: 'payMonth',
  月份: 'payMonth',

  // 港口交通费相关（仅识别包含"港口"的字段）
  portfee: 'portFee',
  港口交通费: 'portFee',
  港口交通: 'portFee',
  港口费: 'portFee',
  港口金额: 'portFee',

  // 备注相关
  remark: 'remark',
  备注: 'remark',
  检验备注: 'remark',
  工作备注: 'remark',
  详细备注: 'remark'
}

function normalizeHeader(header: string) {
  return header.trim().replace(/\s+/g, '').toLowerCase()
}

/**
 * 表头识别规则：按"字段特异性"由高到低依次判断。
 * 关键原则：港口交通费 (portFee) 必须优先于通用津贴 (allowance) 判断，
 * 否则含"金额""补贴"等字眼的港口交通费会被错误地识别成 allowance。
 */
function mapHeader(header: string): keyof ImportedTableARow | undefined {
  const normalized = normalizeHeader(header)
  const direct = headerMap[normalized]
  if (direct) return direct

  // 1) 港口交通费 - 特异性最高，必须优先（仅识别包含"港口"的字段）
  if (normalized.includes('港口')) {
    if (normalized.includes('交通') || normalized.includes('费') || normalized.includes('金额') || normalized.includes('补贴')) {
      return 'portFee'
    }
  }

  // 2) 支付月份
  if ((normalized.includes('支付') || normalized.includes('报销') || normalized.includes('结算')) && normalized.includes('月')) {
    return 'payMonth'
  }

  // 3) 姓名
  if (normalized.includes('姓名') || normalized === 'name' || normalized.includes('申请人') || normalized.includes('检验员')) {
    return 'name'
  }

  // 4) 工作时间
  if (normalized.includes('work') || normalized.includes('工作') || normalized.includes('时间段')) {
    return 'workTime'
  }

  // 5) 工号 / 单号
  if (normalized.includes('job') || normalized.includes('工号') || normalized.includes('单号') || normalized.includes('编号')) {
    return 'id'
  }

  // 6) 日期 - 排除"月份"类（已在 payMonth 处理）
  if ((normalized.includes('date') || normalized.includes('日期') || normalized.includes('检测') || normalized.includes('检验')) && !normalized.includes('月')) {
    return 'testDate'
  }

  // 7) 地点 / 地区
  if (normalized.includes('loc') || normalized.includes('区') || normalized.includes('地点') || normalized.includes('位置')) {
    return 'location'
  }

  // 8) 备注
  if (normalized.includes('remark') || normalized.includes('备注') || normalized.includes('检验备注')) {
    return 'remark'
  }

  // 9) 检验员津贴 - 最后才识别通用的"津贴/补贴/allow"，避免与 portFee 冲突
  if (normalized.includes('allow') || (normalized.includes('津贴') && !normalized.includes('交通')) || (normalized.includes('补贴') && !normalized.includes('交通'))) {
    return 'allowance'
  }

  return undefined
}

function formatCellValue(value: unknown): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'string') return value.trim()
  if (typeof value === 'number') return String(value).trim()
  if (value instanceof Date) {
    const year = value.getFullYear()
    const month = String(value.getMonth() + 1).padStart(2, '0')
    const day = String(value.getDate()).padStart(2, '0')
    return `${year}/${month}/${day}`
  }
  if (Array.isArray(value)) {
    return value.map(item => formatCellValue(item)).join(' ').trim()
  }
  if (typeof value === 'object' && value !== null) {
    const obj = value as Record<string, unknown>
    if ('text' in obj && typeof obj.text === 'string') {
      return obj.text.trim()
    }
    if ('richText' in obj && Array.isArray(obj.richText)) {
      return (obj.richText as Array<Record<string, unknown>>)
        .map(item => (typeof item.text === 'string' ? item.text : ''))
        .join('')
        .trim()
    }
    if ('formula' in obj && 'result' in obj) {
      return formatCellValue(obj.result)
    }
    if ('hyperlink' in obj && 'text' in obj && typeof obj.text === 'string') {
      return obj.text.trim()
    }
  }
  return String(value).trim()
}

interface ParseResult<T> {
  rows: T[]
  error?: string
}

function buildRowFromRaw(raw: Record<string, string | number | null>): ImportedTableARow {
  const region = raw.region !== undefined && raw.region !== null ? String(raw.region).trim() : undefined
  const locationRaw = String(raw.location ?? '').trim()
  
  return {
    id: String(raw.id ?? '').trim(),
    testDate: formatYMD(String(raw.testDate ?? '').trim()),
    location: locationRaw || region || '',
    region,
    workTime: String(raw.workTime ?? '').trim(),
    allowance: Number(raw.allowance ?? 0),
    name: raw.name !== undefined && raw.name !== null ? String(raw.name).trim() : undefined,
    payMonth: raw.payMonth !== undefined && raw.payMonth !== null ? String(raw.payMonth).trim() : undefined,
    portFee: raw.portFee !== undefined && raw.portFee !== null && raw.portFee !== '' ? Number(raw.portFee) : undefined,
    remark: raw.remark !== undefined && raw.remark !== null ? String(raw.remark).trim() : undefined
  }
}

function pad(n: number) {
  return String(n).padStart(2, '0')
}

function formatYMD(input: string): string {
  if (!input) return ''
  // 如果已经是 YYYY/MM/DD 或类似格式，尝试解析
  const cleaned = input.replace(/[.\-年月]/g, '/').replace(/日/g, '').trim()
  const maybeDate = new Date(cleaned)
  if (!Number.isNaN(maybeDate.getTime())) {
    return `${maybeDate.getFullYear()}/${pad(maybeDate.getMonth() + 1)}/${pad(maybeDate.getDate())}`
  }

  // 尝试匹配连续数字形式 YYYYMMDD
  const m = input.match(/(\d{4})\D?(\d{1,2})\D?(\d{1,2})/)
  if (m) {
    return `${m[1]}/${String(m[2]).padStart(2, '0')}/${String(m[3]).padStart(2, '0')}`
  }

  return input
}

function splitCsvLine(line: string): string[] {
  const result: string[] = []
  let current = ''
  let inQuotes = false

  for (let i = 0; i < line.length; i++) {
    const char = line[i]
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"'
        i++
      } else {
        inQuotes = !inQuotes
      }
    } else if (char === ',' && !inQuotes) {
      result.push(current)
      current = ''
    } else {
      current += char
    }
  }

  result.push(current)
  return result.map(cell => cell.trim())
}

function findCsvHeaderRow(lines: string[]): { headers: Array<keyof ImportedTableARow | undefined>; startRow: number } {
  let best: { headers: Array<keyof ImportedTableARow | undefined>; startRow: number; validCount: number; length: number } = { headers: [], startRow: 1, validCount: 0, length: 0 }
  for (let rowIndex = 0; rowIndex < Math.min(lines.length, 50); rowIndex++) {
    const line = lines[rowIndex] ?? ''
    const headerCells = splitCsvLine(line)
    const mapped = headerCells.map(mapHeader)
    const validCount = mapped.filter(Boolean).length
    if (validCount > best.validCount || (validCount === best.validCount && headerCells.length > best.length)) {
      best = { headers: mapped, startRow: rowIndex + 1, validCount, length: headerCells.length }
    }
  }
  const enoughHeaders = best.validCount >= 2 || (best.validCount >= 1 && best.length >= 4)
  return enoughHeaders ? { headers: best.headers, startRow: best.startRow } : { headers: [], startRow: 1 }
}

function parseCsv(content: string): ParseResult<ImportedTableARow> {
  const lines = content.split(/\r?\n/).filter(line => line.trim() !== '')
  if (lines.length < 2) {
    return { rows: [], error: 'CSV 文件内容过短，无法解析' }
  }

  const { headers, startRow } = findCsvHeaderRow(lines)
  if (!headers.length) {
    return { rows: [], error: '未识别到有效表头，请确保包含单号、检验日期、区域、检验员津贴' }
  }

  const rows: ImportedTableARow[] = []
  for (let i = startRow; i < lines.length; i++) {
    const line = lines[i]
    if (!line) continue
    const values = splitCsvLine(line)
    const raw: Record<string, string | number | null> = {}

    values.forEach((value, idx) => {
      const key = headers[idx]
      if (key) {
        raw[key] = key === 'allowance' || key === 'portFee' ? Number(value) : value
      }
    })

    if (raw.testDate && (raw.location || raw.region)) {
      rows.push(buildRowFromRaw(raw))
    }
  }

  if (!rows.length) {
    return { rows: [], error: '未解析到有效 A 表数据行，请检查您的数据内容' }
  }

  return { rows }
}

function getRowValues(row: ExcelJS.Row): unknown[] {
  if (Array.isArray(row.values)) {
    return row.values.slice(1)
  }
  const values: unknown[] = []
  for (const key of Object.keys(row.values || {})) {
    const index = Number(key)
    if (!Number.isNaN(index) && index >= 1) {
      values[index - 1] = (row.values as Record<string, unknown>)[key]
    }
  }
  return values
}

function getRowTextValues(row: ExcelJS.Row): string[] {
  const values: string[] = []
  const maxCell = row.cellCount || (Array.isArray(row.values) ? row.values.length - 1 : 0)
  for (let col = 1; col <= maxCell; col++) {
    const cell = row.getCell(col)
    const text = typeof cell.text === 'string' ? cell.text.trim() : ''
    if (text) {
      values.push(text)
    } else {
      values.push(formatCellValue(cell.value))
    }
  }
  return values
}

function findXlsxHeaderRow(worksheet: ExcelJS.Worksheet): { headers: Array<keyof ImportedTableARow | undefined>; startRow: number } {
  const maxSearch = 20
  let best: { headers: Array<keyof ImportedTableARow | undefined>; startRow: number; validCount: number; length: number } = { headers: [], startRow: 1, validCount: 0, length: 0 }
  for (let rowNumber = 1; rowNumber <= maxSearch; rowNumber++) {
    const row = worksheet.getRow(rowNumber)
    const headerCells = getRowTextValues(row)
    const mapped = headerCells.map(cell => mapHeader(cell))
    const validCount = mapped.filter(Boolean).length
    if (validCount > best.validCount || (validCount === best.validCount && headerCells.length > best.length)) {
      best = { headers: mapped, startRow: rowNumber + 1, validCount, length: headerCells.length }
    }
  }
  const enoughHeaders = best.validCount >= 2 || (best.validCount >= 1 && best.length >= 4)
  return enoughHeaders ? { headers: best.headers, startRow: best.startRow } : { headers: [], startRow: 2 }
}

async function parseXlsx(buffer: ArrayBuffer): Promise<ParseResult<ImportedTableARow>> {
  const workbook = new ExcelJS.Workbook()

  try {
    await workbook.xlsx.load(buffer)

    for (const worksheet of workbook.worksheets) {
      if (!worksheet) continue
      const { headers, startRow } = findXlsxHeaderRow(worksheet)
      if (!headers.length) continue

      const rows: ImportedTableARow[] = []
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber < startRow) return
        const raw: Record<string, string | number | null> = {}

        const values = getRowValues(row)
        values.forEach((value, idx) => {
          const key = headers[idx]
          if (!key) return
          const cell = row.getCell(idx + 1)
          const textValue = typeof cell.text === 'string' && cell.text.trim() ? cell.text.trim() : formatCellValue(value)
          raw[key] = key === 'allowance' || key === 'portFee' ? Number(textValue) : textValue
        })

        if (raw.testDate && (raw.location || raw.region)) {
          rows.push(buildRowFromRaw(raw))
        }
      })

      if (rows.length) {
        return { rows }
      }
    }

    return { rows: [], error: '未解析到有效 A 表数据行，请检查报销审核表内容' }
  } catch (error) {
    console.error('parseXlsx error:', error)
    return { rows: [], error: 'Excel 文件解析失败，请检查格式' }
  }
}

/**
 * 导出数据为Excel文件
 * @param data 要导出的数组数据
 * @param filename 文件名（不含.xlsx后缀）
 */
export async function exportToExcel(data: any[], filename: string, headerInfo?: { title?: string; company?: string; applicantValue?: string; branchValue?: string; claimValue?: string }) {
  if (!data || data.length === 0) {
    console.warn('导出数据为空，无法生成Excel文件')
    return
  }

  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('津贴明细')

  const headerMap: Record<string, string> = {
    jobNumber: 'JOB NUMBER',
    testDate: 'Survey / Testing Date',
    location: 'Location',
    commencedTime: 'Commenced Time',
    completedTime: 'Completed Time',
    mealAllowances: 'Meal Allowances',
    additionalAllowances: 'Additional Allowances',
    pandemicAllowance: '疫情补贴',
    totalHours: 'Total Hours'
  }

  const headers: Array<keyof typeof headerMap> = [
    'jobNumber',
    'testDate',
    'commencedTime',
    'completedTime',
    'location',
    'mealAllowances',
    'additionalAllowances',
    'pandemicAllowance',
    'totalHours'
  ]

  worksheet.columns = headers.map(key => ({
    header: headerMap[key],
    key,
    width: {
      jobNumber: 16,
      testDate: 18,
      commencedTime: 17,
      completedTime: 20,
      location: 12,
      mealAllowances: 12,
      additionalAllowances: 12,
      pandemicAllowance: 12,
      totalHours: 12
    }[key],
    style: {
      alignment: { vertical: 'middle', horizontal: 'center' }
    }
  }))

  // 计算表头与数据起始行索引（当插入自定义表头时需要偏移）
  // 如果有自定义表头，保留 6 行 header 区（标题、公司、APPLICANT、空行、CLAIM、空行），主表表头放在第7行
  const headerRowIndex = headerInfo ? 7 : 1

  // 如果提供了表头信息，在数据上方插入几行说明并绘制横线（通过底部边框）
  if (headerInfo) {
    // 插入六行：标题、公司、APPLICANT 行、空行、CLAIM DURATION 行、空行（用于主表顶部间隔）
    worksheet.insertRow(1, [])
    worksheet.insertRow(1, [])
    worksheet.insertRow(1, [])
    worksheet.insertRow(1, [])
    worksheet.insertRow(1, [])
    worksheet.insertRow(1, [])

    // 合并第一行并写入标题（大字号，居中）
    worksheet.mergeCells(1, 1, 1, headers.length)
    const titleCell = worksheet.getCell(1, 1)
    titleCell.value = headerInfo.title || ''
    titleCell.font = { name: 'Times New Roman', size: 18, bold: true }
    // 居中标题与公司（图片中标题/公司居中）
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
    worksheet.getRow(1).height = 28

    // 第二行公司名（较小字号，居中）
    worksheet.mergeCells(2, 1, 2, headers.length)
    const compCell = worksheet.getCell(2, 1)
    compCell.value = headerInfo.company || ''
    compCell.font = { name: 'Times New Roman', size: 12, italic: false, bold: false }
    compCell.alignment = { horizontal: 'center', vertical: 'middle' }
    worksheet.getRow(2).height = 20

    // 第三行放置 APPLICANT / BRANCH 与对应值
    const applicantRow = 3
    worksheet.getCell(applicantRow, 1).value = 'APPLICANT'
    worksheet.getCell(applicantRow, 2).value = headerInfo.applicantValue || ''
    worksheet.getCell(applicantRow, 3).value = 'BRANCH'
    worksheet.getCell(applicantRow, 4).value = headerInfo.branchValue || ''
    worksheet.getRow(applicantRow).height = 20
    // 样式：标签使用小号字体，值使用常规字体；全部左对齐，去掉下划线（即不绘制底部边框）
    worksheet.getCell(applicantRow, 1).font = { name: 'Times New Roman', size: 10, bold: true }
    worksheet.getCell(applicantRow, 3).font = { name: 'Times New Roman', size: 10, bold: true }
    worksheet.getCell(applicantRow, 2).font = { name: 'Times New Roman', size: 12, bold: false }
    worksheet.getCell(applicantRow, 4).font = { name: 'Times New Roman', size: 12, bold: false }
    worksheet.getCell(applicantRow, 1).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(applicantRow, 2).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(applicantRow, 3).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(applicantRow, 4).alignment = { horizontal: 'left', vertical: 'middle' }

    const blankRow = 4

    // 第五行放置 CLAIM DURATION 单独一行
    const claimRow = 5
    worksheet.getCell(claimRow, 1).value = 'CLAIM DURATION'
    worksheet.mergeCells(claimRow, 2, claimRow, 6)
    worksheet.getCell(claimRow, 2).value = headerInfo.claimValue || ''
    worksheet.getRow(claimRow).height = 26
    worksheet.getCell(claimRow, 1).font = { name: 'Times New Roman', size: 10, bold: true }
    worksheet.getCell(claimRow, 2).font = { name: 'Times New Roman', size: 12, bold: false }
    worksheet.getCell(claimRow, 1).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(claimRow, 2).alignment = { horizontal: 'left', vertical: 'middle' }

    // 去掉 APPLICANT/BRANCH/CLAIM 的下划线（不绘制底部边框），保持页面与导出一致
    // 在第6行添加一个空白行，作为主表与表头的间隔
    worksheet.getRow(6).height = 10

    // 调整这些单元格的对齐
    for (let c = 1; c <= headers.length; c++) {
      worksheet.getCell(applicantRow, c).alignment = { vertical: 'middle', horizontal: 'center' }
      worksheet.getCell(claimRow, c).alignment = { vertical: 'middle', horizontal: 'center' }
    }
    worksheet.getRow(blankRow).height = 10
    worksheet.getRow(claimRow).height = 24
  }

  // 添加数据行（写入数据会追加在表头之后）
  data.forEach(row => {
    const newRow = worksheet.addRow(row)
    newRow.eachCell(cell => {
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      }
    })
  })

  // 设置表头样式（考虑可能插入的自定义表头偏移）
  // 设置列头样式：使用 Times New Roman，非加粗，与图片一致；添加单元格边框并设置高度
  const headerRow = worksheet.getRow(headerRowIndex)
  headerRow.height = 28
  headerRow.eachCell((cell) => {
    cell.font = { name: 'Times New Roman', size: 12, bold: true }
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    }
  })

  // 在数据下面插入统计行（Subtotal / Grandtotal）
  try {
    const sumMeal = data.reduce((s, r) => s + Number(r.mealAllowances || 0), 0)
    const sumAdditional = data.reduce((s, r) => s + Number(r.additionalAllowances || 0), 0)
    const sumPandemic = data.reduce((s, r) => s + Number(r.pandemicAllowance || 0), 0)
    const grandTotal = sumMeal + sumAdditional + sumPandemic

    const insertAt = (worksheet.lastRow ? worksheet.lastRow.number + 1 : headerRowIndex + data.length + 1)

    // Subtotal 行：左侧合并为说明，数值放在第7-9列
    worksheet.insertRow(insertAt, [])
    const subtotalIdx = insertAt
    worksheet.mergeCells(subtotalIdx, 1, subtotalIdx, 6)
    worksheet.getCell(subtotalIdx, 1).value = 'Subtotal in RMB'
    worksheet.getCell(subtotalIdx, 1).font = { name: 'Times New Roman', size: 12, bold: true }
    worksheet.getCell(subtotalIdx, 7).value = sumMeal || 0
    worksheet.getCell(subtotalIdx, 8).value = sumAdditional || 0
    worksheet.getCell(subtotalIdx, 9).value = sumPandemic || 0
    // 添加边框与对齐
    for (let c = 1; c <= headers.length; c++) {
      const cell = worksheet.getCell(subtotalIdx, c)
      cell.alignment = { vertical: 'middle', horizontal: c <= 6 ? 'left' : 'center' }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      }
    }

    // Grandtotal 行：左侧合并为说明，总金额显示在最后一列（Total Hours列）
    worksheet.insertRow(subtotalIdx + 1, [])
    const grandIdx = subtotalIdx + 1
    worksheet.mergeCells(grandIdx, 1, grandIdx, 9)
    worksheet.getCell(grandIdx, 1).value = 'Grandtotal in RMB'
    worksheet.getCell(grandIdx, 1).font = { name: 'Times New Roman', size: 12, bold: true }
    worksheet.getCell(grandIdx, headers.length).value = grandTotal || 0
    // 样式
    for (let c = 1; c <= headers.length; c++) {
      const cell = worksheet.getCell(grandIdx, c)
      cell.alignment = { vertical: 'middle', horizontal: c === headers.length ? 'center' : 'left' }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      }
    }

    // 在统计行之后插入签名区域（两行标签 + 下划线用于签名）
    const sigStart = grandIdx + 2
    worksheet.insertRow(sigStart, [])
    worksheet.insertRow(sigStart + 1, [])

    // Signed by 行
    const signedIdx = sigStart + 2
    worksheet.insertRow(signedIdx, [])
    worksheet.getCell(signedIdx, 1).value = 'Signed by'
    worksheet.getCell(signedIdx, 1).font = { name: 'Times New Roman', size: 12, bold: true }
    worksheet.getCell(signedIdx, 6).value = '(Applicant)'
    worksheet.getCell(signedIdx, 7).value = 'Date'
    // 合并签名单元格并绘制下划线
    worksheet.mergeCells(signedIdx, 2, signedIdx, 4)
    worksheet.getCell(signedIdx, 2).border = { bottom: { style: 'thin' } }
    worksheet.mergeCells(signedIdx, 8, signedIdx, 9)
    worksheet.getCell(signedIdx, 8).border = { bottom: { style: 'thin' } }
    worksheet.getCell(signedIdx, 2).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(signedIdx, 6).alignment = { horizontal: 'center', vertical: 'middle' }

    // Approved by 行
    const approvedIdx = signedIdx + 1
    worksheet.insertRow(approvedIdx, [])
    worksheet.getCell(approvedIdx, 1).value = 'Approved by'
    worksheet.getCell(approvedIdx, 1).font = { name: 'Times New Roman', size: 12, bold: true }
    worksheet.getCell(approvedIdx, 6).value = '(Operating Management)'
    worksheet.getCell(approvedIdx, 7).value = 'Date'
    worksheet.mergeCells(approvedIdx, 2, approvedIdx, 4)
    worksheet.getCell(approvedIdx, 2).border = { bottom: { style: 'thin' } }
    worksheet.mergeCells(approvedIdx, 8, approvedIdx, 9)
    worksheet.getCell(approvedIdx, 8).border = { bottom: { style: 'thin' } }
    worksheet.getCell(approvedIdx, 2).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(approvedIdx, 6).alignment = { horizontal: 'center', vertical: 'middle' }
  } catch (e) {
    // 忽略统计插入错误，继续生成文件
    console.warn('插入统计行或签名区时发生错误', e)
  }

  // 生成并下载文件
  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: 'application/octet-stream' })
  saveAs(blob, `${filename}.xlsx`)
}

async function buildAllowanceWorkbook(data: any[], headerInfo?: { title?: string; company?: string; applicantValue?: string; branchValue?: string; claimValue?: string }) {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('津贴明细')

  const headerMap: Record<string, string> = {
    jobNumber: 'JOB NUMBER',
    testDate: 'Survey / Testing Date',
    location: 'Location',
    commencedTime: 'Commenced Time',
    completedTime: 'Completed Time',
    mealAllowances: 'Meal Allowances',
    additionalAllowances: 'Additional Allowances',
    pandemicAllowance: '疫情补贴',
    totalHours: 'Total Hours'
  }

  const headers: Array<keyof typeof headerMap> = [
    'jobNumber',
    'testDate',
    'commencedTime',
    'completedTime',
    'location',
    'mealAllowances',
    'additionalAllowances',
    'pandemicAllowance',
    'totalHours'
  ]

  worksheet.columns = headers.map(key => ({
    header: headerMap[key],
    key,
    width: {
      jobNumber: 16,
      testDate: 18,
      commencedTime: 17,
      completedTime: 20,
      location: 12,
      mealAllowances: 12,
      additionalAllowances: 12,
      pandemicAllowance: 12,
      totalHours: 12
    }[key],
    style: {
      alignment: { vertical: 'middle', horizontal: 'center' }
    }
  }))

  const headerRowIndex = headerInfo ? 7 : 1

  if (headerInfo) {
    worksheet.insertRow(1, [])
    worksheet.insertRow(1, [])
    worksheet.insertRow(1, [])
    worksheet.insertRow(1, [])
    worksheet.insertRow(1, [])
    worksheet.insertRow(1, [])

    worksheet.mergeCells(1, 1, 1, headers.length)
    const titleCell = worksheet.getCell(1, 1)
    titleCell.value = headerInfo.title || ''
    titleCell.font = { name: 'Times New Roman', size: 18, bold: true }
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
    worksheet.getRow(1).height = 28

    worksheet.mergeCells(2, 1, 2, headers.length)
    const compCell = worksheet.getCell(2, 1)
    compCell.value = headerInfo.company || ''
    compCell.font = { name: 'Times New Roman', size: 12, italic: false, bold: false }
    compCell.alignment = { horizontal: 'center', vertical: 'middle' }
    worksheet.getRow(2).height = 20

    const applicantRow = 3
    worksheet.getCell(applicantRow, 1).value = 'APPLICANT'
    worksheet.getCell(applicantRow, 2).value = headerInfo.applicantValue || ''
    worksheet.getCell(applicantRow, 3).value = 'BRANCH'
    worksheet.getCell(applicantRow, 4).value = headerInfo.branchValue || ''
    worksheet.getRow(applicantRow).height = 20
    worksheet.getCell(applicantRow, 1).font = { name: 'Times New Roman', size: 10, bold: true }
    worksheet.getCell(applicantRow, 3).font = { name: 'Times New Roman', size: 10, bold: true }
    worksheet.getCell(applicantRow, 2).font = { name: 'Times New Roman', size: 12, bold: false }
    worksheet.getCell(applicantRow, 4).font = { name: 'Times New Roman', size: 12, bold: false }
    worksheet.getCell(applicantRow, 1).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(applicantRow, 2).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(applicantRow, 3).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(applicantRow, 4).alignment = { horizontal: 'left', vertical: 'middle' }

    const blankRow = 4
    const claimRow = 5
    worksheet.getCell(claimRow, 1).value = 'CLAIM DURATION'
    worksheet.mergeCells(claimRow, 2, claimRow, 6)
    worksheet.getCell(claimRow, 2).value = headerInfo.claimValue || ''
    worksheet.getRow(claimRow).height = 26
    worksheet.getCell(claimRow, 1).font = { name: 'Times New Roman', size: 10, bold: true }
    worksheet.getCell(claimRow, 2).font = { name: 'Times New Roman', size: 12, bold: false }
    worksheet.getCell(claimRow, 1).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(claimRow, 2).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getRow(6).height = 10

    for (let c = 1; c <= headers.length; c++) {
      worksheet.getCell(applicantRow, c).alignment = { vertical: 'middle', horizontal: 'center' }
      worksheet.getCell(claimRow, c).alignment = { vertical: 'middle', horizontal: 'center' }
    }
    worksheet.getRow(blankRow).height = 10
    worksheet.getRow(claimRow).height = 24
  }

  data.forEach(row => {
    const newRow = worksheet.addRow(row)
    newRow.eachCell(cell => {
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      }
    })
  })

  const headerRow = worksheet.getRow(headerRowIndex)
  headerRow.height = 28
  headerRow.eachCell((cell) => {
    cell.font = { name: 'Times New Roman', size: 12, bold: true }
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    }
  })

  try {
    const sumMeal = data.reduce((s, r) => s + Number(r.mealAllowances || 0), 0)
    const sumAdditional = data.reduce((s, r) => s + Number(r.additionalAllowances || 0), 0)
    const sumPandemic = data.reduce((s, r) => s + Number(r.pandemicAllowance || 0), 0)
    const grandTotal = sumMeal + sumAdditional + sumPandemic
    const insertAt = (worksheet.lastRow ? worksheet.lastRow.number + 1 : headerRowIndex + data.length + 1)
    worksheet.insertRow(insertAt, [])
    const subtotalIdx = insertAt
    worksheet.mergeCells(subtotalIdx, 1, subtotalIdx, 6)
    worksheet.getCell(subtotalIdx, 1).value = 'Subtotal in RMB'
    worksheet.getCell(subtotalIdx, 1).font = { name: 'Times New Roman', size: 12, bold: true }
    worksheet.getCell(subtotalIdx, 7).value = sumMeal || 0
    worksheet.getCell(subtotalIdx, 8).value = sumAdditional || 0
    worksheet.getCell(subtotalIdx, 9).value = sumPandemic || 0
    for (let c = 1; c <= headers.length; c++) {
      const cell = worksheet.getCell(subtotalIdx, c)
      cell.alignment = { vertical: 'middle', horizontal: c <= 6 ? 'left' : 'center' }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      }
    }
    worksheet.insertRow(subtotalIdx + 1, [])
    const grandIdx = subtotalIdx + 1
    worksheet.mergeCells(grandIdx, 1, grandIdx, 9)
    worksheet.getCell(grandIdx, 1).value = 'Grandtotal in RMB'
    worksheet.getCell(grandIdx, 1).font = { name: 'Times New Roman', size: 12, bold: true }
    worksheet.getCell(grandIdx, headers.length).value = grandTotal || 0
    for (let c = 1; c <= headers.length; c++) {
      const cell = worksheet.getCell(grandIdx, c)
      cell.alignment = { vertical: 'middle', horizontal: c === headers.length ? 'center' : 'left' }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      }
    }
    const sigStart = grandIdx + 2
    worksheet.insertRow(sigStart, [])
    worksheet.insertRow(sigStart + 1, [])
    const signedIdx = sigStart + 2
    worksheet.insertRow(signedIdx, [])
    worksheet.getCell(signedIdx, 1).value = 'Signed by'
    worksheet.getCell(signedIdx, 1).font = { name: 'Times New Roman', size: 12, bold: true }
    worksheet.getCell(signedIdx, 6).value = '(Applicant)'
    worksheet.getCell(signedIdx, 7).value = 'Date'
    worksheet.mergeCells(signedIdx, 2, signedIdx, 4)
    worksheet.getCell(signedIdx, 2).border = { bottom: { style: 'thin' } }
    worksheet.mergeCells(signedIdx, 8, signedIdx, 9)
    worksheet.getCell(signedIdx, 8).border = { bottom: { style: 'thin' } }
    worksheet.getCell(signedIdx, 2).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(signedIdx, 6).alignment = { horizontal: 'center', vertical: 'middle' }
    const approvedIdx = signedIdx + 1
    worksheet.insertRow(approvedIdx, [])
    worksheet.getCell(approvedIdx, 1).value = 'Approved by'
    worksheet.getCell(approvedIdx, 1).font = { name: 'Times New Roman', size: 12, bold: true }
    worksheet.getCell(approvedIdx, 6).value = '(Operating Management)'
    worksheet.getCell(approvedIdx, 7).value = 'Date'
    worksheet.mergeCells(approvedIdx, 2, approvedIdx, 4)
    worksheet.getCell(approvedIdx, 2).border = { bottom: { style: 'thin' } }
    worksheet.mergeCells(approvedIdx, 8, approvedIdx, 9)
    worksheet.getCell(approvedIdx, 8).border = { bottom: { style: 'thin' } }
    worksheet.getCell(approvedIdx, 2).alignment = { horizontal: 'left', vertical: 'middle' }
    worksheet.getCell(approvedIdx, 6).alignment = { horizontal: 'center', vertical: 'middle' }
  } catch (e) {
    console.warn('插入统计行或签名区时发生错误', e)
  }

  return workbook
}

/**
 * 构建交通明细Excel工作簿
 * @param data 导出交通数据行数组
 * @returns 返回配置好的Excel工作簿对象
 */

async function buildTrafficWorkbook(data: ExportTrafficRow[]) {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('交通明细')
  // 定义表头映射关系，将英文字段名映射为中文字段名
  const headerMap: Record<string, string> = {
    outdoorDate: '外勤日期',
    payMonth: '支付月份',
    employeeId: '工号',
    name: '姓名',
    region: '地区',
    dock: '码头',
    address: '地址',
    trafficAllowance: '交通津贴金额'
  }
  // 定义表头顺序数组，用于确定列的显示顺序
  const headers: Array<keyof typeof headerMap> = ['outdoorDate', 'payMonth', 'employeeId', 'name', 'region', 'dock', 'address', 'trafficAllowance']
  // 定义各列的样式属性（背景色、字体颜色等）
  const orangeColumns: Array<keyof typeof headerMap> = ['region', 'dock', 'address']
  const redHeaderColumns: Array<keyof typeof headerMap> = ['region', 'dock']
  const tealBorderColumns: Array<keyof typeof headerMap> = ['address']
  const orangeFill: ExcelJS.FillPattern = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFC000' }
  }
  const tealHeaderFill: ExcelJS.FillPattern = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF7FE0D8' }
  }
  const headerRedFont: Partial<ExcelJS.Font> = {
    name: '微软雅黑',
    size: 12,
    bold: true,
    color: { argb: 'FFFF0000' }
  }
  const headerDefaultFont: Partial<ExcelJS.Font> = {
    name: '微软雅黑',
    size: 12,
    bold: true,
    color: { argb: 'FF000000' }
  }
  const cellDefaultFont: Partial<ExcelJS.Font> = {
    name: '微软雅黑',
    size: 12,
    bold: false,
    color: { argb: 'FF000000' }
  }
  const cellAmountFont: Partial<ExcelJS.Font> = {
    name: '微软雅黑',
    size: 12,
    bold: true,
    color: { argb: 'FF000000' }
  }
  // 设置工作表列的配置，包括表头、列宽和样式
  worksheet.columns = headers.map(key => ({
    header: headerMap[key], // 显示的表头名称
    key, // 数据键名
    width: { // 设置各列宽度（按图片比例）
      outdoorDate: 18,
      payMonth: 14,
      employeeId: 13,
      name: 12,
      region: 16,
      dock: 18,
      address: 48,
      trafficAllowance: 22
    }[key],
    style: { // 设置默认样式
      alignment: { vertical: 'middle', horizontal: 'center' }
    }
  }))

  // 如果有数据，则添加数据行
  if (data.length) {
    data.forEach((row, rowIndex) => {
      // 添加新行
      const newRow = worksheet.addRow(row)
      newRow.height = 32
      // 为每个单元格设置样式
      headers.forEach((key, colIndex) => {
        const cell = newRow.getCell(colIndex + 1)
        // 金额列：右对齐、显示两位小数
        if (key === 'trafficAllowance') {
          cell.numFmt = '0.00'
          cell.alignment = { vertical: 'middle', horizontal: 'right' }
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'center' }
        }
        // 金额列加粗，其他列不加粗
        cell.font = key === 'trafficAllowance' ? cellAmountFont : cellDefaultFont
        // 地区/码头/地址列：橙色背景
        if (orangeColumns.includes(key)) {
          cell.fill = orangeFill
          // 地址列：底部加青绿色边框
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: tealBorderColumns.includes(key)
              ? { style: 'medium', color: { argb: 'FF00B050' } }
              : { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          }
        } else {
          // 其他列：白底黑边
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          }
        }
      })
    })
  }

  // 设置表头行样式
  const headerRow = worksheet.getRow(1)
  headerRow.height = 32 // 设置行高
  headers.forEach((key, colIndex) => {
    const cell = headerRow.getCell(colIndex + 1)
    // 地区/码头：红色字体
    if (redHeaderColumns.includes(key)) {
      cell.font = headerRedFont
    } else {
      cell.font = headerDefaultFont
    }
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    // 地址列表头：青绿色背景
    if (tealBorderColumns.includes(key)) {
      cell.fill = tealHeaderFill
    }
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    }
  })

  // 添加金额合计行
  const totalAmount = data.reduce((sum, row) => sum + Number(row.trafficAllowance || 0), 0)
  const totalRowIndex = data.length + 2 // 数据行之后添加一行

  // 在最后一行后添加合计行
  const lastDataRow = worksheet.getRow(data.length + 1)
  const totalRow = worksheet.addRow(['合计', '', '', '', '', '', '', totalAmount])
  totalRow.height = 32

  // 设置合计行样式
  for (let colIndex = 1; colIndex <= headers.length; colIndex++) {
    const cell = totalRow.getCell(colIndex)
    cell.font = {
      name: '微软雅黑',
      size: 12,
      bold: true,
      color: { argb: 'FF000000' }
    }
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'medium', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } }
    }
    // 前7列左对齐
    if (colIndex <= 7) {
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
    } else {
      // 金额列右对齐
      cell.numFmt = '0.00'
      cell.alignment = { vertical: 'middle', horizontal: 'right' }
    }
  }

  return workbook
}

export async function exportTablesZip(
  trafficData: ExportTrafficRow[],
  allowanceData: any[],
  filename: string,
  headerInfo?: { title?: string; company?: string; applicantValue?: string; branchValue?: string; claimValue?: string }
) {
  const zip = new JSZip()
  const trafficWorkbook = await buildTrafficWorkbook(trafficData)
  const allowanceWorkbook = await buildAllowanceWorkbook(allowanceData, headerInfo)

  const [trafficBuffer, allowanceBuffer] = await Promise.all([
    trafficWorkbook.xlsx.writeBuffer(),
    allowanceWorkbook.xlsx.writeBuffer()
  ])

  zip.file('交通费津贴明细.xlsx', trafficBuffer)
  zip.file('Allowance Details.xlsx', allowanceBuffer)

  const content = await zip.generateAsync({ type: 'blob' })
  
  interface FilePickerOptions {
    suggestedName?: string
    types?: Array<{
      description: string
      accept: Record<string, string[]>
    }>
  }
  
  interface FileSystemWritableFileStream {
    write(data: Blob): Promise<void>
    close(): Promise<void>
  }
  
  interface FileSystemFileHandle {
    createWritable(): Promise<FileSystemWritableFileStream>
  }
  
  type SaveFilePickerFn = (options: FilePickerOptions) => Promise<FileSystemFileHandle>
  
  const saveFilePicker = (window as unknown as { showSaveFilePicker?: SaveFilePickerFn }).showSaveFilePicker
  if (saveFilePicker) {
    try {
      const handle = await saveFilePicker({
        suggestedName: `${filename}.zip`,
        types: [
          {
            description: 'ZIP Archive',
            accept: {
              'application/zip': ['.zip']
            }
          }
        ]
      })
      const writable = await handle.createWritable()
      await writable.write(content)
      await writable.close()
      return
    } catch (err) {
      console.log('User cancelled or error in save file picker:', err)
    }
  }
  
  saveAs(content, `${filename}.zip`)
}

/**
 * 单独导出交通明细表为Excel文件
 * @param data 交通明细数据
 * @param filename 导出文件名（不含扩展名）
 */
export async function exportTrafficToExcel(data: ExportTrafficRow[], filename: string) {
  const workbook = await buildTrafficWorkbook(data)
  const buffer = await workbook.xlsx.writeBuffer()
  saveAs(new Blob([buffer]), `${filename}.xlsx`)
}

/**
 * 构建交通津贴汇总Excel工作簿
 * 按图片样式：支付月份日期 + 姓名 + 交通津贴合计(200×行数) + 总计
 * @param trafficData 交通明细数据
 * @returns Excel工作簿
 */
async function buildSummaryWorkbook(trafficData: ExportTrafficRow[]) {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('交通费津贴合计表')

  // 列定义
  worksheet.columns = [
    { header: '', key: 'col1', width: 22 },
    { header: '', key: 'col2', width: 16 },
    { header: '', key: 'col3', width: 22 }
  ]

  // 从交通明细中提取支付月份，转换为"2026/5/25"格式
  const firstRow = trafficData[0]
  let referenceDate = ''
  if (firstRow && firstRow.payMonth) {
    const pm = firstRow.payMonth
    const match = pm.match(/(\d+)年(\d+)月/)
    if (match) {
      referenceDate = `${match[1]}/${match[2]}/25`
    } else {
      // 回退：使用 outdoorDate 去除"日"并格式化为 年/月/日
      const od = firstRow.outdoorDate
      const odMatch = od.match(/(\d+)[\/-](\d+)[\/-](\d+)/)
      if (odMatch) {
        referenceDate = `${odMatch[1]}/${odMatch[2]}/${odMatch[3]}`
      } else {
        referenceDate = pm
      }
    }
  }

  const rowCount = trafficData.length
  const totalAmount = 200 * rowCount

  // 字体样式
  const defaultFont: Partial<ExcelJS.Font> = {
    name: '微软雅黑',
    size: 12,
    bold: false,
    color: { argb: 'FF000000' }
  }
  const boldFont: Partial<ExcelJS.Font> = {
    name: '微软雅黑',
    size: 12,
    bold: true,
    color: { argb: 'FF000000' }
  }

  // 浅蓝背景色（用于数据行）
  const lightBlueFill: ExcelJS.FillPattern = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE1ECF4' }
  }

  const borderStyle: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFBFBFBF' } },
    left: { style: 'thin', color: { argb: 'FFBFBFBF' } },
    bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } },
    right: { style: 'thin', color: { argb: 'FFBFBFBF' } }
  }

  // 行 1：支付月份 + 日期
  const headerRow1 = worksheet.addRow(['支付月份', referenceDate, ''])
  headerRow1.height = 32
  headerRow1.eachCell((cell) => {
    cell.font = boldFont
    cell.alignment = { vertical: 'middle', horizontal: 'left' }
  })

  // 行 2：表头行 Row Labels | 姓名 | 交通津贴合计
  const headerRow2 = worksheet.addRow(['Row Labels', '姓名', '交通津贴合计'])
  headerRow2.height = 32
  headerRow2.eachCell((cell) => {
    cell.font = boldFont
    cell.alignment = { vertical: 'middle', horizontal: 'left' }
    cell.border = borderStyle
  })

  // 行 3：数据行 - 空 | 空 | 总计金额
  const dataRow = worksheet.addRow(['', '', totalAmount])
  dataRow.height = 32
  dataRow.eachCell((cell) => {
    cell.font = defaultFont
    cell.alignment = { vertical: 'middle', horizontal: 'left' }
    cell.fill = lightBlueFill
    cell.border = borderStyle
  })
  // 金额列格式化：右对齐、两位小数
  const amountCell = dataRow.getCell(3)
  amountCell.numFmt = '0.00'
  amountCell.alignment = { vertical: 'middle', horizontal: 'right' }

  // 行 4：总计行
  const totalRow = worksheet.addRow(['总计', '', totalAmount])
  totalRow.height = 32
  totalRow.eachCell((cell) => {
    cell.font = boldFont
    cell.alignment = { vertical: 'middle', horizontal: 'left' }
    cell.fill = lightBlueFill
    cell.border = borderStyle
  })
  const totalAmountCell = totalRow.getCell(3)
  totalAmountCell.numFmt = '0.00'
  totalAmountCell.alignment = { vertical: 'middle', horizontal: 'right' }

  return workbook
}

/**
 * 导出交通津贴汇总表为Excel文件
 * @param trafficData 交通明细数据
 * @param filename 导出文件名（不含扩展名）
 */
export async function exportSummaryToExcel(trafficData: ExportTrafficRow[], filename: string) {
  const workbook = await buildSummaryWorkbook(trafficData)
  const buffer = await workbook.xlsx.writeBuffer()
  saveAs(new Blob([buffer]), `${filename}.xlsx`)
}

/**
 * 从 Excel/CSV 文件导入 A 表数据
 * @param file 上传的文件
 */
export async function importFromExcel(file: File): Promise<ImportedTableARow[]> {
  const extension = file.name.split('.').pop()?.toLowerCase()
  let result: ParseResult<ImportedTableARow>
  if (extension === 'csv') {
    const text = await file.text()
    result = parseCsv(text)
  } else {
    const buffer = await file.arrayBuffer()
    result = await parseXlsx(buffer)
  }

  if (result.error) {
    throw new Error(result.error)
  }

  return result.rows
}

export async function importPortTrafficList(file: File): Promise<Array<{ location: string; region: string; dock: string; address: string; amount: number; name?: string; employeeId?: string }>> {
  const extension = file.name.split('.').pop()?.toLowerCase()
  let result: ParseResult<ImportedPortTrafficRow>
  if (extension === 'csv') {
    const text = await file.text()
    result = parsePortTrafficCsv(text)
  } else {
    const buffer = await file.arrayBuffer()
    result = await parsePortTrafficXlsx(buffer)
  }

  if (result.error) {
    throw new Error(result.error)
  }

  return result.rows
}

/**
 * 港口交通清单表头识别。按"特异性从高到低"判断。
 * 关键：location 与 region/dock 都与"地点/区域/码头"相关，必须用上下文区分。
 * 规则：含"港口/码头/地点位置"判为 location；含"地区"判为 region；含"码头/泊位"判为 dock
 */
function mapPortTrafficHeader(header: string): keyof ImportedPortTrafficRow | undefined {
  const normalized = normalizeHeader(header)

  // 1) 金额 / 费用 — 特异性最高
  if (normalized.includes('金额') || normalized.includes('交通津贴金额') ||
      normalized.includes('费用') || normalized === 'amount' ||
      (normalized.includes('补贴') && !normalized.includes('检验')) ||
      normalized.includes('交通金额') || normalized.includes('津贴金额') ||
      normalized.includes('交通费用')) {
    return 'amount'
  }

  // 2) 地址
  if (normalized.includes('地址') || normalized === 'address' ||
      normalized.includes('详细地址') || normalized.includes('具体地址')) {
    return 'address'
  }

  // 3) 码头 - 最精确的港口字段
  if (normalized.includes('码头') || normalized.includes('泊位') ||
      normalized === 'dock' || normalized.includes('港区') ||
      normalized.includes('泊位名称') || normalized.includes('码头名称')) {
    return 'dock'
  }

  // 4) 地区 / 区域（省市级）
  if (normalized.includes('地区') || normalized === 'region' ||
      normalized.includes('所在地区') || normalized.includes('所属地区') ||
      normalized.includes('城市') || normalized.includes('省市') ||
      normalized.includes('省份') || normalized.includes('所在省市')) {
    return 'region'
  }

  // 5) 地点 / 港口位置（港口名称本身）- 在码头/地址/地区之后判断
  if (normalized.includes('港口') || normalized.includes('地点') ||
      normalized.includes('位置') || normalized.includes('港口位置') ||
      normalized.includes('检验地点') || normalized === 'location' ||
      normalized.includes('港口名称') || normalized.includes('港口地点') ||
      normalized.includes('码头位置') || normalized.includes('停靠港口')) {
    return 'location'
  }

  // 6) 姓名
  if (normalized.includes('姓名') || normalized === 'name' || normalized.includes('申请人')) {
    return 'name'
  }

  // 7) 工号
  if (normalized.includes('id') || normalized.includes('工号') || normalized.includes('员工编号')) {
    return 'employeeId'
  }

  return undefined
}

interface ImportedPortTrafficRow {
  location: string
  region: string
  dock: string
  address: string
  amount: number
  name?: string
  employeeId?: string
}

function buildPortTrafficRowFromRaw(raw: Record<string, string | number | null>): ImportedPortTrafficRow {
  const safeStr = (v: unknown): string => {
    if (v === null || v === undefined) return ''
    const s = String(v).trim()
    return s === 'undefined' || s === 'null' ? '' : s
  }
  const safeNum = (v: unknown): number => {
    if (v === null || v === undefined || v === '') return 0
    const n = Number(v)
    return Number.isFinite(n) ? n : 0
  }
  return {
    location: safeStr(raw.location),
    region: safeStr(raw.region),
    dock: safeStr(raw.dock),
    address: safeStr(raw.address),
    amount: safeNum(raw.amount),
    name: safeStr(raw.name) || undefined,
    employeeId: safeStr(raw.employeeId) || undefined
  }
}

/**
 * 在 CSV/Excel 中搜索表头所在行：检查前 N 行中哪一行包含的可识别字段最多。
 * 与 A 表使用的 findXlsxHeaderRow 逻辑一致。
 */
function findPortTrafficHeaderRow(cells: string[]): { headers: Array<keyof ImportedPortTrafficRow | undefined>; validCount: number } {
  const headers = cells.map(mapPortTrafficHeader)
  const validCount = headers.filter(Boolean).length
  return { headers, validCount }
}

function parsePortTrafficCsv(content: string): ParseResult<ImportedPortTrafficRow> {
  const lines = content.split(/\r?\n/).filter(line => line.trim() !== '')
  if (lines.length < 2) {
    return { rows: [], error: 'CSV 文件内容过短，无法解析港口交通清单' }
  }

  // 搜索前 20 行中哪个是真正的表头行
  let best = { headers: [] as Array<keyof ImportedPortTrafficRow | undefined>, startRow: 0, validCount: 0 }
  for (let r = 0; r < Math.min(lines.length, 20); r++) {
    const cells = splitCsvLine(lines[r] || '')
    const { headers, validCount } = findPortTrafficHeaderRow(cells)
    if (validCount > best.validCount) {
      best = { headers, startRow: r + 1, validCount }
    }
  }
  if (best.validCount < 1) {
    return { rows: [], error: '未识别到有效港口交通清单表头，请确保包含地点、地址、交通津贴金额等字段' }
  }

  const headers = best.headers
  const rows: ImportedPortTrafficRow[] = []
  for (let i = best.startRow; i < lines.length; i++) {
    const line = lines[i] || ''
    const values = splitCsvLine(line)
    const raw: Record<string, string | number | null> = {}
    values.forEach((value, idx) => {
      const key = headers[idx]
      if (!key) return
      raw[key] = key === 'amount' ? Number(value) : value
    })
    if (raw.location || raw.address || raw.region || raw.amount) {
      rows.push(buildPortTrafficRowFromRaw(raw))
    }
  }

  if (!rows.length) {
    return { rows: [], error: '未解析到有效港口交通清单数据行，请检查文件内容' }
  }

  return { rows }
}

async function parsePortTrafficXlsx(buffer: ArrayBuffer): Promise<ParseResult<ImportedPortTrafficRow>> {
  const workbook = new ExcelJS.Workbook()
  try {
    await workbook.xlsx.load(buffer)

    for (const worksheet of workbook.worksheets) {
      if (!worksheet) continue

      // 搜索前 20 行中哪个是真正的表头行（与 A 表 findXlsxHeaderRow 逻辑一致）
      let best = { headers: [] as Array<keyof ImportedPortTrafficRow | undefined>, startRow: 1, validCount: 0 }
      const maxSearch = 20
      const lastRow = worksheet.lastRow?.number ?? 0
      for (let rowNumber = 1; rowNumber <= Math.min(maxSearch, lastRow); rowNumber++) {
        const row = worksheet.getRow(rowNumber)
        const cells = getRowTextValues(row)
        const { headers, validCount } = findPortTrafficHeaderRow(cells)
        if (validCount > best.validCount) {
          best = { headers, startRow: rowNumber + 1, validCount }
        }
      }
      if (best.validCount < 1) continue
      const headers = best.headers

      const rows: ImportedPortTrafficRow[] = []
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber < best.startRow) return
        const raw: Record<string, string | number | null> = {}
        const values = getRowValues(row)
        values.forEach((value, idx) => {
          const key = headers[idx]
          if (!key) return
          raw[key] = key === 'amount' ? Number(formatCellValue(value)) : formatCellValue(value)
        })
        if (raw.location || raw.address || raw.region || raw.amount) {
          rows.push(buildPortTrafficRowFromRaw(raw))
        }
      })

      if (rows.length) {
        return { rows }
      }
    }

    return { rows: [], error: '未解析到有效港口交通清单数据，请检查文件内容' }
  } catch (error) {
    console.error('parsePortTrafficXlsx error:', error)
    return { rows: [], error: '港口交通清单 Excel 解析失败，请检查格式' }
  }
}
