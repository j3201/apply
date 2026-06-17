<template>
  <div class="container">
    <h2>上传报销审核表数据</h2>
    <div class="button-row">
      <el-button type="primary" @click="triggerFileUpload">选择上传 A 表</el-button>
      <input ref="fileInput" type="file" accept=".xlsx,.csv" @change="handleFileChange" style="position:absolute; left:-9999px; opacity:0" />
      <el-button type="warning" @click="triggerPortTrafficUpload">上传港口交通清单</el-button>
      <input ref="portTrafficInput" type="file" accept=".xlsx,.csv" @change="handlePortTrafficFileChange" style="position:absolute; left:-9999px; opacity:0" />
      <el-button type="success" @click="handleExport">导出津贴明细表</el-button>
      <el-button type="info" @click="handleExportTraffic">导出交通明细表</el-button>
      <el-button type="warning" @click="handleExportSummary">导出交通津贴汇总</el-button>
      <el-button type="primary" @click="handleExportAllZip">导出全部（ZIP）</el-button>
    </div>
    <div class="defaults-row">
      <span class="defaults-label">默认工号</span>
      <el-input
        v-model="defaultEmployeeId"
        placeholder="默认工号 (例如 123)"
        style="width: 220px"
        clearable
      />
      <span class="defaults-label">默认姓名</span>
      <el-input
        v-model="defaultName"
        placeholder="默认姓名 (例如 admin)"
        style="width: 220px"
        clearable
      />
      <span class="defaults-hint">（当 A 表中没有工号/姓名时，将使用此处配置的值）</span>
    </div>
    <div class="status-row">
      <span>{{ uploadStatus }}</span>
      <span>{{ portTrafficStatus }}</span>
    </div>

    <div class="tabs-wrapper">
      <el-tabs v-model="activeTab" type="border-card">
        <el-tab-pane label="交通明细表" name="traffic">
          <div class="table-card">
            <div class="table-header-row">
              <h2>交通明细表</h2>
              <div class="table-actions">
                <el-button type="text" @click="showColumnSettings = !showColumnSettings">
                  列设置
                </el-button>
                <el-switch v-model="enableRowEditing" active-text="可编辑" inactive-text="只读" />
              </div>
            </div>
            
            <div v-if="showColumnSettings" class="column-settings-panel">
              <div class="settings-hint">提示：所有字段均可编辑，点击单元格即可修改</div>
              <el-checkbox-group v-model="selectedColumns">
                <el-checkbox
                  v-for="col in trafficColumns"
                  :key="col.prop"
                  :label="col.prop"
                  :checked="col.visible"
                  @change="toggleColumn(col.prop)"
                >
                  {{ col.label }}
                </el-checkbox>
              </el-checkbox-group>
            </div>

            <div v-if="tableTrafficData.length">
              <el-table :data="tableTrafficData" border style="width: 100%;" size="small">
                <template v-for="col in trafficColumns" :key="col.prop">
                  <el-table-column
                    v-if="col.visible"
                    :prop="col.prop"
                    :label="col.label"
                    :width="col.width === 'auto' ? undefined : col.width"
                    :min-width="col.width === 'auto' ? '140' : undefined"
                    :header-cell-style="headerStyle"
                    :show-overflow-tooltip="col.width === 'auto'"
                  >
                    <template #default="{ row }">
                      <el-input-number
                        v-if="col.prop === 'trafficAllowance' && col.editable && enableRowEditing"
                        v-model="row[col.prop]"
                        :min="0"
                        size="small"
                        style="width:100%"
                      />
                      <el-input
                        v-else-if="col.editable && enableRowEditing"
                        v-model="row[col.prop]"
                        size="small"
                        style="width:100%"
                      />
                      <span v-else>{{ row[col.prop] }}</span>
                    </template>
                  </el-table-column>
                </template>
              </el-table>
              <div class="traffic-total-row">
                <span class="traffic-total-label">金额合计：</span>
                <span class="traffic-total-value">{{ trafficTotalAmount.toFixed(2) }}</span>
              </div>
            </div>
            <div v-else class="empty-state">暂无交通明细数据，请先上传 A 表或港口交通清单。</div>
          </div>
        </el-tab-pane>
        <el-tab-pane label="津贴明细表" name="allowance">
          <div class="table-card">
            <div class="table-header-row">
              <h2>自动生成的津贴明细（B表）</h2>
              <div class="table-actions">
                <el-switch v-model="enableAllowanceEditing" active-text="可编辑" inactive-text="只读" />
              </div>
            </div>
            <div class="table-summary" v-if="tableBData.length">
              <div class="summary-row">
                <span class="summary-label">Subtotal in RMB</span>
                <span class="summary-value">Meal Allowances: {{ tableBTotals.sumMeal }}</span>
                <span class="summary-value">Additional Allowances: {{ tableBTotals.sumAdditional }}</span>
                <span class="summary-value">Pandemic Allowance: {{ tableBTotals.sumPandemic }}</span>
              </div>
              <div class="summary-row grandtotal">
                <span class="summary-label">Grandtotal in RMB</span>
                <span class="summary-value">{{ tableBTotals.grandTotal }}</span>
              </div>
            </div>
            <div v-if="tableBData.length">
              <el-table :data="tableBData" border style="width: 100%;" :span-method="mergeBCells" size="small">
                <el-table-column prop="jobNumber" label="JOB NUMBER" width="120" :header-cell-style="headerStyle">
                  <template #default="{ row }">
                    <el-input v-if="enableAllowanceEditing" v-model="row.jobNumber" size="small" style="width:100%" />
                    <span v-else>{{ row.jobNumber }}</span>
                  </template>
                </el-table-column>
                <el-table-column prop="testDate" label="Survey / Testing Date" width="180" :header-cell-style="headerStyle">
                  <template #default="{ row }">
                    <el-input v-if="enableAllowanceEditing" v-model="row.testDate" size="small" style="width:100%" />
                    <span v-else>{{ row.testDate }}</span>
                  </template>
                </el-table-column>
                <el-table-column prop="commencedTime" label="Commenced Time" width="130" :header-cell-style="headerStyle">
                  <template #default="{ row }">
                    <el-input v-if="enableAllowanceEditing" v-model="row.commencedTime" size="small" style="width:100%" />
                    <span v-else>{{ row.commencedTime }}</span>
                  </template>
                </el-table-column>
                <el-table-column prop="completedTime" label="Completed Time" width="130" :header-cell-style="headerStyle">
                  <template #default="{ row }">
                    <el-input v-if="enableAllowanceEditing" v-model="row.completedTime" size="small" style="width:100%" />
                    <span v-else>{{ row.completedTime }}</span>
                  </template>
                </el-table-column>
                <el-table-column prop="location" label="Location" width="120" :header-cell-style="headerStyle">
                  <template #default="{ row }">
                    <el-input v-if="enableAllowanceEditing" v-model="row.location" size="small" style="width:100%" />
                    <span v-else>{{ row.location }}</span>
                  </template>
                </el-table-column>
                <el-table-column prop="mealAllowances" label="Meal Allowances" width="150" :header-cell-style="headerStyle">
                  <template #default="{ row }">
                    <el-input-number v-if="enableAllowanceEditing" v-model="row.mealAllowances" :min="0" size="small" style="width:100%" />
                    <span v-else>{{ row.mealAllowances }}</span>
                  </template>
                </el-table-column>
                <el-table-column prop="additionalAllowances" label="Additional Allowances" width="180" :header-cell-style="headerStyle">
                  <template #default="{ row }">
                    <el-input-number v-if="enableAllowanceEditing" v-model="row.additionalAllowances" :min="0" size="small" style="width:100%" />
                    <span v-else>{{ row.additionalAllowances }}</span>
                  </template>
                </el-table-column>
                <el-table-column prop="pandemicAllowance" label="疫情补贴" width="120" :header-cell-style="headerStyle">
                  <template #default="{ row }">
                    <el-input-number v-if="enableAllowanceEditing" v-model="row.pandemicAllowance" :min="0" size="small" style="width:100%" />
                    <span v-else>{{ row.pandemicAllowance }}</span>
                  </template>
                </el-table-column>
                <el-table-column prop="totalHours" label="Total Hours" width="120" :header-cell-style="headerStyle">
                  <template #default="{ row }">
                    <el-input v-if="enableAllowanceEditing" v-model="row.totalHours" size="small" style="width:100%" />
                    <span v-else>{{ row.totalHours }}</span>
                  </template>
                </el-table-column>
              </el-table>
            </div>
            <div v-else class="empty-state">暂无 B 表数据，请先上传 A 表。</div>
          </div>
        </el-tab-pane>
      </el-tabs>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, watch, nextTick, computed, onMounted } from 'vue'
import { usePortTrafficStore } from '@/stores/portTraffic'
import { exportToExcel, exportTablesZip, exportTrafficToExcel, exportSummaryToExcel, importFromExcel, importPortTrafficList } from '@/utils/excel'

interface TableARow {
  id: string
  testDate: string
  location: string
  workTime: string
  allowance: number
  name?: string
  payMonth?: string
  portFee?: number
  remark?: string
  commencedTime1: string
  completedTime1: string
  commencedTime2: string
  completedTime2: string
}

interface TableBRow {
  jobNumber: string
  testDate: string
  location: string
  commencedTime: string
  completedTime: string
  mealAllowances: number
  additionalAllowances: number
  pandemicAllowance: number
  totalHours: string
}

interface TableTrafficRow {
  outdoorDate: string
  payMonth: string
  employeeId: string
  name: string
  region: string
  dock: string
  address: string
  trafficAllowance: number
}

// ---------------------- A表数据（报销审核表） ----------------------
const tableAData = ref<TableARow[]>([])
const fileInput = ref<HTMLInputElement | null>(null)
const uploadStatus = ref('')

// ---------------------- B表数据（自动生成） ----------------------
const tableBData = ref<TableBRow[]>([])
const tableTrafficData = ref<TableTrafficRow[]>([])
const portTrafficStore = usePortTrafficStore()
const headerApplicant = ref('')
const headerBranch = ref('')
const headerClaimDuration = ref('')
const activeTab = ref('traffic')
const portTrafficInput = ref<HTMLInputElement | null>(null)

// ---------------------- 默认值配置（可自定义） ----------------------
// 当 A 表中缺少工号/姓名时，使用以下默认值
const defaultEmployeeId = ref('123')
const defaultName = ref('admin')

// ---------------------- 交通明细表自定义配置 ----------------------
interface TrafficColumnConfig {
  prop: keyof TableTrafficRow
  label: string
  width: string | number
  visible: boolean
  editable: boolean
}

const trafficColumns = ref<TrafficColumnConfig[]>([
  { prop: 'outdoorDate', label: '外勤日期', width: '130', visible: true, editable: true },
  { prop: 'payMonth', label: '支付月份', width: '120', visible: true, editable: true },
  { prop: 'employeeId', label: '工号', width: '130', visible: true, editable: true },
  { prop: 'name', label: '姓名', width: '130', visible: true, editable: true },
  { prop: 'region', label: '地区', width: '120', visible: true, editable: true },
  { prop: 'dock', label: '码头', width: '140', visible: true, editable: true },
  { prop: 'address', label: '地址', width: 'auto', visible: true, editable: true },
  { prop: 'trafficAllowance', label: '交通津贴金额', width: '140', visible: true, editable: true },
])

const showColumnSettings = ref(false)
const enableRowEditing = ref(true)
const enableAllowanceEditing = ref(false)

const selectedColumns = computed({
  get: () => trafficColumns.value.filter(col => col.visible).map(col => col.prop),
  set: (val) => {
    trafficColumns.value.forEach(col => {
      col.visible = val.includes(col.prop)
    })
  }
})

function toggleColumn(prop: keyof TableTrafficRow) {
  const col = trafficColumns.value.find(c => c.prop === prop)
  if (col) {
    col.visible = !col.visible
  }
}

const tableBTotals = computed(() => {
  const sumMeal = tableBData.value.reduce((sum, row) => sum + Number(row.mealAllowances || 0), 0)
  const sumAdditional = tableBData.value.reduce((sum, row) => sum + Number(row.additionalAllowances || 0), 0)
  const sumPandemic = tableBData.value.reduce((sum, row) => sum + Number(row.pandemicAllowance || 0), 0)
  const grandTotal = sumMeal + sumAdditional + sumPandemic
  return { sumMeal, sumAdditional, sumPandemic, grandTotal }
})

// 交通明细表金额合计
const trafficTotalAmount = computed(() => {
  return tableTrafficData.value.reduce((sum, row) => sum + Number(row.trafficAllowance || 0), 0)
})

// ---------------------- 核心逻辑：A表 / 港口清单 / 默认值变化时，自动生成交通明细表 ----------------------
watch(
  [tableAData, () => portTrafficStore.items, defaultEmployeeId, defaultName],
  () => {
    generateTrafficTable()
    generateBTable()
  },
  { deep: true, immediate: true }
)

// 港口交通费等于此值时新增空白行（根据业务规则：目前是 400）
const PORT_FEE_BLANK_ROW_VALUE = 200

function formatOutdoorDate(value: string) {
  if (!value) return ''
  const cleaned = value.replace(/-/g, '/').replace(/日/g, '').trim()
  const parts = cleaned.split(/\D+/).filter(Boolean)
  if (parts.length >= 3) {
    const year = parts[0]
    const month = String(Number(parts[1]))
    const day = String(Number(parts[2]))
    return `${year}/${month}/${day}`
  }
  return cleaned
}

function formatPayMonthFromDate(value: string) {
  if (!value) return ''
  const cleaned = value.replace(/-/g, '/').replace(/日/g, '').trim()
  const parts = cleaned.split(/\D+/).filter(Boolean)
  if (parts.length >= 2) {
    const year = parts[0]
    const month = String(Number(parts[1]))
    return `${year}年${month}月`
  }
  return cleaned
}

onMounted(() => {
  portTrafficStore.initialize()
})


const portTrafficStatus = computed(() => portTrafficStore.status)

function triggerPortTrafficUpload() {
  portTrafficInput.value?.click()
}

async function handlePortTrafficFileChange(event: Event) {
  const files = (event.target as HTMLInputElement).files
  if (!files?.length) return
  const file = files[0]
  if (!file) return

  try {
    const rows = await portTrafficStore.loadFromFile(file)
    if (rows.length) {
      generateTrafficTable()
    } else {
      // loadFromFile will already set store.status when no data found or an error occurs
    }
  } catch (error) {
    portTrafficStore.status = error instanceof Error ? error.message : '加载港口交通清单失败'
    console.error(error)
  } finally {
    if (portTrafficInput.value) portTrafficInput.value.value = ''
  }
}

function findPortTrafficByLocation(location: string, remark?: string) {
  // 使用 store 中已实现的 精确 + 模糊 查找
  // 若传入了 remark，也将其作为关键词参与匹配
  if (remark) {
    // 以 "检验地点 + 备注" 组合词搜索，提高命中率
    return portTrafficStore.findByLocation(`${location} ${remark}`)
  }
  return portTrafficStore.findByLocation(location)
}

/**
 * 根据 A 表数据 + 港口交通清单 生成交通明细表。
 * 规则：
 *  1) 对每行 A 表数据，提取 外勤日期 / 支付月份 / 工号 / 姓名 — 来自 A 表
 *  2) 若港口交通费 === 400，先新增一行"空白行"（外勤日期/支付月份/工号/姓名一致，其他空，金额=400）
 *  3) 再按 "检验地点 + 备注" 检索港口交通清单：
 *       - 命中：将清单中的 地区/码头/地址/交通津贴金额 完整写入
 *       - 未命中：地区 = 检验地点；码头/地址 = 空；金额 = 港口交通费（A 表值）
 *  姓名：优先 A 表姓名；港口清单姓名仅做补充
 */
function generateTrafficTable() {
  const result: TableTrafficRow[] = []

  tableAData.value.forEach(row => {
    const outdoorDate = formatOutdoorDate(row.testDate)
    const payMonth = (row.payMonth || '').trim() || formatPayMonthFromDate(row.testDate)
    // 工号：固定使用默认值，可自定义修改
    const employeeId = (defaultEmployeeId.value || '').trim() || '123'
    // 姓名：固定使用默认值，可自定义修改
    const rowName = (defaultName.value || '').trim() || 'admin'
    // 备注字段：用于辅助检索港口交通清单
    const remark = (row.remark || '').trim()

    // 规则 1：港口交通费 === 400 时新增一行"空白行"
    // 金额固定为 200
    const portFee = Number(row.portFee !== undefined && row.portFee !== null ? row.portFee : row.allowance)
    if (portFee === PORT_FEE_BLANK_ROW_VALUE) {
      result.push({
        outdoorDate,
        payMonth,
        employeeId,
        name: rowName,
        region: '',
        dock: '',
        address: '',
        trafficAllowance: 200
      })
    }

    // 规则 2：按检验地点（+ 备注）匹配港口交通清单
    //         命中 → 将清单中的 地区/码头/地址/金额 写入交通明细表
    //         金额固定为 200
    const portItem = findPortTrafficByLocation(row.location, remark || undefined)
    if (portItem) {
      const rawRegion = (portItem.region || '').trim()
      const rawDock = (portItem.dock || '').trim()
      const rawAddress = (portItem.address || '').trim()
      const portLocation = (portItem.location || '').trim()

      // region：优先清单 region；为空时用清单 location
      const region = rawRegion || portLocation || ''
      // dock：优先清单 dock；为空时用清单 location
      const dock = rawDock || portLocation || ''
      // address：仅来自清单，为空保持空
      const address = rawAddress
      // 姓名：优先清单中的 name；缺失时使用 A 表 name
      const name = (portItem.name?.trim()) || rowName
      result.push({
        outdoorDate,
        payMonth,
        employeeId,
        name,
        region,
        dock,
        address,
        trafficAllowance: 200
      })
      return
    }

    // 未命中清单 → 地区 = 检验地点，码头/地址为空，金额固定为 200
    result.push({
      outdoorDate,
      payMonth,
      employeeId,
      name: rowName,
      region: row.location?.trim() || '',
      dock: '',
      address: '',
      trafficAllowance: 200
    })
  })

  tableTrafficData.value = result
}

function preserveDateOnly(value: string) {
  if (!value) return ''
  const trimmed = value.trim()
  const match = trimmed.match(/^(.+?)(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)$/)
  return match?.[1]?.trim() ?? trimmed
}

function getWorkDurationHours(value: string): number | null {
  if (!value) return null
  const trimmed = value.trim()

  const rangeMatch = trimmed.match(/(\d{1,2})(?::|\.)(\d{1,2})\s*[-至～~]\s*(\d{1,2})(?::|\.)(\d{1,2})/)
  if (rangeMatch) {
    const start = Number(rangeMatch[1]) + Number(rangeMatch[2]) / 60
    const end = Number(rangeMatch[3]) + Number(rangeMatch[4]) / 60
    return end >= start ? end - start : end + 24 - start
  }

  const hourMatch = trimmed.match(/^(\d{1,2})(?:[:.](\d{1,2}))?$/)
  if (hourMatch) {
    return Number(hourMatch[1]) + (hourMatch[2] ? Number(hourMatch[2]) / 60 : 0)
  }

  return null
}

function splitWorkTime(workTime: string) {
  const durationHours = getWorkDurationHours(workTime)
  const secondSegmentEnd = durationHours !== null && durationHours > 14.5 ? '21:30:00' : '19:00:00'
  return {
    commencedTime1: '6:30:00',
    completedTime1: '13:00:00',
    commencedTime2: '13:30:00',
    completedTime2: secondSegmentEnd
  }
}

function generateBTable() {
  tableBData.value = [] // 清空B表旧数据

  tableAData.value.forEach(rowA => {
    const allowance = Number(rowA.allowance || 0)
    const baseRow = {
      jobNumber: rowA.id,
      testDate: preserveDateOnly(rowA.testDate),
      location: rowA.location
    }

    const row1 = {
      ...baseRow,
      commencedTime: rowA.commencedTime1,
      completedTime: rowA.completedTime1,
      mealAllowances: 70,
      additionalAllowances: 0,
      pandemicAllowance: 0,
      totalHours: '6:30:00'
    }

    const row2 = {
      ...baseRow,
      commencedTime: rowA.commencedTime2,
      completedTime: rowA.completedTime2,
      mealAllowances: 50,
      additionalAllowances: allowance > 120 ? allowance - 120 : 0,
      pandemicAllowance: 0,
      totalHours: rowA.completedTime2 === '21:30:00' ? '8:00:00' : '5:30:00'
    }

    tableBData.value.push(row1)
    tableBData.value.push(row2)
  })
}

function mergeBCells({ row, column, rowIndex }: { row: TableBRow; column: any; rowIndex: number }) {
  const mergeFields = ['jobNumber', 'testDate', 'location']
  if (!mergeFields.includes(column.property)) {
    return [1, 1]
  }

  if (rowIndex > 0) {
    const prev = tableBData.value[rowIndex - 1]
    if (prev && prev.jobNumber === row.jobNumber && prev.testDate === row.testDate && prev.location === row.location) {
      return [0, 0]
    }
  }

  let rowspan = 1
  for (let i = rowIndex + 1; i < tableBData.value.length; i++) {
    const next = tableBData.value[i]
    if (next && next.jobNumber === row.jobNumber && next.testDate === row.testDate && next.location === row.location) {
      rowspan++
    } else {
      break
    }
  }

  return [rowspan, 1]
}

// ---------------------- 辅助方法：添加/删除A表行 ----------------------
function triggerFileUpload() {
  fileInput.value?.click()
}

function buildAItem(row: TableARow): TableARow {
  const segments = splitWorkTime(row.workTime)
  return {
    ...row,
    ...segments,
    name: row.name,
    payMonth: row.payMonth,
    portFee: row.portFee,
    remark: row.remark
  }
}

async function handleFileChange(event: Event) {
  const files = (event.target as HTMLInputElement).files
  if (!files?.length) return
  const file = files[0]
  if (!file) return

  try {
    // Ensure port traffic cache is initialized before parsing A 表
    if (!portTrafficStore.loaded) {
      await portTrafficStore.initialize()
    }

    const rows = await importFromExcel(file)
    tableAData.value = rows.map(row => buildAItem({
      ...row,
      commencedTime1: '6:30:00',
      completedTime1: '13:00:00',
      commencedTime2: '13:30:00',
      completedTime2: '19:00:00'
    }))
    if (rows.length > 0) {
      uploadStatus.value = `已成功上传 ${rows.length} 条 A 表数据，正在生成 B 表...`
      await nextTick()
      if (tableBData.value.length > 0) {
        uploadStatus.value = `已生成 ${tableBData.value.length} 条 B 表数据`
      } else {
        uploadStatus.value = 'A 表已上传，但未生成 B 表数据，请检查 allowance 字段是否为 160'
      }
    } else {
      uploadStatus.value = '未解析到 A 表数据，请检查文件表头是否包含 id、testDate、location、allowance'
    }
  } catch (error) {
    uploadStatus.value = error instanceof Error ? error.message : '上传失败，请检查文件格式或内容'
    console.error(error)
  } finally {
    if (fileInput.value) {
      fileInput.value.value = ''
    }
  }
}

// ---------------------- 导出Excel ----------------------
function handleExport() {
  const headerInfo = {
    title: 'Allowance Monthly Report',
    company: 'Shanghai Orient Intertek Testing Services Company Limited',
    applicantValue: headerApplicant.value,
    branchValue: headerBranch.value,
    claimValue: headerClaimDuration.value
  }
  exportToExcel(tableBData.value, 'Allowance Details', headerInfo)
}

function handleExportTraffic() {
  exportTrafficToExcel(tableTrafficData.value, '交通费津贴明细')
}

function handleExportSummary() {
  exportSummaryToExcel(tableTrafficData.value, '交通费津贴合计表')
}

function handleExportAllZip() {
  const headerInfo = {
    title: 'Allowance Monthly Report',
    company: 'Shanghai Orient Intertek Testing Services Company Limited',
    applicantValue: headerApplicant.value,
    branchValue: headerBranch.value,
    claimValue: headerClaimDuration.value
  }
  exportTablesZip(tableTrafficData.value, tableBData.value, 'Allowance_Tables', headerInfo)
}

const headerStyle = {
  background: '#f5f7fa',
  color: '#333',
  fontWeight: '700',
  fontSize: '13px',
  textAlign: 'center'
}
</script>

<style scoped>
.container {
  padding: 20px;
}
.header-form {
  display: flex;
  flex-direction: column;
  gap: 12px;
  margin-bottom: 12px;
}
.header-row {
  display: flex;
  align-items: flex-end;
  gap: 16px;
}
.field-group {
  display: flex;
  align-items: center;
  gap: 8px;
}
.field-label {
  font-size: 12px;
  color: #333;
  min-width: 80px;
  text-align: right;
  font-weight: 700;
}
.line-field {
  min-width: 200px;
  padding: 0 8px;
  height: 30px;
  line-height: 30px;
}
.no-underline {
  border-bottom: none;
}
.left-align .field-label {
  text-align: left;
}
.left-align .line-field {
  text-align: left;
}
.applicant-group .line-field {
  min-width: 220px;
}
.id-group .line-field {
  min-width: 140px;
}
.branch-group .line-field {
  min-width: 180px;
}
.claim-group .line-field {
  min-width: 360px;
}
.blank-row {
  height: 10px;
}
.input-row {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
}
.input-row input {
  min-width: 180px;
  padding: 4px 8px;
}
.button-row {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  margin-bottom: 12px;
}
.defaults-row {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 10px;
  padding: 10px 14px;
  margin-bottom: 12px;
  background: #f5f9ff;
  border: 1px solid #d9ecff;
  border-radius: 6px;
}
.defaults-label {
  font-size: 13px;
  font-weight: 700;
  color: #303133;
}
.defaults-hint {
  font-size: 12px;
  color: #909399;
  margin-left: 4px;
}
.status-row {
  display: flex;
  flex-wrap: wrap;
  gap: 14px;
  margin-bottom: 20px;
  color: #606266;
}
.tabs-wrapper {
  margin-top: 20px;
}
.table-card {
  width: 100%;
  background: #ffffff;
  border: 1px solid #ebeef5;
  border-radius: 8px;
  padding: 16px;
}
.table-header-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 16px;
}
.table-header-row h2 {
  margin-bottom: 0;
  font-size: 16px;
  font-weight: 700;
}
.table-actions {
  display: flex;
  align-items: center;
  gap: 12px;
}
.column-settings-panel {
  display: flex;
  flex-wrap: wrap;
  gap: 16px;
  padding: 12px 16px;
  margin-bottom: 16px;
  background: #f5f9ff;
  border: 1px solid #d9ecff;
  border-radius: 6px;
}
.settings-hint {
  width: 100%;
  font-size: 12px;
  color: #67c23a;
  font-weight: 500;
}
.table-card h2 {
  margin-bottom: 16px;
  font-size: 16px;
  font-weight: 700;
}
.empty-state {
  color: #909399;
  min-height: 100px;
  display: flex;
  align-items: center;
  justify-content: center;
  border: 1px dashed #dcdfe6;
  border-radius: 6px;
  padding: 20px;
  background: #fafafa;
}
.table-summary {
  margin-bottom: 16px;
  padding: 12px;
  border: 1px solid #ebeef5;
  border-radius: 6px;
  background: #f7f8fa;
}
.summary-row {
  display: flex;
  flex-wrap: wrap;
  gap: 16px;
  align-items: center;
  margin-bottom: 8px;
  font-size: 14px;
}
.summary-row.grandtotal {
  font-weight: 700;
}
.summary-label {
  min-width: 180px;
}
.summary-value {
  min-width: 160px;
}
.table-summary {
  margin: 16px 0;
  padding: 12px;
  border: 1px solid #ebeef5;
  border-radius: 6px;
  background: #f7f8fa;
}
.summary-row {
  display: flex;
  flex-wrap: wrap;
  gap: 16px;
  align-items: center;
  margin-bottom: 8px;
  font-size: 14px;
}
.summary-row.grandtotal {
  font-weight: 700;
}
.summary-label {
  min-width: 180px;
}
.summary-value {
  min-width: 160px;
}
.traffic-total-row {
  display: flex;
  justify-content: flex-end;
  align-items: center;
  gap: 12px;
  margin-top: 12px;
  padding: 12px 16px;
  background: #f0f9eb;
  border: 1px solid #e1f3d8;
  border-radius: 6px;
}
.traffic-total-label {
  font-size: 14px;
  font-weight: 700;
  color: #303133;
}
.traffic-total-value {
  font-size: 16px;
  font-weight: 700;
  color: #67c23a;
}
</style>