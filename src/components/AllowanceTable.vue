<template>
  <div class="container">
    <el-card class="header-card" shadow="hover">
      <template #header>
        <div class="card-header">
          <span class="title">上传报销审核表数据</span>
          <el-tag type="primary" effect="dark" size="small">
            <el-icon><Document /></el-icon>
            A表数据: {{ tableAData.length }} 条
          </el-tag>
        </div>
      </template>

      <div class="button-row">
        <el-upload
          :auto-upload="false"
          :show-file-list="false"
          accept=".xlsx,.csv"
          :on-change="handleFileChange"
        >
          <el-button type="primary">
            <el-icon><Upload /></el-icon>
            选择上传 A 表
          </el-button>
        </el-upload>

        <el-upload
          :auto-upload="false"
          :show-file-list="false"
          accept=".xlsx,.csv"
          :on-change="handlePortTrafficFileChange"
        >
          <el-button type="warning">
            <el-icon><Upload /></el-icon>
            上传港口交通清单
          </el-button>
        </el-upload>

        <el-divider direction="vertical" />

        <el-button type="primary" :disabled="!tableBData.length || !tableTrafficData.length" @click="handleExportAllZip">
          <el-icon><FolderOpened /></el-icon>
          导出全部（ZIP）
        </el-button>
      </div>

      <el-divider />

      <el-form :inline="true" class="defaults-form">
        <el-form-item label="默认工号">
          <el-input
            v-model="defaultEmployeeId"
            placeholder="默认工号 (例如 123)"
            style="width: 200px"
            clearable
          />
        </el-form-item>
        <el-form-item label="默认姓名">
          <el-input
            v-model="defaultName"
            placeholder="默认姓名 (例如 admin)"
            style="width: 200px"
            clearable
          />
        </el-form-item>
        <el-form-item>
          <el-text type="info" size="small">
            （当 A 表中没有工号/姓名时，将使用此处配置的值）
          </el-text>
        </el-form-item>
      </el-form>

      <el-alert
        v-if="uploadStatus || portTrafficStatus"
        :title="uploadStatus || portTrafficStatus"
        :type="uploadStatus.includes('失败') || portTrafficStatus.includes('失败') ? 'error' : 'success'"
        :closable="false"
        show-icon
      />
    </el-card>

    <el-tabs v-model="activeTab" type="border-card" class="main-tabs">
      <el-tab-pane name="traffic">
        <template #label>
          <el-badge :value="tableTrafficData.length" :hidden="!tableTrafficData.length" type="primary">
            <span>交通明细表</span>
          </el-badge>
        </template>

        <el-card class="table-card" shadow="never">
          <template #header>
            <div class="table-header-row">
              <div class="header-left">
                <el-icon class="table-icon"><Location /></el-icon>
                <span>交通明细表</span>
              </div>
              <div class="header-right">
                <el-button type="text" @click="showColumnSettings = !showColumnSettings">
                  <el-icon><Setting /></el-icon>
                  列设置
                </el-button>
                <el-switch
                  v-model="enableRowEditing"
                  active-text="可编辑"
                  inactive-text="只读"
                  size="small"
                />
              </div>
            </div>
          </template>

          <el-collapse-transition>
            <el-card v-if="showColumnSettings" class="column-settings-card" shadow="hover">
              <template #header>
                <div class="settings-header">
                  <el-icon><InfoFilled /></el-icon>
                  <span>列显隐控制</span>
                </div>
              </template>
              <el-text type="info" size="small" class="settings-hint">
                提示：所有字段均可编辑，点击单元格即可修改
              </el-text>
              <el-checkbox-group v-model="selectedColumns" class="column-checkboxes">
                <el-checkbox
                  v-for="col in trafficColumns"
                  :key="col.prop"
                  :label="col.prop"
                  @change="toggleColumn(col.prop)"
                >
                  {{ col.label }}
                </el-checkbox>
              </el-checkbox-group>
            </el-card>
          </el-collapse-transition>

          <el-table
            v-if="tableTrafficData.length"
            :data="tableTrafficData"
            border
            stripe
            size="small"
            class="traffic-table"
          >
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
                    :precision="2"
                    size="small"
                    style="width:100%"
                  />
                  <el-autocomplete
                    v-else-if="col.editable && enableRowEditing && (col.prop === 'region' || col.prop === 'dock')"
                    v-model="row[col.prop]"
                    :fetch-suggestions="(queryString: string, cb: any) => suggestPortTraffic(row, col.prop as 'region' | 'dock', queryString, cb)"
                    :trigger-on-focus="true"
                    clearable
                    size="small"
                    placeholder="输入以检索港口交通清单"
                    style="width:100%"
                    @select="(item: any) => applyPortSuggestion(row, col.prop as 'region' | 'dock', item)"
                  >
                    <template #prefix>
                      <el-icon><Search /></el-icon>
                    </template>
                  </el-autocomplete>
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

          <el-empty v-else description="暂无交通明细数据，请先上传 A 表或港口交通清单" />

          <el-row v-if="tableTrafficData.length" :gutter="16" class="statistics-row">
            <el-col :span="8">
              <el-statistic title="记录数" :value="tableTrafficData.length" />
            </el-col>
            <el-col :span="8">
              <el-statistic title="金额合计" :value="trafficTotalAmount" :precision="2" prefix="¥" />
            </el-col>
            <el-col :span="8">
              <el-statistic title="平均金额" :value="trafficTotalAmount / tableTrafficData.length" :precision="2" prefix="¥" />
            </el-col>
          </el-row>
        </el-card>
      </el-tab-pane>

      <el-tab-pane name="allowance">
        <template #label>
          <el-badge :value="tableBData.length" :hidden="!tableBData.length" type="success">
            <span>津贴明细表</span>
          </el-badge>
        </template>

        <el-card class="table-card" shadow="never">
          <template #header>
            <div class="table-header-row">
              <div class="header-left">
                <el-icon class="table-icon"><Money /></el-icon>
                <span>自动生成的津贴明细（B表）</span>
              </div>
              <div class="header-right">
                <el-switch
                  v-model="enableAllowanceEditing"
                  active-text="可编辑"
                  inactive-text="只读"
                  size="small"
                />
              </div>
            </div>
          </template>

          <el-descriptions v-if="tableBData.length" :column="2" border class="summary-descriptions">
            <el-descriptions-item>
              <template #label>
                <el-icon><Coin /></el-icon>
                Subtotal in RMB
              </template>
              <el-text type="primary">Meal Allowances: {{ tableBTotals.sumMeal }}</el-text>
              <el-divider direction="vertical" />
              <el-text type="primary">Additional Allowances: {{ tableBTotals.sumAdditional }}</el-text>
              <el-divider direction="vertical" />
              <el-text type="primary">Pandemic Allowance: {{ tableBTotals.sumPandemic }}</el-text>
            </el-descriptions-item>
            <el-descriptions-item>
              <template #label>
                <el-icon><Tickets /></el-icon>
                Grandtotal in RMB
              </template>
              <el-text type="success" size="large" tag="b">{{ tableBTotals.grandTotal }}</el-text>
            </el-descriptions-item>
          </el-descriptions>

          <el-table
            v-if="tableBData.length"
            :data="tableBData"
            border
            stripe
            size="small"
            class="allowance-table"
            :span-method="mergeBCells"
          >
            <el-table-column prop="jobNumber" label="JOB NUMBER" width="120" :header-cell-style="headerStyle">
              <template #default="{ row }">
                <el-tag v-if="enableAllowanceEditing" size="small" effect="plain">
                  <el-input v-model="row.jobNumber" size="small" style="width:100%" />
                </el-tag>
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

          <el-empty v-else description="暂无 B 表数据，请先上传 A 表" />
        </el-card>
      </el-tab-pane>
    </el-tabs>
  </div>
</template>

<script setup lang="ts">
import { ref, watch, nextTick, computed, onMounted } from 'vue'
import { ElMessage } from 'element-plus'
import {
  Upload, Download, FolderOpened, Document, Location, Money, Coin, Tickets,
  Setting, InfoFilled
} from '@element-plus/icons-vue'
import { usePortTrafficStore } from '@/stores/portTraffic'
import {
  exportToExcel,
  exportTablesZip,
  exportTrafficToExcel,
  exportSummaryToExcel,
  importFromExcel
} from '@/utils/excel'

interface TableARow {
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
const uploadStatus = ref('')

// ---------------------- B表数据（自动生成） ----------------------
const tableBData = ref<TableBRow[]>([])
const tableTrafficData = ref<TableTrafficRow[]>([])
const portTrafficStore = usePortTrafficStore()
const activeTab = ref('traffic')

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

async function handlePortTrafficFileChange(uploadFile: any) {
  try {
    const rows = await portTrafficStore.loadFromFile(uploadFile.raw)
    if (rows.length) {
      generateTrafficTable()
      ElMessage.success({
        message: `成功加载 ${rows.length} 条港口交通数据`,
        duration: 3000
      })
    }
  } catch (error) {
    portTrafficStore.status = error instanceof Error ? error.message : '加载港口交通清单失败'
    ElMessage.error({
      message: portTrafficStore.status,
      duration: 5000
    })
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

// ---------------------- 交通明细表中"地区/码头"字段的实时检索建议 ----------------------
function queryPortTrafficSuggestions(keyword: string) {
  const list = portTrafficStore.items || []
  if (!list.length) return []
  const kw = String(keyword || '').trim().toLowerCase()
  if (!kw) {
    return list.slice(0, 10).map((item, idx) => ({
      id: `port-${idx}`,
      value: '',
      region: item.region || '',
      dock: item.dock || '',
      address: item.address || '',
      amount: item.amount,
      name: item.name || '',
      _raw: item
    }))
  }
  const matched = list.filter(item => {
    const region = (item.region || '').toLowerCase()
    const dock = (item.dock || '').toLowerCase()
    const address = (item.address || '').toLowerCase()
    const location = (item.location || '').toLowerCase()
    return region.includes(kw) || dock.includes(kw) || address.includes(kw) || location.includes(kw)
  })
  return matched.slice(0, 20).map((item, idx) => ({
    id: `port-${idx}`,
    value: '',
    region: item.region || '',
    dock: item.dock || '',
    address: item.address || '',
    amount: item.amount,
    name: item.name || '',
    _raw: item
  }))
}

function suggestPortTraffic(row: TableTrafficRow, colProp: 'region' | 'dock', queryString: string, callback: any) {
  const suggestions = queryPortTrafficSuggestions(queryString)
  // 候选项的 value 字段是 el-autocomplete 用于显示的文本，根据当前编辑的列变化
  const suggestionsWithValue = suggestions.map(s => ({
    ...s,
    value: colProp === 'region' ? s.region : s.dock
  }))
  callback(suggestionsWithValue)
}

function applyPortSuggestion(row: TableTrafficRow, colProp: 'region' | 'dock', item: any) {
  if (!item || !item._raw) return
  const port = item._raw
  row.region = (port.region || '').trim() || row.region
  row.dock = (port.dock || '').trim() || row.dock
  row.address = (port.address || '').trim() || row.address
  row.trafficAllowance = 200
  if (colProp === 'region') {
    row.region = item.value || row.region
  } else if (colProp === 'dock') {
    row.dock = item.value || row.dock
  }
  ElMessage.success({
    message: `已根据港口清单自动填充：${row.region} / ${row.dock}`,
    duration: 2000
  })
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
  tableBData.value = []

  tableAData.value.forEach(rowA => {
    const allowance = Number(rowA.allowance || 0)
    const baseRow = {
      jobNumber: rowA.id,
      testDate: preserveDateOnly(rowA.testDate),
      location: rowA.region || rowA.location
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

// ---------------------- 辅助方法：上传A表 ----------------------
async function handleFileChange(uploadFile: any) {
  try {
    if (!portTrafficStore.loaded) {
      await portTrafficStore.initialize()
    }

    const rows = await importFromExcel(uploadFile.raw)
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
        ElMessage.success({
          message: `成功导入 ${rows.length} 条 A 表数据，已生成 ${tableBData.value.length} 条 B 表数据`,
          duration: 3000
        })
      } else {
        uploadStatus.value = 'A 表已上传，但未生成 B 表数据，请检查 allowance 字段是否为 160'
        ElMessage.warning({
          message: 'A 表已上传，但未生成 B 表数据',
          duration: 5000
        })
      }
    } else {
      uploadStatus.value = '未解析到 A 表数据，请检查文件表头是否包含 id、testDate、location、allowance'
      ElMessage.error({
        message: '未解析到 A 表数据',
        duration: 5000
      })
    }
  } catch (error) {
    uploadStatus.value = error instanceof Error ? error.message : '上传失败，请检查文件格式或内容'
    ElMessage.error({
      message: uploadStatus.value,
      duration: 5000
    })
  }
}

// ---------------------- 辅助方法：构建A表数据 ----------------------
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

// ---------------------- 导出Excel ----------------------
function handleExport() {
  const headerInfo = {
    title: 'Allowance Monthly Report',
    company: 'Shanghai Orient Intertek Testing Services Company Limited',
    applicantValue: '',
    branchValue: '',
    claimValue: ''
  }
  exportToExcel(tableBData.value, 'Allowance Details', headerInfo)
  ElMessage.success({
    message: '津贴明细表已导出为 Allowance Details.xlsx',
    duration: 3000
  })
}

function handleExportTraffic() {
  exportTrafficToExcel(tableTrafficData.value, '交通费津贴明细')
  ElMessage.success({
    message: '交通明细表已导出为 交通费津贴明细.xlsx',
    duration: 3000
  })
}

function handleExportSummary() {
  exportSummaryToExcel(tableTrafficData.value, '交通费津贴合计表')
  ElMessage.success({
    message: '交通津贴汇总表已导出为 交通费津贴合计表.xlsx',
    duration: 3000
  })
}

function handleExportAllZip() {
  const headerInfo = {
    title: 'Allowance Monthly Report',
    company: 'Shanghai Orient Intertek Testing Services Company Limited',
    applicantValue: '',
    branchValue: '',
    claimValue: ''
  }
  exportTablesZip(tableTrafficData.value, tableBData.value, 'Allowance_Tables', headerInfo)
  ElMessage.success({
    message: '已导出全部表格为 Allowance_Tables.zip',
    duration: 3000
  })
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
  padding: 0;
  min-height: calc(100vh - 40px);
}

.header-card {
  margin-bottom: 20px;
  border-radius: 12px;
  box-shadow: 0 2px 12px rgba(0, 0, 0, 0.08);
  overflow: hidden;
}

.header-card :deep(.el-card__header) {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  padding: 16px 24px;
  border-bottom: none;
}

.header-card :deep(.el-card__header) .card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.header-card :deep(.el-card__header) .title {
  font-size: 18px;
  font-weight: 600;
  color: #ffffff;
}

.header-card :deep(.el-card__body) {
  padding: 20px 24px;
}

.button-row {
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
  align-items: center;
  margin-bottom: 16px;
}

.defaults-form {
  background: #f8fafc;
  padding: 16px 20px;
  border-radius: 8px;
  border: 1px solid #e2e8f0;
}

.defaults-form :deep(.el-form-item) {
  margin-bottom: 0;
}

.defaults-form :deep(.el-form-item__label) {
  font-weight: 600;
  color: #475569;
}

.main-tabs {
  border-radius: 12px;
  overflow: hidden;
  box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
}

.main-tabs :deep(.el-tabs__header) {
  background: #f8fafc;
  border-bottom: 1px solid #e2e8f0;
}

.main-tabs :deep(.el-tabs__item) {
  font-weight: 500;
  color: #64748b;
  padding: 16px 24px;
}

.main-tabs :deep(.el-tabs__item.is-active) {
  color: #667eea;
  font-weight: 600;
}

.main-tabs :deep(.el-tabs__content) {
  padding: 0;
}

.table-card {
  border: none;
  background: #ffffff;
  min-height: 400px;
}

.table-card :deep(.el-card__header) {
  padding: 16px 20px;
  background: #f8fafc;
  border-bottom: 1px solid #e2e8f0;
  border-radius: 8px 8px 0 0;
}

.table-header-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.header-left {
  display: flex;
  align-items: center;
  gap: 10px;
  font-size: 16px;
  font-weight: 600;
  color: #1e293b;
}

.header-right {
  display: flex;
  align-items: center;
  gap: 16px;
}

.table-icon {
  font-size: 20px;
  color: #667eea;
}

.column-settings-card {
  margin-bottom: 16px;
  border-radius: 8px;
  border: 1px solid #e2e8f0;
}

.settings-header {
  display: flex;
  align-items: center;
  gap: 8px;
  font-weight: 600;
  color: #334155;
}

.settings-hint {
  display: block;
  margin-bottom: 12px;
  color: #64748b;
  font-size: 13px;
}

.column-checkboxes {
  display: flex;
  flex-wrap: wrap;
  gap: 20px;
}

.column-checkboxes :deep(.el-checkbox) {
  font-size: 13px;
  color: #475569;
}

.traffic-table,
.allowance-table {
  border-radius: 8px;
  overflow: hidden;
  border: 1px solid #e2e8f0;
}

.traffic-table :deep(.el-table),
.allowance-table :deep(.el-table) {
  --el-table-border-color: #e2e8f0;
}

.traffic-table :deep(.el-table__header-wrapper),
.allowance-table :deep(.el-table__header-wrapper) {
  background: #f8fafc;
}

.traffic-table :deep(.el-table__body tr:hover>td),
.allowance-table :deep(.el-table__body tr:hover>td) {
  background: #f1f5f9;
}

.traffic-table :deep(.el-table__body tr.el-table__row--striped),
.allowance-table :deep(.el-table__body tr.el-table__row--striped) {
  background: #fafafa;
}

.statistics-row {
  margin-top: 16px;
  padding: 20px;
  background: linear-gradient(135deg, #f0fdf4 0%, #ecfdf5 100%);
  border-radius: 8px;
  border: 1px solid #bbf7d0;
}

.statistics-row :deep(.el-statistic__label) {
  font-size: 13px;
  color: #64748b;
  font-weight: 500;
}

.statistics-row :deep(.el-statistic__content) {
  font-size: 24px;
  font-weight: 700;
  color: #1e293b;
}

.statistics-row :deep(.el-statistic__prefix) {
  font-size: 18px;
  color: #22c55e;
}

.summary-descriptions {
  margin-bottom: 16px;
  border-radius: 8px;
  overflow: hidden;
  border: 1px solid #e2e8f0;
}

.summary-descriptions :deep(.el-descriptions__label) {
  font-weight: 600;
  color: #475569;
  background: #f8fafc;
}

.summary-descriptions :deep(.el-descriptions__content) {
  color: #1e293b;
}

:deep(.el-table__empty-text) {
  padding: 60px 0;
  color: #94a3b8;
}

:deep(.el-switch) {
  margin: 0;
}

:deep(.el-button) {
  font-weight: 500;
}

:deep(.el-autocomplete) {
  --el-input-focus-border-color: #667eea;
}
</style>