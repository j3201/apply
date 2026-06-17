import { defineStore } from 'pinia'
import { ref, computed } from 'vue'
import { debounce } from 'lodash-es'
import { importPortTrafficList } from '@/utils/excel'

export interface PortTrafficItem {
  location: string
  region: string
  dock: string
  address: string
  amount: number
  name?: string
  employeeId?: string
  updatedAt?: string
}

interface SearchIndexItem {
  item: PortTrafficItem
  searchText: string
}

export const usePortTrafficStore = defineStore('portTraffic', () => {
  const items = ref<PortTrafficItem[]>([])
  const keyword = ref('')
  const loaded = ref(false)
  const status = ref('未加载港口交通清单')
  const searchIndex = ref<SearchIndexItem[]>([])
  const locationIndex = ref<Map<string, PortTrafficItem>>(new Map())
  const cacheKey = 'portTrafficListCache'

  const buildSearchIndex = (list: PortTrafficItem[]) => {
    searchIndex.value = list.map(item => ({
      item,
      searchText: [item.region, item.dock, item.address, item.location]
        .filter(Boolean)
        .join(' | ')
        .toLowerCase()
    }))

    const locMap = new Map<string, PortTrafficItem>()
    list.forEach(item => {
      if (item.location) {
        locMap.set(item.location.trim().toLowerCase(), item)
      }
    })
    locationIndex.value = locMap
  }

  const filteredItems = computed(() => {
    const query = keyword.value.trim().toLowerCase()
    if (!query) {
      return items.value
    }
    return searchIndex.value
      .filter(entry => entry.searchText.includes(query))
      .map(entry => entry.item)
  })

  const setItems = (list: PortTrafficItem[]) => {
    items.value = list
    loaded.value = true
    status.value = `已加载港口交通清单 ${list.length} 条` 
    buildSearchIndex(list)
    localStorage.setItem(cacheKey, JSON.stringify(list))
  }

  const setKeyword = debounce((value: string) => {
    keyword.value = value
  }, 250)

  const clearKeyword = () => {
    keyword.value = ''
  }

  const clearCache = () => {
    localStorage.removeItem(cacheKey)
    items.value = []
    searchIndex.value = []
    loaded.value = false
    status.value = '已清除港口交通清单缓存'
  }

  const initialize = () => {
    const cached = localStorage.getItem(cacheKey)
    if (cached) {
      try {
        const parsed = JSON.parse(cached) as PortTrafficItem[]
        if (Array.isArray(parsed) && parsed.length) {
          items.value = parsed
          loaded.value = true
          status.value = `从缓存加载 ${parsed.length} 条港口交通清单`
          buildSearchIndex(parsed)
          return
        }
      } catch {
        localStorage.removeItem(cacheKey)
      }
    }
    status.value = '未加载港口交通清单'}

  /**
   * 按"检验地点 + 备注"检索港口交通清单。
   * 策略：精确匹配 > 分词多字段评分匹配 > 简单包含匹配。
   * 重点：优先匹配地区 和码头 字段，命中后将码头/地址/金额写入交通明细表。
   */
  const findByLocation = (location: string) => {
    if (!location) return undefined
    const normalized = location.trim().toLowerCase()
    if (!items.value.length) return undefined

    // 1) 精确匹配 location 字段 - O(1)
    const exact = locationIndex.value.get(normalized)
    if (exact) return exact

    // 2) 将输入分词（按 中文逗号/顿号/空格/英文逗号/斜杠 拆分）
    //    例如 "江西省宜春市, 袁州区 码头" -> ["江西省宜春市", "袁州区", "码头"]
    const tokens = normalized
      .split(/[,，、；;\/\s]+/)
      .map(t => t.trim())
      .filter(t => t && t.length >= 1)

    // 3) 对每条清单记录做评分：
    //    - 重点匹配地区 和码头 字段
    //    - 正/反向包含：字段包含输入词，或输入词包含字段
    //    - 字段优先级：region > dock > location > address
    const tokenCount = tokens.length || 1
    let best: { item: PortTrafficItem; score: number } | null = null

    for (const item of items.value) {
      const itemLoc = (item.location || '').trim().toLowerCase()
      const itemRegion = (item.region || '').trim().toLowerCase()
      const itemDock = (item.dock || '').trim().toLowerCase()
      const itemAddr = (item.address || '').trim().toLowerCase()
      if (!itemLoc && !itemRegion && !itemDock && !itemAddr) continue

      // 快速检查：整串包含 location 字段（最常见场景）
      if (itemLoc && (itemLoc.includes(normalized) || normalized.includes(itemLoc))) {
        return item
      }

      // 评分模式 - 重点匹配 region 和 dock
      let score = 0
      tokens.forEach(tok => {
        if (!tok) return
        // 权重：region 和 dock 最高（3分），location 次之（2分），address 最低（1分）
        // 这样 A 表的"检验地点 + 备注"会优先匹配港口清单的地区和码头
        if (itemRegion && (itemRegion.includes(tok) || tok.includes(itemRegion))) score += 3
        if (itemDock && (itemDock.includes(tok) || tok.includes(itemDock))) score += 3
        if (itemLoc && (itemLoc.includes(tok) || tok.includes(itemLoc))) score += 2
        if (itemAddr && (itemAddr.includes(tok) || tok.includes(itemAddr))) score += 1
      })

      // 归一化 / 数量惩罚：需要至少一定的匹配信号
      const threshold = tokenCount > 1 ? 2 : 2
      if (score >= threshold && (!best || score > best.score)) {
        best = { item, score }
      }
    }

    // 4) 返回最佳候选（若评分足够）
    if (best) return best.item
    return undefined
  }

  const loadFromFile = async (file: File) => {
    status.value = '正在解析港口交通清单...'
    const parsed = await importPortTrafficList(file)
    setItems(parsed)
    return parsed
  }

  return {
    items,
    keyword,
    loaded,
    status,
    filteredItems,
    setItems,
    setKeyword,
    clearKeyword,
    clearCache,
    initialize,
    findByLocation,
    loadFromFile
  }
})
