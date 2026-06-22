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
      .filter(entry => 
        entry.item.location.trim().toLowerCase() === query ||
        entry.item.region.trim().toLowerCase() === query ||
        entry.item.dock.trim().toLowerCase() === query ||
        entry.item.address.trim().toLowerCase() === query
      )
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
   * 按"检验地点"检索港口交通清单。
   * 检验地点由省份名称和码头名称组成，策略：
   * 1. 先按省份名称检索地区字段
   * 2. 再按码头名称检索码头字段
   */
  const findByLocation = (location: string) => {
    if (!location) return undefined
    const normalized = location.trim().toLowerCase()
    if (!items.value.length) return undefined

    // 1) 精准匹配 location 字段
    const exact = locationIndex.value.get(normalized)
    if (exact) return exact

    // 2) 按省份名称检索地区字段
    // 尝试从检验地点中提取省份名称（常见省份名）
    const provinceNames = ['北京', '天津', '河北', '山西', '内蒙古', '辽宁', '吉林', '黑龙江',
      '上海', '江苏', '浙江', '安徽', '福建', '江西', '山东', '河南', '湖北', '湖南',
      '广东', '广西', '海南', '重庆', '四川', '贵州', '云南', '西藏', '陕西', '甘肃',
      '青海', '宁夏', '新疆', '香港', '澳门', '台湾']
    
    for (const province of provinceNames) {
      const provinceLower = province.toLowerCase()
      if (normalized.includes(provinceLower)) {
        const matchByRegion = items.value.find(item => 
          (item.region || '').trim().toLowerCase() === provinceLower
        )
        if (matchByRegion) return matchByRegion
      }
    }

    // 3) 拆分检验地点，尝试分别匹配地区和码头
    // 按常见分隔符拆分：空格、逗号、顿号、斜杠等
    const parts = normalized.split(/[,，、；;\/\s]+/).map(p => p.trim()).filter(p => p)
    
    if (parts.length >= 2) {
      // 尝试组合匹配：第一个部分匹配地区，第二个部分匹配码头
      const regionPart = parts[0]
      const dockPart = parts[1]
      
      const matchByRegionAndDock = items.value.find(item => {
        const itemRegion = (item.region || '').trim().toLowerCase()
        const itemDock = (item.dock || '').trim().toLowerCase()
        return itemRegion === regionPart && itemDock === dockPart
      })
      if (matchByRegionAndDock) return matchByRegionAndDock
    }

    // 4) 按码头名称检索码头字段
    for (const part of parts) {
      const matchByDock = items.value.find(item => 
        (item.dock || '').trim().toLowerCase() === part
      )
      if (matchByDock) return matchByDock
    }

    // 5) 精准匹配 region 字段（完整匹配）
    const matchByRegion = items.value.find(item => 
      (item.region || '').trim().toLowerCase() === normalized
    )
    if (matchByRegion) return matchByRegion

    // 6) 精准匹配 dock 字段（完整匹配）
    const matchByDock = items.value.find(item => 
      (item.dock || '').trim().toLowerCase() === normalized
    )
    if (matchByDock) return matchByDock

    // 7) 精准匹配 address 字段
    const matchByAddress = items.value.find(item => 
      (item.address || '').trim().toLowerCase() === normalized
    )
    if (matchByAddress) return matchByAddress

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
