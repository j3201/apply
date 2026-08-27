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
   * 判断 needle 是否模糊匹配于 haystack（双向包含 + 子串匹配）
   * 返回匹配分数：0 表示不匹配，分数越高匹配度越强
   */
  const fuzzyMatchScore = (needle: string, haystack: string): number => {
    const n = needle.trim()
    const h = haystack.trim()
    if (!n || !h) return 0
    // 完全一致：最高优先级
    if (n === h) return 100
    // 完全一致（忽略大小写）
    if (n.toLowerCase() === h.toLowerCase()) return 95
    // needle 是 haystack 的子串
    if (h.toLowerCase().includes(n.toLowerCase())) return 80 + Math.min(n.length, 20)
    // haystack 是 needle 的子串（例如 needle="广东南沙码头"，haystack="南沙"）
    if (n.toLowerCase().includes(h.toLowerCase())) return 60 + Math.min(h.length, 20)
    return 0
  }

  /**
   * 从检验地点文本中提取可能的"地区关键词"和"码头关键词"
   * 策略：
   *  - 先按常见分隔符拆分
   *  - 再尝试识别省份 + 非省部分（若整体未被分隔符分开）
   */
  const extractRegionDockKeywords = (location: string): { regionKeywords: string[]; dockKeywords: string[] } => {
    const normalized = location.trim().toLowerCase()
    const regionKeywords: string[] = []
    const dockKeywords: string[] = []

    const provinceNames = ['北京', '天津', '河北', '山西', '内蒙古', '辽宁', '吉林', '黑龙江',
      '上海', '江苏', '浙江', '安徽', '福建', '江西', '山东', '河南', '湖北', '湖南',
      '广东', '广西', '海南', '重庆', '四川', '贵州', '云南', '西藏', '陕西', '甘肃',
      '青海', '宁夏', '新疆', '香港', '澳门', '台湾']
    const provinceLowerNames = provinceNames.map(p => p.toLowerCase())

    // 先按分隔符拆
    const parts: string[] = normalized
      .split(/[,，、；;\/\s]+/)
      .map((p: string) => p.trim())
      .filter((p: string): p is string => Boolean(p))

    if (parts.length >= 2) {
      const first = parts[0]!
      // 第一部分若看起来是省份/地区，放入地区关键词；其余放入码头
      if (first && (provinceLowerNames.some(p => first.includes(p)) || parts.length >= 2)) {
        regionKeywords.push(first)
      }
      for (let i = 1; i < parts.length; i++) {
        const part = parts[i]
        if (part) dockKeywords.push(part)
      }
      // 若有多段，也将第一段之后的所有段整体拼接作为码头候选（应对 "南沙 港 一期" → "南沙港一期"）
      const joinedTail = parts.slice(1).filter(Boolean).join('')
      if (joinedTail) dockKeywords.push(joinedTail)
    } else if (parts.length === 1) {
      const only: string = parts[0]!
      if (!only) {
        return { regionKeywords: [], dockKeywords: [] }
      }
      // 尝试从整体中提取省份前缀
      let extractedProvince = ''
      for (const province of provinceLowerNames) {
        if (only.startsWith(province) || only.includes(province)) {
          extractedProvince = province
          break
        }
      }
      if (extractedProvince) {
        regionKeywords.push(extractedProvince)
        const remain = only.replace(extractedProvince, '').trim()
        if (remain) dockKeywords.push(remain)
        // 整体也保留在两个候选中，方便子串匹配
        regionKeywords.push(only)
        dockKeywords.push(only)
      } else {
        // 没有省份信息，两个关键词集合都放入整体，以增加命中率
        regionKeywords.push(only)
        dockKeywords.push(only)
      }
    }

    // 去重
    return {
      regionKeywords: Array.from(new Set(regionKeywords)),
      dockKeywords: Array.from(new Set(dockKeywords))
    }
  }

  /**
   * 按"检验地点"检索港口交通清单（增强模糊匹配版）。
   * 匹配策略（分数高者胜出；任意环节命中最高分即返回）：
   *  1) 精准匹配 location 字段
   *  2) 检验地点中含省份名 → 省份匹配缓存地区字段（精确+模糊）
   *  3) 拆分检验地点为地区+码头 → 与缓存的地区/码头双向模糊匹配
   *  4) 检验地点整体 与 缓存地区 / 缓存码头 双向子串匹配
   *  5) 兜底：address 字段模糊匹配
   */
  const findByLocation = (location: string) => {
    if (!location) return undefined
    const normalized = location.trim().toLowerCase()
    if (!items.value.length) return undefined

    // 1) 精准匹配 location 字段（最高优先级）
    const exact = locationIndex.value.get(normalized)
    if (exact) return exact

    // 预先提取地区/码头关键词
    const { regionKeywords, dockKeywords } = extractRegionDockKeywords(normalized)

    // 2) 省份匹配：若检验地点包含省份名，模糊匹配缓存的地区字段
    const provinceNames = ['北京', '天津', '河北', '山西', '内蒙古', '辽宁', '吉林', '黑龙江',
      '上海', '江苏', '浙江', '安徽', '福建', '江西', '山东', '河南', '湖北', '湖南',
      '广东', '广西', '海南', '重庆', '四川', '贵州', '云南', '西藏', '陕西', '甘肃',
      '青海', '宁夏', '新疆', '香港', '澳门', '台湾']

    interface NormalizedPortItem {
      item: PortTrafficItem
      region: string
      dock: string
      address: string
      loc: string
    }

    // 缓存每项的归一化字段，避免重复计算
    const normalizedItems: NormalizedPortItem[] = items.value.map(item => ({
      item,
      region: (item.region || '').trim().toLowerCase(),
      dock: (item.dock || '').trim().toLowerCase(),
      address: (item.address || '').trim().toLowerCase(),
      loc: (item.location || '').trim().toLowerCase()
    }))

    let bestMatch: { item: PortTrafficItem; score: number } | null = null
    let theBest: PortTrafficItem | undefined = undefined

    const consider = (entry: NormalizedPortItem, score: number) => {
      if (score <= 0) return
      if (!bestMatch || score > bestMatch.score) {
        bestMatch = { item: entry.item, score }
        theBest = entry.item
      }
    }

    for (const entry of normalizedItems) {
      let score = 0

      // —— 3) 地区 + 码头 联合匹配（双向模糊包含） ——
      let regionScore = 0
      for (const kw of regionKeywords) {
        regionScore = Math.max(regionScore, fuzzyMatchScore(kw, entry.region))
      }
      // 将整个检验地点作为地区也匹配一次（应对"广东南沙"这种整体）
      regionScore = Math.max(regionScore, fuzzyMatchScore(normalized, entry.region))

      let dockScore = 0
      for (const kw of dockKeywords) {
        dockScore = Math.max(dockScore, fuzzyMatchScore(kw, entry.dock))
      }
      dockScore = Math.max(dockScore, fuzzyMatchScore(normalized, entry.dock))

      // 地区和码头都命中时加分加权
      if (regionScore > 0 && dockScore > 0) {
        score = Math.max(score, regionScore + dockScore + 20)
      } else {
        score = Math.max(score, regionScore, dockScore)
      }

      // —— 2) 省份精确匹配 ——
      for (const province of provinceNames) {
        const provinceLower = province.toLowerCase()
        if (normalized.includes(provinceLower) && entry.region === provinceLower) {
          score = Math.max(score, 90 + (dockScore > 0 ? dockScore : 0))
        } else if (normalized.includes(provinceLower) && fuzzyMatchScore(provinceLower, entry.region) > 0) {
          score = Math.max(score, 70 + (dockScore > 0 ? Math.floor(dockScore / 2) : 0))
        }
      }

      // —— 4) 整体与缓存地区/码头/location 的模糊匹配 ——
      score = Math.max(score, fuzzyMatchScore(normalized, entry.loc))

      // —— 5) address 模糊匹配（分值稍低） ——
      const addrScore = fuzzyMatchScore(normalized, entry.address)
      if (addrScore > 0) {
        score = Math.max(score, Math.floor(addrScore * 0.7))
      }

      if (score > 0) consider(entry, score)
    }

    return theBest
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
