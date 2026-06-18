<template>
  <div class="search-container">
    <el-card class="search-card" shadow="hover">
      <template #header>
        <div class="card-header">
          <span>
            <el-icon class="header-icon"><Search /></el-icon>
            港口交通清单搜索
          </span>
          <el-tag :type="store.loaded ? 'success' : 'info'" effect="dark" size="small">
            {{ store.loaded ? `已加载 ${store.items.length} 条数据` : '未加载数据' }}
          </el-tag>
        </div>
      </template>

      <el-space wrap class="search-toolbar">
        <el-input
          v-model="query"
          placeholder="请输入地区 / 码头 / 地址 / 名称"
          clearable
          @input="onKeywordChange"
          style="width: 380px"
        >
          <template #prefix>
            <el-icon><Search /></el-icon>
          </template>
        </el-input>

        <el-button type="primary" @click="clearSearch">
          <el-icon><RefreshRight /></el-icon>
          清空
        </el-button>

        <el-button type="success" @click="reloadFromCache">
          <el-icon><Refresh /></el-icon>
          从缓存加载
        </el-button>
      </el-space>

      <el-divider />

      <div class="search-status">
        <el-text type="info">
          <el-icon><InfoFilled /></el-icon>
          {{ status }}
        </el-text>
        <el-text type="primary">
          <el-icon><Document /></el-icon>
          共 {{ filteredItems.length }} 条匹配结果
        </el-text>
      </div>

      <el-table
        :data="filteredItems"
        stripe
        border
        style="width: 100%"
        size="default"
        v-loading="!loaded"
        :default-sort="{ prop: 'updatedAt', order: 'descending' }"
      >
        <el-table-column prop="location" label="地点" width="140" sortable>
          <template #default="{ row }">
            <el-tag type="primary" effect="plain" size="small">{{ row.location }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column prop="region" label="地区" width="120" sortable />
        <el-table-column prop="dock" label="码头" width="140" sortable />
        <el-table-column prop="address" label="地址" min-width="200" show-overflow-tooltip />
        <el-table-column prop="amount" label="金额" width="120" sortable align="right">
          <template #default="{ row }">
            <el-text type="success" tag="b">¥ {{ row.amount.toFixed(2) }}</el-text>
          </template>
        </el-table-column>
        <el-table-column prop="updatedAt" label="更新时间" width="160" sortable>
          <template #default="{ row }">
            <el-text type="info" size="small">{{ row.updatedAt }}</el-text>
          </template>
        </el-table-column>
      </el-table>

      <el-empty v-if="!filteredItems.length && loaded" description="没有找到匹配的数据" />
    </el-card>
  </div>
</template>

<script setup lang="ts">
import { ref, watch, computed } from 'vue'
import { Search, RefreshRight, Refresh, InfoFilled, Document } from '@element-plus/icons-vue'
import { usePortTrafficStore } from '@/stores/portTraffic'

const store = usePortTrafficStore()
const query = ref('')

const status = computed(() => store.status)
const filteredItems = computed(() => store.filteredItems)
const loaded = computed(() => store.loaded)

const onKeywordChange = (value: string) => {
  store.setKeyword(value)
}

const clearSearch = () => {
  query.value = ''
  store.clearKeyword()
}

const reloadFromCache = () => {
  store.initialize()
}

watch(
  () => query.value,
  (value) => {
    store.setKeyword(value)
  }
)

store.initialize()
</script>

<style scoped>
.search-container {
  padding: 16px;
}

.search-card {
  border-radius: 12px;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 16px;
  font-weight: 600;
}

.header-icon {
  margin-right: 8px;
  font-size: 18px;
  color: var(--el-color-primary);
}

.search-toolbar {
  width: 100%;
  margin-bottom: 8px;
}

.search-status {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 16px;
  padding: 12px 16px;
  background: #f5f7fa;
  border-radius: 8px;
}
</style>
