<template>
  <div class="search-container">
    <div class="search-toolbar">
      <el-input
        v-model="query"
        placeholder="请输入地区 / 码头 / 地址 / 名称"
        clearable
        @input="onKeywordChange"
        style="width: 380px"
      />
      <el-button type="primary" @click="clearSearch">清空</el-button>
      <el-button type="success" @click="reloadFromCache">从缓存加载</el-button>
    </div>
    <div class="search-status">
      <span>{{ status }}</span>
      <span>共 {{ filteredItems.length }} 条匹配结果</span>
    </div>
    <el-table
      :data="filteredItems"
      stripe
      style="width: 100%"
      size="small"
      v-loading="!loaded"
    >
      <el-table-column prop="location" label="地点" width="140" />
      <el-table-column prop="region" label="地区" width="120" />
      <el-table-column prop="dock" label="码头" width="140" />
      <el-table-column prop="address" label="地址" />
      <el-table-column prop="amount" label="金额" width="120" />
      <el-table-column prop="updatedAt" label="更新时间" width="140" />
    </el-table>
  </div>
</template>

<script setup lang="ts">
import { ref, watch, computed } from 'vue'
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
  padding: 24px;
  background: #ffffff;
  border-radius: 10px;
  border: 1px solid #ebeef5;
}
.search-toolbar {
  display: flex;
  gap: 12px;
  flex-wrap: wrap;
  margin-bottom: 16px;
}
.search-status {
  display: flex;
  gap: 18px;
  margin-bottom: 18px;
  color: #606266;
}
</style>
