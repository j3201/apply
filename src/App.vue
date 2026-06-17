<template>
  <div id="app">
    <el-tabs v-model="activeTab" type="border-card">
      <el-tab-pane label="报销与津贴管理" name="allowance">
        <AllowanceTable />
      </el-tab-pane>
      <el-tab-pane label="港口交通搜索" name="port-traffic">
        <PortTrafficSearch />
      </el-tab-pane>
    </el-tabs>

    <el-dialog v-model="showUploadPrompt" title="请上传港口交通清单" width="520px">
      <div>
        系统尚未检测到本地缓存的港口交通清单。建议首次打开时上传一次，以便后续自动匹配 A 表中的地点信息。
      </div>
      <template #footer>
        <input ref="initialPortInput" type="file" accept=".xlsx,.csv" style="display:none" @change="handleInitialFileChange" />
        <el-button @click="skipInitial">暂不上传</el-button>
        <el-button type="primary" @click="triggerInitialUpload">立即上传</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted } from 'vue'
import AllowanceTable from './components/AllowanceTable.vue'
import PortTrafficSearch from './components/PortTrafficSearch.vue'
import { usePortTrafficStore } from '@/stores/portTraffic'

const activeTab = ref('allowance')
const showUploadPrompt = ref(false)
const initialPortInput = ref<HTMLInputElement | null>(null)

const portTrafficStore = usePortTrafficStore()

const triggerInitialUpload = () => {
  initialPortInput.value?.click()
}

const skipInitial = () => {
  showUploadPrompt.value = false
}

const handleInitialFileChange = async (e: Event) => {
  const files = (e.target as HTMLInputElement).files
  if (!files?.length) return
  const file = files[0]
  if (!file) return
  try {
    await portTrafficStore.loadFromFile(file)
    showUploadPrompt.value = false
  } catch (err) {
    // store will set status; keep dialog open for retry
    console.error(err)
  } finally {
    if (initialPortInput.value) initialPortInput.value.value = ''
  }
}

onMounted(() => {
  portTrafficStore.initialize()
  if (!portTrafficStore.loaded) {
    showUploadPrompt.value = true
  }
})
</script>
