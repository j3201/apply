<template>
  <div id="app">
    <el-tabs v-model="activeTab" type="border-card" class="main-tabs">
      <el-tab-pane label="报销与津贴管理" name="allowance">
        <AllowanceTable />
      </el-tab-pane>
      <el-tab-pane label="港口交通搜索" name="port-traffic">
        <PortTrafficSearch />
      </el-tab-pane>
    </el-tabs>

    <el-dialog
      v-model="showUploadPrompt"
      title="请上传港口交通清单"
      width="520px"
      :close-on-click-modal="false"
      draggable
    >
      <el-empty v-if="!store.loaded && !store.items.length" description="系统尚未检测到本地缓存的港口交通清单">
        <el-text type="info" size="small">
          建议首次打开时上传一次，以便后续自动匹配 A 表中的地点信息。
        </el-text>
      </el-empty>
      <div v-else class="upload-prompt-content">
        <el-icon class="upload-icon" :size="48"><UploadFilled /></el-icon>
        <p>系统已加载 {{ store.items.length }} 条港口交通数据</p>
        <el-text type="info" size="small">
          您可以继续使用，或上传新文件更新数据
        </el-text>
      </div>
      <template #footer>
        <input ref="initialPortInput" type="file" accept=".xlsx,.csv" style="display:none" @change="handleInitialFileChange" />
        <el-button @click="skipInitial">暂不上传</el-button>
        <el-button type="primary" @click="triggerInitialUpload">
          <el-icon><Upload /></el-icon>
          {{ store.items.length ? '更新数据' : '立即上传' }}
        </el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted } from 'vue'
import { ElMessage } from 'element-plus'
import { UploadFilled, Upload } from '@element-plus/icons-vue'
import AllowanceTable from './components/AllowanceTable.vue'
import PortTrafficSearch from './components/PortTrafficSearch.vue'
import { usePortTrafficStore } from '@/stores/portTraffic'

const activeTab = ref('allowance')
const showUploadPrompt = ref(false)
const initialPortInput = ref<HTMLInputElement | null>(null)

const store = usePortTrafficStore()

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
    const rows = await store.loadFromFile(file)
    showUploadPrompt.value = false
    ElMessage.success({
      message: `成功加载 ${rows.length} 条港口交通数据`,
      duration: 3000
    })
  } catch (err) {
    ElMessage.error({
      message: err instanceof Error ? err.message : '加载港口交通清单失败',
      duration: 5000
    })
  } finally {
    if (initialPortInput.value) initialPortInput.value.value = ''
  }
}

onMounted(() => {
  store.initialize()
  // 延迟显示提示对话框，确保数据加载完成
  setTimeout(() => {
    if (!store.loaded) {
      showUploadPrompt.value = true
    }
  }, 500)
})
</script>

<style scoped>
#app {
  min-height: 100vh;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  padding: 20px;
}

.main-tabs {
  max-width: 1400px;
  margin: 0 auto;
  border-radius: 12px;
  overflow: hidden;
  box-shadow: 0 10px 40px rgba(0, 0, 0, 0.15);
}

.upload-prompt-content {
  text-align: center;
  padding: 20px 0;
}

.upload-icon {
  color: var(--el-color-primary);
  margin-bottom: 16px;
}

.upload-prompt-content p {
  margin: 0 0 8px;
  font-size: 16px;
  color: var(--el-text-color-primary);
}
</style>
