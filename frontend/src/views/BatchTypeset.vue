<template>
  <div class="batch-typeset">
    <div class="page-header">
      <div>
        <h1>批量套版</h1>
        <p class="sub">批量上传 Word 文件，自动排版并打包下载 ZIP + 处理报告。</p>
      </div>
    </div>

    <!-- Setup Form -->
    <div v-if="!currentBatch" class="setup-area">
      <!-- Template Selection -->
      <div class="step-card">
        <div class="step-badge">1</div>
        <div class="step-content">
          <h3>选择模板</h3>
          <el-select
            v-model="selectedTemplateId"
            placeholder="请先选择一个模板"
            style="width: 100%"
          >
            <el-option
              v-for="tpl in templates"
              :key="tpl.template_id"
              :label="`${tpl.template_name} (${tpl.template_id})`"
              :value="tpl.template_id"
            />
          </el-select>
          <div v-if="!templates.length" class="tip">
            还没有模板？请先在
            <router-link to="/templates">模板管理</router-link>
            中上传。
          </div>
        </div>
      </div>

      <!-- File Upload -->
      <div class="step-card">
        <div class="step-badge">2</div>
        <div class="step-content">
          <h3>上传待处理文件</h3>
          <el-upload
            ref="batchUploadRef"
            drag
            multiple
            :auto-upload="false"
            accept=".docx,.doc"
            :on-change="onFileListChange"
          >
            <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
            <div class="el-upload__text">
              将 Word 文件拖到此处或 <em>点击批量上传</em>
            </div>
            <template #tip>
              <div class="el-upload__tip">
                支持同时选择多个 .docx/.doc 文件。
                已选择 {{ pendingFiles.length }} 个文件。
              </div>
            </template>
          </el-upload>

          <!-- File List -->
          <div v-if="pendingFiles.length > 0" class="file-list">
            <el-tag
              v-for="(f, i) in pendingFiles"
              :key="i"
              closable
              @close="removeFile(i)"
              style="margin: 4px"
            >
              {{ f.name }}
            </el-tag>
          </div>
        </div>
      </div>

      <!-- Start Button -->
      <div class="actions">
        <el-button
          type="primary"
          size="large"
          :disabled="!selectedTemplateId || pendingFiles.length === 0"
          :loading="submitting"
          @click="startBatch"
        >
          <el-icon><VideoPlay /></el-icon>
          开始批量处理 ({{ pendingFiles.length }} 个文件)
        </el-button>
      </div>
    </div>

    <!-- Progress View -->
    <div v-else class="progress-area">
      <el-card>
        <template #header>
          <div class="progress-header">
            <span>批量任务: {{ currentBatch.batch_id }}</span>
            <el-tag :type="statusTagType">{{ statusLabel }}</el-tag>
          </div>
        </template>

        <!-- Stats -->
        <div class="stats-row">
          <el-statistic title="总文件数" :value="progress.total" />
          <el-statistic title="已完成" :value="progress.completed" />
          <el-statistic title="成功" :value="progress.success">
            <template #suffix>
              <span style="font-size:14px;color:#67c23a">
                {{ progress.total ? Math.round(progress.success / progress.total * 100) : 0 }}%
              </span>
            </template>
          </el-statistic>
          <el-statistic title="失败" :value="progress.failed">
            <template #suffix>
              <span style="font-size:14px;color:#f56c6c">
                {{ progress.failed > 0 ? '⚠' : '✓' }}
              </span>
            </template>
          </el-statistic>
          <el-statistic title="低置信度" :value="progress.low_confidence" />
        </div>

        <!-- Progress Bar -->
        <el-progress
          :percentage="progress.total ? Math.round(progress.completed / progress.total * 100) : 0"
          :status="progress.status === 'completed' ? 'success' : progress.status === 'failed' ? 'exception' : ''"
          :stroke-width="20"
          style="margin: 20px 0"
        />

        <!-- Download Buttons -->
        <div v-if="progress.status === 'completed'" class="download-actions">
          <el-button type="primary" size="large" @click="downloadZip">
            <el-icon><Download /></el-icon> 下载结果 ZIP
          </el-button>
          <el-button type="success" size="large" @click="downloadReport">
            <el-icon><Document /></el-icon> 下载处理报告 (report.xlsx)
          </el-button>
        </div>

        <!-- Auto-refresh -->
        <div v-if="progress.status === 'processing'" class="auto-refresh">
          <el-icon class="is-loading"><Loading /></el-icon>
          正在自动刷新进度...
        </div>
      </el-card>

      <!-- Task Details -->
      <el-card v-if="batchDetail" style="margin-top: 18px">
        <template #header>文件处理详情</template>
        <el-table :data="batchDetail.tasks" size="small" max-height="400">
          <el-table-column prop="original_filename" label="文件名" min-width="180" />
          <el-table-column label="状态" width="100">
            <template #default="{ row }">
              <el-tag
                :type="row.status === 'completed' ? 'success' : row.status === 'failed' ? 'danger' : row.status === 'processing' ? 'warning' : 'info'"
                size="small"
              >
                {{ statusText(row.status) }}
              </el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="output_filename" label="输出文件" min-width="160" />
          <el-table-column label="低置信度" width="90">
            <template #default="{ row }">
              <el-tag v-if="row.has_low_confidence" type="warning" size="small">是</el-tag>
              <span v-else style="color:#999">否</span>
            </template>
          </el-table-column>
          <el-table-column prop="error_message" label="错误信息" min-width="180">
            <template #default="{ row }">
              <span style="color:#f56c6c;font-size:12px">{{ row.error_message || '-' }}</span>
            </template>
          </el-table-column>
        </el-table>
      </el-card>

      <div class="actions" style="margin-top: 18px">
        <el-button @click="resetBatch">开始新的批量任务</el-button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted, watch } from 'vue'
import { ElMessage } from 'element-plus'
import { templateAPI, typesetAPI } from '../api'

const templates = ref<any[]>([])
const selectedTemplateId = ref('')
const pendingFiles = ref<File[]>([])
const submitting = ref(false)
const currentBatch = ref<any>(null)
const batchDetail = ref<any>(null)
const progress = ref({ total: 0, completed: 0, success: 0, failed: 0, low_confidence: 0, status: 'pending' })

let pollTimer: any = null

onMounted(() => loadTemplates())

onUnmounted(() => {
  if (pollTimer) clearInterval(pollTimer)
})

async function loadTemplates() {
  try {
    const { data } = await templateAPI.list()
    if (data.success) templates.value = data.data || []
  } catch (e: any) {
    ElMessage.error('加载模板失败')
  }
}

function onFileListChange(file: any, fileList: any) {
  pendingFiles.value = fileList.map((f: any) => f.raw).filter(Boolean)
}

function removeFile(index: number) {
  pendingFiles.value.splice(index, 1)
}

async function startBatch() {
  if (!selectedTemplateId.value || pendingFiles.value.length === 0) return
  submitting.value = true
  try {
    const { data } = await typesetAPI.batch(selectedTemplateId.value, pendingFiles.value)
    if (data.success) {
      currentBatch.value = data.data
      progress.value.status = 'processing'
      ElMessage.success(`批量任务已启动: ${data.data.batch_id}`)
      startPolling()
    }
  } catch (e: any) {
    ElMessage.error('启动批量任务失败: ' + (e.response?.data?.detail || e.message))
  } finally {
    submitting.value = false
  }
}

function startPolling() {
  if (pollTimer) clearInterval(pollTimer)
  pollTimer = setInterval(fetchProgress, 2000)
}

async function fetchProgress() {
  if (!currentBatch.value) return
  try {
    const [progressRes, detailRes] = await Promise.all([
      typesetAPI.batchProgress(currentBatch.value.batch_id),
      typesetAPI.batchDetail(currentBatch.value.batch_id),
    ])
    if (progressRes.data.success) {
      Object.assign(progress.value, progressRes.data.data)
    }
    if (detailRes.data.success) {
      batchDetail.value = detailRes.data.data
    }
    if (progress.value.status === 'completed' || progress.value.status === 'failed') {
      if (pollTimer) clearInterval(pollTimer)
      if (progress.value.status === 'completed') {
        ElMessage.success(`批量处理完成！成功: ${progress.value.success}, 失败: ${progress.value.failed}`)
      }
    }
  } catch (e) {
    // Silently retry
  }
}

function downloadZip() {
  if (!currentBatch.value) return
  window.open(typesetAPI.downloadZipUrl(currentBatch.value.batch_id), '_blank')
}

function downloadReport() {
  if (!currentBatch.value) return
  window.open(typesetAPI.downloadReportUrl(currentBatch.value.batch_id), '_blank')
}

function resetBatch() {
  currentBatch.value = null
  batchDetail.value = null
  progress.value = { total: 0, completed: 0, success: 0, failed: 0, low_confidence: 0, status: 'pending' }
  if (pollTimer) clearInterval(pollTimer)
}

const statusLabel = computed(() => {
  const map: Record<string, string> = {
    pending: '等待中', processing: '处理中', completed: '已完成', failed: '失败',
  }
  return map[progress.value.status] || progress.value.status
})

const statusTagType = computed(() => {
  const map: Record<string, string> = {
    pending: 'info', processing: 'warning', completed: 'success', failed: 'danger',
  }
  return map[progress.value.status] || 'info'
})

function statusText(s: string) {
  const map: Record<string, string> = {
    pending: '等待', processing: '处理中', completed: '完成', failed: '失败',
  }
  return map[s] || s
}
</script>

<style scoped>
.batch-typeset { max-width: 900px; margin: 0 auto; }

.page-header h1 { font-size: 28px; margin-bottom: 6px; }
.sub { color: var(--muted); font-size: 14px; }

.setup-area { margin-top: 24px; }

.step-card {
  display: flex;
  gap: 16px;
  margin-bottom: 20px;
  background: rgba(255, 250, 243, 0.7);
  border: 1px solid var(--line);
  border-radius: 16px;
  padding: 20px;
}

.step-badge {
  width: 36px; height: 36px;
  border-radius: 50%;
  background: var(--accent);
  color: #fff;
  display: flex; align-items: center; justify-content: center;
  font-weight: 700; flex-shrink: 0;
}

.step-content { flex: 1; }
.step-content h3 { font-size: 16px; margin-bottom: 10px; }

.tip { font-size: 13px; color: var(--muted); margin-top: 8px; }
.tip a { color: var(--accent); }

.file-list { margin-top: 12px; }

.actions { display: flex; gap: 12px; justify-content: center; margin-top: 24px; }

.progress-header {
  display: flex; justify-content: space-between; align-items: center;
}

.stats-row {
  display: flex; gap: 24px; flex-wrap: wrap;
  justify-content: center;
  padding: 16px 0;
}

.download-actions {
  display: flex; gap: 12px; justify-content: center;
  padding: 16px 0;
}

.auto-refresh {
  text-align: center; color: var(--accent2); padding: 12px;
  display: flex; align-items: center; justify-content: center; gap: 8px;
}
</style>
