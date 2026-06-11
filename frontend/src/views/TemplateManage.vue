<template>
  <div class="template-manage">
    <!-- Header -->
    <div class="page-header">
      <div>
        <h1>模板管理</h1>
        <p class="sub">上传模板 Word，系统自动分析样式并保存。可随时替换或调整块配置。</p>
      </div>
      <el-button type="primary" size="large" @click="showUploadDialog = true">
        <el-icon><Upload /></el-icon> 上传模板
      </el-button>
    </div>

    <!-- Template List -->
    <div v-if="templates.length > 0" class="template-grid">
      <el-card
        v-for="tpl in templates"
        :key="tpl.template_id"
        class="template-card"
        shadow="hover"
      >
        <template #header>
          <div class="card-header">
            <span class="tpl-name">{{ tpl.template_name }}</span>
            <el-tag size="small" type="info">{{ tpl.template_id }}</el-tag>
          </div>
        </template>

        <div class="card-body">
          <div class="meta-item">
            <span class="label">创建时间</span>
            <span>{{ formatDate(tpl.created_at) }}</span>
          </div>
          <div class="meta-item">
            <span class="label">更新时间</span>
            <span>{{ formatDate(tpl.updated_at) }}</span>
          </div>
        </div>

        <div class="card-actions">
          <el-button size="small" @click="viewConfig(tpl.template_id)">
            <el-icon><View /></el-icon> 查看配置
          </el-button>
          <el-button size="small" type="warning" @click="replaceTemplate(tpl.template_id)">
            <el-icon><Refresh /></el-icon> 替换
          </el-button>
          <el-popconfirm
            title="确定删除此模板？"
            @confirm="deleteTpl(tpl.template_id)"
          >
            <template #reference>
              <el-button size="small" type="danger">
                <el-icon><Delete /></el-icon> 删除
              </el-button>
            </template>
          </el-popconfirm>
        </div>
      </el-card>
    </div>

    <!-- Empty State -->
    <el-empty
      v-else
      description="还没有模板，上传一个 Word 模板开始使用"
      :image-size="160"
    />

    <!-- Upload Dialog -->
    <el-dialog v-model="showUploadDialog" title="上传模板 Word" width="500px">
      <el-upload
        ref="uploadRef"
        drag
        :auto-upload="false"
        :limit="1"
        accept=".docx,.doc"
        :on-change="onFileChange"
        :on-remove="onFileRemove"
      >
        <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
        <div class="el-upload__text">
          将模板 Word 拖到此处或 <em>点击上传</em>
        </div>
        <template #tip>
          <div class="el-upload__tip">
            模板 Word 中的样式将被自动提取并保存。
            支持 .docx 格式（.doc 有限支持）。
          </div>
        </template>
      </el-upload>

      <template #footer>
        <el-button @click="showUploadDialog = false">取消</el-button>
        <el-button
          type="primary"
          :loading="uploading"
          :disabled="!pendingFile"
          @click="uploadTemplate"
        >
          {{ uploading ? '分析中...' : '上传并分析' }}
        </el-button>
      </template>
    </el-dialog>

    <!-- Config Dialog -->
    <el-dialog v-model="showConfigDialog" title="模板配置" width="700px">
      <div v-if="currentConfig" class="config-view">
        <el-descriptions :column="2" border size="small">
          <el-descriptions-item label="模板ID">{{ currentConfig.template_id }}</el-descriptions-item>
          <el-descriptions-item label="模板名称">{{ currentConfig.template_name }}</el-descriptions-item>
        </el-descriptions>

        <h3 style="margin: 16px 0 8px">页面设置</h3>
        <el-descriptions v-if="currentConfig.page_settings" :column="3" border size="small">
          <el-descriptions-item label="宽度 (EMU)">{{ currentConfig.page_settings.width_emu }}</el-descriptions-item>
          <el-descriptions-item label="高度 (EMU)">{{ currentConfig.page_settings.height_emu }}</el-descriptions-item>
          <el-descriptions-item label="上边距">{{ currentConfig.page_settings.top_margin_emu }}</el-descriptions-item>
        </el-descriptions>

        <h3 style="margin: 16px 0 8px">样式配置</h3>
        <div v-for="(style, key) in currentConfig.styles" :key="key" class="style-item">
          <div class="style-header">
            <el-tag :type="getTagType(String(key))" size="small">{{ styleLabel(String(key)) }}</el-tag>
            <span class="style-detail">
              {{ style.font_name || '-' }} ·
              {{ style.font_size_pt ? style.font_size_pt + 'pt' : '-' }}
              {{ style.font_bold ? '· 加粗' : '' }}
              {{ style.alignment ? '· ' + style.alignment : '' }}
            </span>
          </div>
        </div>

        <h3 style="margin: 16px 0 8px">块规则</h3>
        <div class="block-rules">
          <div v-for="(rule, key) in currentConfig.block_rules" :key="key" class="rule-item">
            <span class="rule-label">{{ styleLabel(String(key)) }}</span>
            <el-tag
              :type="rule === 'required' ? 'danger' : rule === 'optional' ? 'warning' : 'info'"
              size="small"
            >
              {{ rule === 'required' ? '必须' : rule === 'optional' ? '可选' : '跳过' }}
            </el-tag>
          </div>
        </div>
      </div>
    </el-dialog>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted } from 'vue'
import { ElMessage } from 'element-plus'
import { templateAPI } from '../api'

const templates = ref<any[]>([])
const showUploadDialog = ref(false)
const showConfigDialog = ref(false)
const uploading = ref(false)
const pendingFile = ref<File | null>(null)
const currentConfig = ref<any>(null)

onMounted(() => loadTemplates())

async function loadTemplates() {
  try {
    const { data } = await templateAPI.list()
    if (data.success) {
      templates.value = data.data || []
    }
  } catch (e: any) {
    ElMessage.error('加载模板列表失败: ' + (e.response?.data?.detail || e.message))
  }
}

function onFileChange(file: any) {
  pendingFile.value = file.raw
}

function onFileRemove() {
  pendingFile.value = null
}

async function uploadTemplate() {
  if (!pendingFile.value) return
  uploading.value = true
  try {
    const { data } = await templateAPI.upload(pendingFile.value)
    if (data.success) {
      ElMessage.success('模板上传并分析完成！')
      showUploadDialog.value = false
      pendingFile.value = null
      loadTemplates()
    }
  } catch (e: any) {
    ElMessage.error('模板上传失败: ' + (e.response?.data?.detail || e.message))
  } finally {
    uploading.value = false
  }
}

async function viewConfig(templateId: string) {
  try {
    const { data } = await templateAPI.get(templateId)
    if (data.success) {
      currentConfig.value = data.data
      showConfigDialog.value = true
    }
  } catch (e: any) {
    ElMessage.error('获取配置失败')
  }
}

async function replaceTemplate(templateId: string) {
  const input = document.createElement('input')
  input.type = 'file'
  input.accept = '.docx,.doc'
  input.onchange = async () => {
    const file = input.files?.[0]
    if (!file) return
    try {
      const { data } = await templateAPI.replace(templateId, file)
      if (data.success) {
        ElMessage.success('模板已替换')
        loadTemplates()
      }
    } catch (e: any) {
      ElMessage.error('替换失败: ' + (e.response?.data?.detail || e.message))
    }
  }
  input.click()
}

async function deleteTpl(templateId: string) {
  try {
    await templateAPI.delete(templateId)
    ElMessage.success('模板已删除')
    loadTemplates()
  } catch (e: any) {
    ElMessage.error('删除失败')
  }
}

function formatDate(iso: string) {
  if (!iso) return '-'
  return new Date(iso).toLocaleString('zh-CN')
}

const styleLabels: Record<string, string> = {
  main_title: '主标题',
  article_title: '文章标题',
  section_header: '章节标题',
  body_text: '正文',
  author_name: '作者',
  author_note: '作者说明',
  source: '来源',
  publish_date: '发布日期',
  edition: '版次',
  editor_note: '编者按',
  affiliation: '作者单位',
  tag_label: '标签',
  image: '图片',
  image_caption: '图片说明',
  footer: '页脚',
  other: '其他',
}

function styleLabel(key: string) {
  return styleLabels[key] || key
}

function getTagType(key: string) {
  if (['main_title', 'body_text'].includes(key)) return 'danger'
  if (['image', 'other', 'footer'].includes(key)) return 'info'
  return 'warning'
}
</script>

<style scoped>
.template-manage {
  max-width: 1100px;
  margin: 0 auto;
}

.page-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  margin-bottom: 24px;
}

.page-header h1 {
  font-size: 28px;
  margin-bottom: 6px;
}

.sub {
  color: var(--muted);
  font-size: 14px;
}

.template-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
  gap: 18px;
}

.template-card {
  border-radius: 16px;
  border: 1px solid var(--line);
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.tpl-name {
  font-weight: 700;
  font-size: 15px;
}

.card-body {
  padding: 0 0 12px;
}

.meta-item {
  display: flex;
  justify-content: space-between;
  font-size: 13px;
  color: var(--muted);
  margin-bottom: 4px;
}

.meta-item .label {
  font-weight: 500;
}

.card-actions {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
}

.config-view {
  max-height: 60vh;
  overflow-y: auto;
}

.style-item {
  padding: 6px 0;
  border-bottom: 1px solid #f0f0f0;
}

.style-header {
  display: flex;
  align-items: center;
  gap: 8px;
}

.style-detail {
  font-size: 13px;
  color: var(--muted);
}

.block-rules {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}

.rule-item {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 4px 12px;
  background: #f9f9f9;
  border-radius: 8px;
}

.rule-label {
  font-size: 13px;
  min-width: 60px;
}
</style>
