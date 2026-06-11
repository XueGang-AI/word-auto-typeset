<template>
  <div class="single-typeset">
    <div class="page-header">
      <div>
        <h1>单文件套版</h1>
        <p class="sub">选择一个模板和一个内容文档，自动识别结构并套用模板样式。</p>
      </div>
    </div>

    <div class="form-area">
      <!-- Step 1: Select Template -->
      <div class="step-card">
        <div class="step-badge">1</div>
        <div class="step-content">
          <h3>选择模板</h3>
          <el-select
            v-model="selectedTemplateId"
            placeholder="请先选择一个模板"
            style="width: 100%"
            @change="onTemplateChange"
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

      <!-- Step 2: Upload Content -->
      <div class="step-card">
        <div class="step-badge">2</div>
        <div class="step-content">
          <h3>上传内容文档</h3>
          <el-upload
            ref="contentUploadRef"
            drag
            :auto-upload="false"
            :limit="1"
            accept=".docx,.doc"
            :on-change="onContentChange"
            :on-remove="onContentRemove"
          >
            <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
            <div class="el-upload__text">
              将内容 Word 拖到此处或 <em>点击上传</em>
            </div>
          </el-upload>
        </div>
      </div>

      <!-- Step 3: Preview Structure -->
      <div v-if="analysisResult" class="step-card">
        <div class="step-badge">3</div>
        <div class="step-content">
          <h3>结构识别结果</h3>
          <div v-if="analysisResult.warnings?.length" class="warnings">
            <el-alert
              v-for="(w, i) in analysisResult.warnings"
              :key="i"
              :title="w"
              type="warning"
              show-icon
              :closable="false"
              style="margin-bottom: 6px"
            />
          </div>
          <div class="structure-preview">
            <div
              v-for="article in analysisResult.articles"
              :key="article.article_index"
              class="article-block"
            >
              <div class="article-divider" v-if="analysisResult.articles.length > 1">
                第 {{ article.article_index + 1 }} 篇文章
              </div>
              <div
                v-for="para in article.paragraphs.slice(0, 20)"
                :key="para.index"
                class="para-row"
              >
                <el-tag :type="getParaTagType(para.para_type)" size="small" effect="plain">
                  {{ paraTypeLabel(para.para_type) }}
                </el-tag>
                <span class="para-confidence">
                  <el-tag
                    :type="para.confidence === 'high' ? 'success' : para.confidence === 'medium' ? 'warning' : 'danger'"
                    size="small"
                    effect="light"
                  >
                    {{ para.confidence }}
                  </el-tag>
                </span>
                <span class="para-text">{{ para.text?.substring(0, 80) }}{{ para.text?.length > 80 ? '...' : '' }}</span>
              </div>
              <div v-if="article.paragraphs.length > 20" class="more-hint">
                ... 还有 {{ article.paragraphs.length - 20 }} 个段落
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- Actions -->
      <div class="actions" v-if="pendingContentFile && selectedTemplateId">
        <el-button type="primary" size="large" @click="previewStructure" :loading="analyzing">
          <el-icon><Search /></el-icon>
          {{ analyzing ? '分析中...' : '预览结构' }}
        </el-button>
        <el-button
          type="success"
          size="large"
          @click="doTypeset"
          :loading="typesetting"
          :disabled="!analysisResult"
        >
          <el-icon><Finished /></el-icon>
          {{ typesetting ? '排版中...' : '开始排版并下载' }}
        </el-button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted } from 'vue'
import { ElMessage } from 'element-plus'
import { templateAPI, typesetAPI } from '../api'

const templates = ref<any[]>([])
const selectedTemplateId = ref('')
const pendingContentFile = ref<File | null>(null)
const analysisResult = ref<any>(null)
const analyzing = ref(false)
const typesetting = ref(false)

onMounted(() => loadTemplates())

async function loadTemplates() {
  try {
    const { data } = await templateAPI.list()
    if (data.success) templates.value = data.data || []
  } catch (e: any) {
    ElMessage.error('加载模板失败')
  }
}

function onTemplateChange() {
  analysisResult.value = null
}

function onContentChange(file: any) {
  pendingContentFile.value = file.raw
  analysisResult.value = null
}

function onContentRemove() {
  pendingContentFile.value = null
  analysisResult.value = null
}

async function previewStructure() {
  if (!pendingContentFile.value) return
  analyzing.value = true
  try {
    const { data } = await typesetAPI.analyze(pendingContentFile.value)
    if (data.success) {
      analysisResult.value = data.data
      ElMessage.success('结构分析完成')
    }
  } catch (e: any) {
    ElMessage.error('分析失败: ' + (e.response?.data?.detail || e.message))
  } finally {
    analyzing.value = false
  }
}

async function doTypeset() {
  if (!pendingContentFile.value || !selectedTemplateId.value) return
  typesetting.value = true
  try {
    const response = await typesetAPI.single(selectedTemplateId.value, pendingContentFile.value)
    const blob = response.data
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    const originalName = pendingContentFile.value.name.replace(/\.(docx|doc)$/i, '')
    a.download = `${originalName}_排版后.docx`
    a.click()
    URL.revokeObjectURL(url)
    ElMessage.success('排版完成，文件已下载！')
  } catch (e: any) {
    ElMessage.error('排版失败: ' + (e.response?.data?.detail || e.message))
  } finally {
    typesetting.value = false
  }
}

const paraTypeLabels: Record<string, string> = {
  main_title: '主标题', article_title: '文章标题', section_header: '章节标题',
  body_text: '正文', author_name: '作者', author_note: '作者说明',
  source: '来源', publish_date: '日期', edition: '版次',
  editor_note: '编者按', affiliation: '作者单位', tag_label: '标签',
  image: '图片', image_caption: '图片说明', footer: '页脚', other: '其他',
}

function paraTypeLabel(type: string) {
  return paraTypeLabels[type] || type
}

function getParaTagType(type: string) {
  if (['main_title'].includes(type)) return 'danger'
  if (['body_text'].includes(type)) return ''
  if (['image', 'other', 'footer'].includes(type)) return 'info'
  return 'warning'
}
</script>

<style scoped>
.single-typeset {
  max-width: 900px;
  margin: 0 auto;
}

.page-header h1 { font-size: 28px; margin-bottom: 6px; }
.sub { color: var(--muted); font-size: 14px; }

.form-area { margin-top: 24px; }

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
  width: 36px;
  height: 36px;
  border-radius: 50%;
  background: var(--accent);
  color: #fff;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: 700;
  flex-shrink: 0;
}

.step-content { flex: 1; }
.step-content h3 { font-size: 16px; margin-bottom: 10px; }

.tip { font-size: 13px; color: var(--muted); margin-top: 8px; }
.tip a { color: var(--accent); }

.warnings { margin-bottom: 12px; }

.structure-preview {
  max-height: 400px;
  overflow-y: auto;
  border: 1px solid #eee;
  border-radius: 10px;
  padding: 12px;
}

.article-block { margin-bottom: 8px; }
.article-divider {
  background: rgba(139, 94, 52, 0.08);
  padding: 6px 12px;
  border-radius: 6px;
  font-weight: 600;
  font-size: 14px;
  margin-bottom: 6px;
  color: var(--accent);
}

.para-row {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 4px 0;
  font-size: 13px;
}

.para-text {
  flex: 1;
  color: #555;
}

.more-hint {
  font-size: 12px;
  color: var(--muted);
  text-align: center;
  padding: 8px;
}

.actions {
  display: flex;
  gap: 12px;
  justify-content: center;
  margin-top: 24px;
}
</style>
