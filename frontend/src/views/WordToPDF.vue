<template>
  <div class="word-to-pdf">
    <div class="page-header">
      <div>
        <h1>Word 转 PDF</h1>
        <p class="sub">批量上传 Word 文件，转换为 PDF 并打包下载 ZIP。</p>
      </div>
    </div>

    <div class="form-area">
      <!-- Upload -->
      <div class="step-card">
        <div class="step-badge">1</div>
        <div class="step-content">
          <h3>上传 Word 文件（可多选）</h3>
          <el-upload
            ref="pdfUploadRef"
            drag
            multiple
            :auto-upload="false"
            accept=".doc,.docx,.DOC,.DOCX"
            :on-change="onFileChange"
          >
            <el-icon class="el-icon--upload"><UploadFilled /></el-icon>
            <div class="el-upload__text">
              将 Word 文件拖到此处或 <em>点击批量上传</em>
            </div>
            <template #tip>
              <div class="el-upload__tip">
                已选择 {{ pdfFiles.length }} 个文件
              </div>
            </template>
          </el-upload>

          <div v-if="pdfFiles.length > 0" class="file-list">
            <el-tag
              v-for="(f, i) in pdfFiles"
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

      <!-- Target Names -->
      <div class="step-card">
        <div class="step-badge">2</div>
        <div class="step-content">
          <h3>目标名字列表（可选）</h3>
          <el-input
            v-model="targetNames"
            type="textarea"
            :rows="6"
            placeholder="h1wh202605001&#10;h1wh202605002&#10;h1wh202605003"
          />
          <div class="tip">
            一行一个名字，数量需与上传文件一致。留空则保持原文件名。
          </div>
        </div>
      </div>

      <!-- Convert -->
      <div class="actions">
        <el-button
          type="primary"
          size="large"
          :disabled="pdfFiles.length === 0"
          :loading="converting"
          @click="doConvert"
        >
          <el-icon><Switch /></el-icon>
          {{ converting ? '转换中...' : `开始转换 (${pdfFiles.length} 个文件)` }}
        </el-button>
      </div>

      <!-- Engine Info -->
      <el-alert
        title="转换引擎"
        :description="engineInfo"
        type="info"
        show-icon
        :closable="false"
        style="margin-top: 18px"
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue'
import { ElMessage } from 'element-plus'
import { convertAPI } from '../api'

const pdfFiles = ref<File[]>([])
const targetNames = ref('')
const converting = ref(false)
const engineInfo = ref('需要安装 LibreOffice。可通过 WORD2PDF_SOFFICE 环境变量指定路径。')

function onFileChange(file: any, fileList: any) {
  pdfFiles.value = fileList.map((f: any) => f.raw).filter(Boolean)
}

function removeFile(index: number) {
  pdfFiles.value.splice(index, 1)
}

async function doConvert() {
  if (pdfFiles.value.length === 0) return

  // Validate target names count
  const names = targetNames.value.split('\n').filter(n => n.trim())
  if (names.length > 0 && names.length !== pdfFiles.value.length) {
    ElMessage.error(`目标名字数量 (${names.length}) 与文件数量 (${pdfFiles.value.length}) 不一致`)
    return
  }

  converting.value = true
  try {
    const response = await convertAPI.wordToPdf(pdfFiles.value, targetNames.value)
    const blob = response.data
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'word2pdf_result.zip'
    a.click()
    URL.revokeObjectURL(url)
    ElMessage.success('转换完成，ZIP 已下载！')
  } catch (e: any) {
    const detail = e.response?.data?.detail
    if (detail instanceof Blob) {
      // Try to read blob error
      const text = await detail.text()
      ElMessage.error('转换失败: ' + text)
    } else {
      ElMessage.error('转换失败: ' + (detail || e.message))
    }
  } finally {
    converting.value = false
  }
}
</script>

<style scoped>
.word-to-pdf { max-width: 900px; margin: 0 auto; }

.page-header h1 { font-size: 28px; margin-bottom: 6px; }
.sub { color: var(--muted); font-size: 14px; }

.form-area { margin-top: 24px; }

.step-card {
  display: flex; gap: 16px;
  margin-bottom: 20px;
  background: rgba(255, 250, 243, 0.7);
  border: 1px solid var(--line);
  border-radius: 16px;
  padding: 20px;
}

.step-badge {
  width: 36px; height: 36px;
  border-radius: 50%;
  background: var(--accent2);
  color: #fff;
  display: flex; align-items: center; justify-content: center;
  font-weight: 700; flex-shrink: 0;
}

.step-content { flex: 1; }
.step-content h3 { font-size: 16px; margin-bottom: 10px; }

.tip { font-size: 13px; color: var(--muted); margin-top: 8px; }

.file-list { margin-top: 12px; }

.actions { display: flex; gap: 12px; justify-content: center; margin-top: 24px; }
</style>
