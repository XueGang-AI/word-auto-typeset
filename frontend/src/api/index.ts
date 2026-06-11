import axios from 'axios'

const api = axios.create({
  baseURL: '/api',
  timeout: 120000,
})

// ── Template APIs ────────────────────────────────────────

export const templateAPI = {
  list: () => api.get('/templates/'),
  get: (id: string) => api.get(`/templates/${id}`),
  upload: (file: File) => {
    const fd = new FormData()
    fd.append('file', file)
    return api.post('/templates/upload', fd)
  },
  replace: (id: string, file: File) => {
    const fd = new FormData()
    fd.append('file', file)
    return api.put(`/templates/${id}/replace`, fd)
  },
  delete: (id: string) => api.delete(`/templates/${id}`),
  getConfig: (id: string) => api.get(`/templates/${id}/config`),
  updateConfig: (id: string, config: any) => api.put(`/templates/${id}/config`, config),
}

// ── Typesetting APIs ─────────────────────────────────────

export const typesetAPI = {
  // Single file typesetting — returns the formatted .docx as a blob
  single: (templateId: string, file: File) => {
    const fd = new FormData()
    fd.append('template_id', templateId)
    fd.append('file', file)
    return api.post('/typeset/single/download', fd, {
      responseType: 'blob',
    })
  },

  // Analyze content only (preview)
  analyze: (file: File) => {
    const fd = new FormData()
    fd.append('file', file)
    return api.post('/typeset/analyze', fd)
  },

  // Batch typesetting
  batch: (templateId: string, files: File[]) => {
    const fd = new FormData()
    fd.append('template_id', templateId)
    files.forEach((f) => fd.append('files', f))
    return api.post('/typeset/batch', fd, { timeout: 600000 })
  },

  // Batch progress
  batchProgress: (batchId: string) => api.get(`/typeset/batch/${batchId}/progress`),
  batchDetail: (batchId: string) => api.get(`/typeset/batch/${batchId}`),

  // Download
  downloadZipUrl: (batchId: string) => `/api/typeset/batch/${batchId}/download/zip`,
  downloadReportUrl: (batchId: string) => `/api/typeset/batch/${batchId}/download/report`,
}

// ── Word → PDF APIs ──────────────────────────────────────

export const convertAPI = {
  wordToPdf: (files: File[], targetNames: string, overwrite: boolean = true) => {
    const fd = new FormData()
    files.forEach((f) => fd.append('files', f))
    fd.append('target_names', targetNames)
    fd.append('overwrite', String(overwrite))
    return api.post('/convert/word-to-pdf', fd, {
      responseType: 'blob',
      timeout: 600000,
    })
  },
}

export default api
