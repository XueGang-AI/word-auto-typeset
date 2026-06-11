<template>
  <header class="nav-header">
    <div class="nav-inner">
      <div class="brand" @click="$router.push('/')">
        <span class="logo">📄</span>
        <span class="title">Word 自动套版系统</span>
      </div>
      <nav class="nav-links">
        <router-link to="/templates" class="nav-link" active-class="active">
          <el-icon><Setting /></el-icon>
          模板管理
        </router-link>
        <router-link to="/typeset" class="nav-link" active-class="active">
          <el-icon><Document /></el-icon>
          单文件套版
        </router-link>
        <router-link to="/batch" class="nav-link" active-class="active">
          <el-icon><FolderOpened /></el-icon>
          批量套版
        </router-link>
        <router-link to="/convert" class="nav-link" active-class="active">
          <el-icon><Switch /></el-icon>
          Word 转 PDF
        </router-link>
      </nav>
    </div>
  </header>
</template>

<script setup lang="ts">
</script>

<style scoped>
.nav-header {
  background: rgba(255, 250, 243, 0.92);
  border-bottom: 1px solid rgba(231, 221, 208, 0.85);
  backdrop-filter: blur(12px);
  position: sticky;
  top: 0;
  z-index: 100;
  box-shadow: 0 2px 20px rgba(31, 41, 55, 0.06);
}

.nav-inner {
  max-width: 1200px;
  margin: 0 auto;
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 20px;
  height: 60px;
}

.brand {
  display: flex;
  align-items: center;
  gap: 10px;
  cursor: pointer;
  user-select: none;
}

.logo {
  font-size: 28px;
}

.title {
  font-size: 18px;
  font-weight: 700;
  color: var(--accent);
  letter-spacing: 0.02em;
}

.nav-links {
  display: flex;
  gap: 4px;
}

.nav-link {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 8px 16px;
  border-radius: 10px;
  text-decoration: none;
  color: var(--ink);
  font-size: 14px;
  font-weight: 500;
  transition: all 0.2s;
}

.nav-link:hover {
  background: rgba(139, 94, 52, 0.08);
  color: var(--accent);
}

.nav-link.active {
  background: rgba(139, 94, 52, 0.12);
  color: var(--accent);
  font-weight: 700;
}
</style>
