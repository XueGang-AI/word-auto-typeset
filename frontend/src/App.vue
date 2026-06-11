<template>
  <div id="app-container">
    <NavHeader />
    <main class="main-content">
      <router-view />
    </main>
  </div>
</template>

<script setup lang="ts">
import NavHeader from './components/NavHeader.vue'
</script>

<style>
:root {
  --bg: #f5f1ea;
  --card: #fffaf3;
  --ink: #1f2937;
  --muted: #6b7280;
  --line: #e7ddd0;
  --accent: #8b5e34;
  --accent2: #2f6f73;
  --shadow: 0 18px 50px rgba(31, 41, 55, 0.12);
}

* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

body {
  font-family: 'Aptos', 'Noto Sans SC', 'Microsoft YaHei', sans-serif;
  color: var(--ink);
  background:
    radial-gradient(circle at top left, rgba(139, 94, 52, 0.12), transparent 28%),
    radial-gradient(circle at right top, rgba(47, 111, 115, 0.12), transparent 30%),
    linear-gradient(180deg, #fff, var(--bg));
  min-height: 100vh;
}

#app-container {
  min-height: 100vh;
}

.main-content {
  max-width: 1200px;
  margin: 0 auto;
  padding: 24px 20px 56px;
}
</style>
