import { createRouter, createWebHistory } from 'vue-router'

const router = createRouter({
  history: createWebHistory(),
  routes: [
    {
      path: '/',
      redirect: '/templates',
    },
    {
      path: '/templates',
      name: 'templates',
      component: () => import('../views/TemplateManage.vue'),
      meta: { title: '模板管理' },
    },
    {
      path: '/typeset',
      name: 'typeset',
      component: () => import('../views/SingleTypeset.vue'),
      meta: { title: '单文件套版' },
    },
    {
      path: '/batch',
      name: 'batch',
      component: () => import('../views/BatchTypeset.vue'),
      meta: { title: '批量套版' },
    },
    {
      path: '/convert',
      name: 'convert',
      component: () => import('../views/WordToPDF.vue'),
      meta: { title: 'Word 转 PDF' },
    },
  ],
})

router.beforeEach((to) => {
  document.title = `${to.meta.title || 'Word 自动套版'} - Word 自动套版系统`
})

export default router
