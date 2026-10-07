import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// Проверяем, передана ли специальная переменная для GitHub Pages
const isGitHubPages = process.env.GITHUB_PAGES === 'true';

export default defineConfig({
  // Если это GitHub Pages, используем папку репозитория. Иначе — корень (для Vercel)
  base: isGitHubPages ? '/-OutputForm-/' : '/',
  plugins: [react(), tailwindcss()],
  server: {
    host: "0.0.0.0",
    port: 3000,
    strictPort: true,
    hmr: {
      port: 3000,
    },
  },
})
