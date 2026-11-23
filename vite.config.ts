import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  base: '/data-health-checker/v2-preview/',
  plugins: [react()],
})
