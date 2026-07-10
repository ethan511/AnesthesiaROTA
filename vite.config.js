import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    react({
      jsxRuntime: 'automatic', // This explicitly forces Vite to auto-inject React into your JSX files
    }),
  ],
  build: {
    chunkSizeWarningLimit: 1600,
  },
})
