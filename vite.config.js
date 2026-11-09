import {defineConfig} from 'vite';
import {fileURLToPath} from 'node:url';

const file = name => fileURLToPath(new URL(`./${name}`, import.meta.url));

export default defineConfig({
  base: './',
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: {
        index: file('index.html'),
        redirect: file('redirect.html'),
        relay: file('popup-relay.html'),
      },
    },
  },
});
