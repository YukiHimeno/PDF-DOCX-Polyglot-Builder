import { defineConfig } from 'vite';
import { createReadStream, existsSync, mkdirSync, copyFileSync } from 'node:fs';
import { resolve, basename } from 'node:path';

const root = resolve('node_modules/@matbee/libreoffice-converter');
const files = ['soffice.js','soffice.wasm','soffice.data','soffice.worker.js'];
function officeAssets() {
  return {
    name: 'office-wasm-assets',
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const name = decodeURIComponent((req.url || '').split('?')[0]).split('/').pop();
        if (!req.url?.includes('/office-wasm/') || ![...files, 'browser.worker.global.js'].includes(name)) return next();
        const path = resolve(root, name === 'browser.worker.global.js' ? 'dist' : 'wasm', name);
        if (!existsSync(path)) return next();
        res.setHeader('Cross-Origin-Opener-Policy', 'same-origin');
        res.setHeader('Cross-Origin-Embedder-Policy', 'require-corp');
        res.setHeader('Cross-Origin-Resource-Policy', 'same-origin');
        res.setHeader('Content-Type', name.endsWith('.wasm') ? 'application/wasm' : name.endsWith('.js') ? 'application/javascript' : 'application/octet-stream');
        createReadStream(path).pipe(res);
      });
    },
    writeBundle(options) {
      const dir = resolve(options.dir || 'dist', 'office-wasm'); mkdirSync(dir, { recursive: true });
      for (const name of files.filter(name => name.endsWith('.js'))) copyFileSync(resolve(root, 'wasm', name), resolve(dir, basename(name)));
      copyFileSync(resolve(root, 'dist/browser.worker.global.js'), resolve(dir, 'browser.worker.global.js'));
    },
  };
}
export default defineConfig({ plugins: [officeAssets()], optimizeDeps: { include: ['@matbee/libreoffice-converter/browser'] }, server: { headers: { 'Cross-Origin-Opener-Policy': 'same-origin', 'Cross-Origin-Embedder-Policy': 'require-corp' } }, preview: { headers: { 'Cross-Origin-Opener-Policy': 'same-origin', 'Cross-Origin-Embedder-Policy': 'require-corp' } } });
