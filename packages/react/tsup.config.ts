import { defineConfig } from 'tsup';
import { execSync } from 'child_process';
import { readFileSync } from 'fs';
import { resolve } from 'path';

const pkg = JSON.parse(readFileSync(resolve(__dirname, 'package.json'), 'utf8')) as { version: string };
const commitHash = (() => {
  try { return execSync('git rev-parse --short HEAD', { encoding: 'utf8' }).trim(); } catch { return ''; }
})();

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    react: 'src/react.ts',
    ui: 'src/ui.ts',
    'core-reexport': 'src/core-reexport.ts',
    'headless-reexport': 'src/headless-reexport.ts',
    'core-plugins-reexport': 'src/core-plugins-reexport.ts',
    'mcp-reexport': 'src/mcp-reexport.ts',
  },
  format: ['cjs', 'esm'],
  dts: true,
  splitting: true,
  sourcemap: false,
  clean: true,
  treeshake: true,
  minify: true,
  external: [
    'react',
    'react-dom',
    'prosemirror-commands',
    'prosemirror-dropcursor',
    'prosemirror-history',
    'prosemirror-keymap',
    'prosemirror-model',
    'prosemirror-state',
    'prosemirror-tables',
    'prosemirror-transform',
    'prosemirror-view',
  ],
  injectStyle: false,
  define: {
    __APP_VERSION__: JSON.stringify(pkg.version),
    __COMMIT_HASH__: JSON.stringify(commitHash),
  },
});
