import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from 'tailwindcss';
import autoprefixer from 'autoprefixer';
import path from 'path';
import { execSync } from 'child_process';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);
const pkg = require('../../packages/react/package.json') as { version: string };
const commitHash = (() => {
  try { return execSync('git rev-parse --short HEAD', { encoding: 'utf8' }).trim(); } catch { return ''; }
})();

const monorepoRoot = path.resolve(__dirname, '../..');

async function fetchGitHubStars(): Promise<number | null> {
  try {
    const res = await fetch('https://api.github.com/repos/eigenpal/docx-editor');
    const data = await res.json();
    if (typeof data.stargazers_count === 'number') return data.stargazers_count;
  } catch {}
  return null;
}

export default defineConfig(async () => {
  const stars = await fetchGitHubStars();
  return {
    plugins: [react()],
    root: __dirname,
    resolve: {
      alias: [
        // Resolve package imports to source for live development
        // Order matters: more-specific prefixes before less-specific ones
        {
          find: '@eigenpal/docx-js-editor',
          replacement: path.join(monorepoRoot, 'packages/react/src/index.ts'),
        },
        {
          find: '@eigenpal/docx-core/headless',
          replacement: path.join(monorepoRoot, 'packages/core/src/headless.ts'),
        },
        {
          find: '@eigenpal/docx-core/core-plugins',
          replacement: path.join(monorepoRoot, 'packages/core/src/core-plugins/index.ts'),
        },
        {
          find: '@eigenpal/docx-core/mcp',
          replacement: path.join(monorepoRoot, 'packages/core/src/mcp/index.ts'),
        },
        // Wildcard alias for deep core imports (e.g. @eigenpal/docx-core/utils/docxInput)
        {
          find: /^@eigenpal\/docx-core\/(.+)/,
          replacement: path.join(monorepoRoot, 'packages/core/src/$1'),
        },
        // Exact match for bare @eigenpal/docx-core (must come AFTER the prefix match above)
        {
          find: /^@eigenpal\/docx-core$/,
          replacement: path.join(monorepoRoot, 'packages/core/src/core.ts'),
        },
        { find: '@', replacement: path.join(monorepoRoot, 'packages/react/src') },
        // Deduplicate React to avoid "multiple instances" hook errors
        { find: 'react', replacement: path.join(__dirname, 'node_modules/react') },
        { find: 'react-dom', replacement: path.join(__dirname, 'node_modules/react-dom') },
        { find: 'react/jsx-runtime', replacement: path.join(__dirname, 'node_modules/react/jsx-runtime') },
      ],
    },
    css: {
      postcss: {
        plugins: [
          tailwindcss({ config: path.join(monorepoRoot, 'tailwind.config.js') }),
          autoprefixer(),
        ],
      },
    },
    define: {
      __ENABLE_FRAMEWORK_SWITCHER__: JSON.stringify(
        process.env.ENABLE_FRAMEWORK_SWITCHER === 'true'
      ),
      __GITHUB_STARS__: JSON.stringify(stars),
      __APP_VERSION__: JSON.stringify(pkg.version),
      __COMMIT_HASH__: JSON.stringify(commitHash),
    },
    server: {
      port: 5173,
      open: false,
    },
    build: {
      outDir: 'dist',
    },
  };
});
