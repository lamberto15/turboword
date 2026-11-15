import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    compose: 'src/compose/index.ts',
    fields: 'src/fields/index.ts',
    forms: 'src/forms/index.ts',
    crypto: 'src/crypto/index.ts',
    layout: 'src/layout/index.ts',
    pdf: 'src/render/pdf/index.ts',
    svg: 'src/render/svg/index.ts',
  },
  format: ['esm', 'cjs'],
  dts: true,
  platform: 'neutral',
  target: 'es2022',
  minify: false,
  sourcemap: false,
  clean: true,
  unbundle: true,
});
