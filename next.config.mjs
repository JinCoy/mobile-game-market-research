const staticExport = process.env.PLAYFIELD_STATIC_EXPORT === 'true';

/** @type {import('next').NextConfig} */
const config = {
  ...(process.env.PLAYFIELD_DIST_DIR ? {distDir: process.env.PLAYFIELD_DIST_DIR} : {}),
  ...(staticExport ? {
    output: 'export',
    basePath: process.env.PLAYFIELD_BASE_PATH || '',
    trailingSlash: true,
  } : {}),
};

export default config;
