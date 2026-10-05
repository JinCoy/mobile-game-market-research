const staticExport = process.env.PLAYFIELD_STATIC_EXPORT === 'true';

/** @type {import('next').NextConfig} */
const config = {
  ...(staticExport ? {
    output: 'export',
    basePath: process.env.PLAYFIELD_BASE_PATH || '',
    trailingSlash: true,
  } : {}),
};

export default config;
