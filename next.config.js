/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  // Enables "import X from './file?raw'" -- used by the app itself to embed
  // its own MKIS.jsx source for the Download Centre. Without this rule,
  // Next.js's webpack build silently resolves the import to an empty module,
  // which is why the downloaded MKIS.jsx used to come out blank.
  webpack(config) {
    config.module.rules.push({ resourceQuery: /raw/, type: "asset/source" });
    return config;
  },
};
module.exports = nextConfig;
