import type { NextConfig } from "next";

// Static export: Vercel serves the built files; nothing runs on a server.
const config: NextConfig = {
  output: "export",
  images: { unoptimized: true },
};

export default config;
