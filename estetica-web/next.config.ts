import path from "path";
import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  /** Evita confusión si hay otro `package-lock.json` fuera de esta carpeta. */
  outputFileTracingRoot: path.join(__dirname),
};

export default nextConfig;
