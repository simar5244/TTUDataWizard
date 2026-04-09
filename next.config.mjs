/** @type {import('next').NextConfig} */
const nextConfig = {
  experimental: {
    serverActionsBodySizeLimit: "50mb",
  },
};

export default nextConfig;
