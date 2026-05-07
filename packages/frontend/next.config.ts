import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  reactStrictMode: true,
  transpilePackages: ["dhtmlx-gantt"],
  async redirects() {
    return [
      {
        source: "/irradiation/machine1/:path*",
        destination: "/machine1/:path*",
        permanent: false,
      },
      {
        source: "/irradiation/machine-status",
        destination: "/operations/machine-status",
        permanent: false,
      },
      {
        source: "/irradiation/production",
        destination: "/operations/production",
        permanent: false,
      },
      {
        source: "/admin/db-browser",
        destination: "/db-browser",
        permanent: false,
      },
      {
        source: "/admin/shipment-method-verify",
        destination: "/reports/shipment-method-verify",
        permanent: false,
      },
      {
        source: "/admin/shipment-method-ss",
        destination: "/reports/shipment-method-ss",
        permanent: false,
      },
    ];
  },
};

export default nextConfig;
