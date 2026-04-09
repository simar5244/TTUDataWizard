export { default } from "next-auth/middleware";

export const config = {
  matcher: [
    "/dashboard/:path*",
    "/mapper/:path*",
    "/dashboards/:path*",
    "/settings/:path*",
  ],
};
