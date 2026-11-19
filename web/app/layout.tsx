import "@fontsource-variable/archivo/wdth.css";
import "@fontsource-variable/source-serif-4/opsz.css";
import "./globals.css";
import type { Metadata, Viewport } from "next";
import type { ReactNode } from "react";

export const metadata: Metadata = {
  title: "Gulf solar tariffs",
  description:
    "What drives Saudi Arabia's 1.1 cent solar bids, and which parts of the design Oman's single buyer can copy. An outside-in case on public data.",
};

export const viewport: Viewport = { width: "device-width", initialScale: 1 };

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
