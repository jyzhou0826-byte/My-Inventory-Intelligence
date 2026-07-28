import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({ variable: "--font-geist-sans", subsets: ["latin"] });
const geistMono = Geist_Mono({ variable: "--font-geist-mono", subsets: ["latin"] });

export const metadata: Metadata = {
  title: "RV 庫存決策儀表板",
  description: "My Inventory Intelligence — RV 庫存健康、風險、策略備料與生產計畫分析",
  icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
  openGraph: {
    title: "RV 庫存決策儀表板",
    description: "2026 Q1–Q3 庫存健康、關鍵風險、QoQ 與生產計畫分析",
    images: [{ url: "/og.png", width: 1735, height: 909, alt: "RV 庫存決策儀表板" }],
  },
  twitter: { card: "summary_large_image", images: ["/og.png"] },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="zh-Hant-TW"><body className={`${geistSans.variable} ${geistMono.variable}`}>{children}</body></html>;
}
