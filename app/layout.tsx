import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Taya Tool",
  description:
    "Taya's LinkedIn search builder and expert Excel list creation and update tools.",
  other: {
    "codex-preview": "development",
  },
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ja">
      <body>{children}</body>
    </html>
  );
}
