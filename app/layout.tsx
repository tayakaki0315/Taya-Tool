import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Taya Expert List Builder",
  description: "Turn expert profiles into a formatted Excel list.",
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
