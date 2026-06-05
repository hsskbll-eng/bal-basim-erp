import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "AHİ MATBAA",
  description: "Matbaa yönetim sistemi",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="tr">
      <body>{children}</body>
    </html>
  );
}
