import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "엑셀 가격 일괄 변경기",
  description: "원본 엑셀 구조를 유지하고 가격만 일괄 변경합니다.",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
