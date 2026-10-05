import type { Metadata } from 'next';
import './globals.css';
export const metadata: Metadata = { title: 'PLAYFIELD — Mobile Game Market Research', description: '팀을 위한 모바일 게임 시장 리서치. 순위, 장르, 스토어 비교와 게임 전략을 원자료와 함께 분석합니다.' };
export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="ko"><body>{children}</body></html>;
}
