import type { Metadata } from 'next';
import '@fontsource-variable/inter';
import '@fontsource/ibm-plex-sans/latin-400.css';
import '@fontsource/ibm-plex-sans/latin-500.css';
import '@fontsource/ibm-plex-sans/latin-600.css';
import '@fontsource/ibm-plex-sans/latin-700.css';
import '@fontsource/noto-sans-kr/400.css';
import '@fontsource/noto-sans-kr/500.css';
import '@fontsource/noto-sans-kr/600.css';
import '@fontsource/noto-sans-kr/700.css';
import './globals.css';
import './design-theme.css';
import './research-enhancements.css';
export const metadata: Metadata = { title: 'PLAYFIELD — Mobile Game Market Research', description: '팀을 위한 모바일 게임 시장 리서치. 순위, 장르, 스토어 비교와 게임 전략을 원자료와 함께 분석합니다.' };
export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="ko"><body>{children}</body></html>;
}
