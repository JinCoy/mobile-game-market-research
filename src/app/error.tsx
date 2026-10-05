'use client';
export default function Error({reset}: {reset: () => void}) {return <main className="loading"><h1>자료를 불러오지 못했습니다.</h1><button onClick={reset}>다시 시도</button></main>;}
