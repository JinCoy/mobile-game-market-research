/** Rank is ordinal: longer bars mean a higher placement in the selected Top N. */
export default function RankChart({rank, limit, store, compact = false}: {
  rank?: number;
  limit: number;
  store: 'ios' | 'android';
  compact?: boolean;
}) {
  if (rank === undefined) return <span className="rank-missing">{compact ? '—' : '선택 범위에서 미수집'}</span>;
  const width = (limit - rank + 1) / limit * 100;
  return <span className={`rank-chart ${compact ? 'rank-chart-compact' : ''}`}>
    <span className="rank-chart-value">{compact ? rank : `#${rank}`}</span>
    <span className="rank-chart-track" role="img" aria-label={`Top ${limit} 중 ${rank}위, 막대가 길수록 상위`}>
      <span className={`chart-bar ${store === 'ios' ? 'ios-bar' : 'android-bar'}`} style={{width: `${width}%`}}/>
    </span>
  </span>;
}
