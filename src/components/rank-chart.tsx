/** Placement on an ordinal axis, never a synthetic performance score. */
export default function RankChart({rank, limit, store, compact = false, missing = '순위 미수집'}: {
  rank?: number;
  limit: number;
  store: 'ios' | 'android';
  compact?: boolean;
  missing?: string;
}) {
  if (rank === undefined) return <span className="rank-missing" title={missing}>{compact ? '—' : missing}</span>;
  return <span className={`rank-chart ${compact ? 'rank-chart-compact' : ''}`}>
    <span className="rank-chart-value">{compact ? rank : `#${rank}`}</span>
    <span className="ordinal-track" role="img" aria-label={`1위부터 ${limit}위 공통 순위 축에서 ${rank}위`}>
      <i className={`ordinal-point ${store}`} style={{left: `${(rank-1)/Math.max(1,limit-1)*100}%`}}/>
    </span>
    {!compact && <span className="axis-ticks"><span>1위</span><span>{limit}위</span></span>}
  </span>;
}
