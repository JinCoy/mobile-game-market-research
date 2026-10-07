import {CheckCircle, PencilSimpleLine, WarningDiamond} from '@phosphor-icons/react';
import type {CellStatus} from '@/lib/schema';

/**
 * Data status, kept apart from the brand yellow: each state has its own icon, border style and label.
 * flagged = workbook yellow cell (estimate, reading, needs check) · verified = green cell · input = analyst input or proposal.
 */
export const statusMeta: Record<CellStatus, {label: string; Icon: typeof CheckCircle}> = {
  flagged: {label: '노란 칸 · 추정·판독·확인 필요', Icon: WarningDiamond},
  verified: {label: '확인됨', Icon: CheckCircle},
  input: {label: '분석자 입력·제안', Icon: PencilSimpleLine},
};

export default function StatusBadge({status, label}: {status: CellStatus; label?: string}) {
  const {Icon} = statusMeta[status];
  return <span className={`status-badge status-${status}`}><Icon size={14} weight="bold" aria-hidden="true"/>{label ?? statusMeta[status].label}</span>;
}
