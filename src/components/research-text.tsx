// Emphasis changes presentation only; the original research wording stays intact.
const emphasis = /(Top\s*\d+|D(?:1|7|30|90)\b|\d[\d,.]*(?:\s*[~→]\s*\d[\d,.]*)?\s*(?:%|개국|개|건|위|점|초|분|개월|년|[MB]\+?)|\b(?:IAP|BM|UA|ROAS|ARPU|ARPPU)\b|보상형 광고|전면 광고|광고 중심|광고 수익|인앱결제|인앱 결제|메타 진행|라이브 이벤트|한 손가락|레벨 진행|매치3|미니멀|미확인|미수집|추정)/g;

export default function ResearchText({text}: {text: string}) {
  return <>{text.replace(/[\u{1F000}-\u{1FFFF}\u2600-\u27BF\uFE0F]/gu,'').split(emphasis).map((part,index) => index % 2 === 0 ? part :
    <strong key={index} className={/추정|미확인|미수집/.test(part) ? 'text-caution' : /\d/.test(part) ? 'text-number' : 'text-keyword'}>{part}</strong>)}</>;
}
