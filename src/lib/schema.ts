export type Store = 'ios' | 'android';
export type EvidenceKind = 'observed' | 'secondary' | 'estimate' | 'aiClassification' | 'analystReading' | 'proposal';
export type Source = { file: string; sheet: string; row: number; url?: string | null; range?: string; kind?: EvidenceKind; originalText?: unknown; date?: string | null; urls?: string[] };
export type Game = {
  id: string; name: string; publisher: string; developer: string | null;
  genre: string; subgenre: string | null; missing: boolean;
  gameplay: string | null; monetization: string | null; marketing: string | null;
  artStyle: string | null; audience: string | null; playerAge: string | null;
  marketingHook: string | null;
  downloads: {value: number; kind: string; date: string; store: string; country: string | null} | null;
  /** Google Play official install band and listing month (3-3 F·G). `flagged` lists yellow-cell columns. */
  installRange: {range: string; released: string | null; flagged: string[]; source: Source} | null;
  revenue: number | null; analysis: Source | null; evidence: string | null;
  tags: {label: string; field: string; originalText: string; source: Source; kind: EvidenceKind}[];
};
export type Observation = {
  id: string; gameId: string; name: string; publisher: string; genre: string;
  subgenre: string | null; country: string; store: Store; date: string; year: number;
  rank: number; chart: 'free' | 'grossing'; notes: string | null; source: Source;
  /** The source row has a yellow cell in the workbook (estimate, reading or needs check). */
  flagged: boolean;
};
export type ResearchDocument = {id: string; file: string; sheet: string; rows: (string | number | boolean | null)[][]; formulas: Record<string,string|undefined>; blankStates: Record<string,string|undefined>;
  /** Workbook cell meaning: yellow fill = estimate/reading/needs check, green fill = verified, blue font = analyst input/assumption. */
  cellStatus: Record<string, CellStatus[] | undefined>};
export type CellStatus = 'flagged' | 'verified' | 'input';
export type TocItem = {group: string; sheet: string; description: string | null; status: string | null; note: string | null};
export type ChangelogEntry = {stage: string; date: string | null; sources: string | null; limits: string | null; row: number | null; sheets: {file: string; sheet: string; note: string}[]};
export type Benchmark = {group: string; label: string; metric: string; value: number; unit: string; percentile: string; publicationYear: number; period: string; source: Source};
export type AdCreative = {gameId:string; name:string; adId:string; days:number; hook:string; format:string; tone:string; offer:string; copy:string; actualType:string; requestedType:string; description:string; videoGroup:string; source:Source; readingSource:Source};
export type PlanPhase = {name:string; timing:string; region:string; os:string|null; channel:string; channelRank:string|null; budget:number; cpiId:string|null; cpi:number|null; installs:number|null; retentionRegion:string|null; d7Rate:number|null; d7:number|null; criterion:string; reason:string|null; source:Source; cpiSource:Source|null; retentionSource:Source|null; cpiRegion:string|null; cpiPeriod:string|null};
export type MvpCandidate = {name:string; gameId:string; subgenre:string; score:number; decision:string; roles:string; duration:string; scope:string; caution:string; source:Source};
export type Dataset = {schemaVersion: number; importedAt: string; games: Game[]; observations: Observation[];
  documents: ResearchDocument[]; files: {name: string; sheets: number}[];
  publisherAliases:Record<string,string>;
  advertisingChannels: {name:string; scores:number[]; score:number; rank:number; reason:string; source:Source; public:{name:string; buying:string; reach:string; grade:string; performance:string; genre:string; cost:string; text:string; source:Source}}[];
  channelAxes:string[]; channelWeights:number[];
  adBrandStats:{gameId:string; name:string; active:number; tracked:number; classified:number; longestDays:number; spendEstimate:string; note:string; source:Source}[];
  adHookSamples:{gameId:string; name:string; table:string; label:string; description:string; count:number; sampleSize:number; source:Source}[];
  adCreatives:AdCreative[];
  marketingPlan:{candidate:string; summary:string; source:Source; phases:PlanPhase[]; materials:{name:string; opening:string; reference:string; difficulty:string; source:Source}[]; metrics:{name:string; meaning:string; criterion:string; reference:string; source:Source}[]};
  regionalBenchmarks:Benchmark[];
  osDistribution:{group:string; region:string; android:number; ios:number; other:number; month:string; note:string|null; source:Source}[];
  osPriorities:{region:string; priority:string; reason:string; source:Source}[];
  osGameMarket:{revenue:(string|number)[]; downloads:(string|number)[]; revenueSource:Source; downloadSource:Source};
  mvpAssessments:{name:string; gameId:string; genre:string; subgenre:string; grossingRank:number|string; freeRank:number|string; classification:string; candidate:string; score:number|null; decision:string; source:Source; scoreState:string}[];
  mvpRecommended:MvpCandidate[];
  mvpCriteria:{genre:string; subgenre:string; scores:(number|null)[]; score:number|null; decision:string; reason:string; source:Source}[];
  mvpTargets:{name:string; gameId:string; region:string; genre:string; d1:{criterion:string;target:number;median:number}; d7:{criterion:string;target:number;median:number}; d30:{criterion:string;target:number;median:number}; iapArpu:number|null; adArpu:number|null; iapArppu:number|null; source:Source}[];
  mvpSpecs:{name:string;values:(string|number|null)[];source:Source}[];
  seaAnalysis:{text:string;evidence:string;kind:string;source:Source}[];
  toc: {file: string; items: TocItem[]}[];
  changelog: ChangelogEntry[];
  documentReferences:{file:string;sheet:string;originalFile:string;originalSheet:string}[];
};
