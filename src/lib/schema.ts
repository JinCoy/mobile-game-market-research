export type Store = 'ios' | 'android';
export type Source = { file: string; sheet: string; row: number; url?: string | null };
export type Game = {
  id: string; name: string; publisher: string; developer: string | null;
  genre: string; subgenre: string | null; missing: boolean;
  gameplay: string | null; monetization: string | null; marketing: string | null;
  artStyle: string | null; audience: string | null; playerAge: string | null;
  marketingHook: string | null;
  downloads: {value: number; kind: string; date: string; store: string; country: string | null} | null;
  revenue: number | null; analysis: Source | null; evidence: string | null;
};
export type Observation = {
  id: string; gameId: string; name: string; publisher: string; genre: string;
  subgenre: string | null; country: string; store: Store; date: string; year: number;
  rank: number; chart: 'free' | 'grossing'; notes: string | null; source: Source;
};
export type ResearchDocument = {id: string; file: string; sheet: string; rows: (string | number | boolean | null)[][]};
export type Dataset = {schemaVersion: number; importedAt: string; games: Game[]; observations: Observation[];
  documents: ResearchDocument[]; files: {name: string; sheets: number}[]};
