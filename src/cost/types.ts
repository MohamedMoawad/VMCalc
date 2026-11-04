// The nouns used by the page.
// A price row lives in the configuration list.
// A saved sheet lives in the calculations list.

export type PricingType = 'PerUnit' | 'Fixed' | 'PercentDiscount' | 'PercentOfNet';

export type SheetStatus = 'Draft' | 'Submitted';

export interface PriceItem {
  id: number;
  title: string;
  category: string;
  unit: string;
  unitPrice: number;
  pricingType: PricingType;
  optional: boolean;
  active: boolean;
  sortOrder: number;
  defaultQuantity: number;
  details: string;
}

// One row on the open sheet: the price, plus what the user typed.
export interface LineInput {
  item: PriceItem;
  included: boolean;
  quantity: number;
}

// One row frozen at save time, so a later price change does not rewrite it.
export interface ResultLine {
  configItemId: number;
  title: string;
  category: string;
  unit: string;
  pricingType: PricingType;
  quantity: number;
  unitPrice: number;
  amount: number;
  included: boolean;
  details: string;
}

export interface CostResult {
  lines: ResultLine[];
  subtotal: number;
  discountAmount: number;
  netAmount: number;
  taxAmount: number;
  totalAmount: number;
  warnings: string[];
}

export interface SheetInput {
  title: string;
  requesterName: string;
  department: string;
  requestDate: string;
  notes: string;
  currency: string;
  status: SheetStatus;
  lines: LineInput[];
}

export interface SavedSheet {
  id: number;
  title: string;
  requesterName: string;
  department: string;
  requestDate: string;
  notes: string;
  currency: string;
  status: SheetStatus;
  subtotal: number;
  discountAmount: number;
  taxAmount: number;
  totalAmount: number;
  lines: ResultLine[];
  savedAt?: string;
}

export interface Workspace {
  ready: boolean;
  canSetup: boolean;
  message?: string;
  prices: PriceItem[];
  saved: SavedSheet[];
  pageUrl?: string;
  configListUrl?: string;
  savedListUrl?: string;
}

// What the page asks for. SharePoint and the browser preview both implement this.
export interface CostStore {
  readonly mode: 'sharepoint' | 'preview';
  readonly currency: string;
  load(): Promise<Workspace>;
  saveSheet(input: SheetInput): Promise<SavedSheet>;
  savePrices(items: readonly PriceItem[], removedIds: readonly number[]): Promise<void>;
  setup(): Promise<string>;
  resetDemo?(): Promise<void>;
}

export const PRICING_TYPES: readonly PricingType[] = [
  'PerUnit',
  'Fixed',
  'PercentDiscount',
  'PercentOfNet'
];

export const PRICING_LABEL: Record<PricingType, string> = {
  PerUnit: 'Per unit',
  Fixed: 'Fixed amount',
  PercentDiscount: 'Percent discount',
  PercentOfNet: 'Percent of net'
};

export function isPricingType(value: string): value is PricingType {
  return PRICING_TYPES.indexOf(value as PricingType) !== -1;
}
