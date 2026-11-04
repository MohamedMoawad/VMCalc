import { calculateCost } from '../src/cost/formula';
import { copyPrices, SAMPLE_PRICES } from '../src/cost/samplePrices';
import { CostStore, PriceItem, SavedSheet, SheetInput, Workspace } from '../src/cost/types';

// Same screen as SharePoint, but the two lists are saved in this browser.

const STORAGE_KEY = 'cost-sheet.v1';

interface LocalDb {
  prices: PriceItem[];
  saved: SavedSheet[];
  nextPriceId: number;
  nextSheetId: number;
}

export class BrowserStore implements CostStore {
  public readonly mode = 'preview' as const;
  public readonly currency: string;

  public constructor(currency: string) {
    this.currency = /^[A-Za-z]{3}$/.test(currency) ? currency.toUpperCase() : 'QAR';
  }

  public async load(): Promise<Workspace> {
    const db = this.read();
    return {
      ready: true,
      canSetup: true,
      prices: db.prices.map((item) => ({ ...item })),
      saved: db.saved.map((sheet) => ({ ...sheet, lines: sheet.lines.map((line) => ({ ...line })) }))
    };
  }

  public async saveSheet(input: SheetInput): Promise<SavedSheet> {
    const db = this.read();
    const result = calculateCost(input.lines);
    const sheet: SavedSheet = {
      id: db.nextSheetId,
      title: input.title.trim(),
      requesterName: input.requesterName.trim(),
      department: input.department.trim(),
      requestDate: input.requestDate,
      notes: input.notes.trim(),
      currency: this.currency,
      status: input.status,
      subtotal: result.subtotal,
      discountAmount: result.discountAmount,
      taxAmount: result.taxAmount,
      totalAmount: result.totalAmount,
      lines: result.lines,
      savedAt: new Date().toISOString()
    };
    db.nextSheetId += 1;
    db.saved = [sheet, ...db.saved];
    this.write(db);
    return sheet;
  }

  public async savePrices(items: readonly PriceItem[], removedIds: readonly number[]): Promise<void> {
    const db = this.read();
    const removed = removedIds.filter((id) => id > 0);
    const next: PriceItem[] = [];
    items.forEach((item) => {
      if (removed.indexOf(item.id) !== -1) {
        return;
      }
      if (item.id > 0) {
        next.push({ ...item });
        return;
      }
      next.push({ ...item, id: db.nextPriceId });
      db.nextPriceId += 1;
    });
    db.prices = next;
    this.write(db);
  }

  public async setup(): Promise<string> {
    return 'In SharePoint this creates the two lists and SitePages/Cost-Calculator.aspx.';
  }

  public async resetDemo(): Promise<void> {
    this.write(this.fresh());
  }

  private fresh(): LocalDb {
    return { prices: copyPrices(SAMPLE_PRICES), saved: [], nextPriceId: 100, nextSheetId: 1 };
  }

  private read(): LocalDb {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      const created = this.fresh();
      this.write(created);
      return created;
    }
    try {
      const parsed = JSON.parse(raw) as LocalDb;
      if (!parsed.prices || !parsed.saved) {
        return this.fresh();
      }
      return parsed;
    } catch {
      return this.fresh();
    }
  }

  private write(db: LocalDb): void {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(db));
  }
}
