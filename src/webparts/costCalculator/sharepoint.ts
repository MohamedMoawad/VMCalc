import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { calculateCost } from '../../cost/formula';
import { copyPrices, SAMPLE_PRICES } from '../../cost/samplePrices';
import {
  CostStore,
  isPricingType,
  PriceItem,
  PricingType,
  ResultLine,
  SavedSheet,
  SheetInput,
  SheetStatus,
  Workspace
} from '../../cost/types';

// This file only talks to SharePoint.
// The math is in src/cost/formula.ts. The screen is src/cost/CostSheet.tsx.
//
// Two lists:
//   Cost Configuration  = price rows
//   Cost Calculations   = saved sheets

const CONFIG_LIST = 'Cost Configuration';
const SAVED_LIST = 'Cost Calculations';

interface Column {
  name: string;
  title: string;
  kind: 'text' | 'note' | 'number' | 'boolean' | 'choice' | 'date';
  choices?: string[];
  defaultValue?: string;
}

const CONFIG_COLUMNS: Column[] = [
  { name: 'Category', title: 'Category', kind: 'text' },
  { name: 'Unit', title: 'Unit', kind: 'text' },
  { name: 'UnitPrice', title: 'Unit price', kind: 'number' },
  { name: 'PricingType', title: 'Pricing type', kind: 'choice', choices: ['PerUnit', 'Fixed', 'PercentDiscount', 'PercentOfNet'], defaultValue: 'PerUnit' },
  { name: 'Optional', title: 'Optional', kind: 'boolean', defaultValue: '0' },
  { name: 'Active', title: 'Active', kind: 'boolean', defaultValue: '1' },
  { name: 'SortOrder', title: 'Sort order', kind: 'number' },
  { name: 'DefaultQuantity', title: 'Default quantity', kind: 'number' },
  { name: 'Details', title: 'Details', kind: 'note' }
];

class SpError extends Error {
  public constructor(message: string, public readonly status: number) {
    super(message);
    this.name = 'SpError';
  }
}

const SAVED_COLUMNS: Column[] = [
  { name: 'RequesterName', title: 'Requester', kind: 'text' },
  { name: 'Department', title: 'Department', kind: 'text' },
  { name: 'RequestDate', title: 'Request date', kind: 'date' },
  { name: 'Notes', title: 'Notes', kind: 'note' },
  { name: 'Currency', title: 'Currency', kind: 'text' },
  { name: 'LineItemsJson', title: 'Line items', kind: 'note' },
  { name: 'Subtotal', title: 'Subtotal', kind: 'number' },
  { name: 'DiscountAmount', title: 'Discount', kind: 'number' },
  { name: 'TaxAmount', title: 'Tax', kind: 'number' },
  { name: 'TotalAmount', title: 'Total', kind: 'number' },
  { name: 'Status', title: 'Status', kind: 'choice', choices: ['Draft', 'Submitted'], defaultValue: 'Draft' }
];

export class SharePointStore implements CostStore {
  public readonly mode = 'sharepoint' as const;
  public readonly currency: string;
  private readonly configList: string;
  private readonly savedList: string;

  private readonly webPartId: string;
  private readonly pageName: string;

  public constructor(
    private readonly context: WebPartContext,
    options: { configListTitle: string; calculationsListTitle: string; currency: string; pageName?: string; webPartId: string }
  ) {
    this.configList = (options.configListTitle || '').trim() || CONFIG_LIST;
    this.savedList = (options.calculationsListTitle || '').trim() || SAVED_LIST;
    this.currency = /^[A-Za-z]{3}$/.test(options.currency || '') ? options.currency.toUpperCase() : 'QAR';
    this.webPartId = options.webPartId;
    const page = (options.pageName || 'Cost-Calculator.aspx').trim() || 'Cost-Calculator.aspx';
    this.pageName = page.toLowerCase().endsWith('.aspx') ? page.replace(/\s+/g, '-') : `${page.replace(/\s+/g, '-')}.aspx`;
  }

  private get web(): string {
    return this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
  }

  public async load(): Promise<Workspace> {
    const config = await this.readList(this.configList);
    if (!config.ok) {
      return emptyWorkspace(config.canSetup, config.message);
    }
    const saved = await this.readList(this.savedList);
    if (!saved.ok) {
      return emptyWorkspace(saved.canSetup, saved.message, config.viewUrl);
    }
    const missing = (await this.missingColumns(this.configList, CONFIG_COLUMNS))
      .concat(await this.missingColumns(this.savedList, SAVED_COLUMNS));
    if (missing.length) {
      return emptyWorkspace(true, 'The cost lists are missing columns. Choose Create lists and page.', config.viewUrl, saved.viewUrl);
    }
    return {
      ready: true,
      canSetup: true,
      prices: await this.loadPrices(),
      saved: await this.loadSaved(),
      configListUrl: config.viewUrl,
      savedListUrl: saved.viewUrl,
      pageUrl: await this.findPage()
    };
  }

  public async saveSheet(input: SheetInput): Promise<SavedSheet> {
    const result = calculateCost(input.lines);
    const id = await this.addItem(this.savedList, {
      Title: input.title.trim(),
      RequesterName: input.requesterName.trim(),
      Department: input.department.trim(),
      RequestDate: input.requestDate ? `${input.requestDate.slice(0, 10)}T00:00:00Z` : undefined,
      Notes: input.notes.trim(),
      Currency: this.currency,
      LineItemsJson: JSON.stringify(result.lines),
      Subtotal: result.subtotal,
      DiscountAmount: result.discountAmount,
      TaxAmount: result.taxAmount,
      TotalAmount: result.totalAmount,
      Status: input.status
    });
    return {
      id,
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
      lines: result.lines
    };
  }

  public async savePrices(items: readonly PriceItem[], removedIds: readonly number[]): Promise<void> {
    for (const id of removedIds) {
      if (id > 0) {
        await spPost(this.context, `${listUrl(this.web, this.configList)}/items(${id})`, {}, {
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
        });
      }
    }
    for (const item of items) {
      const fields = priceFields(item);
      if (item.id > 0) {
        await this.mergeItem(this.configList, item.id, fields);
      } else {
        await this.addItem(this.configList, fields);
      }
    }
  }

  public async setup(): Promise<string> {
    const notes: string[] = [];
    if (await this.ensureList(this.configList, 'Prices for the cost calculator. Change a price or add a row.')) {
      notes.push(`Created ${this.configList}.`);
    }
    if (await this.ensureList(this.savedList, 'Saved cost sheets.')) {
      notes.push(`Created ${this.savedList}.`);
    }
    const added = (await this.ensureColumns(this.configList, CONFIG_COLUMNS))
      .concat(await this.ensureColumns(this.savedList, SAVED_COLUMNS));
    if (added.length) {
      notes.push('Added the columns.');
    }
    if (await this.seedIfEmpty()) {
      notes.push('Added sample prices. Replace them with yours.');
    }
    try {
      const page = await this.ensurePage();
      if (page) {
        notes.push('Calculator page is ready.');
      }
    } catch (reason: unknown) {
      const message = reason instanceof Error ? reason.message : 'The page was not created.';
      notes.push(`Lists are ready. ${message}`);
    }
    return notes.join(' ') || 'Setup finished.';
  }

  private async loadPrices(): Promise<PriceItem[]> {
    const data = await spGet<{ value?: RawPrice[] }>(this.context,
      `${listUrl(this.web, this.configList)}/items?$select=Id,Title,Category,Unit,UnitPrice,PricingType,Optional,Active,SortOrder,DefaultQuantity,Details&$top=500`);
    const items: PriceItem[] = [];
    (data.value || []).forEach((row) => {
      const item = toPrice(row);
      if (item) {
        items.push(item);
      }
    });
    return items;
  }

  private async loadSaved(): Promise<SavedSheet[]> {
    const data = await spGet<{ value?: RawSaved[] }>(this.context,
      `${listUrl(this.web, this.savedList)}/items?$select=Id,Title,RequesterName,Department,RequestDate,Notes,Currency,LineItemsJson,Subtotal,DiscountAmount,TaxAmount,TotalAmount,Status,Created&$orderby=Id desc&$top=200`);
    return (data.value || []).map(toSaved);
  }

  private async readList(title: string): Promise<{ ok: true; viewUrl?: string } | { ok: false; canSetup: boolean; message: string }> {
    try {
      const list = await spGet<{ DefaultViewUrl?: string }>(this.context, `${listUrl(this.web, title)}?$select=DefaultViewUrl,ItemCount`);
      return { ok: true, viewUrl: absolute(this.web, list.DefaultViewUrl) };
    } catch (reason: unknown) {
      const status = reason instanceof SpError ? reason.status : 0;
      if (status === 404) {
        return { ok: false, canSetup: true, message: `The list “${title}” is not on this site yet.` };
      }
      if (status === 403) {
        return { ok: false, canSetup: false, message: 'You do not have permission to read the cost lists.' };
      }
      throw reason;
    }
  }

  private async missingColumns(title: string, columns: Column[]): Promise<string[]> {
    const data = await spGet<{ value?: { InternalName: string }[] }>(this.context, `${listUrl(this.web, title)}/fields?$select=InternalName`);
    const names = (data.value || []).map((field) => field.InternalName);
    return columns.filter((column) => names.indexOf(column.name) === -1).map((column) => column.name);
  }

  // Returns true when the list was created now.
  private async ensureList(title: string, description: string): Promise<boolean> {
    try {
      await spGet(this.context, `${listUrl(this.web, title)}?$select=Id`);
      return false;
    } catch (reason: unknown) {
      if (!(reason instanceof SpError) || reason.status !== 404) {
        throw reason;
      }
    }
    await spPost(this.context, `${this.web}/_api/web/lists`, {
      __metadata: { type: 'SP.List' },
      BaseTemplate: 100,
      Title: title,
      Description: description
    });
    return true;
  }

  private async ensureColumns(title: string, columns: Column[]): Promise<string[]> {
    const missing = await this.missingColumns(title, columns);
    const list = await spGet<{ Id: string }>(this.context, `${listUrl(this.web, title)}?$select=Id`);
    const added: string[] = [];
    for (const column of columns) {
      if (missing.indexOf(column.name) === -1) {
        continue;
      }
      await spPost(this.context, `${listUrl(this.web, title)}/fields`, columnBody(column));
      if (column.title !== column.name) {
        await spPost(this.context, `${listUrl(this.web, title)}/fields/getbyinternalnameortitle('${column.name}')`, {
          __metadata: { type: 'SP.Field' },
          Title: column.title
        }, { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' }).catch(() => undefined);
      }
      await spPost(this.context, `${this.web}/_api/web/lists(guid'${list.Id.replace(/[{}]/g, '')}')/DefaultView/ViewFields/AddViewField`, {
        strField: column.name
      }).catch(() => undefined);
      added.push(column.title);
    }
    return added;
  }

  private async seedIfEmpty(): Promise<boolean> {
    const list = await spGet<{ ItemCount?: number }>(this.context, `${listUrl(this.web, this.configList)}?$select=ItemCount`);
    if ((list.ItemCount || 0) > 0) {
      return false;
    }
    for (const sample of copyPrices(SAMPLE_PRICES)) {
      await this.addItem(this.configList, priceFields(sample));
    }
    return true;
  }

  private async addItem(title: string, fields: Record<string, unknown>): Promise<number> {
    const list = await spGet<{ ListItemEntityTypeFullName?: string }>(this.context, `${listUrl(this.web, title)}?$select=ListItemEntityTypeFullName`);
    const created = await spPost<{ d?: { Id?: number; ID?: number } }>(this.context, `${listUrl(this.web, title)}/items`, {
      __metadata: { type: list.ListItemEntityTypeFullName },
      ...fields
    });
    const id = created.d && (created.d.Id || created.d.ID);
    if (!id) {
      throw new SpError('SharePoint did not return an id for the new row.', 500);
    }
    return id;
  }

  private async mergeItem(title: string, id: number, fields: Record<string, unknown>): Promise<void> {
    const list = await spGet<{ ListItemEntityTypeFullName?: string }>(this.context, `${listUrl(this.web, title)}?$select=ListItemEntityTypeFullName`);
    await spPost(this.context, `${listUrl(this.web, title)}/items(${id})`, {
      __metadata: { type: list.ListItemEntityTypeFullName },
      ...fields
    }, { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' });
  }

  // Creates SitePages/Cost-Calculator.aspx with this web part on it.
  private async ensurePage(): Promise<string | undefined> {
    const existing = await this.findPage();
    if (existing) {
      return existing;
    }
    const instanceId = createGuid();
    const created = await spPost<Record<string, unknown>>(this.context, `${this.web}/_api/sitepages/Pages/CreateFullPageApp`, {
      webPartDataAsJson: JSON.stringify({
        id: this.webPartId,
        instanceId,
        title: 'Cost Calculator',
        description: 'Calculate a cost and save it',
        dataVersion: '1.0',
        properties: {
          configListTitle: this.configList,
          calculationsListTitle: this.savedList,
          currency: this.currency,
          pageName: 'Cost-Calculator.aspx'
        }
      }),
      title: this.pageName.replace(/\.aspx$/i, ''),
      addToQuickLaunch: true
    });
    return readUrl(created) || await this.findPage();
  }

  private async findPage(): Promise<string | undefined> {
    const root = this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, '');
    const path = `${root}/SitePages/${this.pageName}`;
    try {
      const file = await spGet<{ ServerRelativeUrl?: string }>(this.context,
        `${this.web}/_api/web/getfilebyserverrelativeurl('${path.replace(/'/g, "''")}')?$select=ServerRelativeUrl`);
      return absolute(this.web, file.ServerRelativeUrl);
    } catch (reason: unknown) {
      if (reason instanceof SpError && reason.status === 404) {
        return undefined;
      }
      return undefined;
    }
  }
}

// Reads come back as plain JSON: { value: [...] }.
async function spGet<T>(context: WebPartContext, url: string): Promise<T> {
  const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
    headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' }
  });
  return read<T>(response);
}

// Writes use verbose JSON because new list items need __metadata.
async function spPost<T>(context: WebPartContext, url: string, body: unknown, extra?: Record<string, string>): Promise<T> {
  const response = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      ...extra
    },
    body: JSON.stringify(body)
  });
  return read<T>(response);
}

async function read<T>(response: SPHttpClientResponse): Promise<T> {
  const text = await response.text();
  if (!response.ok) {
    throw new SpError(cleanMessage(text, response.status), response.status);
  }
  if (!text) {
    return {} as T;
  }
  return JSON.parse(text) as T;
}

function listUrl(web: string, title: string): string {
  return `${web}/_api/web/lists/getbytitle('${title.replace(/'/g, "''")}')`;
}

function columnBody(column: Column): Record<string, unknown> {
  const kind = column.kind === 'note' ? 3
    : column.kind === 'date' ? 4
      : column.kind === 'choice' ? 6
        : column.kind === 'boolean' ? 8
          : column.kind === 'number' ? 9
            : 2;
  const metadata = column.kind === 'choice' ? 'SP.FieldChoice'
    : column.kind === 'number' ? 'SP.FieldNumber'
      : column.kind === 'note' ? 'SP.FieldMultiLineText'
        : column.kind === 'date' ? 'SP.FieldDateTime'
          : 'SP.Field';
  return {
    __metadata: { type: metadata },
    Title: column.name,
    FieldTypeKind: kind,
    Required: false,
    ...(column.choices ? { Choices: { results: column.choices }, DefaultValue: column.defaultValue } : {}),
    ...(column.kind === 'boolean' && column.defaultValue ? { DefaultValue: column.defaultValue } : {}),
    ...(column.kind === 'note' ? { RichText: false, NumberOfLines: 6 } : {}),
    ...(column.kind === 'date' ? { DisplayFormat: 0 } : {})
  };
}

function priceFields(item: PriceItem): Record<string, unknown> {
  return {
    Title: item.title.trim(),
    Category: item.category.trim() || 'Other',
    Unit: item.unit.trim() || 'each',
    UnitPrice: item.unitPrice,
    PricingType: item.pricingType,
    Optional: item.optional,
    Active: item.active,
    SortOrder: item.sortOrder,
    DefaultQuantity: item.defaultQuantity,
    Details: item.details.trim()
  };
}

interface RawPrice {
  Id?: number;
  Title?: string;
  Category?: string;
  Unit?: string;
  UnitPrice?: number;
  PricingType?: string;
  Optional?: boolean;
  Active?: boolean;
  SortOrder?: number;
  DefaultQuantity?: number;
  Details?: string;
}

function toPrice(row: RawPrice): PriceItem | undefined {
  if (!row.Id || !(row.Title || '').trim()) {
    return undefined;
  }
  const pricingType: PricingType = row.PricingType && isPricingType(row.PricingType) ? row.PricingType : 'PerUnit';
  return {
    id: row.Id,
    title: (row.Title || '').trim(),
    category: row.Category || 'Other',
    unit: row.Unit || 'each',
    unitPrice: numberOr(row.UnitPrice, 0),
    pricingType,
    optional: !!row.Optional,
    active: row.Active !== false,
    sortOrder: numberOr(row.SortOrder, 0),
    defaultQuantity: numberOr(row.DefaultQuantity, 1),
    details: row.Details || ''
  };
}

interface RawSaved {
  Id?: number;
  Title?: string;
  RequesterName?: string;
  Department?: string;
  RequestDate?: string;
  Notes?: string;
  Currency?: string;
  LineItemsJson?: string;
  Subtotal?: number;
  DiscountAmount?: number;
  TaxAmount?: number;
  TotalAmount?: number;
  Status?: string;
  Created?: string;
}

function toSaved(row: RawSaved): SavedSheet {
  const status: SheetStatus = row.Status === 'Submitted' ? 'Submitted' : 'Draft';
  return {
    id: row.Id || 0,
    title: row.Title || 'Cost sheet',
    requesterName: row.RequesterName || '',
    department: row.Department || '',
    requestDate: (row.RequestDate || '').slice(0, 10),
    notes: row.Notes || '',
    currency: row.Currency || 'QAR',
    status,
    subtotal: numberOr(row.Subtotal, 0),
    discountAmount: numberOr(row.DiscountAmount, 0),
    taxAmount: numberOr(row.TaxAmount, 0),
    totalAmount: numberOr(row.TotalAmount, 0),
    lines: parseLines(row.LineItemsJson),
    savedAt: row.Created
  };
}

function parseLines(json: string | undefined): ResultLine[] {
  if (!json) {
    return [];
  }
  try {
    const parsed: unknown = JSON.parse(json);
    if (!Array.isArray(parsed)) {
      return [];
    }
    return parsed.filter((line) => !!line && typeof line === 'object') as ResultLine[];
  } catch {
    return [];
  }
}

function emptyWorkspace(canSetup: boolean, message: string, configListUrl?: string, savedListUrl?: string): Workspace {
  return { ready: false, canSetup, message, prices: [], saved: [], configListUrl, savedListUrl };
}

function numberOr(value: unknown, fallback: number): number {
  const parsed = typeof value === 'number' ? value : Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function absolute(web: string, path: string | undefined): string | undefined {
  if (!path) {
    return undefined;
  }
  if (path.indexOf('http') === 0) {
    return path;
  }
  return `${new URL(web).origin}${path.charAt(0) === '/' ? path : `/${path}`}`;
}

function readUrl(payload: Record<string, unknown>): string | undefined {
  const direct = payload.AbsoluteUrl || payload.Url;
  if (typeof direct === 'string') {
    return direct;
  }
  const wrapped = payload.d;
  if (wrapped && typeof wrapped === 'object') {
    return readUrl(wrapped as Record<string, unknown>);
  }
  return undefined;
}

function createGuid(): string {
  const cryptoApi = globalThis.crypto;
  if (cryptoApi && typeof cryptoApi.randomUUID === 'function') {
    return cryptoApi.randomUUID();
  }
  return `cost-${Date.now()}`;
}

function cleanMessage(body: string, status: number): string {
  try {
    const parsed = JSON.parse(body) as { error?: { message?: { value?: string } | string }; 'odata.error'?: { message?: { value?: string } } };
    const verbose = parsed['odata.error'] && parsed['odata.error'].message && parsed['odata.error'].message.value;
    const modern = parsed.error && (typeof parsed.error.message === 'string' ? parsed.error.message : parsed.error.message && parsed.error.message.value);
    const message = verbose || modern || '';
    if (message) {
      return message.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
    }
  } catch {
    // The body was not JSON.
  }
  const cleaned = body.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  return cleaned || `SharePoint returned status ${status}.`;
}
