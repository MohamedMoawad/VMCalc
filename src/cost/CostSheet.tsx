import * as React from 'react';
import { calculateCost, formatMoney, formatPercent, sortPrices } from './formula';
import './cost-app.css';
import { copyPrices } from './samplePrices';
import {
  CostStore,
  isPricingType,
  LineInput,
  PriceItem,
  PRICING_LABEL,
  PRICING_TYPES,
  PricingType,
  ResultLine,
  SavedSheet,
  SheetInput,
  SheetStatus,
  Workspace
} from './types';

// The page has three tabs.
// Calculate: type quantities, see the amount, save.
// Prices:    change a price or add a row (the configuration list).
// Saved:     sheets already stored. Their totals do not change if prices change later.

export function CostSheet(props: {
  store: CostStore;
  userName: string;
  theme: 'light' | 'dark';
}): React.ReactElement {
  const [workspace, setWorkspace] = React.useState<Workspace | undefined>(undefined);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [notice, setNotice] = React.useState<string | undefined>(undefined);
  const [tab, setTab] = React.useState<'calculate' | 'prices' | 'saved'>('calculate');
  const [busy, setBusy] = React.useState(false);

  function reload(): void {
    setLoading(true);
    props.store.load().then((data) => {
      setWorkspace(data);
      setError(undefined);
      setLoading(false);
    }).catch((reason: unknown) => {
      setError(textOf(reason));
      setLoading(false);
    });
  }

  React.useEffect(() => {
    reload();
  }, [props.store]);

  function runSetup(): void {
    setBusy(true);
    props.store.setup().then((message) => {
      setNotice(message);
      return props.store.load();
    }).then((data) => {
      setWorkspace(data);
      setBusy(false);
      setTab('calculate');
    }).catch((reason: unknown) => {
      setError(textOf(reason));
      setBusy(false);
    });
  }

  return (
    <div className={props.theme === 'dark' ? 'cc-app cc-dark' : 'cc-app'}>
      <div className="cc-wrap">
        <header className="cc-header-row">
          <div>
            <p className="cc-kicker">Cost sheet</p>
            <h1>Calculate a cost</h1>
            <p className="cc-lead">
              Type a quantity. The amount updates immediately.
              Save stores the sheet, including the prices used, on the calculations list.
            </p>
          </div>
          <button className="cc-btn secondary" type="button" onClick={reload}>Reload</button>
        </header>

        {props.store.mode === 'preview' ? (
          <div className="cc-preview">
            Preview only: the lists are stored in this browser.
            In SharePoint the same screen uses the Cost Configuration and Cost Calculations lists.
            {props.store.resetDemo ? (
              <> <button className="cc-linkish" type="button" onClick={() => {
                props.store.resetDemo?.().then(reload).catch((reason: unknown) => setError(textOf(reason)));
              }}>Reset sample data</button></>
            ) : null}
          </div>
        ) : null}

        {notice ? <div className="cc-banner" role="status"><span>{notice}</span>
          <button className="cc-linkish" type="button" onClick={() => setNotice(undefined)}>Dismiss</button></div> : null}
        {error ? <div className="cc-error" role="alert">{error}</div> : null}
        {loading && !workspace ? <p>Loading the cost sheet…</p> : null}

        {workspace && !workspace.ready ? (
          <section className="cc-panel">
            <h2>Set up the lists</h2>
            <p>{workspace.message || 'This site does not have the cost lists yet.'}</p>
            <ul>
              <li>Cost Configuration — the prices. Change a price or add a row here.</li>
              <li>Cost Calculations — one row each time someone saves.</li>
              <li>Site Pages / Cost-Calculator.aspx — the page people open.</li>
            </ul>
            {workspace.canSetup ? (
              <button className="cc-btn" type="button" disabled={busy} onClick={runSetup}>
                {busy ? 'Creating…' : 'Create lists and page'}
              </button>
            ) : null}
          </section>
        ) : null}

        {workspace && workspace.ready ? (
          <>
            <div className="cc-tabs" role="tablist">
              {(['calculate', 'prices', 'saved'] as const).map((name) => (
                <button key={name} className="cc-tab" type="button" role="tab" aria-selected={tab === name} onClick={() => setTab(name)}>
                  {name === 'calculate' ? 'Calculate' : name === 'prices' ? 'Prices' : 'Saved'}
                </button>
              ))}
            </div>
            <div className="cc-panel">
              {tab === 'calculate' ? (
                <CalculateTab
                  prices={workspace.prices}
                  currency={props.store.currency}
                  userName={props.userName}
                  onSave={(input) => props.store.saveSheet(input).then((saved) => {
                    setNotice(`Saved “${saved.title}” as #${saved.id}.`);
                    return props.store.load().then(setWorkspace);
                  })}
                />
              ) : null}
              {tab === 'prices' ? (
                <PricesTab
                  prices={workspace.prices}
                  currency={props.store.currency}
                  listUrl={workspace.configListUrl}
                  onSave={(items, removedIds) => props.store.savePrices(items, removedIds).then(() => {
                    setNotice('Prices saved.');
                    return props.store.load().then(setWorkspace);
                  })}
                />
              ) : null}
              {tab === 'saved' ? (
                <SavedTab saved={workspace.saved} listUrl={workspace.savedListUrl} />
              ) : null}
            </div>
          </>
        ) : null}
      </div>
    </div>
  );
}

function CalculateTab(props: {
  prices: PriceItem[];
  currency: string;
  userName: string;
  onSave: (input: SheetInput) => Promise<void>;
}): React.ReactElement {
  const active = sortPrices(props.prices.filter((item) => item.active));
  const [title, setTitle] = React.useState('');
  const [requester, setRequester] = React.useState(props.userName);
  const [department, setDepartment] = React.useState('');
  const [requestDate, setRequestDate] = React.useState(today());
  const [notes, setNotes] = React.useState('');
  const [quantity, setQuantity] = React.useState<Record<number, number>>({});
  const [included, setIncluded] = React.useState<Record<number, boolean>>({});
  const [formError, setFormError] = React.useState<string | undefined>(undefined);
  const [saving, setSaving] = React.useState(false);

  const lines: LineInput[] = active.map((item) => ({
    item,
    included: included[item.id] !== undefined ? included[item.id] : !item.optional,
    quantity: quantity[item.id] !== undefined ? quantity[item.id] : item.defaultQuantity
  }));
  const result = calculateCost(lines);

  function save(status: SheetStatus): void {
    if (!title.trim() || !requester.trim()) {
      setFormError('Enter a name for the sheet and the requester.');
      return;
    }
    setSaving(true);
    setFormError(undefined);
    const input: SheetInput = {
      title: title.trim(),
      requesterName: requester.trim(),
      department: department.trim(),
      requestDate,
      notes: notes.trim(),
      currency: props.currency,
      status,
      lines
    };
    props.onSave(input).then(() => {
      setTitle('');
      setNotes('');
      setSaving(false);
    }).catch((reason: unknown) => {
      setFormError(textOf(reason));
      setSaving(false);
    });
  }

  let lastCategory = '';
  return (
    <div className="cc-calc">
      <div>
        <div className="cc-form">
          <label className="cc-field">Sheet *<input value={title} onChange={(event) => setTitle(event.target.value)} /></label>
          <label className="cc-field">Requester *<input value={requester} onChange={(event) => setRequester(event.target.value)} /></label>
          <label className="cc-field">Department<input value={department} onChange={(event) => setDepartment(event.target.value)} /></label>
          <label className="cc-field">Date<input type="date" value={requestDate} onChange={(event) => setRequestDate(event.target.value)} /></label>
          <label className="cc-field cc-notes">Notes<textarea rows={2} value={notes} onChange={(event) => setNotes(event.target.value)} /></label>
        </div>
        {!active.length ? <p>No active prices yet. Add one on the Prices tab.</p> : (
          <div className="cc-table-wrap">
            <table className="cc-table">
              <thead>
                <tr><th>Include</th><th>Item</th><th>Qty</th><th className="cc-num">Unit price</th><th className="cc-num">Amount</th></tr>
              </thead>
              <tbody>
                {lines.map((line) => {
                  const showCategory = line.item.category !== lastCategory;
                  lastCategory = line.item.category;
                  const shown = result.lines.find((entry) => entry.configItemId === line.item.id);
                  return (
                    <React.Fragment key={line.item.id}>
                      {showCategory ? <tr className="cc-group"><td colSpan={5}>{line.item.category}</td></tr> : null}
                      <tr>
                        <td>
                          <input
                            type="checkbox"
                            checked={line.included}
                            disabled={!line.item.optional}
                            aria-label={`Include ${line.item.title}`}
                            onChange={(event) => setIncluded({ ...included, [line.item.id]: event.target.checked })}
                          />
                        </td>
                        <td>
                          <div className="cc-item-title">{line.item.title}</div>
                          <div className="cc-item-meta">{PRICING_LABEL[line.item.pricingType]}{line.item.details ? ` · ${line.item.details}` : ''}</div>
                        </td>
                        <td>
                          {line.item.pricingType === 'PerUnit' ? (
                            <input
                              className="cc-qty"
                              type="number"
                              min={0}
                              step="0.01"
                              aria-label={`Quantity for ${line.item.title}`}
                              value={line.quantity}
                              disabled={!line.included}
                              onChange={(event) => setQuantity({
                                ...quantity,
                                [line.item.id]: event.target.value === '' ? 0 : Number(event.target.value)
                              })}
                            />
                          ) : '—'}
                        </td>
                        <td className="cc-num">
                          {isPercent(line.item.pricingType) ? formatPercent(line.item.unitPrice) : formatMoney(line.item.unitPrice, props.currency)}
                        </td>
                        <td className="cc-num">{line.included ? formatMoney(shown ? shown.amount : 0, props.currency) : '—'}</td>
                      </tr>
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
        <details className="cc-formula">
          <summary>How this amount is calculated</summary>
          <p>Per unit = quantity × unit price. Fixed = the price once. Those rows are the subtotal.</p>
          <p>Discount = a percent of the subtotal. Tax = a percent of the amount after discount.</p>
          <p>Total = subtotal − discount + tax. Saving keeps a copy of the prices used.</p>
        </details>
      </div>
      <aside className="cc-summary" aria-live="polite">
        <h2>Amount</h2>
        <Sum label="Subtotal" value={formatMoney(result.subtotal, props.currency)} />
        <Sum label="Discount" value={formatMoney(result.discountAmount, props.currency)} />
        <Sum label="After discount" value={formatMoney(result.netAmount, props.currency)} />
        <Sum label="Tax" value={formatMoney(result.taxAmount, props.currency)} />
        <p className="cc-total">{formatMoney(result.totalAmount, props.currency)}</p>
        {result.warnings.map((warning) => <p key={warning}>{warning}</p>)}
        {formError ? <p>{formError}</p> : null}
        <button className="cc-btn" type="button" disabled={saving} onClick={() => save('Submitted')}>{saving ? 'Saving…' : 'Save to list'}</button>
        <button className="cc-btn secondary" type="button" disabled={saving} onClick={() => save('Draft')}>Save as draft</button>
      </aside>
    </div>
  );
}

function PricesTab(props: {
  prices: PriceItem[];
  currency: string;
  listUrl?: string;
  onSave: (items: PriceItem[], removedIds: number[]) => Promise<void>;
}): React.ReactElement {
  const [rows, setRows] = React.useState<PriceItem[]>(() => copyPrices(props.prices));
  const [removed, setRemoved] = React.useState<number[]>([]);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [saving, setSaving] = React.useState(false);
  const nextTempId = React.useRef(-1);

  React.useEffect(() => {
    setRows(copyPrices(props.prices));
    setRemoved([]);
  }, [props.prices]);

  function patch(id: number, change: Partial<PriceItem>): void {
    setRows((current) => current.map((row) => row.id === id ? { ...row, ...change } : row));
  }

  function addRow(): void {
    const id = nextTempId.current;
    nextTempId.current -= 1;
    const sortOrder = rows.reduce((max, row) => Math.max(max, row.sortOrder), 0) + 10;
    setRows([...rows, {
      id,
      title: 'New item',
      category: 'Service',
      unit: 'each',
      unitPrice: 0,
      pricingType: 'PerUnit',
      optional: false,
      active: true,
      sortOrder,
      defaultQuantity: 1,
      details: ''
    }]);
  }

  function save(): void {
    const problem = validatePrices(rows);
    if (problem) {
      setError(problem);
      return;
    }
    setSaving(true);
    setError(undefined);
    props.onSave(rows.map((row) => ({ ...row, title: row.title.trim(), unit: row.unit.trim() || 'each' })), removed)
      .then(() => setSaving(false))
      .catch((reason: unknown) => {
        setError(textOf(reason));
        setSaving(false);
      });
  }

  return (
    <div>
      <p className="cc-lead">These rows are the configuration list. Change a price or add an item, then save. Amounts are in {props.currency}.</p>
      {props.listUrl ? <p><a href={props.listUrl} target="_blank" rel="noreferrer">Open the list in SharePoint</a></p> : null}
      <button className="cc-btn" type="button" onClick={addRow}>Add item</button>
      <div className="cc-cards">
        {rows.map((row) => (
          <article className="cc-card" key={row.id}>
            <label className="cc-field">Name<input value={row.title} onChange={(event) => patch(row.id, { title: event.target.value })} /></label>
            <label className="cc-field">Category<input value={row.category} onChange={(event) => patch(row.id, { category: event.target.value })} /></label>
            <label className="cc-field">Pricing
              <select value={row.pricingType} onChange={(event) => {
                if (isPricingType(event.target.value)) {
                  patch(row.id, { pricingType: event.target.value, unit: defaultUnit(event.target.value, row.unit) });
                }
              }}>
                {PRICING_TYPES.map((type) => <option key={type} value={type}>{PRICING_LABEL[type]}</option>)}
              </select>
            </label>
            <label className="cc-field">Unit<input value={row.unit} onChange={(event) => patch(row.id, { unit: event.target.value })} /></label>
            <label className="cc-field">{isPercent(row.pricingType) ? 'Percent' : 'Unit price'}
              <input type="number" min={0} max={isPercent(row.pricingType) ? 100 : undefined} step="0.01" value={row.unitPrice}
                onChange={(event) => patch(row.id, { unitPrice: Number(event.target.value) })} />
            </label>
            <label className="cc-field">Default quantity
              <input type="number" min={0} step="1" value={row.defaultQuantity}
                onChange={(event) => patch(row.id, { defaultQuantity: Number(event.target.value) })} />
            </label>
            <label className="cc-field">Sort order
              <input type="number" min={0} step="1" value={row.sortOrder}
                onChange={(event) => patch(row.id, { sortOrder: Number(event.target.value) })} />
            </label>
            <label className="cc-field cc-card-wide">Details
              <input value={row.details} onChange={(event) => patch(row.id, { details: event.target.value })} />
            </label>
            <div className="cc-checks cc-card-wide">
              <label><input type="checkbox" checked={row.optional} onChange={(event) => patch(row.id, { optional: event.target.checked })} /> Optional</label>
              <label><input type="checkbox" checked={row.active} onChange={(event) => patch(row.id, { active: event.target.checked })} /> Active</label>
              <button className="cc-btn danger small" type="button" onClick={() => {
                setRows(rows.filter((item) => item.id !== row.id));
                if (row.id > 0) {
                  setRemoved(removed.indexOf(row.id) === -1 ? [...removed, row.id] : removed);
                }
              }}>Remove</button>
            </div>
          </article>
        ))}
      </div>
      {error ? <p className="cc-error">{error}</p> : null}
      <button className="cc-btn" type="button" disabled={saving} onClick={save}>{saving ? 'Saving…' : 'Save prices'}</button>
    </div>
  );
}

function SavedTab(props: { saved: SavedSheet[]; listUrl?: string }): React.ReactElement {
  const [openId, setOpenId] = React.useState<number | undefined>(undefined);
  const open = props.saved.find((sheet) => sheet.id === openId);
  return (
    <div>
      <p className="cc-lead">Each saved row keeps the prices from the moment it was stored.</p>
      {props.listUrl ? <p><a href={props.listUrl} target="_blank" rel="noreferrer">Open the list in SharePoint</a></p> : null}
      {!props.saved.length ? <p>Nothing has been saved yet.</p> : (
        <table className="cc-table">
          <thead><tr><th>Ref</th><th>Sheet</th><th>Requester</th><th>Status</th><th className="cc-num">Total</th></tr></thead>
          <tbody>
            {props.saved.map((sheet) => (
              <tr key={sheet.id} className="cc-click-row" onClick={() => setOpenId(sheet.id)}>
                <td>#{sheet.id}</td>
                <td>{sheet.title}</td>
                <td>{sheet.requesterName}</td>
                <td>{sheet.status}</td>
                <td className="cc-num">{formatMoney(sheet.totalAmount, sheet.currency)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      {open ? <SavedDetail sheet={open} onClose={() => setOpenId(undefined)} /> : null}
    </div>
  );
}

function SavedDetail(props: { sheet: SavedSheet; onClose: () => void }): React.ReactElement {
  const sheet = props.sheet;
  return (
    <article className="cc-panel" style={{ marginTop: 16 }}>
      <h2>{sheet.title}</h2>
      <p className="cc-muted">#{sheet.id} · {sheet.requesterName}{sheet.department ? ` · ${sheet.department}` : ''}</p>
      {sheet.notes ? <p>{sheet.notes}</p> : null}
      <table className="cc-table">
        <thead><tr><th>Item</th><th className="cc-num">Amount</th></tr></thead>
        <tbody>
          {sheet.lines.filter((line) => line.included).map((line: ResultLine) => (
            <tr key={`${line.configItemId}-${line.title}`}>
              <td>{line.title}</td>
              <td className="cc-num">{formatMoney(line.amount, sheet.currency)}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <p className="cc-total" style={{ color: 'inherit', fontSize: '1.8rem' }}>{formatMoney(sheet.totalAmount, sheet.currency)}</p>
      <button className="cc-btn secondary" type="button" onClick={props.onClose}>Close</button>
    </article>
  );
}

function Sum(props: { label: string; value: string }): React.ReactElement {
  return <div className="cc-sum-row"><span>{props.label}</span><strong>{props.value}</strong></div>;
}

function validatePrices(rows: readonly PriceItem[]): string | undefined {
  const names: string[] = [];
  for (const row of rows) {
    const title = row.title.trim();
    if (!title) {
      return 'Every price item needs a name.';
    }
    if (names.indexOf(title.toLowerCase()) !== -1) {
      return `Two items are named “${title}”.`;
    }
    names.push(title.toLowerCase());
    if (row.unitPrice < 0 || (isPercent(row.pricingType) && row.unitPrice > 100)) {
      return `${title} has a price that this sheet cannot use.`;
    }
  }
  return undefined;
}

function isPercent(pricingType: PricingType): boolean {
  return pricingType === 'PercentDiscount' || pricingType === 'PercentOfNet';
}

function defaultUnit(pricingType: PricingType, current: string): string {
  if (isPercent(pricingType)) {
    return '%';
  }
  return current === '%' ? 'each' : current;
}

function today(): string {
  const now = new Date();
  const month = now.getMonth() + 1;
  const day = now.getDate();
  return `${now.getFullYear()}-${month < 10 ? '0' : ''}${month}-${day < 10 ? '0' : ''}${day}`;
}

function textOf(reason: unknown): string {
  if (reason instanceof Error && reason.message) {
    return reason.message;
  }
  return 'Something went wrong. Try again.';
}
