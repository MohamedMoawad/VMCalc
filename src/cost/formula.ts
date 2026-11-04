import { CostResult, LineInput, PriceItem, PricingType, ResultLine } from './types';

// Category order on the sheet. Anything else is treated like "Other".
const CATEGORY_RANK: Record<string, number> = {
  Labor: 0,
  Material: 1,
  Service: 2,
  Fee: 3,
  Other: 4,
  Discount: 5,
  Tax: 6
};

// The workbook was not included, so each price row says which rule to use:
//   Per unit          quantity × unit price
//   Fixed             unit price once
//   Percent discount  percent of the subtotal (stops at zero)
//   Percent of net    percent of what remains after discount (tax)
// Total = subtotal − discount + tax.

export function roundMoney(value: number): number {
  if (!Number.isFinite(value)) {
    return 0;
  }
  const sign = value < 0 ? -1 : 1;
  const rounded = Number(Math.round(Number(`${Math.abs(value)}e2`)) + 'e-2');
  return sign * rounded;
}

export function calculateCost(inputs: readonly LineInput[]): CostResult {
  const warnings: string[] = [];
  const rows = inputs
    .filter((row) => row.item.active)
    .map((row) => cleanRow(row, warnings))
    .sort((a, b) => comparePrice(a.item, b.item));

  const amount: number[] = [];
  let subtotal = 0;

  // 1. Subtotal is only quantity × price, and fixed fees.
  for (let index = 0; index < rows.length; index++) {
    const row = rows[index];
    const isCharge = row.item.pricingType === 'PerUnit' || row.item.pricingType === 'Fixed';
    if (!row.included || !isCharge) {
      amount[index] = 0;
      continue;
    }
    const lineAmount = row.item.pricingType === 'Fixed'
      ? roundMoney(row.item.unitPrice)
      : roundMoney(row.quantity * row.item.unitPrice);
    amount[index] = lineAmount;
    subtotal = roundMoney(subtotal + lineAmount);
  }

  // 2. Each discount is a percent of that subtotal, in row order.
  let discountLeft = subtotal;
  let discountAmount = 0;
  for (let index = 0; index < rows.length; index++) {
    const row = rows[index];
    if (row.item.pricingType !== 'PercentDiscount' || !row.included) {
      continue;
    }
    const wanted = roundMoney(subtotal * row.item.unitPrice / 100);
    const applied = Math.min(wanted, discountLeft);
    if (applied < wanted) {
      warnings.push('Discount was limited so it cannot be more than the subtotal.');
    }
    amount[index] = applied;
    discountLeft = roundMoney(discountLeft - applied);
    discountAmount = roundMoney(discountAmount + applied);
  }

  // 3. Each tax row is a percent of the amount after discount.
  const netAmount = roundMoney(subtotal - discountAmount);
  let taxAmount = 0;
  for (let index = 0; index < rows.length; index++) {
    const row = rows[index];
    if (row.item.pricingType !== 'PercentOfNet' || !row.included) {
      continue;
    }
    const applied = roundMoney(netAmount * row.item.unitPrice / 100);
    amount[index] = applied;
    taxAmount = roundMoney(taxAmount + applied);
  }

  return {
    lines: rows.map((row, index) => toLine(row, row.included ? amount[index] : 0)),
    subtotal,
    discountAmount,
    netAmount,
    taxAmount,
    totalAmount: roundMoney(netAmount + taxAmount),
    warnings: unique(warnings)
  };
}

export function sortPrices(items: readonly PriceItem[]): PriceItem[] {
  return [...items].sort(comparePrice);
}

export function formatMoney(amount: number, currency: string): string {
  const code = /^[A-Z]{3}$/i.test(currency) ? currency.toUpperCase() : 'QAR';
  try {
    return new Intl.NumberFormat('en-GB', {
      style: 'currency',
      currency: code,
      currencyDisplay: 'code',
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(roundMoney(amount));
  } catch {
    return `${roundMoney(amount).toFixed(2)} ${code}`;
  }
}

export function formatPercent(value: number): string {
  return `${roundMoney(value)}%`;
}

function comparePrice(a: PriceItem, b: PriceItem): number {
  const rankA = CATEGORY_RANK[a.category] ?? 4;
  const rankB = CATEGORY_RANK[b.category] ?? 4;
  if (rankA !== rankB) {
    return rankA - rankB;
  }
  if (a.sortOrder !== b.sortOrder) {
    return a.sortOrder - b.sortOrder;
  }
  return a.title.localeCompare(b.title);
}

function cleanRow(row: LineInput, warnings: string[]): LineInput {
  let quantity = row.quantity;
  let unitPrice = row.item.unitPrice;
  if (!Number.isFinite(quantity) || quantity < 0) {
    warnings.push(`${row.item.title}: quantity cannot be negative, so it was treated as zero.`);
    quantity = 0;
  }
  if (!Number.isFinite(unitPrice) || unitPrice < 0) {
    warnings.push(`${row.item.title}: price cannot be negative, so it was treated as zero.`);
    unitPrice = 0;
  }
  if (isPercent(row.item.pricingType) && unitPrice > 100) {
    warnings.push(`${row.item.title}: percent was limited to 100.`);
    unitPrice = 100;
  }
  return { included: row.included, quantity, item: { ...row.item, unitPrice } };
}

function toLine(row: LineInput, amount: number): ResultLine {
  return {
    configItemId: row.item.id,
    title: row.item.title,
    category: row.item.category,
    unit: row.item.unit,
    pricingType: row.item.pricingType,
    quantity: row.item.pricingType === 'PerUnit' ? row.quantity : 1,
    unitPrice: row.item.unitPrice,
    amount,
    included: row.included,
    details: row.item.details
  };
}

function isPercent(pricingType: PricingType): boolean {
  return pricingType === 'PercentDiscount' || pricingType === 'PercentOfNet';
}

function unique(warnings: string[]): string[] {
  const seen: string[] = [];
  warnings.forEach((warning) => {
    if (seen.indexOf(warning) === -1) {
      seen.push(warning);
    }
  });
  return seen;
}
