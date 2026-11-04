import { calculateCost } from './formula';
import { SAMPLE_PRICES } from './samplePrices';
import { LineInput, PriceItem } from './types';

function item(partial: Partial<PriceItem> & Pick<PriceItem, 'id' | 'title' | 'pricingType' | 'unitPrice'>): PriceItem {
  return {
    category: 'Service',
    unit: 'each',
    optional: false,
    active: true,
    sortOrder: partial.id * 10,
    defaultQuantity: 1,
    details: '',
    ...partial
  };
}

function input(config: PriceItem, quantity: number, included = true): LineInput {
  return { item: config, quantity, included };
}

describe('calculateCost', () => {
  it('prices the sample labor rows as quantity times unit price', () => {
    const labor = SAMPLE_PRICES.filter((row) => row.category === 'Labor' && !row.optional);
    const breakdown = calculateCost(labor.map((row) => input(row, row.defaultQuantity)));
    expect(breakdown.subtotal).toBe(8720);
    expect(breakdown.discountAmount).toBe(0);
    expect(breakdown.taxAmount).toBe(0);
    expect(breakdown.totalAmount).toBe(8720);
  });

  it('adds fixed amounts once and ignores an excluded optional row', () => {
    const hours = item({ id: 1, title: 'Hours', pricingType: 'PerUnit', unitPrice: 100, unit: 'hour', category: 'Labor' });
    const visit = item({ id: 2, title: 'Visit', pricingType: 'Fixed', unitPrice: 750, category: 'Service', optional: true });
    const materials = item({ id: 3, title: 'Materials', pricingType: 'Fixed', unitPrice: 1500, category: 'Material', optional: true });
    const breakdown = calculateCost([
      input(hours, 2),
      input(visit, 1, true),
      input(materials, 1, false)
    ]);
    expect(breakdown.subtotal).toBe(950);
    expect(breakdown.totalAmount).toBe(950);
    expect(breakdown.lines.find((line) => line.title === 'Materials')?.amount).toBe(0);
  });

  it('takes discount from the subtotal and tax from the net amount', () => {
    const hours = item({ id: 1, title: 'Hours', pricingType: 'PerUnit', unitPrice: 100, unit: 'hour', category: 'Labor', sortOrder: 10 });
    const discount = item({ id: 2, title: 'Discount', pricingType: 'PercentDiscount', unitPrice: 10, unit: '%', category: 'Discount', sortOrder: 80 });
    const tax = item({ id: 3, title: 'Tax', pricingType: 'PercentOfNet', unitPrice: 5, unit: '%', category: 'Tax', sortOrder: 90 });
    const breakdown = calculateCost([
      input(tax, 1),
      input(hours, 87.2),
      input(discount, 1)
    ]);
    expect(breakdown.subtotal).toBe(8720);
    expect(breakdown.discountAmount).toBe(872);
    expect(breakdown.netAmount).toBe(7848);
    expect(breakdown.taxAmount).toBe(392.4);
    expect(breakdown.totalAmount).toBe(8240.4);
    expect(breakdown.lines.map((line) => line.title)).toEqual(['Hours', 'Discount', 'Tax']);
  });

  it('applies several discounts to the original subtotal without going below zero', () => {
    const fee = item({ id: 1, title: 'Fee', pricingType: 'Fixed', unitPrice: 1000, category: 'Fee', sortOrder: 10 });
    const first = item({ id: 2, title: 'First', pricingType: 'PercentDiscount', unitPrice: 60, category: 'Discount', sortOrder: 20 });
    const second = item({ id: 3, title: 'Second', pricingType: 'PercentDiscount', unitPrice: 50, category: 'Discount', sortOrder: 30 });
    const breakdown = calculateCost([input(fee, 1), input(second, 1), input(first, 1)]);
    expect(breakdown.lines.find((line) => line.title === 'First')?.amount).toBe(600);
    expect(breakdown.lines.find((line) => line.title === 'Second')?.amount).toBe(400);
    expect(breakdown.discountAmount).toBe(1000);
    expect(breakdown.netAmount).toBe(0);
    expect(breakdown.warnings.some((warning) => warning.indexOf('limited') !== -1)).toBe(true);
  });

  it('adds each tax row as a percent of the same net amount', () => {
    const fee = item({ id: 1, title: 'Fee', pricingType: 'Fixed', unitPrice: 200, category: 'Fee' });
    const discount = item({ id: 2, title: 'Discount', pricingType: 'PercentDiscount', unitPrice: 50, category: 'Discount' });
    const local = item({ id: 3, title: 'Local', pricingType: 'PercentOfNet', unitPrice: 5, category: 'Tax', sortOrder: 30 });
    const other = item({ id: 4, title: 'Other', pricingType: 'PercentOfNet', unitPrice: 5, category: 'Tax', sortOrder: 40 });
    const breakdown = calculateCost([input(fee, 1), input(discount, 1), input(local, 1), input(other, 1)]);
    expect(breakdown.netAmount).toBe(100);
    expect(breakdown.taxAmount).toBe(10);
    expect(breakdown.totalAmount).toBe(110);
  });

  it('skips inactive rows and limits a percent that is over 100', () => {
    const hidden = item({ id: 1, title: 'Hidden', pricingType: 'Fixed', unitPrice: 999, active: false });
    const tax = item({ id: 2, title: 'Tax', pricingType: 'PercentOfNet', unitPrice: 150, category: 'Tax' });
    const fee = item({ id: 3, title: 'Fee', pricingType: 'Fixed', unitPrice: 80, category: 'Fee' });
    const breakdown = calculateCost([input(hidden, 1), input(tax, 1), input(fee, 1)]);
    expect(breakdown.lines.some((line) => line.title === 'Hidden')).toBe(false);
    expect(breakdown.taxAmount).toBe(80);
    expect(breakdown.totalAmount).toBe(160);
  });

  it('rounds each line to cents before summing', () => {
    const hours = item({ id: 1, title: 'Hours', pricingType: 'PerUnit', unitPrice: 10.1, unit: 'hour', category: 'Labor' });
    const breakdown = calculateCost([input(hours, 3)]);
    expect(breakdown.lines[0].amount).toBe(30.3);
    expect(breakdown.totalAmount).toBe(30.3);
  });
});
