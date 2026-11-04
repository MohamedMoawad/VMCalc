import { PriceItem } from './types';

// Starting rows. Replace the prices after the list is created.
// 8 × 450 + 16 × 320 = 8,720 before any optional row, discount, or tax.
export const SAMPLE_PRICES: readonly PriceItem[] = [
  price(1, 'Senior specialist', 'Labor', 'hour', 450, 'PerUnit', false, 8, 10, 'Sample rate. Change it on Prices.'),
  price(2, 'Specialist', 'Labor', 'hour', 320, 'PerUnit', false, 16, 20, 'Sample rate. Change it on Prices.'),
  price(3, 'Coordinator', 'Labor', 'hour', 220, 'PerUnit', true, 4, 30, 'Optional. Tick the row to include it.'),
  price(4, 'Site visit', 'Service', 'visit', 750, 'Fixed', true, 1, 40, 'Charged once when the row is included.'),
  price(5, 'Materials allowance', 'Material', 'lot', 1500, 'Fixed', true, 1, 50, 'Lump sum. Replace the amount with yours.'),
  price(6, 'Discount', 'Discount', '%', 0, 'PercentDiscount', true, 1, 80, 'Percent of the subtotal.'),
  price(7, 'Tax', 'Tax', '%', 0, 'PercentOfNet', true, 1, 90, 'Percent of the amount after discount.')
];

export function copyPrices(items: readonly PriceItem[]): PriceItem[] {
  return items.map((item) => ({ ...item }));
}

function price(
  id: number,
  title: string,
  category: string,
  unit: string,
  unitPrice: number,
  pricingType: PriceItem['pricingType'],
  optional: boolean,
  defaultQuantity: number,
  sortOrder: number,
  details: string
): PriceItem {
  return {
    id,
    title,
    category,
    unit,
    unitPrice,
    pricingType,
    optional,
    active: true,
    sortOrder,
    defaultQuantity,
    details
  };
}
