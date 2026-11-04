# Cost calculator

Read the code in this order:

1. `src/cost/formula.ts` — the amount: quantity × price, then discount, then tax.
2. `src/cost/CostSheet.tsx` — the page: type quantities, see the total, save.
3. `src/webparts/costCalculator/sharepoint.ts` — writes the two SharePoint lists.

`preview/browserStore.ts` is the same page with the lists stored in the browser, so you can try it without a tenant.

A SharePoint Framework (SPFx) 1.24 React web part for SharePoint Online. Someone opens a site page, enters quantities, and sees the amount update. Saving writes the sheet to a list. Prices and extra rows live in a configuration list, so they can be changed without redeploying the app.

The Excel file mentioned with the request was not in the project. The page uses the cost-sheet formula below. Each configuration row says how it is priced, which is how a workbook is mapped onto the page:

| Spreadsheet row | Pricing type | Amount |
| --- | --- | --- |
| Quantity × rate | Per unit | quantity × unit price |
| Lump sum | Fixed amount | unit price once |
| Discount % | Percent discount | percent of the subtotal, in sort order, and never below zero |
| Tax or other % | Percent of net | percent of the amount after discount |

Subtotal is only the per-unit and fixed rows. Total = subtotal − discount + tax. Each line is rounded to cents before it is added. A saved sheet stores the prices used at that moment, so a later price change does not rewrite it.

Sample rows ship with the list (QAR): 8 hours at 450 and 16 hours at 320 total **8,720.00 QAR** before any optional row, discount, or tax.

## Try it in the browser

The preview uses the same screen as the web part and stores both lists in this browser.

```bash
npm install
npm install --prefix preview
npm run preview
```

Open http://127.0.0.1:43127

## Deploy to SharePoint Online

Requirements: Node.js 22.14 or newer, and permission to upload a package to the tenant app catalog and add an app to a site.

```bash
npm install
npm run build
```

The deployment file is `sharepoint/solution/cost-calculator.sppkg`. Upload that file to the app catalog and deploy it. Add the app to the site. Adding the app creates two lists:

- **Cost Configuration** — prices and items
- **Cost Calculations** — saved sheets

Open the site, edit a page (or the new page created below), and add the **Cost Calculator** web part. The first time the lists have no columns, choose **Create lists and page**. That adds the columns, sample prices when the configuration list is empty, and `SitePages/Cost-Calculator.aspx`, a full-page app page that hosts this web part.

To debug against a real site, set `initialPage` in `config/serve.json` to that page and run `npm start`. Load the debug scripts from the SharePoint bar. The hosted workbench is retired, so use the site page.

Web part properties, if the list titles or currency should differ:

- Configuration list title
- Calculations list title
- Currency code (default `QAR`)
- Page file name (default `Cost-Calculator.aspx`)

## Change a price or add an item

On the page, open **Prices**, edit the row or choose **Add item**, then **Save prices**. The same rows are the Cost Configuration list, so they can also be edited in the SharePoint list. **Active** hides a row without deleting it. **Optional** lets the person filling the sheet leave it out.

Use **Reload** after changing the list in SharePoint so the page picks up the new rows.

## Lists

**Cost Configuration:** Title, Category, Unit, Unit price, Pricing type (`PerUnit`, `Fixed`, `PercentDiscount`, `PercentOfNet`), Optional, Active, Sort order, Default quantity, Details.

**Cost Calculations:** Title, Requester, Department, Request date, Notes, Currency, Line items (JSON snapshot), Subtotal, Discount, Tax, Total, Status (`Draft` or `Submitted`).

Creating lists and the page needs permission to manage lists and Site Pages. The page shows the SharePoint error when that permission is missing.

## Formula checks

`npm run build` runs the unit tests for the formula, including the sample labor total, discount-then-tax, stacked discounts capped at the subtotal, and cent rounding.
