# TikTok Stock Sync Tool

Mister Mobile Singapore — generates a TikTok Seller Center bulk stock-update
file from the POS Masterlist, using a persistent SKU match registry.
Companion to the Shopee Stock Sync Tool, same workflow and registry design.

Single file: everything lives in `app.py`.

## Deploy

Push `app.py` and `requirements.txt` to a GitHub repo, then
<https://share.streamlit.io> → **Create app** → your repo, branch `main`,
main file `app.py`.

Run locally: `pip install -r requirements.txt && streamlit run app.py`

## Inputs

| Upload | What it is | Required |
|---|---|---|
| POS Masterlist | Your POS stock report (`stock_report_DD-MM-YYYY.xlsx`) | Yes |
| TikTok Bulk Stock Inventory | Seller Center batch-edit export | Yes |
| SKU Registry | The match registry the app itself produces | From run 2 onward |
| Seed mapping *(optional)* | `POS stock ID and Tiktok Variation ID.xlsx` | First run only |

**POS Masterlist** — two-row header; Available Qty is read from the first
`Available Quantity` column (the company-wide Total block). Duplicate Stock
Type IDs are summed.

**TikTok file** — worksheet `Template`; row 1 machine keys, row 3 labels, data
from row 6. `sku_id` is the join key (same value as TikTok Variation ID).

## Stock update rules

**Locked Matches** — TikTok Quantity = sum of Available Qty of every linked
Masterlist SKU. A linked ID missing from today's POS report counts as 0 (sold
out) but stays in the registry, so the lock is never lost.

**Match Review** — TikTok listings not yet locked. Quantities left unchanged.
To promote a row, fill `Corrected Masterlist ID` and put a linking word
(`Linked`, `Confirmed`, `Locked`) in `Reviewer Decision`, then re-upload.

**New Masterlist SKUs** — POS SKUs with stock that aren't locked or classified.
Each needs a decision. Decisions persist through the exported registry.

## Reviewing

In the app: the Review tab ranks candidates against Seller SKU, Variation
Option and Product name. Pasted SKU IDs are rejected if not in the TikTok file.

In Excel, on the **New Masterlist SKUs** sheet:

| Column | What to do |
|---|---|
| H · Link to TikTok SKU ID | Paste a SKU ID from the *Match Review* sheet's column D. Excel warns on unknown IDs; you can override. |
| I · Reviewer Decision | Dropdown: `Linked (fill col H)` · `Not Selling in TikTok` · `Not on TikTok yet` |
| J · Notes | Free text, ignored by the app. |

Fill H only when the decision is `Linked (fill col H)`. Freebies, gift sets and
old used stock usually want `Not Selling in TikTok` with H blank.

The hidden `_TikTokSKUs` sheet powers the column-H warning — don't delete it.

## Outputs

**`Tiktok_Stock_Update_DD-MM-YYYY.xlsx`** — the original workbook re-saved with
only the `quantity` cells of locked matches changed. All worksheets, formatting
and other columns preserved.

**`Tiktok_Match_Review_UPDATED_DD-MM-YYYY.xlsx`** — the updated registry.
Save it and upload it next run.

The TikTok download stays disabled until every New Masterlist SKU has a
decision. The registry download is always available.

## Branding

Sampled from mistermobile.com.sg: `#FFEB00` yellow masthead, `#111111` black
nav, `#333333` text. Edit the `:root` block and the `MM_YELLOW` / `MM_BLACK`
constants in `app.py` to change it.

## Daily workflow

1. Export today's POS stock report and the TikTok batch-edit file.
2. Upload both, plus yesterday's SKU Registry.
3. Clear the Review tab (or fill columns H and I in Excel).
4. Download both files. Upload the TikTok one to Seller Center; keep the
   registry for tomorrow.
