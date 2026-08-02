# TikTok Stock Sync Tool

Mister Mobile Singapore — generates a TikTok Seller Center bulk stock-update file
from the POS Masterlist, using a persistent SKU match registry.
Companion to the Shopee Stock Sync Tool, same workflow and same registry design.

---

## Deploy on Streamlit Community Cloud

1. Create a GitHub repo (e.g. `tiktok-stock-sync`) and push these files:

   ```
   app.py            <- the whole tool, single file
   requirements.txt
   .streamlit/config.toml
   README.md
   .gitignore
   ```

   There is no `core.py` any more — everything lives in `app.py` so the two
   files can never fall out of sync on GitHub.

2. Go to <https://share.streamlit.io> → **New app** → pick the repo,
   branch `main`, main file `app.py` → **Deploy**.

No secrets or API keys are needed — everything runs on uploaded files.

### Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## Inputs

| Upload | What it is | Required |
|---|---|---|
| **POS Masterlist** | Your POS stock report (`stock_report_DD-MM-YYYY.xlsx`) | Yes |
| **TikTok Bulk Stock Inventory** | Seller Center batch-edit export | Yes |
| **SKU Registry** | The match registry the app itself produces | Yes, from run 2 onward |
| **Seed mapping** *(optional)* | `POS stock ID and Tiktok Variation ID.xlsx` — bootstrap only | First run only |

### POS Masterlist format
Two-row header. Row 1: `Stock Type ID · Category · Brand · Model · Color · Total · <branches>`.
Row 2: sub-headers. **Available Qty** is read from the first `Available Quantity`
column (the company-wide Total block). Duplicate Stock Type IDs are summed.

### TikTok file format
Worksheet **`Template`**. Row 1 holds machine keys (`product_id`, `sku_id`,
`quantity`, `seller_sku`…), row 3 holds human labels, data starts at **row 6**.
`sku_id` is the join key — the same value as *TikTok Variation ID*.

---

## Stock update rules

**Locked Matches** — TikTok Quantity = sum of `Available Qty` of every linked
Masterlist SKU. If a linked Masterlist ID is missing from today's POS report it
counts as **0** (the item sold out), and the ID is still kept in the registry so
the lock is never lost. Logged as *info*, not an error.

**Match Review** — every TikTok listing not yet locked to a Masterlist ID.
Quantities are **left completely unchanged**. To promote a row, fill
`Corrected Masterlist ID` and put a linking word (`Linked`, `Confirmed`,
`Locked`) in `Reviewer Decision`, then re-upload.

**New Masterlist SKUs** — POS SKUs with `Available Qty > 0` that are not in
Locked Matches, Not Selling in TikTok, or Not on TikTok Yet. Each needs a
decision. Decisions persist through the exported registry.

**Not Selling in TikTok / Not on TikTok Yet** — SKUs you've classified. They are
excluded from future review.

---

## Reviewing

Two ways to do it, and they read the same registry file.

### In the app
The Review tab ranks TikTok candidates against **Seller SKU**, **Variation
Option** and **Product name**. Pick a suggestion or paste a SKU ID — a pasted ID
is rejected if it isn't in the uploaded TikTok file.

### In Excel
On the **New Masterlist SKUs** sheet:

| Column | What to do |
|---|---|
| **H · Link to TikTok SKU ID** | Paste a SKU ID from the *Match Review* sheet's column D. Excel warns if the ID isn't real — you can override. |
| **I · Reviewer Decision** | Dropdown: `Linked (fill col H)` · `Not Selling in TikTok` · `Not on TikTok yet` |
| **J · Notes** | Free text, ignored by the app. |

Only fill column H when the decision is `Linked (fill col H)`. Freebies, gift
sets and old used stock usually want `Not Selling in TikTok` with H left blank.

The dropdown is backed by a hidden `_TikTokSKUs` worksheet holding every valid
SKU ID. Don't delete it — that's what powers the column-H warning.

---

## Outputs

**`Tiktok_Stock_Update_DD-MM-YYYY.xlsx`** — the original workbook reopened and
re-saved with **only** the `quantity` cells of locked matches modified. All
worksheets, formatting, row order and every other column are preserved. Verified
against the real 3,420-row file: only quantity cells differed.

**`Tiktok_Match_Review_UPDATED_DD-MM-YYYY.xlsx`** — the updated registry, same
sheet layout as `Shopee_Match_Review`, in Mister Mobile colours.
**Save this and upload it next run.**

### Export gate
The TikTok download stays disabled until every New Masterlist SKU has a
decision, a registry or seed mapping is present, and all required worksheets
exist. The registry download is always available so review progress is never lost.

---

## Branding

Colours sampled from mistermobile.com.sg: **`#FFEB00`** yellow masthead,
**`#111111`** black nav bar, **`#333333`** body text.

- App: yellow header over a black strip, black buttons with yellow text, yellow
  tab underline, metric cards with a yellow left rule.
- Excel: yellow title band on Summary, black header rows with yellow text,
  pale-yellow fills on the columns you fill in.

To change the palette, edit the `:root` block and the `MM_YELLOW` / `MM_BLACK`
constants in `app.py`, plus `[theme]` in `.streamlit/config.toml`.

---

## Daily workflow

1. Export today's POS stock report and the TikTok batch-edit file.
2. Upload both, plus yesterday's SKU Registry.
3. Clear the *Review New Masterlist SKUs* tab (or fill columns H and I in Excel).
4. Download **both** files. Upload the TikTok file to Seller Center; keep the
   registry for tomorrow.
