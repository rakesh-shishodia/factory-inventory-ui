# Factory Inventory

An internal stock and picking app for 3DPrintronics. The new version runs on Cloudflare Workers, D1, and Queues. The original Apps Script and Google Sheets implementation remains at the repository root for reference.

**Current delivery: a working local demo including supplier-backed receiving/allocation/picking, variation-aware integration code, read-only pilot reviews, and an inactive opening-stock staging endpoint. The picker is not operating the real store.** No real pilot items have been imported or aligned to Ecwid. Exact opening-order import, approved alignment/activation and an operator reconciliation workflow are still required before production use.

## Run locally

Use Node **24 LTS** (Node 22 or newer required), then:

```sh
npm ci
npm run setup:demo
npm run dev
```

Open <http://127.0.0.1:8787>. Demo sign-in is automatic on loopback only. Sample data stays in `.wrangler/state`; rerunning setup preserves existing demo picks and movements. Local stock updates use the queue but never contact Ecwid. The QR scanner dependency is bundled locally, so the app does not load a script from a public CDN.

The demo contains two Paid orders (`DEMO-1001`, `DEMO-1002`) and one Awaiting Payment order (`DEMO-1003`). Try `NUT-M3`, `BOLT-M3-12`, or `BEARING-608` in the item lookup. Camera scanning requires a supported browser and permission; manual entry also works. The preferred QR content is the item's unique SKU, including a variation's own SKU. Legacy `SKU||LOCATION` codes still work. `SKU|NUMERIC_COMBINATION_ID|LOCATION` also works when that ID matches the mapped variation; a size/colour label in the middle field is rejected because it is not a stable identity. Enter the variation's unique SKU instead. Scanning does not create or change mappings.

Run `npm run setup:variation-demo` to add `DEMO-BOLT-M6-20` and `DEMO-BOLT-M6-30`, two independently stocked sizes under the same fictitious parent. Their sample orders are `DEMO-VARIATION-PAID` and `DEMO-VARIATION-AWAITING`. This additive fixture preserves existing stock movements and order progress when rerun.

## Inventory rules

For stock-limited items, `available = physical_on_hand − outstanding_reservations`. The table below describes that mode. Supplier-backed items use the separate allocation rules below.

| Event | Physical stock | Reserved | Ecwid inventory write |
| --- | --- | --- | --- |
| Ecwid order accepted as Paid/Awaiting Payment | No change | Increases | None: Ecwid handles checkout |
| Awaiting Payment changes to Paid | No change | No change | None; picking becomes available |
| Paid order picked | Decreases | Decreases | None: avoids double deduction |
| Email/offline sale outside Ecwid | Decreases | No change | Negative quantity delta |
| In-house use | Decreases | No change | Negative quantity delta |
| New stock received | Increases | No change | Positive quantity delta |
| Order cancelled before picking | No change | Released | None: Ecwid restores it |

Paid orders must also have an eligible fulfillment state (`AWAITING_PROCESSING` or `PROCESSING`). For this store, **Ready for Pickup means already picked and packed, equivalent to Shipped**; it does not claim customer delivery. It is not pickable and its status change creates no second stock deduction. Already-tracked orders with missing picks are quarantined instead of inventing physical movements or releasing reservations. Picking never marks an order shipped. Each line tracks its remaining quantity and supports partial picking. Email orders already created in Ecwid must use the Ecwid picking flow, otherwise stock would be deducted twice.

This version supports single-unit base products and independently stocked variations with their own unique SKUs. Each variation is an app item mapped to an Ecwid product ID, combination ID, and exact option selection; its physical stock is separate from sibling sizes. It must have its own tracked whole-unit quantity, an enabled parent, no extra non-variation options, and no bundle/component relationships. The parent container is not counted as a second stock item when its SKU equals a child's SKU. Two independent targets sharing a SKU remain ambiguous and blocked. Confirm that one Ecwid quantity unit really means one physical piece; a unique SKU alone does not prove that it is not a pack or measured length.

Orders containing unsupported/unmapped lines, edited order lines, cancellation/refund after picking, deleted mappings, and missing picks are flagged for review. A known supplier shortage alone does not block other mapped lines in a mixed order. Physical returns are **not** normal restocks: Ecwid may already have restored their sellable stock. A return/reconciliation workflow must handle that distinction before such cases are used live.

## What is implemented

- Mobile and desktop screens for order picking, stock movements, inventory search, activity, and sync health.
- Immutable opening-balance and movement ledgers. A SQL write records movement, changes physical stock, updates picked quantity, and creates pending Ecwid work atomically.
- Client operation UUIDs retained across uncertain responses and reloads. The server deduplicates exact retries and rejects changed payloads using the same UUID.
- Database guards against unpaid orders, excess picks, unavailable stock, and unresolved item issues, including concurrent requests.
- Signed Ecwid stock deltas, a durable D1 outbox, Queue delivery, and scheduled redelivery.
- Verified Ecwid webhook signatures, deduplicated durable inbox, fresh order snapshots, and resumable paginated order polling.
- Cloudflare Access JWT verification with administrator roles. Demo authentication works only on loopback and cannot open a deployed API.
- An opening-stock preview that validates source balances, exact base/variation mappings, per-SKU physical units, reservations, duplicates, and eligibility.
- An administrator-only, atomic/idempotent opening-stock staging endpoint. New items remain inactive and quarantined; it does not import orders, change Ecwid or activate inventory.
- A local, read-only workbook review with source-row references, source-quality checks, and explicit workbook/app rollout groups.
- Explicit supplier-backed mode, complete receipts, immutable assignment/release events, allocation-capped Paid-only picks and mixed-order progress. Supplier physical movements never create Ecwid stock writes.

## Phased rollout: one movement ledger per item

The existing workbook's **Balance means physical stock held in the factory**. It is not a promise that all of that stock is available for new orders. Existing unpicked commitments, including partial picks, must be reconciled separately before the initial Ecwid alignment. Awaiting Payment orders reserve stock but are not pickable until Paid.

During the rollout, workbook-managed and app-managed items coexist:

- Until approved migration, **all items remain workbook-managed**, including source-review candidates.
- After a specific item is migrated, the app owns every physical movement for it: website picking, email sales, internal use and receipts. Do not also record those same movements in the workbook's stock ledger or let a later sheet import overwrite the app balance.
- Unmapped items, meters, other unsupported units, bundles, variations without independent stock, and unresolved rows continue in the workbook. Independently stocked single-unit variations may be migrated after review. `NA` is an unmapped placeholder, not proof that Ecwid has no corresponding listing. Do not combine these rows under a common SKU or create Ecwid products automatically.
- If workbook-recorded manufacturing/BOM activity consumes an app-managed component, its physical consumption must also be recorded once in the **app's** movement ledger. Do not rely on a legacy workbook formula to change app inventory. Agree this staff procedure before migrating components used in manufacturing.
- Record an approved per-item migration list, source snapshot, start time, physical opening balance and existing order commitments at each cutover. Reconcile only that approved set; never push the whole workbook to Ecwid after partial rollout.

The current picker deliberately blocks orders containing unsupported/unmapped lines. A mixed-order reconciliation process remains a go-live requirement; source candidacy does not promise that mixed orders can already be split between the two systems.

### Supplier-backed workflow: implemented locally, not activated

The [supplier-backed inventory design](docs/supplier-backed-workflow.md) separates factory stock, assigned stock and supplier demand while preserving unlimited Ecwid selling policy. Migration `0005` and the app implement complete receipts (including exact-order purchases), explicit per-order allocation/release, Paid-only picks, mixed-order progress and cancellation safeguards. Real items still need a separately reviewed exact SKU/product/variation mapping and opening count; attribute identification is deferred. There is no production activation or supplier onboarding endpoint yet.

For these items, `free = on_hand − allocated`. Ordering creates demand, not shelf stock or an automatic assignment. Receive the entire delivery in **Stock movement → Restock**, then use **Pick orders → Assign to this order**. Awaiting Payment orders can hold assignments but cannot be picked until Paid. Supplier receipts, picks, email sales and internal use leave Ecwid unchanged; the latter two can consume only free stock. A receipt with free stock shows **awaiting assignment**, not falsely **awaiting supplier**. Aggregate uncovered demand is a shelf-coverage indicator, not an instruction to buy goods.

An unpicked cancellation releases assigned stock without changing shelf stock. A cancellation after any pick, or a terminal status with missing picks, holds the order for review. Picking never sets an Ecwid shipment status. Supplier identity, option or finite/unlimited policy drift remains blocked for review, not silently remapped.

Run the fictitious supplier rehearsal in its own database, without changing the original demo:

```sh
npm run setup:supplier-demo
npm run dev:supplier-demo
```

Open <http://127.0.0.1:8793>. This uses `fixtures/supplier-demo/wrangler.jsonc`, a separate dummy D1 ID and `.wrangler/supplier-demo-v2`, blank credentials, demo mode and disabled live sync. Do not substitute production credentials or data. `SUP-DEMO-PAID` contains two supplier lines plus one stocked line; `SUP-DEMO-AWAITING` tests the unpaid gate. Receive `SUP-DEMO-PART-A`, then assign its units to an exact order. All names/IDs/counts are fabricated. Setup is additive/idempotent and preserves earlier rehearsal movements; it does not reset them.

The user has confirmed one-to-one units for all 76 reviewed pilot matches, including names containing “Set”. The six identified supplier lines are awaiting MISUMI receipt; their incomplete workbook rows are not approved opening counts. This supersedes the earlier all-on-shelf assumption for those lines only. The original review is historical, not a renewed stock confirmation. Private business clarification and updated status files are under `import-data/pilot-review-20260922/verified/`. `python3 scripts/prepare-workflow-review.py --help` describes the hash-bound, report-only builder. It performs no import, activation, stock write or attribute lookup. The finite-stock exception for `TNP40-M5` still needs rechecking before activation.

## Review the existing stock workbook (step 1)

This is an **offline source review**, not a live Google Sheets connection. Drive-hosted `.xlsm` files can be downloaded unchanged. The local reader does not execute macros, convert or edit the workbook, recalculate formulas, call Ecwid, or write app inventory. It validates the established `Stock Sheet` row-3 layout and reads saved values, including leading-zero SKU strings. It checks the expected balance/movement formulas, their saved results, and saved Material Register/BOM totals. Those checks are not a fresh recalculation in Excel and do not prove the physical count; refresh/review the workbook before cutover.

Use Node 24 and Python 3.11+ (standard library only; no Python packages required):

```sh
mkdir -p import-data
npm run import:source -- /path/to/stock.xlsm --out import-data/review-YYYYMMDD --source-ref 'https://docs.google.com/spreadsheets/d/FILE_ID/edit'
```

The output directory must be new and its parent must already exist. `--python /path/to/python3` (or `INVENTORY_PYTHON`) selects a Python runtime. `--source-modified-at <ISO timestamp>` records the Drive version timestamp. The output contains:

- `review.html`: searchable offline review, defaulting to source exceptions.
- `source-snapshot.json` and `source-review.json`: original cell values/formulas, workbook hash and row references, plus review reasons.
- `candidates.json`: only source-eligible whole-unit rows, for later Ecwid matching. It intentionally says `ecwid_checked:false` and `reservations_confirmed:false`.

The three groups are **Candidate** (source checks only), **Keep in workbook** (outside this phase or unmapped), and **Needs review** (source issues). None means migrated or approved. A duplicate SKU/Store ID is checked against the full source, including excluded rows. Negative/fractional piece balances, manual stock formula overrides, invalid register quantities and missing formula results are held for review, never coerced to zero, rounded, combined or silently fixed. Other-unit fractional stock stays in the workbook. The reviewed source is never modified.

Keep these operational files in gitignored `import-data/`; never put them in `public/` or commit them. The CLI creates private local output files and refuses existing output directories or directories under deployable public assets.

## Read-only Ecwid connection (step 2)

Store ID: **2442119**. Keep the Worker in demo mode with both live flags disabled during this stage. The local connection check uses a separate `.env.ecwid-readonly` file; it does not start the Worker, import orders, write the database or change stock.

1. Open [Ecwid's developer dashboard](https://my.ecwid.com/#develop-apps) for this store and use a **dedicated inventory custom app**, not an app another integration depends on.
2. Under **Manage App → Access Scopes**, select only `read_catalog` and `read_orders` for this phase. Default app scopes may include writes. If scopes change, Ecwid requires reinstalling that app, which invalidates its old token. See [Ecwid's configuration guide](https://docs.ecwid.com/get-started/add-more-features-to-your-custom-app).
3. Put the **secret access token** from the app's Details page in the blank `ECWID_TOKEN=` entry in `.env.ecwid-readonly`. Do not paste it in chat, command arguments or frontend files. The file is gitignored; `.env.example` contains no secret. On a new checkout, copy `.env.example` to `.env.ecwid-readonly` and make it private with `chmod 600 .env.ecwid-readonly`.
4. Run `npm run ecwid:check`. This sends exactly two GET requests to Ecwid, checking products and orders with only `total,count,offset` returned. It uses no customer details, follows no redirects, performs no writes and makes no automatic retries. A successful check verifies read access, **not** that the token has no write permissions; inspect the app scopes too.

No webhook URL, client secret or OAuth callback is needed for this check. The secret token remains local. Reference: [Ecwid authentication](https://docs.ecwid.com/get-started/make-your-first-api-request), [product search](https://docs.ecwid.com/api-reference/rest-api/products/search-products), [order search](https://docs.ecwid.com/api-reference/rest-api/orders/search-orders).

The catalogue export is a preliminary match, not a current API snapshot or cutover approval. The adapter now verifies options, independently stocked variations and bundle relationships (`compositeParents` / `compositeComponents`). Bundle/component relationships, extra options, absent stock tracking and incomplete eligibility stay blocked. CSV exports cannot establish all these facts. Outstanding order commitments must also be reconciled separately.

Compare a native Ecwid catalogue CSV with the workbook candidates offline:

```sh
npm run import:catalogue -- /path/to/catalog.csv --candidates import-data/review-YYYYMMDD/candidates.json --store-id 2442119 --out import-data/catalogue-review-YYYYMMDD
```

Use a new output directory beneath gitignored `import-data/`. The report joins option/variation/file records by product ID, matches text SKUs without removing leading zeros, and flags ambiguous identities rather than choosing a first match. Independently tracked variants with whole quantities and exact option selections are preliminary matches, not excluded solely for being variations. `review.html` is searchable; `catalogue-review.json` retains both source hashes and workbook row references. The output is not an import payload. It always leaves `ready_count:0`; a CSV-only report cannot confirm live eligibility or reservations. Neither source file is changed.

### Fresh catalogue and open-order review

With the read-only credentials configured, collect every catalogue and order page:

```sh
npm run ecwid:snapshot -- --out import-data/ecwid-YYYYMMDD
npm run import:catalogue -- /path/to/catalog.csv --candidates import-data/review-YYYYMMDD/candidates.json --store-id 2442119 --live-catalog import-data/ecwid-YYYYMMDD/catalog.json --out import-data/catalogue-live-review-YYYYMMDD
```

Both output directories must be new. The snapshot command uses GET-only requests and writes private local review files, never the inventory database or Ecwid stock. `catalog.json` records the store, collection interval, parent products and independent `stock_targets`, including variation IDs and option selections. The catalogue review preserves the live snapshot hash alongside the CSV/workbook hashes and flags conflicts between the export and live identity. It still approves no items.

`orders-review.json` and `orders-review.html` show Paid/Awaiting Payment orders with nonterminal fulfillment states, plus ambiguous payment states as unpickable review exceptions. **Ordered quantity is not confirmed unpicked quantity.** Previously picked quantities are deliberately blank and must be reconciled against physical movements and stale order statuses. Awaiting Payment remains unpickable until Paid. Collection has a fixed order-creation cutoff but is not a transactionally frozen store snapshot; new orders and edits can occur while pages are read. Take a fresh, reviewed snapshot during the later approved cutover freeze.

## Opening-stock preview

### Approved pilot scope

The selected scope is the union of category IDs `40854040`, `40854051`, `12115348`, `48083853`, `40865024`, `159132005`, `159154502`, plus **Fasteners & Spacers** (`7135088`) and all its subcategories. Category selection is not approval of missing balances, unit conversions, stock-tracking changes or live stock writes. Count parent products separately from independently stocked variation SKUs.

`npm run ecwid:categories -- --out import-data/new-category-snapshot` reads category hierarchy and product membership only. It explicitly includes disabled categories, which the API otherwise omits, without enabling them. Keep this snapshot alongside the read-only catalogue and full workbook review. The offline `scripts/review-pilot-scope.py` report includes every selected stock target, including targets absent from the workbook; it never substitutes zero for an unknown quantity.

`python3 scripts/prepare-pilot-approval.py --help` describes the focused approval-list builder. Supply the verified pilot review and its exact CSV, full source review, catalogue, category scope, orders and physical-pick confirmation snapshots. It verifies input hashes and reproduces the pilot review before creating `approval.html` and `approval-list.json` in a new private directory. This is a read-only proposal, not an import format or approval form. Prior user clarification for the 32 matched nut/screw variations is display context tied to the exact reviewed snapshot; it never changes importer unit flags or approves opening balances. Sets/assemblies and stock shortfalls remain held. A changed snapshot does not automatically inherit that clarification. Run its regression checks with `python3 -m unittest discover -s tests -p 'test_*.py'`.

### Preview quantities

The authoritative stock sheet supplies **physical shelf quantities**. Open, unpicked Ecwid quantities must be reconciled separately, including any orders partly picked before this app starts. The preview sets the proposed Ecwid quantity to `Balance − unpicked reservations`.

```sh
npm run import:preview -- fixtures/opening-stock.csv fixtures/catalog.json fixtures/reservations.json
```

For real exports, put files in the gitignored `import-data/` directory. The CSV requires `SKU` and `Balance` headers; `Name`, `Scan_Code`, and `Location` are optional. A row can become READY only with `Single_Unit_Confirmed` set to exactly `true`, confirming that one Ecwid quantity unit equals one physical piece. JSON source rows use `single_unit_confirmed:true`. Do not infer this from the SKU, product name or the workbook's `Pcs` label. Leading-zero SKUs are preserved. Blank, negative, fractional, or invalid balances are blocked rather than converted to zero.

`catalog.json` can be the complete `READONLY_CATALOGUE` envelope produced above, or an array of reviewed stock targets for fixtures/manual review. Each target includes `id` (parent product ID), `combinationId` (variation ID or `null`), `sku`, `name`, `quantity`, `unlimited`, `enabled`, `hasOptions`, `hasVariations`, `variationOptions`, `hasBundleRelationships`, `hasExtraOptions`, and `eligibilityVerified`. A leaf variation has `hasOptions:true` and `hasVariations:false`; only the parent container has child variations. Missing eligibility fields are blocked. The envelope must be complete, its counts and UTC interval must be valid, and its store must match the reservation manifest's `store_id` when supplied. Include `store_id:"2442119"` in real reservation manifests to guard against mixing stores.

The reservation manifest identifies the source export, confirms physical balance semantics, and explicitly confirms one aggregated unpicked quantity per SKU (see the sample fixture). Snapshot order quantities are not automatically converted into this manifest.

The command prints JSON and exits with `0` for a valid preview, `2` for blocked rows/global issues, and `1` for invalid files. Reports include a source hash, proposed quantities, and differences. The authenticated admin endpoint `POST /api/import/preview` accepts the same fields (`csv` or parsed `rows`, plus `catalog` and manifest fields). Neither route changes a database or Ecwid stock.

After obtaining a complete read-only Ecwid catalogue snapshot and reconciling unpicked quantities, the same command accepts a reviewed copy of the source candidates without losing original workbook row references. Keep the original candidates untouched; add per-row unit confirmations only after checking the actual physical/store mapping:

```sh
npm run import:preview -- import-data/reviewed-candidates.json import-data/ecwid-YYYYMMDD/catalog.json import-data/reservations.json
```

The reservation manifest must explicitly confirm quantities and must not conflict with the candidates' source reference or catalogue store. The report retains `snapshot_source_hash` separately from its combined preview hash, and `catalogue_snapshot` records the live store and collection interval. The combined hash includes the original catalogue envelope and its timestamps. Each line preserves `ecwid_product_id`, `ecwid_combination_id` and canonical `ecwid_option_signature`. Real candidates must not be compared with the demo fixtures. Keep catalogue/order read-only access now; catalogue write access and live flags belong to the later approved cutover. A saved preview alone never authorizes stock updates.

### Inactive opening-stock staging

`POST /api/import/stage` is administrator-only and same-origin. It requires a configured store and both `INVENTORY_ENABLED=false` and `LIVE_SYNC_ENABLED=false`. Send `{operation_id, expected_hash, confirm_staging:true, input, scope}`. `input` is the original preview input, including a complete store-matched `READONLY_CATALOGUE`, workbook `snapshot_source_hash`, explicit unit confirmations and reservations. `expected_hash` must equal the freshly recomputed preview's `source_hash`; `scope` lists each exact `{sku,ecwid_product_id,ecwid_combination_id,ecwid_option_signature}` once. The bounded request size is 10 MB to accommodate the full variation catalogue.

Every selected row must be READY. At most 200 new stock targets are accepted per batch; all nonzero reservations are rejected until exact opening order-line import is implemented. This endpoint relies on the administrator's reviewed source/scope; it does not independently prove current order completeness or category membership. It cannot be used as cutover approval. It preserves existing rows and rolls back the entire batch on any conflicting SKU, scan code or target.

The batch stores immutable audit hashes, exact identities and opening balances in one D1 transaction. Identical operation retries return the existing result. Imported items remain `active=0` with an open staging issue, preventing even demo movements. No outbox entry, reservation, Ecwid request or activation is created. A `STAGED` result is **not aligned or ready for factory use**. Do not mix real staging rows into the demo database; use a separate clean database for an approved rehearsal/cutover. There is deliberately no generic activate or stock-overwrite endpoint.

## Live setup and cutover still required

The checked-in Wrangler configuration is **local-only** with a placeholder D1 ID, no public route, and live writes disabled. Do not deploy it unchanged.

1. Provision distinct staging/production D1 databases and queues, configure their IDs/names and the intended hostname, and apply migrations to an empty target database. Use a separate config/environment so demo data never enters production.
2. Protect the app with Cloudflare Access. Configure `ACCESS_TEAM_DOMAIN` (full `https://…cloudflareaccess.com` URL), `ACCESS_AUD`, and comma-separated `ADMIN_EMAILS`. Keep only `/api/webhooks/ecwid` outside browser sign-in; its Ecwid signature is verified by the Worker.
3. Set `ECWID_MODE=live`, `ECWID_STORE_ID`, and secrets `ECWID_TOKEN` and `ECWID_CLIENT_SECRET` through Wrangler's secret prompts. Use Ecwid scopes `read_catalog`, `update_catalog`, and `read_orders`. Keep `INVENTORY_ENABLED=false` and `LIVE_SYNC_ENABLED=false` during preparation.
4. Validate source stock and open orders, resolve duplicates and unsupported mappings, and review a fresh preview. Decide/record how already-picked quantities are carried across the cutover. Complete exact opening-order import and reviewed alignment/activation; the inactive staging endpoint alone is insufficient.
5. Briefly freeze checkout and all stock movements, disable the old Apps Script queue/triggers, and back up the sheet, Ecwid quantities/open orders, and target DB. Create verified items first, then opening balances and open-order/pick state. Do not enable polling/webhooks before item mapping is ready: unmapped order lines are intentionally quarantined.
6. Set Ecwid quantities once to the approved available quantities while activity is frozen, read them back, and resolve any mismatch. Complete the initial full order scan and confirm no unresolved supported-item issues. Enable order/product webhooks, scheduled processing, inventory recording, and live delta delivery only after this validation.
7. Pilot a small set of simple SKUs with the actual picker device, then resume normal activity. An old snapshot is not a safe rollback after new sales; use audited reconciliation instead.

Ecwid writes are deliberately independent of local database commits. A slow network can briefly leave sellable stock behind physical activity. There is no distributed transaction with checkout, so online sales can race an offline movement. Live use needs the pilot/reconciliation checks above; the app exposes pending work and exceptions instead of claiming every local save is already reflected in Ecwid.

## Sync exceptions

`PENDING → PROCESSING → APPLIED` is the normal path. Confirmed request rejection becomes `BLOCKED`; a timeout, server error, malformed confirmation, or stale processing claim becomes `UNKNOWN`. Unknown writes are never automatically resent. A documented Ecwid `429` rejection is retried after its cooldown. Item-level processing is serialized and blocked items do not starve unrelated products.

An administrator must reconcile unknown writes using the movement record, Ecwid order/stock history, and physical counts. A current quantity alone cannot prove whether a write occurred because another order may have changed it. Do not change `UNKNOWN` to `PENDING` unless non-application is established. This build displays exceptions but intentionally has no unreviewed "retry anyway" or arbitrary balance overwrite control. An audited resolution/physical-return interface remains a pre-production task.

The latest 100 movements/orders and 200 matching items are returned for interactive views. Inventory search is server-side; exact Order ID lookup refreshes/imports that order in live mode. Ecwid background polling is paginated independently and retains its cursor. Local scheduled events require manual triggering; Queues work in the local emulator.

## Development and verification

```sh
npm test
npm run test:extractor
npm run test:catalogue
npm run typecheck
npm run build
```

`build` bundles assets and performs a Wrangler **dry run**, not deployment. Tests use real SQLite transactions/triggers, mocked Ecwid HTTP responses, and real JWT cryptography. Browser checks exercise the app on local D1/Queues. `npm run types` regenerates configuration types; `.dev.vars.example` lists optional local secret keys, while real `.dev.vars` files are gitignored.

| Location | Purpose |
| --- | --- |
| `public/` | New picker UI and self-hosted scanner |
| `src/index.ts`, `src/auth.ts` | Worker routes, staff identity, request controls |
| `src/inventory.ts`, `src/supplier.ts`, `migrations/` | Stock ledger, reservations, supplier assignments and atomic rules |
| `src/ecwid.ts`, `src/sync.ts` | Ecwid adapter, webhook inbox, order sync, outbox |
| `src/opening-import.ts`, `scripts/preview-import.ts` | Opening-stock review report |
| `src/source-review.ts`, `scripts/extract_stock_workbook.py`, `scripts/review-stock-workbook.ts` | Read-only workbook screening and staged candidate export |
| `fixtures/`, `tests/` | Local sample data and regression tests |
| `Code.js`, root `index.html`/`style.css`, `InventoryOneTimeSync.gs` | Preserved legacy implementation |

Platform behavior was checked against official [Workers guidance](https://developers.cloudflare.com/workers/best-practices/workers-best-practices/), [D1 transactions](https://developers.cloudflare.com/d1/worker-api/d1-database/), [Access token verification](https://developers.cloudflare.com/cloudflare-one/access-controls/applications/http-apps/authorization-cookie/validating-json/), and [Ecwid inventory rules](https://support.ecwid.com/hc/en-us/articles/207099919-Managing-inventory).
