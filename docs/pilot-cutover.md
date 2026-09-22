# Pilot cutover: operator runbook

## Current boundary

On 2026-09-22, a login-test deployment was published with a separate production database and queue. Cloudflare Access protects its entire hostname with email-code login, the exact approved staff list and the user-approved one-month session. Anonymous requests to the page, static assets and API redirect to Access. The user has confirmed a successful approved-user login. Actual factory-phone camera, picking and receiving checks remain pending.

This is not a completed production cutover: no real pilot items have been staged or activated, no pilot stock-alignment write has changed Ecwid, and stock operations are not available for factory use. All three enabling flags remain false, no Ecwid credentials have been uploaded, and no cron is configured. Empty counters mean **not imported**, not zero factory stock. The production database contains eight verified schema migrations, zero items, orders, movements, outbox entries, cutover batches or workbook targets, and no foreign-key errors. Resource identifiers, deployment version, URL and verification receipts are retained privately under gitignored `import-data/`.

The local implementation now includes a private request assembler (`scripts/prepare-cutover-request.ts`) and operator CLI (`scripts/cutover-operator.ts`) around the guarded alignment service. There is no public alignment endpoint or arbitrary-SQL operator command. Recent-update polling and extended workbook-only profile policies are implemented and tested locally; migration `0008` is applied and verified in production, but the updated Worker code is not yet deployed. The schema upgrade did not import inventory or change Ecwid. Do not replace the guarded tools with hand-written SQL or ad hoc stock writes.

## Approved scope and stock meaning

- The 76 reviewed stock-limited pilot matches are single-unit targets, including independently stocked variations and names containing “Set”. Preserve exact SKU, product ID, variation ID and canonical option selection.
- The six supplier-backed MISUMI targets are excluded: `CB-40-2SLOT`, `LP-M6AS-20`, `SMB-2020-SET`, `SMB-2040-SET`, `TNP40-M5`, `TNP40-M6`. Their incomplete workbook records are not physical opening counts. Their separate receiving/allocation workflow must be onboarded later; do not change their Ecwid unlimited/finite policy in this cutover.
- Workbook **Balance** means physical factory stock. For each approved stock-limited target, the authorized Ecwid opening quantity is **Balance minus exact unpicked order commitments**, not raw Balance.
- Paid and Awaiting Payment orders both reserve stock. Only Paid orders in `AWAITING_PROCESSING` or `PROCESSING` are pickable. `READY_FOR_PICKUP` means already picked and packed, equivalent to Shipped for inventory handling, and cannot be picked again.
- Other targets remain workbook-managed, including the profile products specifically approved for this treatment. Registration records a reviewed identity/policy only; it does not import their physical balances, adjust their Ecwid quantities or authorize app movements. It must not turn arbitrary unknown options or every future profile into an automatic exception.

These are approval boundaries, not permission to reuse old data. Fresh identities, source balances and outstanding order lines must still pass every check below. A changed or newly discovered target does not inherit approval automatically.

### Explicit workbook-only option policies

The existing default is `sku_source:TARGET` with `option_policy:EXACT`; omitting these fields retains that strict behavior. For an explicitly approved profile variation, the independent choices are:

- `sku_source:TARGET` with `option_policy:STOCK_SELECTION_PLUS_OPAQUE_EXTRAS`: use the exact variation's own SKU while treating extra options as workbook-only information.
- `sku_source:PARENT_IF_VARIATION_BLANK` with `option_policy:STOCK_SELECTION_PLUS_OPAQUE_EXTRAS`: use the exact parent SKU only when the current catalogue proves that this variation's SKU is blank.

Both extended forms still require an exact product ID, variation ID and nonempty canonical stock-option signature. Repeated inherited display SKUs are allowed only across explicitly reviewed fallback variations of that same parent. App-item collisions, missing/contradictory stock selections and unknown variations remain blocked. A new order or profile must still match the reviewed registry; the current approval is not a blanket profile/category rule.

The version-2 `READONLY_ORDERS` snapshot carries only the stock-option subset plus `workbookOptionsEvidence:{kind:'OPAQUE_OPTIONS_V1',sha256}` for opaque extras. Its fingerprint includes actual extra CHOICE values, free text and file-option data; those raw extra values and URLs are not written into the snapshot. Version 1 cannot carry this evidence. Explicit scope approval is still required, and live parsing does not trust an incoming evidence field: hashes are recomputed from actual Ecwid options. A change in options, quantity or identity quarantines the order instead of silently remapping it.

Ecwid defines order-line `digital` as the presence of downloadable attachments; it is not sufficient to decide whether a reviewed item is physical factory stock. An exact, explicitly approved workbook-managed line may carry `digital:true` or `digital:false`. Preserve and hash that boolean rather than rewriting it; later drift still quarantines the order. App-managed lines continue to require `digital:false`, and no download URLs are imported. See [Ecwid's order field reference](https://docs.ecwid.com/api-reference/rest-api/orders/search-orders).

Migration `0008_workbook_opaque_options.sql` preserves existing registry IDs, references and review records while adding the narrow option policies. It has been rehearsed on populated local workerd D1 and applied to production through a guarded atomic file import after backup. Do not rerun it or the empty-database bootstrap. These policies do not relax app-item eligibility or create workbook-target stock, reservations, movements or Ecwid writes. User confirmations for current outside-app lines and zero previous picks belong in the private, hash-bound evidence, not in generic documentation or inferred confirmations for later orders.

## Prepare before the stock pause

1. Verify the existing isolated production Worker/D1/Queue and private `wrangler.production.jsonc`; do not recreate resources that are already provisioned. For a new environment only, copy `wrangler.production.example.jsonc`, replace placeholders and provision distinct resources. Never reuse the demo database or deploy a template unchanged. Finish tests, review incremental migrations and deploy the intended code while every live flag is still false. Record its actual deployed version for the operator.
2. Verify Cloudflare Access, `ACCESS_TEAM_DOMAIN` and `ACCESS_AUD` on the entire hostname. Use the authorized staff email list and administrator subset from the private approval record, not committed documentation. Put every permitted account in `STAFF_EMAILS`, including administrators, and the administrator subset in `ADMIN_EMAILS`. Both lists use exact addresses, not whole-domain admission. The Worker also verifies the Access JWT; a valid token outside `STAFF_EMAILS` is rejected. Test the actual factory phone before any stock pause.
3. Keep `INVENTORY_ENABLED=false`, `LIVE_SYNC_ENABLED=false` and `ORDER_SYNC_ENABLED=false`. All three are independent safeguards. Do not enable polling/webhook ingestion before opening mappings/orders are in place. Do not publish a demo-authenticated service as the factory app.
4. Prepare the exact approved pilot scope, workbook-managed identity registry and one-to-one unit confirmations. Review mixed orders ahead of time. Resolve unknown or contradictory lines; do not silently omit them.
5. Prepare a private backup location beneath gitignored `import-data/`, and backups of the authoritative workbook, Ecwid catalogue/open orders and target database. Keep tokens out of command arguments, source files, public assets, chat and Git. Do not disable unrelated workbook processes: identify and stop only integrations that could also write stock for the pilot.
6. Configure and read-only-test the operator below. Time complete catalogue/order collection, exact-target readbacks and deployment checks against current data before requesting a pause; then rehearse the full flow using isolated test data. There is no GET-only full-cutover rehearsal command: `stage`, `begin`, `rows` and `finish` are not read-only tests. The whole real flow must fit the 15-minute freeze with margin. Do not extend it with a clock override or ask the factory to stay paused while missing infrastructure is built.
7. Set the queue consumer's `max_batch_size` to **1** before enabling processing. Scheduled polling is deliberately budgeted together with five webhook and five stock queue hints; a ten-message consumer batch can exceed a Free Worker invocation's subrequest budget. Keep cron absent/disabled until activation; do not enable the template schedule accidentally during a login-only deployment.

### Empty-database schema bootstrap

The initial remote `wrangler d1 migrations apply` failed on the first migration with `incomplete input`; inspection confirmed rollback, an empty journal and no business schema. Do not remove or split the stock-safety triggers to work around this error.

For a **new, empty database only**, `scripts/prepare-d1-bootstrap.ts` creates a private local SQL artifact containing the complete original migrations byte-for-byte, in numeric order, with the standard Wrangler journal entries. It makes no database or network calls. Guards reject an existing business schema, nonempty journal or incompatible journal schema before importing business objects. The output directory must already exist with mode `0700`; output creation is exclusive with mode `0600`.

```sh
node --import tsx scripts/prepare-d1-bootstrap.ts \
  --migrations ./migrations \
  --output /absolute/private-directory/bootstrap.sql
```

Review the artifact hashes and exact target configuration, verify the remote database is empty, and test the complete artifact against an isolated local database. Then import the file through Wrangler's atomic SQL-file import path (`d1 execute DATABASE --remote --config CONFIG --file FILE`). The 2026-09-22 production bootstrap used this path successfully without changing any original migration. Do not rerun it on the initialized database or use it to upgrade an existing database.

After import, verify the migration journal, `migrations list`, schema objects, zero business rows and `PRAGMA foreign_key_check`. The initial seven-migration bootstrap had 16 business tables, 48 triggers, 2 views and 13 explicit indexes, excluding platform tables and the migration journal. Those counts describe the login-test baseline, not later schema changes. Preserve the artifact hash and import receipt. New migrations require a reviewed incremental upgrade and updated schema checks; do not rerun bootstrap on the initialized production database.

The subsequent production `0008` upgrade used a reviewed guarded atomic SQL-file import, with a pre-upgrade backup and an independently verified populated local runtime rehearsal. Post-upgrade checks confirmed **8 migrations, 16 business tables, 48 triggers, 2 views and 14 explicit indexes**, zero business rows and zero foreign-key errors. Its private import receipt is the migration record; it is not evidence of stock cutover or deployment of the newer Worker code.

### Order freshness is a launch requirement

The new `pollOrders` uses a fixed update-time window, a five-second settling interval and a two-minute overlap. It includes recently changed old orders regardless of payment/fulfillment status. Each invocation applies at most three orders; a complete second pass must match IDs and status/content signatures before the watermark advances. A database lease serializes pollers and fences stale progress writes. Changed pages restart the same window; failures retain progress, and more than 1,000 orders in a window is an explicit review condition.

Deploy and verify this recent feed before enabling its schedule. Verify creation, an old Awaiting Payment order becoming Paid, cancellation and fulfilled-without-picks review, and inspect the recent-window status/watermark. The configured cron interval is not a promise of that same freshness: large windows span invocations. The UI must expose pending, stale or failed synchronization. Exact order lookup is not an automatic feed for every reservation. Historical reconciliation is a slow fallback, limited to two orders after a successful recent window that applied none; it is not the launch freshness test.

Signed Ecwid webhooks can supplement polling. If configured, add `ECWID_CLIENT_SECRET` securely and exempt only the exact `/api/webhooks/ecwid` endpoint from Access browser sign-in; the Worker must still verify Ecwid's signature, store and event identity. Keep all staff app/API routes protected. Test actual signed delivery and failure recovery before relying on webhooks. The current login-only deployment has no webhook secret/bypass or running cron; enable order ingestion only at the approved post-activation boundary with `ORDER_SYNC_ENABLED=true`.

## Private operator configuration

Keep operator files beneath gitignored `import-data/`, not under `public/`. Use a dedicated directory with mode `0700`, regular configuration/request files with mode `0600`, and no sibling `.env*` or `.dev.vars*` files: Wrangler can load those implicitly. Store the credential file elsewhere and reference it by path. The operator refuses broad-readable files, final-path symlinks, unknown configuration fields and unintended bindings. Do not commit accounts, database/version IDs, staff emails, tokens or operational snapshots.

Example `import-data/operator/operator.json` (replace all uppercase placeholders; the example email is not an approval):

```json
{
  "schema_version": 1,
  "account_id": "ACCOUNT_ID",
  "worker_name": "factory-inventory",
  "database_id": "DATABASE_UUID",
  "database_name": "inventory-production",
  "store_id": "STORE_ID",
  "actor": "admin@example.com",
  "expected_version_id": "DEPLOYED_VERSION_UUID",
  "credentials_file": "../../.env.ecwid-readonly",
  "wrangler_config": "wrangler.operator.json"
}
```

The credential file contains only the intended `ECWID_STORE_ID` and secret `ECWID_TOKEN`. The token needs verified catalogue-write permission for alignment; a successful read check alone does not establish that. The separate `wrangler.operator.json` must contain exactly these fields, with identifiers matching the operator configuration:

```json
{
  "name": "factory-inventory-operator",
  "account_id": "ACCOUNT_ID",
  "compatibility_date": "2026-09-22",
  "compatibility_flags": ["nodejs_compat"],
  "workers_dev": false,
  "preview_urls": false,
  "d1_databases": [{
    "binding": "DB",
    "database_name": "inventory-production",
    "database_id": "DATABASE_UUID",
    "remote": true
  }]
}
```

Its name must be the production Worker name followed by `-operator`. This is a narrowly scoped remote D1 proxy configuration, not a deployable app configuration: no `main`, assets, queues, cron, routes, variables or extra databases. The operator uses official Wrangler authentication, captures the bearer in memory and disables credential logging; never run a token-printing command into a terminal transcript yourself. It creates no public alignment service.

Before mutation, the operator verifies the actual Cloudflare account/database, current single-version deployment, store, Access settings, administrator allowlists and three false live flags. Immutable version metadata is checked once per session; current deployment and the real freeze clock are checked again immediately before writes. A different deployment stops mutation rather than silently changing the pinned version.

For a read-only status check, create private `status.json` containing only the intended `operation_id`, then run:

```sh
npm run cutover:operate -- status \
  --config import-data/operator/operator.json \
  --request import-data/operator/status.json
```

`status` reads the fixed journal query without an Ecwid request or database proxy. It can report version/flag drift after verifying the current deployment's exact store/database/admin authority. `safe_for_cutover:false` and `resubmission_authorized:false` are not permissions to retry. An unknown or held row still requires reconciliation; wrong identity/authority remains an error.

## Fresh evidence during the freeze

Pause new pilot sales and every pilot physical movement: picking, receipts, email sales and internal use. Record an explicit UTC `freeze.started_at`. The user must confirm that the workbook is current and that physical balances can be used. The importer never treats an unknown or blank balance as zero.

Collect new read-only catalogue and order snapshots while paused:

```sh
npm run ecwid:snapshot -- --out import-data/ecwid-FRESH-UNIQUE-NAME
```

The output directory must not already exist. This command reads via GET and writes only private local evidence:

- `catalog.json`: complete `READONLY_CATALOGUE`, exact stock targets and collection interval.
- `orders-review.json` / `orders-review.html`: human review, with prior picks deliberately unconfirmed.
- `opening-orders.json`: complete version-2 `READONLY_ORDERS`, normalized pending orders and full line identities, explicit digital evidence, creation cutoff, checked-order count, pending-order count and line count. Reviewed workbook-only extra options use the canonical stock subset plus opaque fingerprint described above; unmatched or contradictory selections remain blocked. Snapshot creation does not approve a policy.

Refresh/download the authoritative workbook unchanged and run the source review. Bind its exact SHA-256 and source reference into the request. Confirm **zero previous physical picks for every exact opening order line**, including explicitly workbook-managed lines. This first importer does not support historical partial picks. If any line was already picked, stop and reconcile it; do not set its confirmation to zero or put the quantity back into a snapshot to make validation pass.

Both complete API snapshots must start at or after the confirmed freeze start. The freeze and snapshots have a **15-minute** freshness window. A request containing incomplete counts, ambiguous statuses, missing lines, stale evidence or future timestamps is rejected. Completion metadata is required even when there are no pending orders. If the window expires, obtain a fresh pause confirmation, snapshots and reviewed request; do not merely edit timestamps or hash fields.

If a batch has already been staged or any alignment row attempted, first reconcile that existing batch. A new UUID is not permission to duplicate its items or to resend its remote writes. The current service has no automatic freshness-reset or restaging shortcut for an existing cutover.

## Preview and stage, without live writes

### Assemble a private, hash-bound request

The approved scope file is `APPROVED_PILOT_SCOPE`, `schema_version:1`, with the intended `store_id`, a `review_reference` and an exact `scope` array. Each pilot identity is `{sku,ecwid_product_id,ecwid_combination_id,ecwid_option_signature}`. It is a reviewed allowlist, not a new category query or permission to add unmatched rows.

Prepare the questions before execution, using account-neutral paths below and substituting the exact authorized administrator:

```sh
npm run cutover:prepare -- \
  --stock import-data/source-FRESH/candidates.json \
  --catalog import-data/ecwid-FRESH/catalog.json \
  --orders import-data/ecwid-FRESH/opening-orders.json \
  --scope import-data/approved-pilot-scope.json \
  --actor admin@example.com \
  --out import-data/cutover-confirmations-NEW \
  --draft
```

The draft has `false`/`null` confirmations and is intentionally non-executable. Review its instructions, every pilot unit, every exact zero-prior-pick line and each outside-app workbook identity/policy. Supply the current physical-count/staging approvals and actual freeze time; do not bulk-convert unanswered questions into approvals. Artifact digests bind confirmations to the exact workbook candidates, catalogue, orders and scope. New evidence requires renewed confirmation, not copied hashes.

Assemble only after the current evidence has been reviewed:

```sh
npm run cutover:prepare -- \
  --stock import-data/source-FRESH/candidates.json \
  --catalog import-data/ecwid-FRESH/catalog.json \
  --orders import-data/ecwid-FRESH/opening-orders.json \
  --scope import-data/approved-pilot-scope.json \
  --actor admin@example.com \
  --out import-data/cutover-prepared-NEW \
  --confirmations import-data/reviewed-confirmations.json
```

Outputs are `cutover-request.json`, `cutover-review.json` and `artifact-manifest.json`. The assembler checks the authoritative preview against the real clock and inserts its `review_hash` as `expected_hash`; it does not connect to a database, stage, activate or contact Ecwid. Keep the generated operation UUID and original files for exact retries. Optional `--operation-id` must identify the same intended operation, never a workaround for a conflict. Output directories must be new and private; existing evidence is never overwritten.

Use the operator for the final authority-checked preview:

```sh
npm run cutover:operate -- preview \
  --config import-data/operator/operator.json \
  --request import-data/cutover-prepared-NEW/cutover-request.json
```

After reviewing a successful receipt and only while the same real freeze remains valid, run `stage` with those same paths. Unlike `preview`, `stage` **writes the inactive opening data to production D1**, but it does not write Ecwid stock.

### Staging contract

The equivalent authenticated, same-origin administrator routes remain:

- `POST /api/cutover/preview`
- `POST /api/cutover/stage`

Both require the configured store and all three flags explicitly `false`, and bound the request body to 10 MB. The canonical request type is `OpeningCutoverRequest` in `src/opening-cutover.ts`; the fields are:

- `operation_id`: a new UUID retained for exact retries.
- `confirm_staging:true`, `physical_counts_confirmed:true` and `freeze:{confirmed:true,started_at}`.
- `input`: original opening-stock preview data, complete catalogue, workbook source hash/reference, physical balances, unit confirmations and reservation totals.
- `scope`: every exact approved app target once.
- `orders`: the complete `opening-orders.json` envelope.
- `line_confirmations`: every `{order_id,ecwid_line_id,previously_picked_quantity:0}` exactly once.
- `workbook_scope`: explicitly reviewed outside-app targets, with exact SKU/product/variation/stock-option signature and name, plus explicit `sku_source`/`option_policy` when using the narrow profile policy above. Do not attach opaque evidence to an app-managed target or infer workbook classification from the evidence alone.
- `expected_hash`: use the preview's returned `review_hash` when staging. This is not the legacy stock-only preview hash.

The preview recomputes eligibility and totals. Every app reservation must equal the sum of the exact confirmed unpicked lines. Every workbook-managed target must match an unambiguous exact identity in the complete catalogue; a caller cannot invent an exclusion to bypass an unknown line. Workbook unit eligibility is not converted into app unit eligibility.

Staging atomically inserts up to 200 new app targets, 200 pending orders and 2,000 order lines, opening ledgers, explicit workbook identities and immutable cutover audit records. A conflict anywhere rolls the entire transaction back. Existing items/orders are not overwritten. UUID retries must match the original reviewed content and administrator.

After staging:

- App items are inactive and carry `OPENING_CUTOVER_STAGED` issues.
- Physical stock and exact unpicked commitments are loaded, but no stock can be picked or received through the live app.
- `opening_cutover_rows` records physical, unpicked, target quantity, prior Ecwid quantity and exact identity.
- No Ecwid request, stock outbox entry or activation is performed.

The older `POST /api/import/stage` remains available for stock-only staging but still rejects **all nonzero reservations**. Do not use it as a substitute for opening-order import or as an activation shortcut.

## Guarded alignment and activation

This workflow uses the private CLI and existing guarded service, not an app control or public HTTP endpoint. Create a private alignment request containing exactly `{operation_id,expected_hash,freeze}` copied from the successfully staged request; `freeze` retains `{confirmed:true,started_at}`. Do not send the full staging request to alignment commands or replace its timestamps.

Run each step separately and inspect its receipt before proceeding:

```sh
npm run cutover:operate -- begin \
  --config import-data/operator/operator.json \
  --request import-data/operator/alignment.json
```

`beginCutoverAlignment` must be given the exact staged operation/review identity and current freeze evidence. Its preflight checks the complete live pending-order set and local staged state. Each row's subsequent preflight checks its current Ecwid stock identity/quantity before a write. A missing order is not presumed fulfilled; changed lines, payment/fulfillment status, stock policy, quantity or mapping require a new review. Local physical balances, reservations, quarantine state and absence of intervening movements must also agree.

`alignCutoverRow` handles one exact app-managed target at a time. The journal progresses from `PENDING` to `PROCESSING` **before** a stock write. A target already at its approved quantity can be verified as a read-only no-op. The remote write is confined to that exact product or variation; supplier and workbook-only targets are excluded. A returned success alone is insufficient: read back stock and retain verification evidence.

After `begin` returns `ALIGNING`, explicitly process the reviewed batch:

```sh
npm run cutover:operate -- rows \
  --config import-data/operator/operator.json \
  --request import-data/operator/alignment.json
```

`rows` opens one remote D1 session and discovers only the 1–200 durable rows belonging to the exact staged store/actor/hash/freeze. It accepts no arbitrary item list, skips `VERIFIED` rows, rejects any held row before dispatch and processes remaining rows sequentially. It emits sanitized per-row receipts and stops at the first error or batch review state. It never automatically retries a write or activates the batch. `row` with `--item-id UUID` is the single-row equivalent; it does not relax journal or scope checks. Do not launch parallel operators.

An uncertain or rejected row moves the whole batch to `REVIEW` and stops further alignment. **Never resend an already attempted `PROCESSING` or `UNKNOWN` write.** A timeout may occur after Ecwid applied it, and a current quantity by itself is not proof that it did or did not happen. Reconcile using the journal, Ecwid history, order changes and physical evidence. There is deliberately no blind “retry anyway” operation or `REVIEW`-to-`ALIGNING` reset.

`finishAndActivateCutover` performs fresh complete-order and stock readbacks again, checks that every row is verified and that no local movement or unresolved conflict invalidated the opening state, then completes the guarded local activation transaction. The state sequence is `STAGED → ALIGNING → ALIGNED → ACTIVE`; uncertainty leaves `REVIEW`, never a partial success advertised as active.

Only after every row has a verified receipt, and while the original freeze is still valid:

```sh
npm run cutover:operate -- finish \
  --config import-data/operator/operator.json \
  --request import-data/operator/alignment.json
```

The bounded session reduces repeated authentication/metadata work, but it does not extend the freeze or remove complete-order/readback checks. Read-only timing and a full isolated rehearsal are prerequisites, not reasons to use stale evidence. If a process dies, inspect `status` and durable row evidence before deciding whether untouched pending rows can continue; never interpret a rerunnable command as permission to repeat an uncertain PUT.

Activation must seed `orders_tracking_started` from the **original freeze timestamp**, not the activation time. That protects the gap between opening collection and the first poll: orders created and completed in that interval must still be reviewed. Activating a database batch does not toggle deployment flags or automatically publish a hostname.

## Publish, verify and resume

Only after verified activation and production authentication checks should the operator enable `ORDER_SYNC_ENABLED`, `INVENTORY_ENABLED` and `LIVE_SYNC_ENABLED` in the intended deployment. Configure the reviewed cron schedule, queue batch size of one and necessary Worker secrets through secure Wrangler prompts. A token used by the local operator is not automatically a deployed Worker secret. Verify a completed recent-update window and actual creation/payment/cancellation behavior; test signed webhook delivery too if configured. Inspect newly discovered issues before resuming pilot sales/movements. The slow historical scan alone is not a launch acceptance test.

Verify approved staff can sign in, unauthorized identities cannot, and the HTTPS link works on an actual factory phone. Test camera permission/QR scanning and manual lookup, an approved Paid-order pick, a receipt, and an offline/internal movement as appropriate; verify the physical ledger and applicable Ecwid result. A Paid-order pick must not deduct online stock a second time. Keep Awaiting Payment and terminal orders unpickable. A successful login or local demo does not satisfy these tests, and none has been claimed as a completed production stock test.

For app-managed targets, staff stop duplicate workbook stock movements at the cutover boundary. They continue workbook entries for targets outside the app. A mixed order labels those outside lines **Workbook-managed** and never offers an app pick/assignment for them. “All app-managed items picked” is not confirmation that the entire order is ready to ship; staff must check the workbook lines separately. Picking does not automatically update Ecwid shipment status.

Preserve the evidence, receipt, original freeze time, verified quantities, staff sign-in checks and final rollout decision. After activity resumes, an old workbook snapshot is not a safe rollback: new sales and physical movements need audited reconciliation rather than an overwrite.
