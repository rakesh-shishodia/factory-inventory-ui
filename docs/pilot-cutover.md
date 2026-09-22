# Pilot cutover: operator runbook

## Current boundary

On 2026-09-22, the reviewed pilot cutover became ACTIVE in the existing protected Worker with its separate production database and queue. The reduced worker workflow is now deployed but stock recording remains paused: `INVENTORY_ENABLED=false`, `LIVE_SYNC_ENABLED=false` and `ORDER_SYNC_ENABLED=false`. The configured five-minute schedule only recovers already committed stock-outbox work; it never polls orders. Cloudflare Access remains on the entire hostname with email-code login, the exact approved staff list and the user-approved one-month session. Actual factory-phone camera, picking and receiving checks remain pending.

All 76 approved stock-limited targets remain aligned to workbook Balance minus exact unpicked commitments and active in D1; pausing the Worker did not roll back their balances or opening evidence. Seven opening orders and 38 lines remain loaded, while the six MISUMI targets and other registered exclusions remain workbook-managed. The guarded recovery skipped 12 already verified targets, aligned the remaining 64 untouched targets and activated only after a fresh full catalogue/order check. At activation the database had eight verified migrations, zero opening movements, zero pending outbox work and no open synchronization issue.

The reduced app presents one worker screen: scan or enter SKU, fetch current stock and location, choose quantity and reason, optionally enter notes, enter an exact Order ID only for Store/Ecwid picks, and submit. It does not run scheduled or background order polling. The private request assembler and operator CLI remain for audit/recovery history; there is still no public alignment endpoint or arbitrary-SQL operator command. Do not restage or realign the completed batch to deploy this UI.

The activation contains 76 verified active pilot targets, seven opening orders and 38 lines. The recovery's fresh preflight matched the full reviewed order set and all exact stock identities before any remaining write. A private backup was captured and `0009_simple_movement_safety.sql` was applied as the ninth migration; post-upgrade checks found both safety triggers, zero movements/outbox/issues and zero foreign-key errors. Actual factory-phone pick and restock tests remain required.

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

1. Verify the existing isolated production Worker/D1/Queue, deployed code and private `wrangler.production.jsonc`; do not recreate resources that are already provisioned. For a new environment only, copy `wrangler.production.example.jsonc`, replace placeholders and provision distinct resources. Never reuse the demo database or deploy a template unchanged. If code changes again, finish tests, review any incremental migrations and deploy while every live flag is still false. Record the actual current deployed version for the operator.
2. Verify Cloudflare Access, `ACCESS_TEAM_DOMAIN` and `ACCESS_AUD` on the entire hostname. Use the authorized staff email list and administrator subset from the private approval record, not committed documentation. Put every permitted account in `STAFF_EMAILS`, including administrators, and the administrator subset in `ADMIN_EMAILS`. Both lists use exact addresses, not whole-domain admission. The Worker also verifies the Access JWT; a valid token outside `STAFF_EMAILS` is rejected. Test the actual factory phone before any stock pause.
3. Keep `INVENTORY_ENABLED=false`, `LIVE_SYNC_ENABLED=false` and `ORDER_SYNC_ENABLED=false`. All three are independent safeguards. Do not enable polling/webhook ingestion before opening mappings/orders are in place. Do not publish a demo-authenticated service as the factory app.
4. Prepare the exact approved pilot scope, workbook-managed identity registry and one-to-one unit confirmations. Review mixed orders ahead of time. Resolve unknown or contradictory lines; do not silently omit them.
5. Prepare a private backup location beneath gitignored `import-data/`, and backups of the authoritative workbook, Ecwid catalogue/open orders and target database. Keep tokens out of command arguments, source files, public assets, chat and Git. Do not disable unrelated workbook processes: identify and stop only integrations that could also write stock for the pilot.
6. Configure and read-only-test the operator below. Time complete catalogue/order collection, exact-target readbacks and deployment checks against current data before requesting a pause; then rehearse the full flow using isolated test data. There is no GET-only full-cutover rehearsal command: `stage`, `begin`, `rows` and `finish` are not read-only tests. The whole real flow must fit the 15-minute freeze with margin. Do not extend it with a clock override or ask the factory to stay paused while missing infrastructure is built.
7. Preserve and verify the deployed queue consumer's `max_batch_size` of **1** before enabling processing. Keep the old order-polling schedule absent during preparation. After the reduced outbox-only scheduled handler is deployed, restore its five-minute recovery trigger; it must never poll or import orders. Current production has this recovery trigger configured.

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

The subsequent production `0008` upgrade used a reviewed guarded atomic SQL-file import, with a pre-upgrade backup and an independently verified populated local runtime rehearsal. Post-upgrade checks confirmed **8 migrations, 16 business tables, 48 triggers, 2 views and 14 explicit indexes**, zero business rows and zero foreign-key errors. Its private import receipt records the schema upgrade; it does not itself prove Worker configuration or stock cutover.

Migration `0009_simple_movement_safety.sql` stores the worker's adjustment reason, requires notes for Adjust Up/Down and prevents a second stock-limited non-pick movement while the same item has a `PENDING` or `PROCESSING` Ecwid delta. It was applied through a state-guarded atomic SQL-file import after a private production backup because the ordinary Wrangler migration path again rejected trigger SQL before changing schema. Post-upgrade checks confirmed nine migrations, both triggers, 76 unchanged active items and no foreign-key error. Do not rerun the bootstrap or any applied migration.

### Exact-order freshness in the reduced pilot

For a Store/Ecwid pick, the worker enters the exact Order ID. With `ORDER_SYNC_ENABLED=true`, the server refreshes only that order from Ecwid immediately before the pick and then enforces the existing exact line identity, Paid status, eligible fulfillment status, remaining quantity, reservation and shelf-stock guards. Awaiting Payment, `READY_FOR_PICKUP` and other terminal orders remain unpickable. While the deployment is paused, explicit refresh refuses to run.

The scheduled handler does not invoke `pollOrders`. Production runs it every five minutes only to retry committed stock-delta outbox work if the initial Queue hint was missed. No background feed discovers new orders, payment changes or cancellations; staff discover them by entering the relevant Order ID. The older polling/webhook implementation remains preserved for possible separately reviewed future work, but it is not part of this reduced pilot and no webhook secret or Access bypass is configured.

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

### Recovering an expired, partially verified batch

Use `recover` only when the original batch remains `ALIGNING`, at least one row is `VERIFIED`, every other row is still untouched `PENDING`, and there are no `PROCESSING`, `UNKNOWN` or `BLOCKED` rows. It does not restage, change the original freeze, or provide an uncertainty override. Obtain a new explicit confirmation that **store-wide Ecwid ordering and every pilot physical movement are paused**. Create a new private request containing exactly:

```json
{
  "operation_id": "THE-EXISTING-OPERATION-UUID",
  "expected_hash": "THE-EXISTING-REVIEW-HASH",
  "recovery_id": "A-NEW-RECOVERY-UUID",
  "recovery_freeze": {
    "confirmed": true,
    "started_at": "CURRENT-UTC-PAUSE-START"
  }
}
```

Run one trusted session; do not split it into ad hoc row commands:

```sh
npm run cutover:operate -- recover \
  --config import-data/operator/operator.json \
  --request import-data/operator/recovery.json
```

Before the first PUT, the command verifies the pinned deployment and disabled flags, every live target identity/policy/quantity, and the complete live order set. A VERIFIED row must still equal its target; a PENDING row must still equal its original expected Ecwid quantity. It emits a hash-bound preflight receipt, skips all VERIFIED rows without replay, and processes only PENDING rows through the existing durable `PROCESSING → VERIFIED/UNKNOWN/BLOCKED` journal. It stops at the first failure and never retries an uncertain write.

The recovery lease is 30 minutes from the exact newly confirmed pause. The command keeps at least five minutes for the final 76-target and complete-order readback, checks the live deployment immediately before each changed-row PUT and before activation, and leaves the batch `ALIGNING` on a pure pre-write/activation lease expiry so another explicitly approved pause can revalidate it. Final activation still writes `orders_tracking_started` from the **original** batch freeze. Retain the private request, preflight evidence hash, row receipts and final receipt together.

Activation seeded `orders_tracking_started` from the **original freeze timestamp**, not the activation time, preserving the audit boundary for later reconciliation. The reduced pilot does not advance that boundary through scheduled polling. Activating a database batch does not toggle deployment flags or automatically publish a hostname.

## Publish, verify and resume

The opening activation and alignment are complete; do not repeat them for the reduced UI. The private backup, guarded migration `0009` import and paused reduced deployment are complete. Keep all three live flags false while inspecting the build; retain the five-minute outbox-only recovery schedule, queue batch size of one, Cloudflare Access and its exact allowlists. Recheck the 76 active item balances, opening orders, unresolved issues, outbox and Ecwid token before the supervised smoke window. Any secret change must use secure Wrangler prompts/private stdin, never arguments or logs.

For a supervised smoke window, enable `ORDER_SYNC_ENABLED`, `INVENTORY_ENABLED` and `LIVE_SYNC_ENABLED` only after those checks. `ORDER_SYNC_ENABLED` authorizes exact staff-requested order refresh; it does not start background polling. Keep the configured cron limited to stock-outbox recovery. A mismatch, uncertain Ecwid write or unexpected database state means pause again and reconcile; never overwrite stock from the old workbook or rerun cutover to repair runtime activity.

Verify approved staff can sign in, unauthorized identities cannot, and the HTTPS link works on an actual factory phone. Test camera permission/QR scanning and manual lookup, an approved Paid-order pick, a receipt, and an offline/internal movement as appropriate; verify the physical ledger and applicable Ecwid result. A Paid-order pick must not deduct online stock a second time. Keep Awaiting Payment and terminal orders unpickable. A successful login or local demo does not satisfy these tests, and none has been claimed as a completed production stock test.

For app-managed targets, staff stop duplicate workbook stock movements at the cutover boundary. They continue workbook entries for targets outside the app. A mixed order labels those outside lines **Workbook-managed** and never offers an app pick/assignment for them. “All app-managed items picked” is not confirmation that the entire order is ready to ship; staff must check the workbook lines separately. Picking does not automatically update Ecwid shipment status.

Preserve the evidence, receipt, original freeze time, verified quantities, staff sign-in checks and final rollout decision. After activity resumes, an old workbook snapshot is not a safe rollback: new sales and physical movements need audited reconciliation rather than an overwrite.
