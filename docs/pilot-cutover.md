# Pilot cutover: operator runbook

## Current boundary

The app, exact opening-order importer and workbook/app mixed-order handling are implemented locally. This is not a record of a completed production cutover: no real pilot items have been staged or activated, no pilot stock-alignment write has changed Ecwid, and there is no operational mobile URL yet. Production Cloudflare Access is not enabled for the pilot account.

The one-time alignment workflow is a guarded service (`src/cutover-alignment.ts`), separate from the Worker routes. Its entry points are `beginCutoverAlignment`, `alignCutoverRow` and `finishAndActivateCutover`. There is no public alignment endpoint or operator CLI. A reviewed operator integration and production access/configuration are required before using it with real data. Do not replace that integration with hand-written SQL or ad hoc stock writes.

## Approved scope and stock meaning

- The 76 reviewed stock-limited pilot matches are single-unit targets, including independently stocked variations and names containing “Set”. Preserve exact SKU, product ID, variation ID and canonical option selection.
- The six supplier-backed MISUMI targets are excluded: `CB-40-2SLOT`, `LP-M6AS-20`, `SMB-2020-SET`, `SMB-2040-SET`, `TNP40-M5`, `TNP40-M6`. Their incomplete workbook records are not physical opening counts. Their separate receiving/allocation workflow must be onboarded later; do not change their Ecwid unlimited/finite policy in this cutover.
- Workbook **Balance** means physical factory stock. For each approved stock-limited target, the authorized Ecwid opening quantity is **Balance minus exact unpicked order commitments**, not raw Balance.
- Paid and Awaiting Payment orders both reserve stock. Only Paid orders in `AWAITING_PROCESSING` or `PROCESSING` are pickable. `READY_FOR_PICKUP` means already picked and packed, equivalent to Shipped for inventory handling, and cannot be picked again.
- Other targets remain workbook-managed. Registration of an outside-pilot target records identity only; it does not import its physical balance, adjust its Ecwid quantity or authorize app movements.

These are approval boundaries, not permission to reuse old data. Fresh identities, source balances and outstanding order lines must still pass every check below. A changed or newly discovered target does not inherit approval automatically.

## Prepare before the stock pause

1. Copy `wrangler.production.example.jsonc` to gitignored `wrangler.production.jsonc`, then provision and enter separate production Worker/D1/Queue resources and the intended hostname. The template keeps live mode configured but all three enabling flags false, with Access fields blank and a placeholder database ID. Apply migrations to the clean target database; never reuse the demo database or deploy either placeholder configuration unchanged.
2. Configure Cloudflare Access, `ACCESS_TEAM_DOMAIN` and `ACCESS_AUD`. Use the authorized staff email list and administrator subset from the private approval record, not committed documentation. Put every permitted account in `STAFF_EMAILS`, including administrators, and the administrator subset in `ADMIN_EMAILS`. Both lists use exact email addresses, not whole-domain admission. The Worker also verifies the Access JWT; a valid token outside `STAFF_EMAILS` is rejected.
3. Keep `INVENTORY_ENABLED=false`, `LIVE_SYNC_ENABLED=false` and `ORDER_SYNC_ENABLED=false`. All three are independent safeguards. Do not enable polling/webhook ingestion before opening mappings/orders are in place. Do not publish a demo-authenticated service as the factory app.
4. Prepare the exact approved pilot scope, workbook-managed identity registry and one-to-one unit confirmations. Review mixed orders ahead of time. Resolve unknown or contradictory lines; do not silently omit them.
5. Prepare a private backup location beneath gitignored `import-data/`, and backups of the authoritative workbook, Ecwid catalogue/open orders and target database. Keep tokens out of command arguments, source files, public assets, chat and Git. Do not disable unrelated workbook processes: identify and stop only integrations that could also write stock for the pilot.
6. Confirm that the operator integration for the alignment service and the authenticated mobile deployment are ready before requesting a short pause. If infrastructure or operator tooling is still missing, do not ask the factory to remain paused indefinitely.

### Order freshness is a launch requirement

Configure Ecwid order webhooks and verify signed delivery into the durable inbox before factory-floor use with the current polling code. Exempt only the exact `/api/webhooks/ecwid` endpoint from Cloudflare Access browser sign-in; the Worker still checks the Ecwid signature, store and event identity. Keep all staff app/API routes protected. Enable ingestion only at the approved cutover boundary with `ORDER_SYNC_ENABLED=true`.

The cron runs every five minutes but `pollOrders` reads only two pages of up to 100 orders per invocation. At the reviewed 8,298 historical orders, a full scan takes approximately 42 runs / 210 minutes, and its fixed creation cutoff can defer newer orders until the next scan. This fallback does **not** provide five-minute reservation freshness. Do not launch with historical polling as the sole automatic order feed. Either configure and test the webhooks or first implement and verify a recent-update polling path. Opening stock alignment and fresh-on-pick order lookup do not remove this requirement.

## Fresh evidence during the freeze

Pause new pilot sales and every pilot physical movement: picking, receipts, email sales and internal use. Record an explicit UTC `freeze.started_at`. The user must confirm that the workbook is current and that physical balances can be used. The importer never treats an unknown or blank balance as zero.

Collect new read-only catalogue and order snapshots while paused:

```sh
npm run ecwid:snapshot -- --out import-data/ecwid-FRESH-UNIQUE-NAME
```

The output directory must not already exist. This command reads via GET and writes only private local evidence:

- `catalog.json`: complete `READONLY_CATALOGUE`, exact stock targets and collection interval.
- `orders-review.json` / `orders-review.html`: human review, with prior picks deliberately unconfirmed.
- `opening-orders.json`: complete `READONLY_ORDERS`, normalized pending orders and full line identities, explicit digital evidence, creation cutoff, checked-order count, pending-order count and line count. Unsupported option contents are redacted, not silently accepted.

Refresh/download the authoritative workbook unchanged and run the source review. Bind its exact SHA-256 and source reference into the request. Confirm **zero previous physical picks for every exact opening order line**, including explicitly workbook-managed lines. This first importer does not support historical partial picks. If any line was already picked, stop and reconcile it; do not set its confirmation to zero or put the quantity back into a snapshot to make validation pass.

Both complete API snapshots must start at or after the confirmed freeze start. The freeze and snapshots have a **15-minute** freshness window. A request containing incomplete counts, ambiguous statuses, missing lines, stale evidence or future timestamps is rejected. Completion metadata is required even when there are no pending orders. If the window expires, obtain a fresh pause confirmation, snapshots and reviewed request; do not merely edit timestamps or hash fields.

If a batch has already been staged or any alignment row attempted, first reconcile that existing batch. A new UUID is not permission to duplicate its items or to resend its remote writes. The current service has no automatic freshness-reset or restaging shortcut for an existing cutover.

## Preview and stage, without live writes

Use authenticated, same-origin administrator requests:

- `POST /api/cutover/preview`
- `POST /api/cutover/stage`

Both require the configured store and all three flags explicitly `false`, and bound the request body to 10 MB. The canonical request type is `OpeningCutoverRequest` in `src/opening-cutover.ts`; the fields are:

- `operation_id`: a new UUID retained for exact retries.
- `confirm_staging:true`, `physical_counts_confirmed:true` and `freeze:{confirmed:true,started_at}`.
- `input`: original opening-stock preview data, complete catalogue, workbook source hash/reference, physical balances, unit confirmations and reservation totals.
- `scope`: every exact approved app target once.
- `orders`: the complete `opening-orders.json` envelope.
- `line_confirmations`: every `{order_id,ecwid_line_id,previously_picked_quantity:0}` exactly once.
- `workbook_scope`: explicitly reviewed outside-app targets, with exact SKU/product/variation/options and name.
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

This section describes the separately guarded operator service, not an exposed app control. There is no operational alignment CLI or public HTTP endpoint yet.

`beginCutoverAlignment` must be given the exact staged operation/review identity and current freeze evidence. Its preflight checks the complete live pending-order set and local staged state. Each row's subsequent preflight checks its current Ecwid stock identity/quantity before a write. A missing order is not presumed fulfilled; changed lines, payment/fulfillment status, stock policy, quantity or mapping require a new review. Local physical balances, reservations, quarantine state and absence of intervening movements must also agree.

`alignCutoverRow` handles one exact app-managed target at a time. The journal progresses from `PENDING` to `PROCESSING` **before** a stock write. A target already at its approved quantity can be verified as a read-only no-op. The remote write is confined to that exact product or variation; supplier and workbook-only targets are excluded. A returned success alone is insufficient: read back stock and retain verification evidence.

An uncertain or rejected row moves the whole batch to `REVIEW` and stops further alignment. **Never resend an already attempted `PROCESSING` or `UNKNOWN` write.** A timeout may occur after Ecwid applied it, and a current quantity by itself is not proof that it did or did not happen. Reconcile using the journal, Ecwid history, order changes and physical evidence. There is deliberately no blind “retry anyway” operation or `REVIEW`-to-`ALIGNING` reset.

`finishAndActivateCutover` performs fresh complete-order and stock readbacks again, checks that every row is verified and that no local movement or unresolved conflict invalidated the opening state, then completes the guarded local activation transaction. The state sequence is `STAGED → ALIGNING → ALIGNED → ACTIVE`; uncertainty leaves `REVIEW`, never a partial success advertised as active.

Activation must seed `orders_tracking_started` from the **original freeze timestamp**, not the activation time. That protects the gap between opening collection and the first poll: orders created and completed in that interval must still be reviewed. Activating a database batch does not toggle deployment flags or automatically publish a hostname.

## Publish, verify and resume

Only after verified activation and production authentication checks should the operator enable `ORDER_SYNC_ENABLED`, `INVENTORY_ENABLED` and `LIVE_SYNC_ENABLED` in the intended deployment. Verify signed webhook delivery and processing for new orders and payment/fulfillment changes; the slow historical scan alone is not a launch acceptance test. Verify a full order poll and inspect newly discovered issues before resuming pilot sales/movements. Verify approved staff can sign in, unauthorized identities cannot, and the production HTTPS link works from an actual factory phone for picking and adding stock. Use a controlled approved movement and confirm its ledger and applicable Ecwid outcome; a local demo is not that validation.

For app-managed targets, staff stop duplicate workbook stock movements at the cutover boundary. They continue workbook entries for targets outside the app. A mixed order labels those outside lines **Workbook-managed** and never offers an app pick/assignment for them. “All app-managed items picked” is not confirmation that the entire order is ready to ship; staff must check the workbook lines separately. Picking does not automatically update Ecwid shipment status.

Preserve the evidence, receipt, original freeze time, verified quantities, staff sign-in checks and final rollout decision. After activity resumes, an old workbook snapshot is not a safe rollback: new sales and physical movements need audited reconciliation rather than an overwrite.
