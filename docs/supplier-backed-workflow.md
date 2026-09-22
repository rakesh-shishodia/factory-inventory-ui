# Supplier-backed inventory and mixed-order picking

Status: core workflow implemented and locally verified, 22 September 2026. No real inventory has been imported or activated, no supplier order has been placed, and Ecwid remains unchanged. Exact real-item onboarding, opening-order import and approved cutover remain pending.

## Confirmed business rules

- All 76 reviewed pilot SKUs have one-to-one units: one workbook unit is one Ecwid sale unit. Within the selected pilot categories, “Set” in a name means one complete sale unit, not a multiplier. Existing identity, options, eligibility and balance checks still apply.
- MISUMI goods always arrive at the factory before dispatch. Staff sometimes order the exact customer requirement and sometimes buy extra.
- The six supplier-sourced SKUs already identified in the review have pending order goods awaiting receipt. The earlier blanket “all pending items are on shelves” statement is superseded for those exact lines. It is not a statement that every supplier SKU has zero physical stock.
- Supplier workbook entries are incomplete. They cannot establish a trustworthy opening balance, even where the saved value is zero or positive.
- Identify supplier-backed items with an explicit, administrator-reviewed SKU/product/variation mapping. Attribute-based identification is deferred: no attribute name, reader, tag writer or auto-classification is introduced now.
- Unit confirmation is not opening-balance approval, inventory activation or authority for a live stock update.

The private clarification record at `import-data/pilot-review-20260922/verified/business-clarification.json` and status review under `import-data/pilot-review-20260922/verified/workflow/` bind these statements to the original review, catalogue and order hashes. The original files remain historical evidence, not current all-on-shelf confirmation.

## Two inventory modes

| Rule | Stock-limited item | Supplier-backed / unlimited-online item |
| --- | --- | --- |
| Physical stock | App ledger after approved cutover | App ledger after separately verified opening count |
| Ecwid policy | Finite, tracked stock | Unlimited; do not send shelf quantities |
| New order | Existing outstanding-reservation logic | Creates demand; allocation is a separate fact |
| Receipt | Physical increase and existing Ecwid delta | Physical increase only |
| Ecwid order pick | Physical deduction, no second Ecwid deduction | Physical deduction, no Ecwid stock write |
| Email sale / internal use | Deduct only uncommitted stock, with Ecwid delta | Deduct only free physical stock, no Ecwid stock write |
| Supply shortage | Existing review rules | Normal line-level “Awaiting supplier”, not negative shelf stock |

Unlimited is a selling policy, not a claim that stock exists at the factory. Ecwid documents it as always available for purchase. [Ecwid inventory guide](https://support.ecwid.com/hc/en-us/articles/4402494390162-Intro-to-product-inventory)

The saved snapshot has one policy mismatch: `TNP40-M5` was tracked with quantity 4,797, whereas the other five identified supplier targets were unlimited. Preserve that observation. Recheck the exact variation before activation and obtain a separate decision for any setting change; do not automatically switch it or write 4,797 as factory stock.

## Quantities and invariants

For an active supplier-backed item:

- `on_hand`: physically on the shelf, established by an approved opening count plus audited receipts minus audited outward movements.
- `allocated`: physical units assigned to open order lines but not yet picked.
- `free = on_hand - allocated`.
- Per order line, `remaining = ordered - picked` and `unallocated_demand = remaining - allocated_to_line`.
- Enforce `on_hand >= 0`, `0 <= allocated <= on_hand`, and `0 <= allocated_to_line <= remaining` in the same database write as each change.

Demand may exceed physical stock. Allocation may not. Incoming goods, promised goods and customer demand never increase `on_hand`.

Unallocated demand is not automatically awaiting supplier receipt. If goods have arrived but remain free, show “Awaiting assignment” and the available free stock. Only when no free stock is available does an unallocated line necessarily await supply. Where some free stock exists but cannot cover all lines, staff assigns it explicitly; do not claim the same free units for several lines. For one item with a verified balance, aggregate uncovered demand is `max(0, sum(open-line unallocated_demand) - free)`. Label it as aggregate demand not covered by current shelf stock, with Paid and Awaiting Payment demand shown separately; it is not a per-line allocation or purchasing instruction. A receipt without assignment must reduce this aggregate uncovered demand while leaving line allocation unchanged.

“Awaiting supply” is not “quantity to purchase”: goods may already be ordered from MISUMI. The first version records an optional supplier reference/note and leaves purchasing manual. It must not generate purchase orders or claim to know unrecorded goods in transit.

Unknown opening stock is a blocking state, not numeric zero. A staff-confirmed zero is valid. Do not activate an unknown item with a database default of zero.

## Staff workflow

### 1. Register and activate the item

An administrator selects the exact SKU, Ecwid parent ID, optional variation ID and canonical option selection. Supplier and location are explicit fields. Validate uniqueness, one-to-one units, supported options and absence of component-consumption relationships; never infer eligibility from the name or unlimited setting alone.

Capture a verified physical opening count and cutover time. Keep the item inactive until identity, stock policy, count and outstanding order lines are reviewed. Mode changes after activation require reconciliation and an audit entry; do not reinterpret historical movements.

### 2. Receive goods

The receiver scans the SKU, enters the entire accepted quantity and a delivery reference/note, and saves once with an operation ID. Exact-order purchases must be recorded just like extra purchases.

Receipt increases physical stock only. It does not mark any item picked, paid, packed or shipped. In this build, receiving and assignment are separate audited operations: received stock remains free until staff assigns it on the order screen. A combined receipt-and-assignment operation is not implemented.

Do not introduce automatic FIFO, paid-order priority or supplier purchasing. Staff chooses the order allocation; the screen shows payment status and remaining need. A future allocation policy can be added separately.

### 3. Allocate and pick

Staff can allocate free stock to an exact open order line. Allocation may hold stock for an Awaiting Payment order, but it cannot make that order pickable. No two orders may receive the same physical unit.

Picking requires all of:

1. Paid order with eligible fulfillment status (`AWAITING_PROCESSING` or `PROCESSING`).
2. Valid, active exact mapping and no relevant unresolved inventory/order issue.
3. Quantity no greater than the line's remaining quantity, its allocation, or physical stock.

A pick atomically reduces physical stock and allocation and increases picked quantity. It creates no Ecwid stock outbox entry. Partial receipts and partial picks are allowed; unfilled quantities remain visible.

Example, with a separately verified opening count of zero:

| Action | On shelf | Allocated | Free | Ordered | Picked | Unallocated demand |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Paid order for 20 | 0 | 0 | 0 | 20 | 0 | 20 |
| Receive 25; explicitly allocate 20 | 25 | 20 | 5 | 20 | 0 | 0 |
| Pick the allocated 20 | 5 | 0 | 5 | 20 | 20 | 0 |

### 4. Handle a mixed order

Distinguish a known, mapped supplier line awaiting receipt from an unmapped or contradictory line.

- A supplier shortage alone is a line-level state. Safely mapped, available stock-limited lines in the same Paid order may be picked.
- The order can show “Partly picked — awaiting supplier”. The supplier line remains unpicked, and the order cannot be shown as fully picked.
- Unmapped/unsupported lines, edited identities, inconsistent snapshots, cancellation after picking and terminal statuses with missing picks remain order-wide review conditions in the first version. Do not simply remove the existing whole-order guard or mark ignored lines complete.
- “Fully picked” requires every required line to be completed. It is not automatically “Packed”, “Shipped” or “Ready for Pickup”. No new Ecwid order-status write is part of this design.
- Ready for Pickup remains the store's picked-and-packed terminal status, equivalent to Shipped for inventory purposes. An incoming terminal status with missing local picks must trigger reconciliation, not fabricated picks.

### 5. Handle cancellation, errors and returns

- Before any picking: release allocations and close cancelled demand. Physical stock does not increase. Buying/cancelling goods from a supplier is a separate manual action.
- After partial picking: hold the order for reconciliation. Keep original movements and do not silently release reservations or recreate stock.
- Stock returns require an explicit record of an actual physical return. They are not inferred from an Ecwid refund/cancellation and are not normal duplicate restocks. The existing return-workflow gap remains a go-live restriction until implemented and tested.
- Correct a receipt or allocation with a linked audited reversal/correction, not by editing or deleting its original event. Do not reverse received units that have already been allocated or consumed without reconciliation.
- Offline/email sales and internal use can consume only free supplier stock. An email order already created in Ecwid must use its order-pick flow to avoid double recording.

## Implementation boundary and integration plan

The following describes implementation boundaries. The core schema, runtime, UI and isolated tests are implemented; real-item onboarding, opening-order import and production activation are not.

1. **Schema and ledger guards.** Add an explicit inventory mode and an immutable per-line allocation/release/pick ledger, with a current allocation projection. Extend movement audit data with the mode used when the movement was created. Supplier receipts/outward movements must have `ecwid_quantity_delta = 0`; stock-limited events retain their existing rules. Ensure every supplier operation produces zero Ecwid outbox records, including retries.
2. **Opening-stock and mapping validation — partially implemented.** Database guards require an explicit physical opening count before a supplier item can activate; demo fixtures supply deliberately fabricated counts. A reviewed production supplier-onboarding/activation path is still required, including unlimited-online policy checks instead of finite stock. The existing staging endpoint rejects nonzero reservations and does not import real order lines; that limitation is not resolved by this build.
3. **Order synchronization.** Store exact order line demand and distinguish normal supplier shortage from order integrity review. Catalogue reconciliation must validate according to the explicit mode. An unexpected finite/unlimited policy change remains reviewable, not an automatic mode switch.
4. **Picking and receiving UI.** Add supplier receipt and allocation controls. Show on shelf, assigned, free and awaiting supply separately; keep unpaid picking disabled. Mixed-order totals include every line.
5. **Isolated verification and opening orders.** Test against fictitious data in a separate local database first. Then prepare a fresh, reviewed real opening-order import with supplier receipt status represented separately from demand and physical allocations. No live stock alignment until separately approved.

Current integration points: `src/domain.ts`, `src/inventory.ts`, `src/supplier.ts`, `src/sync.ts`, migration `0005_supplier_inventory.sql`, and the receiving/picking screens. Earlier migrations are retained. The new rehearsal uses a separate local database; real stock and the original demo database are not migrated as part of this verification. `src/opening-import.ts` and `src/opening-apply.ts` still implement stock-limited preview/inactive staging only.

All coupled ledger/projection changes must be atomic and guarded in the database, not a read-then-write check in application code. D1 supports transactional batches that roll back when a statement fails. Guard failures must raise an error rather than silently produce a partial no-op. [Cloudflare D1 batch documentation](https://developers.cloudflare.com/d1/worker-api/d1-database/#batch)

Receipts, assignments, releases and picks need operation-ID idempotency: identical retries return the original result; a changed payload or actor under the same ID is rejected. Preserve existing unknown-outcome protections. No in-memory lock or external Ecwid call belongs inside the physical inventory transaction.

## Acceptance scenarios for implementation

| Scenario | Required result |
| --- | --- |
| 20 ordered, no verified opening count | Item inactive; no invented zero, allocation or pick |
| Verified zero; receive exactly 20; allocate and pick 20 | One receipt, one pick; balance zero; no Ecwid stock writes |
| Receive 25 for demand 20 | Allocate 20 only when selected; surplus five remains free |
| Receive eight for demand 20 | At most eight allocated/pickable; twelve still awaiting supply |
| Receive 25 for demand 20 without allocation | Shelf/free stock 25; line demand 20 awaiting assignment, aggregate uncovered demand zero; not falsely awaiting supplier |
| Two order lines compete for the last unit | Exactly one allocation succeeds |
| Two pickers consume the last allocated unit | Exactly one pick succeeds |
| Retry a receipt/allocation/pick after a lost response | No duplicate stock or picked quantity |
| Reuse operation ID with changed quantity or actor | Reject without stock change |
| Cancellation races an allocation or pick | Revalidate order state atomically; only a valid transition commits, with no stranded or duplicated allocation |
| Awaiting Payment order with allocation | Stock held, picking blocked; Paid transition permits eligible picks |
| Paid mixed order, supplier line not received | Stocked lines may be picked; order remains incomplete |
| Mixed order also contains an unsupported line | Whole order remains in review; no fabricated completion |
| Cancel an unpicked supplier line with allocation | Release allocation; shelf quantity unchanged |
| Cancel/refund after partial picking | Reconciliation hold; no automatic physical return |
| Mark terminal remotely before all lines picked | Review issue; no generated physical movements |
| Supplier email/internal use tries to consume allocated stock | Reject; only free stock may leave |
| Exact variation/options mismatch or finite/unlimited policy drift | Review; no stock write or silent remap |
| Stock-limited regression cases | Existing reservation, outbox and payment behavior unchanged |

## Verification completed on 22 September 2026

- 662 application tests across 20 files, including SQL ledger guards, HTTP boundaries, sync defenses and UI retry/eligibility helpers; TypeScript checks passed.
- 92 existing workbook/catalogue/review regression tests passed. Worker/asset build completed as a dry run, with no deployment.
- Actual local D1 migration/fixture setup succeeded with matching expected trigger/view names and no foreign-key errors. Original demo database counts and schema remained unchanged.
- Browser rehearsal: receive 25 fabricated supplier units; explicitly assign/pick 20; five remain on shelf. A stocked line in a mixed order can be picked while another supplier line awaits receipt, and the whole order remains incomplete.
- Browser rehearsal: assign five surplus units to an Awaiting Payment order; picking stays locked. Release two; shelf stock stays five, assigned stock becomes three and free stock two. Both physical movements and assignment events appear in their audit histories, with zero Ecwid outbox entries.
- A separate clean final rehearsal database is served on `http://127.0.0.1:8793`. Earlier fabricated test states are retained separately. No real pilot mapping, order or stock was imported. Physical-device camera scanning and factory-floor use are not validated by these browser checks.

## Deferred work

Attribute identifiers and catalogue-wide automatic tagging; supplier API integration or automated purchasing; full purchase-order accounting; direct-to-customer supplier delivery; unit conversions and component/BOM deduction; automatic order-status updates; production deployment and live stock alignment.

Next: review the isolated rehearsal, then implement exact opening-order import and the separately approved onboarding/cutover process. Actual opening counts, fresh order snapshots and the finite-stock exception need resolution before activating real inventory. Linked physical receipt corrections and return reconciliation are also still pending; use neither arbitrary balance edits nor a generic restock to simulate them.
