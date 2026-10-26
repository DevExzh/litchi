//! Host-vs-leaf differential parity harness for the `litchi-iwa` exit
//! (record 0363 §4; usage in record 0369).
//!
//! Every case runs one operation, expressed once, through the host route
//! (`litchi_iwa::{numbers,pages,keynote}::Draft`) and through the leaf route
//! (`litchi_{numbers,pages,keynote}::Package`) on the same fixture bytes and
//! asserts the tier the case declares:
//!
//! - **Tier A**: the two outputs are byte-identical, the whole ZIP included;
//! - **Tier B**: every ZIP member and every IWA object the host left
//!   untouched is byte-identical in the leaf output, and every object the
//!   host changed is equal in the leaf output on every known protobuf field;
//!   unknown-field and ZIP-layout differences are reported, not failed.
//!
//! Both tiers also require identical reads of both outputs, a strict reopen
//! of both through the leaf `Package`, equal object identifiers, and (when
//! the leaf route returns one) an inverse that restores the input bytes.
//! Each case declares whether its host route has native Apple evidence
//! (owner decision D1); `tools/iwa_exit_ledger.py` reads that declaration
//! and requires a `docs/unified-ops/0363-mac-queue.md` row for every
//! migrated route without it.
//!
//! The harness lives in the host's tests so that W3 deletes it with the host.

mod archive_diff;
mod cost_probes;
mod harness;
mod known_fields;

mod keynote_cases;
mod numbers_cases;
mod pages_cases;
mod self_test;
