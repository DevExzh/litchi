//! Change 0649 measurement-only counters. NEVER COMMITTED.
//!
//! Deterministic per-operation counts for the opened-PPTX edit attribution:
//! how many MCE preprocessing passes run, over how many bytes, how many of
//! them rewrite rather than borrow, and the same for part decodes, shape-scene
//! parses and complete-package fingerprints.
use std::sync::atomic::{AtomicU64, Ordering};

macro_rules! counters {
    ($($name:ident),* $(,)?) => {
        $(pub static $name: AtomicU64 = AtomicU64::new(0);)*
        /// Every counter, in declaration order.
        #[must_use]
        pub fn snapshot() -> Vec<(&'static str, u64)> {
            vec![$((stringify!($name), $name.load(Ordering::Relaxed))),*]
        }
        /// Zero every counter.
        pub fn reset() {
            $($name.store(0, Ordering::Relaxed);)*
        }
    };
}

counters!(
    MCE_CALLS,
    MCE_INPUT_BYTES,
    MCE_REWRITES,
    MCE_REWRITE_INPUT_BYTES,
    MCE_REWRITE_OUTPUT_BYTES,
    MCE_ELEMENTS,
    MCE_NS_EMISSIONS,
    PART_DECODES,
    PART_DECODE_BYTES,
    SCENE_READS,
    SCENE_READ_BYTES,
    CAPTURES,
    FINGERPRINTS,
    FINGERPRINT_BYTES,
    PACKAGES_EQUAL,
    SITE_PRES_ROOT,
    SITE_PRES_SLIDE_REFS,
    SITE_SLIDE_ROOT_NAME,
    SITE_SLIDE_CSLD_NAME,
    SITE_NOTES_SNAPSHOT,
    SITE_PRES_ROOT_BYTES,
    SITE_PRES_SLIDE_REFS_BYTES,
    SITE_SLIDE_ROOT_NAME_BYTES,
    SITE_SLIDE_CSLD_NAME_BYTES,
    SITE_NOTES_SNAPSHOT_BYTES,
);

/// Add `value` to `counter`.
pub fn bump(counter: &AtomicU64, value: u64) {
    counter.fetch_add(value, Ordering::Relaxed);
}

/// Input lengths of every MCE pass, in call order. A length identifies the OPC
/// member almost uniquely, so the trace maps each pass onto a part name.
pub static MCE_TRACE: std::sync::Mutex<Vec<u64>> = std::sync::Mutex::new(Vec::new());

/// Record one MCE pass length.
pub fn trace(length: u64) {
    if let Ok(mut trace) = MCE_TRACE.lock() {
        trace.push(length);
    }
}

/// Take and clear the trace.
#[must_use]
pub fn take_trace() -> Vec<u64> {
    MCE_TRACE.lock().map(|mut t| std::mem::take(&mut *t)).unwrap_or_default()
}
