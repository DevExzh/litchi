//! Isolated allocator-instrumented benchmark entry point.
//!
//! The shared harness library is forbid-safe. Only this target owns the
//! process-global allocator wrapper, which lives in the reusable support
//! module so isolated benchmark bins report the same metric semantics.

#[path = "support/counting_allocator.rs"]
mod allocator;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    litchi_perf_baseline::allocation_metrics::enable();
    let mut args = std::env::args_os();
    let _executable = args.next();
    match args.next() {
        Some(selector) if selector == std::ffi::OsStr::new("retention") => {
            litchi_perf_baseline::pptx_retention::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("provider-lifecycle") => {
            litchi_perf_baseline::pptx_provider_lifecycle::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("docx-provider-lifecycle") => {
            litchi_perf_baseline::docx_provider_lifecycle::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("docx-section-retention") => {
            litchi_perf_baseline::docx_section_retention::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("docx-edit-provider-cold") => {
            litchi_perf_baseline::docx_edit_provider::run_cold_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("docx-edit-provider") => {
            litchi_perf_baseline::docx_edit_provider::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("docx-managed-edit") => {
            litchi_perf_baseline::docx_managed_edit::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("docx-managed-read-ahead") => {
            litchi_perf_baseline::docx_managed_read_ahead::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("docx-raw-codec") => {
            litchi_perf_baseline::docx_raw_codec::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("docx-owned-attribution") => {
            litchi_perf_baseline::docx_owned_attribution::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("xlsx-raw-identity") => {
            litchi_perf_baseline::xlsx_raw_identity::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("pptx-pair-lifecycle") => {
            litchi_perf_baseline::pptx_pair_lifecycle::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("pptx-notes-memo") => {
            litchi_perf_baseline::pptx_notes_memo::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("odp-append-attribution") => {
            litchi_perf_baseline::odp_append_attribution::run_from_args(args)
        },
        _ => litchi_perf_baseline::run(),
    }
}
