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
        Some(selector) if selector == std::ffi::OsStr::new("docx-managed-read-ahead") => {
            litchi_perf_baseline::docx_managed_read_ahead::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("pptx-pair-lifecycle") => {
            litchi_perf_baseline::pptx_pair_lifecycle::run_from_args(args)
        },
        Some(selector) if selector == std::ffi::OsStr::new("odp-append-attribution") => {
            litchi_perf_baseline::odp_append_attribution::run_from_args(args)
        },
        _ => litchi_perf_baseline::run(),
    }
}
