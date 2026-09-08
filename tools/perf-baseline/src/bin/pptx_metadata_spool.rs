//! Compare ordinary and explicit metadata-spool PPTX creation.

#[cfg(feature = "allocator-metrics")]
#[path = "support/counting_allocator.rs"]
mod allocator;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    #[cfg(feature = "allocator-metrics")]
    litchi_perf_baseline::allocation_metrics::enable();

    litchi_perf_baseline::pptx_metadata_spool::run_from_args(std::env::args_os().skip(1))
}
