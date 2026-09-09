#[cfg(feature = "allocator-metrics")]
#[path = "support/counting_allocator.rs"]
mod allocator;

use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    #[cfg(feature = "allocator-metrics")]
    litchi_perf_baseline::allocation_metrics::enable();
    litchi_perf_baseline::docx_bounded_tail_append_compare::run_from_args(
        std::env::args_os().skip(1),
    )
}
