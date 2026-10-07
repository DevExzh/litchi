//! Isolated allocator-instrumented managed XLSX compression diagnostic.
//!
//! The allocator wrapper is deliberately included only by this target.  The
//! route reports operation-scoped allocation counters and never reports native
//! timing, throughput, RSS, or adoption results.

#[path = "support/counting_allocator.rs"]
mod allocator;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    litchi_perf_baseline::allocation_metrics::enable();
    let mut arguments = std::env::args_os();
    let _executable = arguments.next();
    litchi_perf_baseline::measure_xlsx_managed_compression_allocations(arguments)
}
