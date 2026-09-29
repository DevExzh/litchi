//! Isolated operation-allocation observer entry point.
//!
//! The only unsafe boundary is the copied counting global allocator.  The
//! workload itself is a safe module and owns the corpus and verification
//! checks.

#[allow(dead_code)]
mod allocation_metrics;
mod counting_allocator;
mod workload;

fn main() {
    allocation_metrics::enable();
    workload::run_from_args();
}
