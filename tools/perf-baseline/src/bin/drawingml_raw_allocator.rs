#[path = "support/drawingml_counting_allocator.rs"]
mod allocator;

fn main() -> std::io::Result<()> {
    litchi_perf_baseline::drawingml_raw_probe::run(Some(&allocator::OBSERVER))
}
