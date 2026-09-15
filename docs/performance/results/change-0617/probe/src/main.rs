//! Timing and callgrind entry point: no allocator wrapper is installed.

fn main() -> Result<(), cfb_save_probe::BoxError> {
    cfb_save_probe::run()
}
