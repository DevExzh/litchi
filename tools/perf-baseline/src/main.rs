#![forbid(unsafe_code)]

use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    litchi_perf_baseline::run()
}
