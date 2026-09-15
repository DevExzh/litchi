//! Change 0628 probe D: the callgrind isolation-pair body.
//!
//! Opens and republishes one fixture `iterations` times so that profiling
//! `iterations = N` and `iterations = N + M` and differencing the totals gives
//! the instruction cost of M open+save round trips, with the process start-up,
//! the file read and the profiler's own fixed cost cancelled out.
//!
//! Usage: open_save_iters <fixture> <iterations>
fn main() {
    let fixture = std::env::args().nth(1).expect("fixture");
    let iterations: usize = std::env::args()
        .nth(2)
        .and_then(|value| value.parse().ok())
        .expect("iterations");
    let bytes = std::fs::read(&fixture).expect("fixture");
    let mut checksum = 0u64;
    for _ in 0..iterations {
        let package = litchi_opc::OpcPackage::from_bytes(&bytes).expect("open");
        let output = litchi_opc::PackageWriter::to_bytes(&package).expect("save");
        checksum = checksum.wrapping_add(output.len() as u64);
    }
    println!("iterations={iterations} checksum={checksum}");
}
