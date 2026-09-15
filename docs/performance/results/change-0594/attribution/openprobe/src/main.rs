// Isolation-pair probe for change 0594's attribution: open one OOXML package
// from a filesystem path N times. `perf stat` or callgrind over two runs with
// different N, differenced and divided by the difference, gives the per-open
// cost of the region under study with process startup cancelled out.
use litchi_opc::SourceBackedPackage;

fn main() {
    let mut args = std::env::args().skip(1);
    let path = args.next().expect("fixture path");
    let repeat: usize = args
        .next()
        .expect("repeat count")
        .parse()
        .expect("repeat count is a number");
    let mut checksum: usize = 0;
    for _ in 0..repeat {
        let package = SourceBackedPackage::from_path(&path).expect("open");
        checksum = checksum.wrapping_add(package.iter_parts().count());
        std::hint::black_box(&package);
    }
    println!("repeat={repeat} parts_seen={checksum}");
}
