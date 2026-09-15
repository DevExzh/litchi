use litchi_core::{Budget, Limits, Resource, ReadAt, FileSource, OwnedSource};
use std::hint::black_box;
use std::time::Instant;

fn bench<F: FnMut()>(name: &str, iters: u64, mut f: F) {
    // warm-up
    for _ in 0..(iters / 10).max(1) { f(); }
    let mut best = f64::MAX;
    for _ in 0..5 {
        let start = Instant::now();
        for _ in 0..iters { f(); }
        let ns = start.elapsed().as_nanos() as f64 / iters as f64;
        if ns < best { best = ns; }
    }
    println!("{name:<48} {best:>9.1} ns/op (best of 5 x {iters})");
}

fn main() {
    let path = std::env::args().nth(1).expect("fixture path");
    let limits = Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let root = Budget::root("root", limits);
    let child = root.child("child", limits);
    let grandchild = child.child("grandchild", limits);
    let n = 2_000_000;
    bench("Budget::consume depth1 (root)", n, || { black_box(root.consume(Resource::Work, 1)).unwrap(); });
    bench("Budget::consume depth3 (grandchild)", n, || { black_box(grandchild.consume(Resource::Work, 1)).unwrap(); });
    bench("Budget::reserve+drop depth1 (root)", n, || { drop(black_box(root.reserve(Resource::Memory, 1).unwrap())); });
    bench("Budget::reserve+drop depth3 (grandchild)", n, || { drop(black_box(grandchild.reserve(Resource::Memory, 1).unwrap())); });

    let file = FileSource::open(&path).expect("open");
    let owned = OwnedSource::new(std::fs::read(&path).expect("read"));
    let mut buf = vec![0u8; 4096];
    let m = 200_000;
    bench("FileSource::version() (mutex+fstat)", m, || { black_box(file.version().unwrap()); });
    bench("FileSource::len() (fstat)", m, || { black_box(file.len().unwrap()); });
    bench("FileSource::read_at 512 B (pread)", m, || { black_box(file.read_at(4096, &mut buf[..512]).unwrap()); });
    bench("FileSource::read_at 4 KiB (pread)", m, || { black_box(file.read_at(4096, &mut buf).unwrap()); });
    bench("OwnedSource::version()", n, || { black_box(owned.version().unwrap()); });
    bench("OwnedSource::read_at 4 KiB (memcpy)", n, || { black_box(owned.read_at(4096, &mut buf).unwrap()); });
    let v1 = file.version().unwrap();
    bench("SourceVersion == compare", n, || { black_box(black_box(v1) == black_box(v1)); });
}
