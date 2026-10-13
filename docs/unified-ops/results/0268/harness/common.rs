use std::time::Instant;

pub fn median(mut v: Vec<u128>) -> u128 {
    v.sort_unstable();
    v[v.len() / 2]
}

pub fn time<F: FnMut()>(iters: usize, mut f: F) -> u128 {
    // warm-up
    for _ in 0..3 {
        f();
    }
    let mut samples = Vec::with_capacity(iters);
    for _ in 0..iters {
        let start = Instant::now();
        f();
        samples.push(start.elapsed().as_nanos());
    }
    median(samples)
}

pub fn report(name: &str, ns: u128) {
    println!("{name}\t{ns}");
}

pub fn digest(bytes: &[u8]) -> u64 {
    // FNV-1a
    let mut h: u64 = 0xcbf29ce484222325;
    for b in bytes {
        h ^= u64::from(*b);
        h = h.wrapping_mul(0x100000001b3);
    }
    h
}
