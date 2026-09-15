//! Callgrind isolation pair: `open_part_loop N <fixture> <partname>` performs N
//! iterations of "open the package, read one Part", reusing one source. Profile
//! N and N+M and divide the difference by M for the per-iteration cost.
use std::sync::Arc;

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::{PackURI, ReadLimits, SourceBackedPackage, SourceCacheLimits, SourceReadPolicy};

struct Bytes(Vec<u8>);

impl ReadAt for Bytes {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.0.len() as u64)
    }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        if start >= self.0.len() {
            return Ok(0);
        }
        let take = output.len().min(self.0.len() - start);
        output[..take].copy_from_slice(&self.0[start..start + take]);
        Ok(take)
    }
    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x600, 1))
    }
}

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let iterations: usize = args[0].parse().expect("iterations");
    let bytes = std::fs::read(&args[1]).expect("fixture");
    let part = PackURI::new(args[2].as_str()).expect("partname");
    let source: Arc<dyn ReadAt> = Arc::new(Bytes(bytes));
    let mut total = 0usize;
    for _ in 0..iterations {
        let package =
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy(
                Arc::clone(&source),
                ReadLimits::default(),
                SourceCacheLimits::default(),
                SourceReadPolicy::exact(),
            )
            .expect("open");
        let data = package.part(&part).expect("part").data().expect("data");
        total += std::hint::black_box(data.as_bytes().len());
    }
    println!("iterations={iterations} total={total}");
}
