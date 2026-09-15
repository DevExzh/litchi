//! Determinism check for `OleWriter::write_to` with explicitly created storages.
//!
//! `write_to` iterates `self.storages: HashSet<Vec<String>>` to call
//! `DirectoryBuilder::add_storage_path`, which assigns each new entry
//! `sid = entries.len()`; entries are then serialized in SID order. Rust's
//! default hasher is seeded per process, so the directory image — and therefore
//! the whole file — may depend on hash iteration order across runs.
//!
//! Prints the output length and a FNV-1a digest of the bytes. Run it twice in
//! separate processes and compare.

use litchi_cfb::writer::OleWriter;
use std::io::Cursor;

fn fnv1a(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    hash
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let count: usize = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "8".to_string())
        .parse()?;
    let mut writer = OleWriter::new();
    for index in 0..count {
        let name = format!("Storage{index:02}");
        writer.create_storage(&[name.as_str()])?;
        writer.create_stream(&[name.as_str(), "Payload"], format!("payload-{index}").as_bytes())?;
    }
    let mut output = Vec::new();
    writer.write_to(&mut Cursor::new(&mut output))?;
    println!("{{\"storages\":{count},\"bytes\":{},\"fnv1a\":\"{:016x}\"}}", output.len(), fnv1a(&output));
    Ok(())
}
