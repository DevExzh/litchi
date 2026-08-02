//! Demonstrate the LZFu compressed-RTF helpers: `is_compressed_rtf`, `compress`, `decompress`.
//!
//! The `compress`/`decompress` pair implements the algorithm used by Outlook for
//! the `PR_RTF_COMPRESSED` MAPI property. This example round-trips a small RTF
//! fragment through both compressed (LZFu) and uncompressed framings.
//!
//! Run from the workspace root:
//!
//! ```bash
//! cargo run -p litchi-rtf --example compressed_rtf
//! ```

use litchi_rtf::transport::{compress, decompress, is_compressed_rtf};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let original: &[u8] = b"{\\rtf1\\ansi{\\fonttbl\\f0\\fswiss Helvetica;}\\f0\\pard \
Hello compressed RTF! This text is repeated. Hello compressed RTF! \
This text is repeated.\\par}";

    println!("Original RTF size: {} bytes", original.len());
    println!(
        "is_compressed_rtf(original) -> {} (raw RTF has no header)",
        is_compressed_rtf(original)
    );

    // ----- LZFu compressed framing -----
    let compressed = compress(original, true)?;
    println!("\nLZFu framing");
    println!("{}", "-".repeat(60));
    println!("Compressed payload size: {} bytes", compressed.len());
    println!(
        "is_compressed_rtf(compressed) -> {}",
        is_compressed_rtf(&compressed)
    );

    let round_trip = decompress(&compressed)?;
    assert_eq!(
        round_trip, original,
        "LZFu round-trip should restore the exact bytes"
    );
    println!(
        "Decompressed size: {} bytes (matches original)",
        round_trip.len()
    );

    // ----- Uncompressed framing -----
    // The `compress = false` mode still wraps the data in the 16-byte
    // compressed-RTF header, but skips LZFu encoding.
    let stored = compress(original, false)?;
    println!("\nUncompressed framing");
    println!("{}", "-".repeat(60));
    println!("Stored payload size: {} bytes", stored.len());
    println!(
        "is_compressed_rtf(stored) -> {}",
        is_compressed_rtf(&stored)
    );

    let stored_round_trip = decompress(&stored)?;
    assert_eq!(
        stored_round_trip, original,
        "Uncompressed round-trip should restore the exact bytes"
    );
    println!(
        "Decompressed size: {} bytes (matches original)",
        stored_round_trip.len()
    );

    println!("\nAll round-trip assertions passed.");
    Ok(())
}
