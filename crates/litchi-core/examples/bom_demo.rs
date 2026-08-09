//! Demonstrates Byte Order Mark (BOM) handling in `litchi-core`.
//!
//! Walks through every variant of [`BomKind`], synthesises a small payload
//! with that BOM prefixed, then uses [`strip_bom`] to detect the BOM and
//! [`write_bom`] to emit one. For each case the program prints the BOM
//! detected, the body length after stripping, and a hex preview of the
//! first few bytes.
//!
//! Run with:
//! ```bash
//! cargo run -p litchi-core --example bom_demo
//! ```
//!
//! No CLI arguments are required — the example is fully self-contained.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example exists to print its BOM walkthrough"
)]

use litchi_core::{
    BomKind, UTF8_BOM, UTF16_BE_BOM, UTF16_LE_BOM, UTF32_BE_BOM, UTF32_LE_BOM, strip_bom, write_bom,
};
use std::io::Cursor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== litchi-core BOM demo ===\n");

    print_bom_constants();
    println!();

    // Body fed alongside each BOM. ASCII for UTF-8; for UTF-16/UTF-32 we
    // encode "Hi" in the matching width so the hex preview looks plausible.
    // The detector itself only inspects the prefix, so any opaque trailer
    // works.
    let body_ascii = b"Hello, BOM!";
    demo_round_trip("UTF-8", BomKind::Utf8, body_ascii)?;
    demo_round_trip("UTF-16 LE", BomKind::Utf16Le, &encode_utf16_le("Hi"))?;
    demo_round_trip("UTF-16 BE", BomKind::Utf16Be, &encode_utf16_be("Hi"))?;
    demo_round_trip("UTF-32 LE", BomKind::Utf32Le, &encode_utf32_le("Hi"))?;
    demo_round_trip("UTF-32 BE", BomKind::Utf32Be, &encode_utf32_be("Hi"))?;

    println!();
    demo_no_bom(b"no-bom-here, just plain ASCII")?;

    Ok(())
}

/// Print the raw constants exposed at the crate root, plus each variant's
/// `as_bytes()` and `len()` accessors.
fn print_bom_constants() {
    println!("Public BOM constants:");
    println!("  UTF8_BOM     = {}", hex(&UTF8_BOM));
    println!("  UTF16_LE_BOM = {}", hex(&UTF16_LE_BOM));
    println!("  UTF16_BE_BOM = {}", hex(&UTF16_BE_BOM));
    println!("  UTF32_LE_BOM = {}", hex(&UTF32_LE_BOM));
    println!("  UTF32_BE_BOM = {}", hex(&UTF32_BE_BOM));

    println!("\nBomKind accessors:");
    for kind in [
        BomKind::Utf8,
        BomKind::Utf16Le,
        BomKind::Utf16Be,
        BomKind::Utf32Le,
        BomKind::Utf32Be,
    ] {
        println!(
            "  {kind:?}: {} bytes, prefix = {}",
            kind.len(),
            hex(kind.as_bytes())
        );
    }
}

/// Build `[BOM | body]`, write it via `write_bom`, then read it back via
/// `strip_bom` and report what was detected.
fn demo_round_trip(
    label: &str,
    kind: BomKind,
    body: &[u8],
) -> Result<(), Box<dyn std::error::Error>> {
    // 1. Build the payload using `write_bom`.
    let mut payload: Vec<u8> = Vec::with_capacity(kind.len() + body.len());
    write_bom(&mut payload, kind)?;
    payload.extend_from_slice(body);

    // 2. Detect it back via `strip_bom`, which seeks past the BOM on success.
    let mut cursor = Cursor::new(&payload);
    let detected = strip_bom(&mut cursor)?;

    // 3. Whatever remains is the body.
    let mut remaining = Vec::new();
    std::io::Read::read_to_end(&mut cursor, &mut remaining)?;

    println!("--- {label} ---");
    println!(
        "  full payload     : {} bytes  preview = {}",
        payload.len(),
        hex_preview(&payload, 12)
    );
    match detected {
        Some((found, consumed)) => {
            println!("  detected BOM     : {found:?} ({consumed} bytes consumed)");
            assert_eq!(found, kind, "round trip mismatch for {kind:?}");
        },
        None => println!("  detected BOM     : <none>"),
    }
    println!(
        "  body after strip : {} bytes  preview = {}",
        remaining.len(),
        hex_preview(&remaining, 12)
    );
    println!();
    Ok(())
}

/// Show that `strip_bom` returns `Ok(None)` and does NOT advance the cursor
/// when the input lacks a BOM.
fn demo_no_bom(body: &[u8]) -> Result<(), Box<dyn std::error::Error>> {
    let mut cursor = Cursor::new(body);
    let detected = strip_bom(&mut cursor)?;
    let pos_after = std::io::Seek::stream_position(&mut cursor)?;

    println!("--- no BOM ---");
    println!(
        "  input            : {} bytes  preview = {}",
        body.len(),
        hex_preview(body, 12)
    );
    println!("  detected BOM     : {detected:?}");
    println!("  cursor position  : {pos_after} (should be 0 — strip_bom rewinds on miss)");
    Ok(())
}

// --- helpers -----------------------------------------------------------------

fn hex(bytes: &[u8]) -> String {
    const DIGITS: &[u8; 16] = b"0123456789ABCDEF";
    let mut s = String::with_capacity(bytes.len() * 3);
    for (i, b) in bytes.iter().enumerate() {
        if i > 0 {
            s.push(' ');
        }
        s.push(char::from(DIGITS[usize::from(b >> 4)]));
        s.push(char::from(DIGITS[usize::from(b & 0x0F)]));
    }
    s
}

fn hex_preview(bytes: &[u8], max: usize) -> String {
    let take = bytes.len().min(max);
    let mut out = hex(&bytes[..take]);
    if bytes.len() > take {
        out.push_str(" ...");
    }
    out
}

fn encode_utf16_le(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(s.len() * 2);
    for unit in s.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn encode_utf16_be(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(s.len() * 2);
    for unit in s.encode_utf16() {
        out.extend_from_slice(&unit.to_be_bytes());
    }
    out
}

fn encode_utf32_le(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(s.len() * 4);
    for ch in s.chars() {
        out.extend_from_slice(&(ch as u32).to_le_bytes());
    }
    out
}

fn encode_utf32_be(s: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(s.len() * 4);
    for ch in s.chars() {
        out.extend_from_slice(&(ch as u32).to_be_bytes());
    }
    out
}
