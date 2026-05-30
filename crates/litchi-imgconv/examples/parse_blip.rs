//! Parse an OfficeArt BLIP record and convert it to PNG.
//!
//! A BLIP (Binary Large Image or Picture) record is the wrapper Microsoft
//! Office uses to embed images inside `.doc` / `.xls` / `.ppt` streams. This
//! example shows the two-step API:
//!
//!   1. [`litchi_imgconv::blip::Blip::parse`] decodes the OfficeArt record
//!      header, the optional UID(s), and the embedded picture payload.
//!   2. [`litchi_imgconv::convert_blip_to_png`] dispatches to the right
//!      backend (EMF/WMF/PICT decoder, or pass-through for already-raster
//!      formats) and produces a modern PNG.
//!
//! Because this repository does not ship a raw BLIP record fixture, the
//! example synthesises a minimal but spec-valid `BlipPNG` record (record
//! type `0xF01E`, see [MS-ODRAW] §2.2.23) by wrapping the bundled
//! `test-data/images/png/lena.png` with an 8-byte OfficeArt header, a
//! 16-byte zero UID, and the `0xFF` marker byte. The same code path works
//! on real BLIP bytes extracted from an Office document.
//!
//! # Run
//!
//! ```bash
//! # Synthesize a BlipPNG from the bundled PNG and convert it back out:
//! cargo run -p litchi-imgconv --example parse_blip
//!
//! # Or pass a real BLIP record file (raw bytes starting at the OfficeArt
//! # record header — i.e. starting with the 4-bit version + 12-bit
//! # instance + 16-bit record type + 32-bit length):
//! cargo run -p litchi-imgconv --example parse_blip -- path/to/blip.bin out.png
//! ```

use std::path::PathBuf;

use litchi_imgconv::blip::Blip;
use litchi_imgconv::{BlipType, convert_blip_to_png};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let blip_input = args.next();
    let output: PathBuf = args
        .next()
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("blip_out.png"));

    // Either load a real BLIP record from disk, or synthesise one inline
    // from a bundled PNG so the example is runnable out of the box.
    let blip_bytes: Vec<u8> = match blip_input {
        Some(path) => {
            println!("reading BLIP record from {path}");
            std::fs::read(&path)?
        },
        None => {
            let png_path = PathBuf::from("test-data/images/png/lena.png");
            if !png_path.exists() {
                return Err(format!(
                    "no BLIP path given and the bundled fallback does not exist: {} \
                     (run from repo root)",
                    png_path.display()
                )
                .into());
            }
            let png = std::fs::read(&png_path)?;
            println!(
                "no BLIP path given - synthesising a BlipPNG from {} ({} bytes)",
                png_path.display(),
                png.len()
            );
            synthesize_blip_png(&png)
        },
    };

    // 1. Parse the OfficeArt BLIP record. This is zero-copy: `blip` borrows
    //    from `blip_bytes` for its picture data.
    let blip = Blip::parse(&blip_bytes)?;

    // 2. Inspect the parsed record.
    let kind = blip.blip_type();
    println!("\n--- BLIP metadata ---");
    println!("kind             : {:?}", kind);
    if let Some(t) = kind {
        println!("extension        : .{}", t.extension());
        println!("is_metafile      : {}", t.is_metafile());
    }
    match &blip {
        Blip::Metafile(m) => {
            println!(
                "header           : version={} instance=0x{:03X} type=0x{:04X} length={}",
                m.header.version, m.header.instance, m.header.record_type, m.header.length
            );
            println!("uid              : {}", hex16(&m.uid));
            if let Some(u) = m.secondary_uid {
                println!("secondary uid    : {}", hex16(&u));
            }
            println!("uncompressed_size: {} bytes", m.uncompressed_size);
            println!("compressed_size  : {} bytes", m.compressed_size);
            println!(
                "bounds           : ({}, {}) -> ({}, {})",
                m.bounds.0, m.bounds.1, m.bounds.2, m.bounds.3
            );
            println!("size (EMU)       : {} x {}", m.size_emu.0, m.size_emu.1);
            println!(
                "compression / filter: 0x{:02X} / 0x{:02X} (compressed = {})",
                m.compression,
                m.filter,
                m.is_compressed()
            );
            println!("picture_data len : {} bytes", m.picture_data.len());
        },
        Blip::Bitmap(b) => {
            println!(
                "header           : version={} instance=0x{:03X} type=0x{:04X} length={}",
                b.header.version, b.header.instance, b.header.record_type, b.header.length
            );
            println!("uid              : {}", hex16(&b.uid));
            println!("marker           : 0x{:02X}", b.marker);
            println!("picture_data len : {} bytes", b.picture_data.len());
        },
    }

    // 3. Convert to PNG. For an already-PNG bitmap BLIP this is essentially a
    //    decode/re-encode round-trip; for EMF/WMF/PICT the raster is rendered.
    let png = convert_blip_to_png(&blip, None, None)?;
    std::fs::write(&output, &png)?;
    println!(
        "\nconverted to PNG : {} ({} bytes)",
        output.display(),
        png.len()
    );
    Ok(())
}

/// Build a minimal `BlipPNG` OfficeArt record (`0xF01E`) wrapping `png_data`.
///
/// Layout (see MS-ODRAW §2.2.23 OfficeArtBlipPNG):
///   - 8-byte record header: ver/inst (u16 LE) | type (u16 LE) | length (u32 LE)
///   - 16-byte primary UID (we use all zeros - the parser does not validate it)
///   - 1-byte marker (0xFF means "external", any value parses)
///   - PNG file bytes
fn synthesize_blip_png(png_data: &[u8]) -> Vec<u8> {
    // For BlipPNG the canonical instance is 0x6E0 (no secondary UID); version is 0.
    // ver_inst layout: low 4 bits = version, upper 12 bits = instance.
    let version: u16 = 0x0;
    let instance: u16 = 0x6E0;
    let ver_inst: u16 = (instance << 4) | (version & 0x0F);
    let record_type: u16 = BlipType::Png as u16; // 0xF01E
    let body_len = (16 + 1 + png_data.len()) as u32;

    let mut out = Vec::with_capacity(8 + body_len as usize);
    out.extend_from_slice(&ver_inst.to_le_bytes());
    out.extend_from_slice(&record_type.to_le_bytes());
    out.extend_from_slice(&body_len.to_le_bytes());
    out.extend_from_slice(&[0u8; 16]); // primary UID
    out.push(0xFF); // marker
    out.extend_from_slice(png_data);
    out
}

fn hex16(b: &[u8; 16]) -> String {
    let mut s = String::with_capacity(32);
    for byte in b {
        s.push_str(&format!("{:02x}", byte));
    }
    s
}
