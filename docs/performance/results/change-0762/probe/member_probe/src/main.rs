//! Member-level differential and cross-process determinism probe for change
//! 0762, built once against each tree (`Cargo.toml.in`, `@TREE@`).
//!
//! `digest` builds the harness's streaming DOCX, XLSX and PPTX corpora (tiny,
//! medium, large) and prints, per corpus, the archive's SHA-256 and length
//! and, per member, its name, uncompressed length, the SHA-256 of its
//! uncompressed payload and its compressed size. Two trees agree on every
//! member line exactly when they publish the same members with the same
//! uncompressed bytes.
//!
//! `split SEED` builds the large DOCX corpus handing every run's text to
//! `write_text` in SEED-dependent pieces, and the large XLSX and PPTX corpora,
//! and prints each archive's SHA-256; run in separate processes with
//! different seeds, identical lines mean the bytes depend on neither the
//! caller's write sizes nor the process.

use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits};
use litchi_docx::{StreamingDocumentLimits, StreamingDocumentWriter};
use litchi_opc::phys_pkg::OwnedPhysPkgReader;
use litchi_pptx::{
    StreamingPresentationLimits, StreamingPresentationOptions, StreamingPresentationWriter,
    TextBoxSpec,
};
use litchi_xlsx::{StreamingCell, StreamingCellValue, StreamingWorkbookLimits, StreamingWorkbookWriter};
use sha2::{Digest, Sha256};
use std::num::{NonZeroU64, NonZeroUsize};

type Res<T> = Result<T, Box<dyn std::error::Error>>;

fn hex(bytes: &[u8]) -> String {
    Sha256::digest(bytes).iter().map(|b| format!("{b:02x}")).collect()
}

fn streaming_context(memory: u64, input: u64, output: u64, objects: u64, work: u64) -> Res<ExecutionContext> {
    let one = NonZeroUsize::new(1).ok_or("one")?;
    let in_flight = NonZeroU64::new(memory.max(1)).ok_or("in flight")?;
    let limits = ExecutionLimits::new(one, one, in_flight, 0)?;
    let (_source, token) = CancellationSource::pair();
    Ok(ExecutionContext::new(
        Budget::root("probe", Limits::new(memory, input, output, objects, 32, work)),
        token,
        limits,
    ))
}

fn docx_text(index: usize) -> String {
    format!("litchi-perf-docx-streaming-v1-{index:06}-café-<&>")
}

/// `seed == 0` writes each run's text whole.
fn build_docx(paragraphs: usize, seed: u64) -> Res<Vec<u8>> {
    let texts: Vec<String> = (0..paragraphs).map(docx_text).collect();
    let input: u64 = texts.iter().map(|t| t.len() as u64).sum();
    let max_run = texts.iter().map(String::len).max().unwrap_or(0) as u64;
    let document_xml = (max_run * 8 + 256) * paragraphs as u64 + 16 * 1024;
    let output = document_xml + 256 * 1024;
    let p = paragraphs as u64;
    let limits = StreamingDocumentLimits::new(input, output, p, p, 1, max_run, document_xml, document_xml, 64);
    // Splitting a run's text into k pieces charges the same Work (one unit
    // per character) but k scans; the budget below is the harness's.
    let context = streaming_context(64, input, output, p * 2 + 16, p * 4 + input + 32)?;
    let mut writer = StreamingDocumentWriter::new(Vec::new(), context, limits)?;
    let mut state = seed;
    for text in &texts {
        writer.start_paragraph()?;
        writer.start_run()?;
        if seed == 0 {
            writer.write_text(text)?;
        } else {
            let mut rest = text.as_str();
            while !rest.is_empty() {
                state ^= state << 13;
                state ^= state >> 7;
                state ^= state << 17;
                // Seeds 1 and 2 cut every run into 1- and 2-byte pieces; any
                // other seed into pieces of 1 to 23 bytes.
                let piece = match seed {
                    1 => 1,
                    2 => 2,
                    _ => 1 + (state % 23) as usize,
                };
                let mut take = piece.min(rest.len());
                while !rest.is_char_boundary(take) {
                    take += 1;
                }
                writer.write_text(&rest[..take])?;
                rest = &rest[take..];
            }
        }
        writer.finish_run()?;
        writer.finish_paragraph()?;
    }
    Ok(writer.finish()?)
}

fn xlsx_text(row: usize) -> String {
    format!("litchi-perf-streaming-xlsx-row-{row:06}-café-<&>")
}

fn build_xlsx(rows: usize) -> Res<Vec<u8>> {
    let r = rows as u64;
    let cells = r * 4;
    let max_sheet = r * 512 + 4 * 1024;
    let max_output = max_sheet * 2 + 64 * 1024;
    let objects = r + cells + 16;
    let work = objects * 2;
    let limits = StreamingWorkbookLimits::new(u32::try_from(rows)?, cells, 256, 4 * 1024, max_sheet, max_output);
    let context = streaming_context(4 * 1024, 0, max_output, objects, work)?;
    let mut writer = StreamingWorkbookWriter::new(Vec::new(), context, limits)?;
    for row in 1..=rows {
        let text = xlsx_text(row);
        let number = u32::try_from(row)?;
        writer.write_row(
            number,
            [
                StreamingCell::new(1, StreamingCellValue::Number(f64::from(number))),
                StreamingCell::new(2, StreamingCellValue::Text(&text)),
                StreamingCell::new(3, StreamingCellValue::Bool(row % 2 == 0)),
                StreamingCell::new(4, StreamingCellValue::Blank),
            ],
        )?;
    }
    Ok(writer.finish()?)
}

fn pptx_text(index: usize) -> String {
    format!("litchi-perf-pptx-streaming-v1-{index:06}-café-&<>")
}

fn build_pptx(slides: usize) -> Res<Vec<u8>> {
    let texts: Vec<String> = (0..slides).map(pptx_text).collect();
    let input: usize = texts.iter().map(String::len).sum();
    let max_text = texts.iter().map(String::len).max().unwrap_or(0);
    let limits = StreamingPresentationLimits {
        max_slides: slides,
        max_text_boxes_per_slide: 1,
        max_text_bytes_per_box: max_text,
        max_total_text_bytes: input,
        max_slide_xml_bytes: max_text * 5 + 16 * 1024,
        max_output_bytes: 4 * 1024 * 1024 + slides as u64 * 8 * 1024,
    };
    let mut writer = StreamingPresentationWriter::with_options(
        Vec::new(),
        slides,
        StreamingPresentationOptions::standard(),
        limits,
    )?;
    for text in &texts {
        let mut slide = writer.start_slide(None)?;
        slide.write_text_box(TextBoxSpec::new(text, 914_400, 914_400, 7_315_200, 914_400))?;
        writer = slide.finish()?;
    }
    Ok(writer.finish()?)
}

fn compressed_sizes(archive: &[u8]) -> Res<Vec<u64>> {
    let reader = soapberry_zip::ZipArchive::from_slice(archive)?;
    let mut entries = reader.entries();
    let mut sizes = Vec::new();
    while let Some(record) = entries.next_entry()? {
        sizes.push(record.compressed_size_hint());
    }
    Ok(sizes)
}

fn digest() -> Res<()> {
    let corpora: Vec<(&str, Vec<u8>)> = vec![
        ("docx-tiny", build_docx(64, 0)?),
        ("docx-medium", build_docx(8_192, 0)?),
        ("docx-large", build_docx(131_072, 0)?),
        ("xlsx-tiny", build_xlsx(64)?),
        ("xlsx-medium", build_xlsx(8_192)?),
        ("xlsx-large", build_xlsx(131_072)?),
        ("pptx-tiny", build_pptx(8)?),
        ("pptx-medium", build_pptx(256)?),
        ("pptx-large", build_pptx(8_192)?),
    ];
    for (name, archive) in corpora {
        let reader = OwnedPhysPkgReader::from_bytes(archive.clone())?;
        let names = reader.member_names()?;
        let sizes = compressed_sizes(&archive)?;
        println!("archive {name} bytes={} sha256={} members={}", archive.len(), hex(&archive), names.len());
        let mut compressed_total = 0;
        for (member, size) in names.iter().zip(&sizes) {
            let payload = reader.read_member(member)?;
            compressed_total += size;
            println!("member {name} {member} uncompressed={} sha256={}", payload.len(), hex(&payload));
            println!("packed {name} {member} compressed={size}");
        }
        println!("packed-total {name} compressed={compressed_total}");
    }
    Ok(())
}

fn split(seed: u64) -> Res<()> {
    println!("docx-large seed={seed} sha256={}", hex(&build_docx(131_072, seed)?));
    println!("xlsx-large sha256={}", hex(&build_xlsx(131_072)?));
    println!("pptx-large sha256={}", hex(&build_pptx(8_192)?));
    Ok(())
}

fn main() -> Res<()> {
    let args: Vec<String> = std::env::args().collect();
    match args.get(1).map(String::as_str) {
        Some("digest") => digest(),
        Some("split") => split(args.get(2).ok_or("seed")?.parse()?),
        _ => Err("usage: member_probe digest | split SEED".into()),
    }
}
