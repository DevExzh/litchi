//! Compression-level, chunk-size and sync-flush probe for change 0762.
//!
//! Builds the harness's streaming DOCX, XLSX and PPTX corpora with the real
//! writers, extracts every member's uncompressed payload, and recompresses
//! the members with zlib-rs (through flate2, as soapberry-zip does) under the
//! batched protocol of change 0762: fixed input chunks cut at absolute member
//! offsets, one codec call per chunk with a full 32 KiB output buffer, and a
//! final `Finish` with or without the old pre-finish sync flush. One codec is
//! reused across members (`reset`, then `set_level` when the level differs),
//! as the archive writer does. Sizes are exact; times are the minimum and the
//! median of repeated passes over all members of one corpus, on one pinned
//! core. Every configuration's output is inflated and compared once.

use flate2::{Compress, Compression, Decompress, FlushCompress, FlushDecompress, Status};
use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits};
use litchi_docx::{StreamingDocumentLimits, StreamingDocumentWriter};
use litchi_opc::phys_pkg::OwnedPhysPkgReader;
use litchi_pptx::{
    StreamingPresentationLimits, StreamingPresentationOptions, StreamingPresentationWriter,
    TextBoxSpec,
};
use litchi_xlsx::{StreamingCell, StreamingCellValue, StreamingWorkbookLimits, StreamingWorkbookWriter};
use std::num::{NonZeroU64, NonZeroUsize};
use std::time::Instant;

type Res<T> = Result<T, Box<dyn std::error::Error>>;

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

fn build_docx(paragraphs: usize) -> Res<Vec<u8>> {
    let texts: Vec<String> = (0..paragraphs).map(docx_text).collect();
    let input: u64 = texts.iter().map(|t| t.len() as u64).sum();
    let max_run = texts.iter().map(String::len).max().unwrap_or(0) as u64;
    let document_xml = (max_run * 8 + 256) * paragraphs as u64 + 16 * 1024;
    let output = document_xml + 256 * 1024;
    let p = paragraphs as u64;
    let limits = StreamingDocumentLimits::new(input, output, p, p, 1, max_run, document_xml, document_xml, 64);
    let context = streaming_context(64, input, output, p * 2 + 16, p * 4 + input + 32)?;
    let mut writer = StreamingDocumentWriter::new(Vec::new(), context, limits)?;
    for text in &texts {
        writer.start_paragraph()?;
        writer.start_run()?;
        writer.write_text(text)?;
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

struct Member {
    payload: Vec<u8>,
}

fn members(archive: &[u8]) -> Res<Vec<Member>> {
    let reader = OwnedPhysPkgReader::from_bytes(archive.to_vec())?;
    let mut out = Vec::new();
    for name in reader.member_names()? {
        out.push(Member { payload: reader.read_member(&name)? });
    }
    Ok(out)
}

struct Codec {
    compress: Compress,
    level: u32,
    used: bool,
    buffer: Box<[u8]>,
}

impl Codec {
    fn new(level: u32) -> Self {
        Self {
            compress: Compress::new(Compression::new(level), false),
            level,
            used: false,
            buffer: vec![0; 32 * 1024].into_boxed_slice(),
        }
    }

    fn call(&mut self, input: &[u8], flush: FlushCompress, out: &mut Vec<u8>) -> (Status, usize, usize) {
        let (before_in, before_out) = (self.compress.total_in(), self.compress.total_out());
        let status = self.compress.compress(input, &mut self.buffer, flush).expect("compress");
        let consumed = (self.compress.total_in() - before_in) as usize;
        let produced = (self.compress.total_out() - before_out) as usize;
        out.extend_from_slice(&self.buffer[..produced]);
        (status, consumed, produced)
    }

    /// One member under the batched protocol; returns its compressed length.
    fn member(&mut self, payload: &[u8], level: u32, chunk: usize, sync: bool, out: &mut Vec<u8>) -> usize {
        if self.used {
            self.compress.reset();
        }
        self.used = true;
        if level != self.level {
            self.compress.set_level(Compression::new(level)).expect("set_level");
            self.level = level;
        }
        let start = out.len();
        for piece in payload.chunks(chunk) {
            let mut input = piece;
            while !input.is_empty() {
                let (_, consumed, produced) = self.call(input, FlushCompress::None, out);
                assert!(consumed > 0 || produced > 0, "no progress");
                input = &input[consumed..];
            }
        }
        if sync {
            self.call(&[], FlushCompress::Sync, out);
            loop {
                let (_, _, produced) = self.call(&[], FlushCompress::None, out);
                if produced == 0 {
                    break;
                }
            }
        }
        loop {
            let (status, _, _) = self.call(&[], FlushCompress::Finish, out);
            if status == Status::StreamEnd {
                break;
            }
        }
        out.len() - start
    }
}

fn inflate(data: &[u8], expected: usize) -> Vec<u8> {
    let mut decompress = Decompress::new(false);
    let mut out = Vec::with_capacity(expected + 16);
    let status = decompress
        .decompress_vec(data, &mut out, FlushDecompress::Finish)
        .expect("inflate");
    assert_eq!(status, Status::StreamEnd);
    out
}

const CLASSES: [(&str, usize, usize); 3] = [
    ("lt1KiB", 0, 1024),
    ("1-16KiB", 1024, 16 * 1024),
    ("ge16KiB", 16 * 1024, usize::MAX),
];

fn class_of(len: usize) -> usize {
    CLASSES.iter().position(|(_, lo, hi)| len >= *lo && len < *hi).unwrap_or(0)
}

fn percentile(sorted: &[f64], q: f64) -> f64 {
    let index = ((sorted.len() - 1) as f64 * q).round() as usize;
    sorted[index]
}

fn real(root: &str) -> Res<()> {
    // Every member of every real OOXML fixture under `root`, grouped by the
    // package's extension; members of packages that fail to open are skipped.
    let mut files = Vec::new();
    let mut stack = vec![std::path::PathBuf::from(root)];
    while let Some(dir) = stack.pop() {
        for entry in std::fs::read_dir(&dir)? {
            let path = entry?.path();
            if path.is_dir() {
                stack.push(path);
            } else if let Some(ext) = path.extension().and_then(|e| e.to_str()) {
                if ["docx", "xlsx", "pptx"].contains(&ext) {
                    files.push((ext.to_owned(), path));
                }
            }
        }
    }
    files.sort();
    println!("format\tlevel\tpackages\tmembers\tuncompressed\tcompressed\tmin_ms\tclass_sizes");
    for format in ["docx", "xlsx", "pptx"] {
        let mut payloads: Vec<Vec<u8>> = Vec::new();
        let mut packages = 0;
        for (ext, path) in &files {
            if ext != format {
                continue;
            }
            let bytes = std::fs::read(path)?;
            let Ok(reader) = OwnedPhysPkgReader::from_bytes(bytes) else { continue };
            let Ok(names) = reader.member_names() else { continue };
            let mut ok = Vec::new();
            let mut failed = false;
            for name in names {
                match reader.read_member(&name) {
                    Ok(payload) => ok.push(payload),
                    Err(_) => { failed = true; break; }
                }
            }
            if failed { continue; }
            packages += 1;
            payloads.extend(ok);
        }
        let uncompressed: usize = payloads.iter().map(Vec::len).sum();
        // (label, level, chunk, sync): levels 1-6 under the batched protocol,
        // and the base protocol for a member written in one call (one codec
        // call, then the pre-finish sync flush) at level 6.
        let mut configs: Vec<(String, u32, usize, bool)> =
            (1..=6u32).map(|level| (format!("{level}"), level, 16 * 1024, false)).collect();
        configs.push(("base6".to_owned(), 6, usize::MAX, true));
        for (label, level, chunk, sync) in configs {
            let mut out = Vec::with_capacity(uncompressed);
            let mut codec = Codec::new(level);
            let mut class_sizes = [(0usize, 0usize, 0usize); 3];
            let mut total = 0;
            for payload in &payloads {
                let size = codec.member(payload, level, chunk, sync, &mut out);
                let class = class_of(payload.len());
                class_sizes[class].0 += 1;
                class_sizes[class].1 += payload.len();
                class_sizes[class].2 += size;
                total += size;
            }
            let mut best = f64::MAX;
            for _ in 0..3 {
                let mut codec = Codec::new(level);
                out.clear();
                let started = Instant::now();
                for payload in &payloads {
                    std::hint::black_box(codec.member(payload, level, chunk, sync, &mut out));
                }
                best = best.min(started.elapsed().as_secs_f64() * 1e3);
            }
            let classes: Vec<String> = CLASSES
                .iter()
                .zip(class_sizes)
                .map(|((label, _, _), (count, raw, packed))| format!("{label}:{count}:{raw}:{packed}"))
                .collect();
            println!("{format}\t{label}\t{packages}\t{}\t{uncompressed}\t{total}\t{best:.3}\t{}", payloads.len(), classes.join(","));
        }
    }
    Ok(())
}

fn main() -> Res<()> {
    let args: Vec<String> = std::env::args().collect();
    if args.get(1).map(String::as_str) == Some("real") {
        return real(args.get(2).map(String::as_str).unwrap_or("."));
    }
    let only = args.get(1).cloned();
    let corpora: Vec<(&str, Box<dyn Fn() -> Res<Vec<u8>>>, usize)> = vec![
        ("docx-medium", Box::new(|| build_docx(8_192)), 30),
        ("docx-large", Box::new(|| build_docx(131_072)), 8),
        ("xlsx-medium", Box::new(|| build_xlsx(8_192)), 30),
        ("xlsx-large", Box::new(|| build_xlsx(131_072)), 8),
        ("pptx-medium", Box::new(|| build_pptx(256)), 30),
        ("pptx-large", Box::new(|| build_pptx(8_192)), 8),
    ];
    // (level, chunk, sync)
    let mut configs: Vec<(u32, usize, bool)> = (1..=6).map(|level| (level, 16 * 1024, false)).collect();
    configs.push((6, 16 * 1024, true));
    configs.push((6, 4 * 1024, false));
    configs.push((6, 64 * 1024, false));
    configs.push((6, 1 << 30, false));
    println!("corpus\tlevel\tchunk\tsync\tmembers\tuncompressed\tcompressed\tarchive_estimate\tbase_archive\tmin_ms\tmedian_ms\tclass_sizes");
    for (name, build, repeats) in corpora {
        if only.as_deref().is_some_and(|o| !name.starts_with(o)) {
            continue;
        }
        let archive = build()?;
        let members = members(&archive)?;
        // Base member data size from the base writer's archive: the
        // difference between the archive and its members' compressed sizes
        // is ZIP framing, which no level changes.
        let base_compressed: usize = {
            let mut codec_total = 0usize;
            let zip = soapberry_sizes(&archive)?;
            for size in zip {
                codec_total += size;
            }
            codec_total
        };
        let framing = archive.len() - base_compressed;
        let uncompressed: usize = members.iter().map(|m| m.payload.len()).sum();
        for &(level, chunk, sync) in &configs {
            let mut codec = Codec::new(level);
            let mut out = Vec::with_capacity(uncompressed);
            let mut sizes = vec![0usize; members.len()];
            for (index, member) in members.iter().enumerate() {
                sizes[index] = codec.member(&member.payload, level, chunk, sync, &mut out);
            }
            // Verify every member once.
            let mut offset = 0;
            for (member, size) in members.iter().zip(&sizes) {
                let inflated = inflate(&out[offset..offset + size], member.payload.len());
                assert_eq!(inflated, member.payload, "round trip");
                offset += size;
            }
            let compressed: usize = sizes.iter().sum();
            let mut class_sizes = [(0usize, 0usize, 0usize); 3];
            for (member, size) in members.iter().zip(&sizes) {
                let class = class_of(member.payload.len());
                class_sizes[class].0 += 1;
                class_sizes[class].1 += member.payload.len();
                class_sizes[class].2 += size;
            }
            let mut times = Vec::new();
            for _ in 0..repeats {
                let mut codec = Codec::new(level);
                out.clear();
                let started = Instant::now();
                for member in &members {
                    std::hint::black_box(codec.member(&member.payload, level, chunk, sync, &mut out));
                }
                times.push(started.elapsed().as_secs_f64() * 1e3);
            }
            times.sort_by(|a, b| a.partial_cmp(b).unwrap());
            let classes: Vec<String> = CLASSES
                .iter()
                .zip(class_sizes)
                .map(|((label, _, _), (count, raw, packed))| format!("{label}:{count}:{raw}:{packed}"))
                .collect();
            println!(
                "{name}\t{level}\t{chunk}\t{sync}\t{}\t{uncompressed}\t{compressed}\t{}\t{}\t{:.3}\t{:.3}\t{}",
                members.len(),
                framing + compressed,
                archive.len(),
                times[0],
                percentile(&times, 0.5),
                classes.join(",")
            );
        }
    }
    Ok(())
}

/// Compressed member sizes from the central directory of `archive`.
fn soapberry_sizes(archive: &[u8]) -> Res<Vec<usize>> {
    let reader = soapberry_zip::ZipArchive::from_slice(archive)?;
    let mut entries = reader.entries();
    let mut sizes = Vec::new();
    while let Some(record) = entries.next_entry()? {
        sizes.push(usize::try_from(record.compressed_size_hint())?);
    }
    Ok(sizes)
}
