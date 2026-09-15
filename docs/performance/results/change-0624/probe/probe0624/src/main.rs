//! Scratch probe for change 0624: the gate measurement for parallel deflate of
//! the changed member set of an OOXML save (survey item CORE-4 / SAVE-6).
//!
//! The gate has two halves, and this probe measures both:
//!
//!   1. *shape*   — how many members a realistic edit regenerates, and how big
//!                  they are. `census` drives the real editor routes and the
//!                  `OpcPackage` publication route and reports, per fixture,
//!                  the regenerated member set with uncompressed sizes. A
//!                  member is "regenerated" when its compressed payload in the
//!                  published archive differs from the source archive's, which
//!                  is exactly the preservation writer's `Regenerate` action.
//!   2. *size*    — what share of a save's wall time the deflate of that
//!                  changed set is (`savesplit`), and what parallel deflate
//!                  could win on such a set at widths 1/2/4/8 (`deflatescale`).
//!
//! `deflatescale` is the threshold measurement the frozen design needs: it
//! runs the *same* flate2 level-6 codec the writers run, serially and through
//! a caller-owned rayon pool, over member sets of a stated count and size.
//!
//! Subcommands:
//!   census <root>                                  corpus census, one line per fixture/scenario
//!   censusone <file> <scenario>                    one fixture, listing every changed member
//!   dump <file> <scenario> <outdir>                write the changed members' *uncompressed*
//!                                                  payloads to <outdir> (input for deflatescale)
//!   savesplit <file> <scenario> <warmup> <samples> publish ns and deflate-only ns per save
//!   deflatescale <dir> <warmup> <samples> <w..>    serial vs pooled deflate of the payloads in <dir>
//!   synthscale <count> <bytes> <kind> <warmup> <samples> <w..>
//!                                                  the same on synthetic XML-like payloads
//!   synthwb <sheets> <rows> <cols> <outfile>          rebuild a `tools/perf-baseline` XLSX corpus shape
//!   pptxcreate <slides> <outfile>                  author a deck and write it (all members regenerated)
//!
//! Scenarios: noop | addrel | addrelN:<k> | addrelall | xlsxcells:<k> | xlsxpercent:<p> | xlsxspread:<p>

use std::collections::BTreeMap;
use std::env;
use std::io::Write as _;
use std::path::{Path, PathBuf};
use std::time::Instant;

use flate2::Compression;
use flate2::write::DeflateEncoder;
use litchi_opc::package::OpcPackage;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::Part;
use litchi_opc::pkgwriter::PackageWriter;

const EXTENSIONS: &[&str] = &[
    "xlsx", "xlsm", "xltx", "xltm", "docx", "docm", "dotx", "dotm", "pptx", "pptm", "potx", "ppsx",
    "xlsb",
];

// ---------------------------------------------------------------------------
// Minimal ZIP central-directory reader.
//
// The probe only ever reads archives litchi itself published (and the fixtures
// those were published from), so this handles exactly the shapes those carry:
// a central directory located from the EOCD, with ZIP64 promotion read from the
// ZIP64 EOCD when the 32-bit fields are saturated.
// ---------------------------------------------------------------------------

#[derive(Clone)]
struct Member {
    name: String,
    method: u16,
    uncompressed: u64,
    compressed: u64,
    crc: u32,
    local_offset: u64,
}

fn u16_at(data: &[u8], at: usize) -> u16 {
    u16::from_le_bytes([data[at], data[at + 1]])
}

fn u32_at(data: &[u8], at: usize) -> u32 {
    u32::from_le_bytes([data[at], data[at + 1], data[at + 2], data[at + 3]])
}

fn u64_at(data: &[u8], at: usize) -> u64 {
    let mut buffer = [0u8; 8];
    buffer.copy_from_slice(&data[at..at + 8]);
    u64::from_le_bytes(buffer)
}

fn read_central_directory(data: &[u8]) -> Result<Vec<Member>, String> {
    let eocd = (0..data.len().saturating_sub(21))
        .rev()
        .find(|&at| data[at..].starts_with(&[0x50, 0x4b, 0x05, 0x06]))
        .ok_or_else(|| "no EOCD".to_owned())?;
    let mut entries = u16_at(data, eocd + 10) as u64;
    let mut directory = u32_at(data, eocd + 16) as u64;
    if entries == 0xffff || directory == 0xffff_ffff {
        let locator = (0..eocd)
            .rev()
            .find(|&at| data[at..].starts_with(&[0x50, 0x4b, 0x06, 0x07]))
            .ok_or_else(|| "no ZIP64 locator".to_owned())?;
        let zip64_eocd = u64_at(data, locator + 8) as usize;
        entries = u64_at(data, zip64_eocd + 32);
        directory = u64_at(data, zip64_eocd + 48);
    }
    let mut at = usize::try_from(directory).map_err(|_| "directory offset".to_owned())?;
    let mut members = Vec::with_capacity(entries as usize);
    for _ in 0..entries {
        if !data[at..].starts_with(&[0x50, 0x4b, 0x01, 0x02]) {
            return Err(format!("bad central record at {at}"));
        }
        let crc = u32_at(data, at + 16);
        let mut compressed = u32_at(data, at + 20) as u64;
        let mut uncompressed = u32_at(data, at + 24) as u64;
        let name_len = u16_at(data, at + 28) as usize;
        let extra_len = u16_at(data, at + 30) as usize;
        let comment_len = u16_at(data, at + 32) as usize;
        let mut local_offset = u32_at(data, at + 42) as u64;
        let name = String::from_utf8_lossy(&data[at + 46..at + 46 + name_len]).into_owned();
        let extra = &data[at + 46 + name_len..at + 46 + name_len + extra_len];
        // ZIP64 extended information, in the fixed order the spec mandates.
        let mut cursor = 0usize;
        while cursor + 4 <= extra.len() {
            let tag = u16_at(extra, cursor);
            let size = u16_at(extra, cursor + 2) as usize;
            if tag == 0x0001 {
                let field = &extra[cursor + 4..(cursor + 4 + size).min(extra.len())];
                let mut offset = 0usize;
                if uncompressed == 0xffff_ffff && offset + 8 <= field.len() {
                    uncompressed = u64_at(field, offset);
                    offset += 8;
                }
                if compressed == 0xffff_ffff && offset + 8 <= field.len() {
                    compressed = u64_at(field, offset);
                    offset += 8;
                }
                if local_offset == 0xffff_ffff && offset + 8 <= field.len() {
                    local_offset = u64_at(field, offset);
                }
            }
            cursor += 4 + size;
        }
        members.push(Member {
            name,
            method: u16_at(data, at + 10),
            uncompressed,
            compressed,
            crc,
            local_offset,
        });
        at += 46 + name_len + extra_len + comment_len;
    }
    Ok(members)
}

/// The compressed payload bytes of one member, located from its local header.
fn payload<'a>(data: &'a [u8], member: &Member) -> Result<&'a [u8], String> {
    let at = usize::try_from(member.local_offset).map_err(|_| "local offset".to_owned())?;
    if !data[at..].starts_with(&[0x50, 0x4b, 0x03, 0x04]) {
        return Err(format!("bad local header for {}", member.name));
    }
    let name_len = u16_at(data, at + 26) as usize;
    let extra_len = u16_at(data, at + 28) as usize;
    let start = at + 30 + name_len + extra_len;
    let end = start
        .checked_add(usize::try_from(member.compressed).map_err(|_| "compressed".to_owned())?)
        .ok_or_else(|| "payload end".to_owned())?;
    data.get(start..end)
        .ok_or_else(|| format!("payload out of range for {}", member.name))
}

/// Inflate a Deflate member (or return a stored member verbatim).
fn inflate(member: &Member, raw: &[u8]) -> Result<Vec<u8>, String> {
    match member.method {
        0 => Ok(raw.to_vec()),
        8 => {
            let mut out = Vec::with_capacity(member.uncompressed as usize);
            let mut decoder = flate2::write::DeflateDecoder::new(&mut out);
            decoder
                .write_all(raw)
                .and_then(|()| decoder.finish().map(|_| ()))
                .map_err(|error| format!("inflate {}: {error}", member.name))?;
            Ok(out)
        },
        other => Err(format!("method {other} for {}", member.name)),
    }
}

/// Members of `after` whose compressed payload differs from `before`'s, i.e.
/// the members the writer regenerated rather than copied through.
fn changed(before: &[u8], after: &[u8]) -> Result<Vec<(Member, Vec<u8>)>, String> {
    let source: BTreeMap<String, Member> = read_central_directory(before)?
        .into_iter()
        .map(|member| (member.name.clone(), member))
        .collect();
    let mut out = Vec::new();
    for member in read_central_directory(after)? {
        let raw = payload(after, &member)?;
        let same = source.get(&member.name).is_some_and(|original| {
            original.crc == member.crc
                && original.compressed == member.compressed
                && original.method == member.method
                && payload(before, original).is_ok_and(|bytes| bytes == raw)
        });
        if !same {
            let plain = inflate(&member, raw)?;
            out.push((member, plain));
        }
    }
    Ok(out)
}

// ---------------------------------------------------------------------------
// Scenarios
// ---------------------------------------------------------------------------

/// Publish one scenario, returning (source bytes, published bytes).
fn run_scenario(file: &Path, scenario: &str) -> Result<(Vec<u8>, Vec<u8>), String> {
    let source = std::fs::read(file).map_err(|error| format!("read: {error}"))?;
    let published = if let Some(rest) = scenario.strip_prefix("xlsxcells:") {
        let count: usize = rest.parse().map_err(|_| "cell count".to_owned())?;
        xlsx_edit(file, CellBudget::Fixed(count))?
    } else if let Some(rest) = scenario.strip_prefix("xlsxpercent:") {
        let percent: f64 = rest.parse().map_err(|_| "percent".to_owned())?;
        xlsx_edit(file, CellBudget::Percent(percent))?
    } else if let Some(rest) = scenario.strip_prefix("xlsxspread:") {
        let percent: f64 = rest.parse().map_err(|_| "percent".to_owned())?;
        xlsx_spread_edit(file, percent)?
    } else {
        let mut package = OpcPackage::open(file).map_err(|error| format!("open: {error:?}"))?;
        opc_mutate(&mut package, scenario);
        PackageWriter::to_bytes(&package).map_err(|error| format!("publish: {error:?}"))?
    };
    Ok((source, published))
}

fn opc_mutate(package: &mut OpcPackage, scenario: &str) -> usize {
    match scenario {
        "noop" => 0,
        "addrel" => add_relationships(package, 1),
        "addrelall" => add_relationships(package, usize::MAX),
        other => match other.strip_prefix("addrelN:") {
            Some(count) => add_relationships(package, count.parse().expect("addrelN count")),
            None => panic!("unknown scenario {other}"),
        },
    }
}

fn add_relationships(package: &mut OpcPackage, limit: usize) -> usize {
    let mut names: Vec<PackURI> = package
        .iter_parts()
        .filter(|part| !part.rels().is_empty())
        .map(Part::partname)
        .cloned()
        .collect();
    names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut touched = 0;
    for (index, partname) in names.iter().take(limit).enumerate() {
        if let Ok(part) = package.get_part_mut(partname) {
            let _relationship = part.rels_mut().add_relationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                    .to_owned(),
                "https://example.invalid/probe".to_owned(),
                format!("rIdProbe0624x{index}"),
                true,
            );
            touched += 1;
        }
    }
    touched
}

enum CellBudget {
    Fixed(usize),
    Percent(f64),
}

/// `tools/perf-baseline`'s XLSX corpus, rebuilt here so this probe can census
/// the harness shapes without building the harness. `build_xlsx_corpus`
/// (`tools/perf-baseline/src/lib.rs:20590`) authors `sheet_count` sheets of
/// `row_count` x `column_count` deterministic integers; `XlsxShape`
/// (`:795-829`) gives tiny = 3x8x8, medium = 4x32x32 and dense-wide = 2x256x256.
fn synth_workbook(sheets: usize, rows: usize, columns: usize) -> Result<Vec<u8>, String> {
    let workbook =
        litchi_xlsx::Workbook::new().map_err(|error| format!("new workbook: {error:?}"))?;
    let mut edit = workbook
        .edit()
        .map_err(|error| format!("edit: {error:?}"))?;
    for sheet in 0..sheets {
        let name = if sheet == 0 {
            "Sheet1".to_owned()
        } else {
            format!("Bench{sheet:02}")
        };
        if sheet == 0 {
            let mut target = edit
                .sheet(name.as_str())
                .map_err(|error| format!("edit sheet: {error:?}"))?
                .ok_or_else(|| format!("sheet {name} missing"))?;
            fill(
                &mut |reference: &str, value: i32| {
                    target
                        .set(reference, value)
                        .map(|_| ())
                        .map_err(|error| format!("set {reference}: {error:?}"))
                },
                rows,
                columns,
            )?;
        } else {
            let mut target = edit
                .add(name.as_str())
                .map_err(|error| format!("add sheet: {error:?}"))?;
            fill(
                &mut |reference: &str, value: i32| {
                    target
                        .set(reference, value)
                        .map(|_| ())
                        .map_err(|error| format!("set {reference}: {error:?}"))
                },
                rows,
                columns,
            )?;
        }
    }
    let commit = edit
        .commit()
        .map_err(|error| format!("commit: {error:?}"))?;
    commit
        .workbook()
        .to_bytes()
        .map_err(|error| format!("to_bytes: {error:?}"))
}

/// The post-commit workbook of an `xlsx*` scenario, so a save can be timed on
/// its own without the open, the edit or the commit in the interval.
fn committed_workbook(file: &Path, scenario: &str) -> Result<litchi_xlsx::Workbook, String> {
    if let Some(rest) = scenario.strip_prefix("xlsxspread:") {
        let percent: f64 = rest.parse().map_err(|_| "percent".to_owned())?;
        return xlsx_spread_commit(file, percent);
    }
    let budget = if let Some(rest) = scenario.strip_prefix("xlsxcells:") {
        CellBudget::Fixed(rest.parse().map_err(|_| "cell count".to_owned())?)
    } else if let Some(rest) = scenario.strip_prefix("xlsxpercent:") {
        CellBudget::Percent(rest.parse().map_err(|_| "percent".to_owned())?)
    } else {
        return Err(format!("scenario {scenario} has no workbook route"));
    };
    xlsx_commit(file, budget)
}

/// Write a dense `rows` x `columns` integer grid through one cell setter.
fn fill(
    set: &mut dyn FnMut(&str, i32) -> Result<(), String>,
    rows: usize,
    columns: usize,
) -> Result<(), String> {
    for row in 0..rows {
        for column in 0..columns {
            let reference = cell_reference(row, column);
            let value = i32::try_from((row * columns + column) % 1_000_000)
                .map_err(|_| "cell value".to_owned())?;
            set(reference.as_str(), value)?;
        }
    }
    Ok(())
}

/// A1-style reference for a zero-based row and column.
fn cell_reference(row: usize, column: usize) -> String {
    let mut letters = Vec::new();
    let mut index = column;
    loop {
        letters.push(b'A' + (index % 26) as u8);
        if index < 26 {
            break;
        }
        index = index / 26 - 1;
    }
    letters.reverse();
    format!("{}{}", String::from_utf8(letters).expect("ascii"), row + 1)
}

/// The harness's one-percent update distribution: `ceil(total/100)` cells laid
/// out evenly across *every* sheet, exactly as `xlsx_one_percent_updates`
/// (`tools/perf-baseline/src/lib.rs:21758`) computes them. This is the only
/// realistic scenario in the program that touches more than one worksheet.
fn xlsx_spread_edit(file: &Path, percent: f64) -> Result<Vec<u8>, String> {
    xlsx_spread_commit(file, percent)?
        .to_bytes()
        .map_err(|error| format!("to_bytes: {error:?}"))
}

fn xlsx_spread_commit(file: &Path, percent: f64) -> Result<litchi_xlsx::Workbook, String> {
    let workbook =
        litchi_xlsx::Workbook::open(file).map_err(|error| format!("xlsx open: {error:?}"))?;
    let names: Vec<String> = workbook
        .sheets()
        .map(|sheet| sheet.name().to_owned())
        .collect();
    let mut counts = Vec::new();
    for name in &names {
        let count = workbook
            .sheet(name.as_str())
            .map_err(|error| format!("sheet: {error:?}"))?
            .ok_or_else(|| "sheet missing".to_owned())?
            .cells(litchi_sheet::Rect::ALL)
            .map_err(|error| format!("cells: {error:?}"))?
            .count();
        counts.push(count);
    }
    let total: usize = counts.iter().sum();
    let updates = (((total as f64) * percent / 100.0).ceil() as usize).max(1);
    let mut edit = workbook
        .edit()
        .map_err(|error| format!("edit: {error:?}"))?;
    let mut applied = 0usize;
    for (sheet_index, name) in names.iter().enumerate() {
        let count = counts[sheet_index];
        if count == 0 {
            continue;
        }
        // Each sheet gets its share of the budget, laid out row-major from A1
        // over a 16-column lattice, as `xlsx_edit` does for the single-sheet
        // scenarios.
        let share = ((updates * count) / total.max(1)).max(1);
        let mut target = edit
            .sheet(name.as_str())
            .map_err(|error| format!("edit sheet: {error:?}"))?
            .ok_or_else(|| "edit sheet missing".to_owned())?;
        for index in 0..share {
            let reference = cell_reference(index / 16, index % 16);
            target
                .set(reference.as_str(), format!("p0624-{sheet_index}-{index}"))
                .map_err(|error| format!("set {reference}: {error:?}"))?;
            applied += 1;
        }
    }
    if applied == 0 {
        return Err("no cell was updated".to_owned());
    }
    let commit = edit
        .commit()
        .map_err(|error| format!("commit: {error:?}"))?;
    Ok(commit.workbook().clone())
}

/// The real `litchi-xlsx` value-editor route: set `k` cells on the first sheet
/// and publish. `Percent` sizes `k` from the sheet's own used-cell count.
fn xlsx_edit(file: &Path, budget: CellBudget) -> Result<Vec<u8>, String> {
    xlsx_commit(file, budget)?
        .to_bytes()
        .map_err(|error| format!("to_bytes: {error:?}"))
}

fn xlsx_commit(file: &Path, budget: CellBudget) -> Result<litchi_xlsx::Workbook, String> {
    let workbook =
        litchi_xlsx::Workbook::open(file).map_err(|error| format!("xlsx open: {error:?}"))?;
    let sheet_name = workbook
        .sheets()
        .next()
        .map(|sheet| sheet.name().to_owned())
        .ok_or_else(|| "no sheet".to_owned())?;
    let used = workbook
        .sheet(sheet_name.as_str())
        .map_err(|error| format!("sheet: {error:?}"))?
        .ok_or_else(|| "sheet missing".to_owned())?
        .cells(litchi_sheet::Rect::ALL)
        .map_err(|error| format!("cells: {error:?}"))?
        .count();
    let count = match budget {
        CellBudget::Fixed(count) => count,
        CellBudget::Percent(percent) => (((used as f64) * percent / 100.0).ceil() as usize).max(1),
    };
    let mut edit = workbook
        .edit()
        .map_err(|error| format!("edit: {error:?}"))?;
    {
        let mut sheet = edit
            .sheet(sheet_name.as_str())
            .map_err(|error| format!("edit sheet: {error:?}"))?
            .ok_or_else(|| "edit sheet missing".to_owned())?;
        for index in 0..count {
            let row = index / 16 + 1;
            let column = index % 16;
            let reference = format!("{}{}", (b'A' + column as u8) as char, row);
            sheet
                .set(reference.as_str(), format!("p0624-{index}"))
                .map_err(|error| format!("set {reference}: {error:?}"))?;
        }
    }
    let commit = edit
        .commit()
        .map_err(|error| format!("commit: {error:?}"))?;
    Ok(commit.workbook().clone())
}

// ---------------------------------------------------------------------------
// Deflate timing
// ---------------------------------------------------------------------------

const THRESHOLD: u64 = 256 * 1024;

fn deflate_one(payload: &[u8]) -> usize {
    let mut encoder = DeflateEncoder::new(Vec::new(), Compression::default());
    encoder.write_all(payload).expect("deflate");
    encoder.finish().expect("finish").len()
}

fn deflate_serial(payloads: &[Vec<u8>]) -> usize {
    payloads.iter().map(|payload| deflate_one(payload)).sum()
}

fn deflate_pooled(pool: &rayon::ThreadPool, payloads: &[Vec<u8>]) -> usize {
    use rayon::prelude::*;
    pool.install(|| {
        payloads
            .par_iter()
            .map(|payload| deflate_one(payload))
            .sum()
    })
}

fn percentile(sorted: &[u128], fraction: f64) -> u128 {
    if sorted.is_empty() {
        return 0;
    }
    let at = ((sorted.len() as f64 - 1.0) * fraction).round() as usize;
    sorted[at.min(sorted.len() - 1)]
}

fn report(label: &str, mut samples: Vec<u128>) -> u128 {
    samples.sort_unstable();
    let mean = samples.iter().sum::<u128>() / samples.len().max(1) as u128;
    let p50 = percentile(&samples, 0.50);
    println!(
        "{label}\tn={}\tp50={p50}\tmean={mean}\tp95={}\tp99={}",
        samples.len(),
        percentile(&samples, 0.95),
        percentile(&samples, 0.99),
    );
    p50
}

fn load_payloads(dir: &Path) -> Vec<Vec<u8>> {
    let mut files: Vec<PathBuf> = std::fs::read_dir(dir)
        .expect("payload dir")
        .flatten()
        .map(|entry| entry.path())
        .filter(|path| path.is_file())
        .collect();
    files.sort();
    files
        .into_iter()
        .map(|path| std::fs::read(path).expect("payload"))
        .collect()
}

/// Synthetic payloads shaped like the members a save actually deflates.
fn synth(count: usize, bytes: usize, kind: &str) -> Vec<Vec<u8>> {
    (0..count)
        .map(|index| {
            let mut out = Vec::with_capacity(bytes + 64);
            match kind {
                "xml" => {
                    out.extend_from_slice(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
                    let mut row = 0u64;
                    while out.len() < bytes {
                        row += 1;
                        let _ = write!(
                            out,
                            "<row r=\"{row}\" spans=\"1:6\"><c r=\"A{row}\" t=\"inlineStr\"><is><t>member {index} row {row} value</t></is></c><c r=\"B{row}\"><v>{}</v></c></row>",
                            row.wrapping_mul(2_654_435_761) % 1_000_000
                        );
                    }
                    out.extend_from_slice(b"</sheetData></worksheet>");
                },
                "random" => {
                    // Incompressible: the deflate upper bound per byte.
                    let mut state = 0x243f_6a88_85a3_08d3u64 ^ (index as u64);
                    while out.len() < bytes {
                        state ^= state << 13;
                        state ^= state >> 7;
                        state ^= state << 17;
                        out.extend_from_slice(&state.to_le_bytes());
                    }
                    out.truncate(bytes);
                },
                other => panic!("unknown payload kind {other}"),
            }
            out
        })
        .collect()
}

fn scale(payloads: &[Vec<u8>], warmup: u32, samples: u32, widths: &[usize]) {
    let total: usize = payloads.iter().map(Vec::len).sum();
    let over = payloads
        .iter()
        .filter(|p| p.len() as u64 >= THRESHOLD)
        .count();
    println!(
        "set\tmembers={}\ttotal_bytes={total}\tmembers_ge_256KiB={over}\tmax_bytes={}",
        payloads.len(),
        payloads.iter().map(Vec::len).max().unwrap_or(0),
    );
    let mut sink = 0usize;
    for _ in 0..warmup {
        sink = sink.wrapping_add(deflate_serial(payloads));
    }
    let mut serial = Vec::with_capacity(samples as usize);
    for _ in 0..samples {
        let start = Instant::now();
        sink = sink.wrapping_add(deflate_serial(payloads));
        serial.push(start.elapsed().as_nanos());
    }
    let serial_p50 = report("serial", serial);
    for &width in widths {
        let pool = rayon::ThreadPoolBuilder::new()
            .num_threads(width)
            .build()
            .expect("pool");
        for _ in 0..warmup {
            sink = sink.wrapping_add(deflate_pooled(&pool, payloads));
        }
        let mut timings = Vec::with_capacity(samples as usize);
        for _ in 0..samples {
            let start = Instant::now();
            sink = sink.wrapping_add(deflate_pooled(&pool, payloads));
            timings.push(start.elapsed().as_nanos());
        }
        let p50 = report(&format!("width{width}"), timings);
        let speedup = serial_p50 as f64 / p50.max(1) as f64;
        println!(
            "scaling\twidth={width}\tspeedup={speedup:.3}\tefficiency={:.3}",
            speedup / width as f64
        );
    }
    eprintln!("sink {sink}");
}

// ---------------------------------------------------------------------------

fn collect(root: &Path, files: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(root) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect(&path, files);
        } else if path
            .extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| EXTENSIONS.contains(&extension.to_ascii_lowercase().as_str()))
        {
            files.push(path);
        }
    }
}

fn census_line(file: &Path, scenario: &str) -> String {
    match run_scenario(file, scenario) {
        Err(error) => format!("error\t{error}"),
        Ok((source, published)) => match changed(&source, &published) {
            Err(error) => format!("diff-error\t{error}"),
            Ok(members) => {
                let members_total = read_central_directory(&published)
                    .map(|entries| entries.len())
                    .unwrap_or(0);
                let bytes: u64 = members.iter().map(|(member, _)| member.uncompressed).sum();
                let max = members
                    .iter()
                    .map(|(member, _)| member.uncompressed)
                    .max()
                    .unwrap_or(0);
                let over = members
                    .iter()
                    .filter(|(member, _)| member.uncompressed >= THRESHOLD)
                    .count();
                format!(
                    "ok\tmembers={members_total}\tchanged={}\tchanged_bytes={bytes}\tmax_changed={max}\tchanged_ge_256KiB={over}",
                    members.len()
                )
            },
        },
    }
}

fn main() {
    let arguments: Vec<String> = env::args().skip(1).collect();
    let usage = "usage: probe0624 census <root> <scenario..> | censusone <file> <scenario> | dump <file> <scenario> <outdir> | savesplit <file> <scenario> <warmup> <samples> | deflatescale <dir> <warmup> <samples> <width..> | synthscale <count> <bytes> <kind> <warmup> <samples> <width..> | pptxcreate <slides> <outfile>";
    match arguments.first().map(String::as_str) {
        Some("census") => {
            let root = arguments.get(1).expect(usage);
            let scenarios = &arguments[2..];
            let mut files = Vec::new();
            collect(Path::new(root), &mut files);
            files.sort();
            for file in files {
                for scenario in scenarios {
                    println!(
                        "{}\t{scenario}\t{}",
                        file.display(),
                        census_line(&file, scenario)
                    );
                }
            }
        },
        Some("censusone") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let scenario = arguments.get(2).expect(usage);
            let (source, published) = run_scenario(&file, scenario).expect("scenario");
            let members = changed(&source, &published).expect("diff");
            println!("published_bytes\t{}", published.len());
            println!(
                "published_members\t{}",
                read_central_directory(&published).expect("cd").len()
            );
            println!("changed_members\t{}", members.len());
            for (member, plain) in &members {
                println!(
                    "changed\t{}\tuncompressed={}\tcompressed={}\tmethod={}\tinflated={}",
                    member.name,
                    member.uncompressed,
                    member.compressed,
                    member.method,
                    plain.len()
                );
            }
        },
        Some("dump") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let scenario = arguments.get(2).expect(usage);
            let outdir = PathBuf::from(arguments.get(3).expect(usage));
            let (source, published) = run_scenario(&file, scenario).expect("scenario");
            let members = changed(&source, &published).expect("diff");
            std::fs::create_dir_all(&outdir).expect("outdir");
            for (index, (member, plain)) in members.iter().enumerate() {
                let safe = member.name.replace(['/', '\\'], "_");
                let path = outdir.join(format!("{index:04}-{safe}"));
                std::fs::write(&path, plain).expect("write payload");
                println!("{}\t{}", path.display(), plain.len());
            }
        },
        Some("savesplit") => {
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let scenario = arguments.get(2).expect(usage);
            let warmup: u32 = arguments.get(3).expect(usage).parse().expect("warmup");
            let samples: u32 = arguments.get(4).expect(usage).parse().expect("samples");
            // The changed set, established once from a first publication.
            let (source, published) = run_scenario(&file, scenario).expect("scenario");
            let members = changed(&source, &published).expect("diff");
            let payloads: Vec<Vec<u8>> = members.iter().map(|(_, plain)| plain.clone()).collect();
            println!(
                "changed_members\t{}\tchanged_bytes\t{}\tge_256KiB\t{}",
                payloads.len(),
                payloads.iter().map(Vec::len).sum::<usize>(),
                payloads
                    .iter()
                    .filter(|p| p.len() as u64 >= THRESHOLD)
                    .count()
            );
            // Publish timing. Both routes mutate once and publish repeatedly,
            // as change 0618's probe does, so the timed interval is the save
            // and nothing else.
            let mut sink = 0u64;
            let opc = !scenario.starts_with("xlsx");
            if !opc {
                let workbook = committed_workbook(&file, scenario).expect("commit");
                for _ in 0..warmup {
                    sink += workbook.to_bytes().expect("publish").len() as u64;
                }
                let mut timings = Vec::with_capacity(samples as usize);
                for _ in 0..samples {
                    let start = Instant::now();
                    let bytes = workbook.to_bytes().expect("publish");
                    timings.push(start.elapsed().as_nanos());
                    sink += bytes.len() as u64;
                }
                report("publish", timings);
            } else if opc {
                let mut package = OpcPackage::open(&file).expect("open");
                opc_mutate(&mut package, scenario);
                for _ in 0..warmup {
                    sink += PackageWriter::to_bytes(&package).expect("publish").len() as u64;
                }
                let mut timings = Vec::with_capacity(samples as usize);
                for _ in 0..samples {
                    let start = Instant::now();
                    let bytes = PackageWriter::to_bytes(&package).expect("publish");
                    timings.push(start.elapsed().as_nanos());
                    sink += bytes.len() as u64;
                }
                report("publish", timings);
            } else {
                for _ in 0..warmup {
                    sink += run_scenario(&file, scenario).expect("scenario").1.len() as u64;
                }
                let mut timings = Vec::with_capacity(samples as usize);
                for _ in 0..samples {
                    let start = Instant::now();
                    let bytes = run_scenario(&file, scenario).expect("scenario").1;
                    timings.push(start.elapsed().as_nanos());
                    sink += bytes.len() as u64;
                }
                report("operation", timings);
            }
            // Deflate-only timing over the same changed set.
            for _ in 0..warmup {
                sink += deflate_serial(&payloads) as u64;
            }
            let mut timings = Vec::with_capacity(samples as usize);
            for _ in 0..samples {
                let start = Instant::now();
                sink += deflate_serial(&payloads) as u64;
                timings.push(start.elapsed().as_nanos());
            }
            report("deflate", timings);
            eprintln!("sink {sink}");
        },
        Some("deflatescale") => {
            let dir = PathBuf::from(arguments.get(1).expect(usage));
            let warmup: u32 = arguments.get(2).expect(usage).parse().expect("warmup");
            let samples: u32 = arguments.get(3).expect(usage).parse().expect("samples");
            let widths: Vec<usize> = arguments[4..]
                .iter()
                .map(|width| width.parse().expect("width"))
                .collect();
            scale(&load_payloads(&dir), warmup, samples, &widths);
        },
        Some("synthscale") => {
            let count: usize = arguments.get(1).expect(usage).parse().expect("count");
            let bytes: usize = arguments.get(2).expect(usage).parse().expect("bytes");
            let kind = arguments.get(3).expect(usage);
            let warmup: u32 = arguments.get(4).expect(usage).parse().expect("warmup");
            let samples: u32 = arguments.get(5).expect(usage).parse().expect("samples");
            let widths: Vec<usize> = arguments[6..]
                .iter()
                .map(|width| width.parse().expect("width"))
                .collect();
            scale(&synth(count, bytes, kind), warmup, samples, &widths);
        },
        Some("pptxbench") => {
            // The authored save through `StreamingArchiveWriter`: build the
            // deck once, then publish it `iterations` times. Every member is
            // regenerated on this path (change 0607: an authored package has
            // no source archive, so the plan has no `Copy` action available).
            let slides: usize = arguments.get(1).expect(usage).parse().expect("slides");
            let iterations: u32 = arguments.get(2).expect(usage).parse().expect("iterations");
            let mut package = authored_package(slides);
            let mut sink = 0u64;
            for _ in 0..iterations {
                sink += package.to_bytes().expect("to_bytes").len() as u64;
            }
            println!("{sink}");
        },
        Some("pptxsplit") => {
            // Publish wall time against the deflate of the same member set.
            let slides: usize = arguments.get(1).expect(usage).parse().expect("slides");
            let warmup: u32 = arguments.get(2).expect(usage).parse().expect("warmup");
            let samples: u32 = arguments.get(3).expect(usage).parse().expect("samples");
            let mut package = authored_package(slides);
            let bytes = package.to_bytes().expect("to_bytes");
            let members = read_central_directory(&bytes).expect("cd");
            let payloads: Vec<Vec<u8>> = members
                .iter()
                .map(|member| {
                    inflate(member, payload(&bytes, member).expect("payload")).expect("inflate")
                })
                .collect();
            println!(
                "changed_members\t{}\tchanged_bytes\t{}\tge_256KiB\t{}",
                payloads.len(),
                payloads.iter().map(Vec::len).sum::<usize>(),
                payloads
                    .iter()
                    .filter(|p| p.len() >= THRESHOLD as usize)
                    .count()
            );
            let mut sink = 0u64;
            for _ in 0..warmup {
                sink += package.to_bytes().expect("to_bytes").len() as u64;
            }
            let mut timings = Vec::with_capacity(samples as usize);
            for _ in 0..samples {
                let start = Instant::now();
                let published = package.to_bytes().expect("to_bytes");
                timings.push(start.elapsed().as_nanos());
                sink += published.len() as u64;
            }
            report("publish", timings);
            for _ in 0..warmup {
                sink += deflate_serial(&payloads) as u64;
            }
            let mut timings = Vec::with_capacity(samples as usize);
            for _ in 0..samples {
                let start = Instant::now();
                sink += deflate_serial(&payloads) as u64;
                timings.push(start.elapsed().as_nanos());
            }
            report("deflate", timings);
            eprintln!("sink {sink}");
        },
        Some("savebench") => {
            // Publish only, N times, from an already-mutated package. This is
            // the shape change 0618's probe uses for callgrind and `perf`
            // isolation pairs: profile N and N+M and difference the totals.
            let file = PathBuf::from(arguments.get(1).expect(usage));
            let scenario = arguments.get(2).expect(usage);
            let iterations: u32 = arguments.get(3).expect(usage).parse().expect("iterations");
            let mut sink = 0u64;
            if scenario.starts_with("xlsx") {
                let workbook = committed_workbook(&file, scenario).expect("commit");
                for _ in 0..iterations {
                    sink += workbook.to_bytes().expect("publish").len() as u64;
                }
            } else {
                let mut package = OpcPackage::open(&file).expect("open");
                opc_mutate(&mut package, scenario);
                for _ in 0..iterations {
                    sink += PackageWriter::to_bytes(&package).expect("publish").len() as u64;
                }
            }
            println!("{sink}");
        },
        Some("synthwb") => {
            let sheets: usize = arguments.get(1).expect(usage).parse().expect("sheets");
            let rows: usize = arguments.get(2).expect(usage).parse().expect("rows");
            let columns: usize = arguments.get(3).expect(usage).parse().expect("columns");
            let out = PathBuf::from(arguments.get(4).expect(usage));
            let bytes = synth_workbook(sheets, rows, columns).expect("synth workbook");
            std::fs::write(&out, &bytes).expect("write workbook");
            let members = read_central_directory(&bytes).expect("cd");
            println!("bytes\t{}", bytes.len());
            println!("members\t{}", members.len());
            for member in &members {
                println!(
                    "member\t{}\tuncompressed={}",
                    member.name, member.uncompressed
                );
            }
        },
        Some("pptxcreate") => {
            let slides: usize = arguments.get(1).expect(usage).parse().expect("slides");
            let out = PathBuf::from(arguments.get(2).expect(usage));
            let bytes = authored_deck(slides);
            std::fs::write(&out, &bytes).expect("write deck");
            println!("bytes\t{}", bytes.len());
            println!(
                "members\t{}",
                read_central_directory(&bytes).expect("cd").len()
            );
            let mut sizes: Vec<u64> = read_central_directory(&bytes)
                .expect("cd")
                .iter()
                .map(|member| member.uncompressed)
                .collect();
            sizes.sort_unstable();
            println!("total_uncompressed\t{}", sizes.iter().sum::<u64>());
            println!("max_member\t{}", sizes.last().copied().unwrap_or(0));
            println!(
                "members_ge_256KiB\t{}",
                sizes.iter().filter(|&&size| size >= THRESHOLD).count()
            );
        },
        _ => {
            eprintln!("{usage}");
            std::process::exit(2);
        },
    }
}

/// Change 0607's authored deck, unchanged, so the member shape matches the
/// figures that record and change 0618 retain.
fn authored_deck(slides: usize) -> Vec<u8> {
    authored_package(slides).to_bytes().expect("to_bytes")
}

fn authored_package(slides: usize) -> litchi_pptx::Package {
    const TITLE_WIDTH: usize = 24;
    let mut package = litchi_pptx::Package::new().expect("package");
    {
        let presentation = package.presentation_mut().expect("presentation");
        for index in 0..slides {
            let slide = presentation.add_slide().expect("slide");
            let title = format!("S{index:04} rev {:06}", 0);
            slide.set_title(&format!("{title:<TITLE_WIDTH$}"));
            for box_index in 0..3 {
                slide.add_text_box(
                    &format!(
                        "Slide {index:04} line {box_index} - a sentence of ordinary presentation body text that a real deck would carry."
                    ),
                    914_400,
                    1_828_800 + 914_400 * i64::try_from(box_index).expect("i64"),
                    7_315_200,
                    914_400,
                );
            }
            if index % 4 == 3 {
                slide.set_notes(&format!("Speaker notes for slide {index:04}."));
            }
        }
    }
    package
}
