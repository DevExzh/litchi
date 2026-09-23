//! Cross-leg mutation differential for change 0746 (scratch measurement code).
//!
//! For every XLS file named on the command line, this derives a fixed set of
//! record-level mutations of the Workbook stream (duplicated, colliding,
//! truncated, dropped, swapped and bit-flipped records inside one worksheet
//! substream, stray String records, out-of-range XF indexes, repeated Formula
//! companions and cell columns outside the BIFF8 grid), repackages each into
//! a single-stream CFB with every BoundSheet8 offset repointed, and prints the
//! outcome of four owners on it:
//!
//! - `cell_values::Snapshot::from_bytes` (exact refusal, or a digest of every
//!   editable cell's storage, style and value),
//! - `comments::Snapshot::from_bytes` and `sheet_visibility::Snapshot::from_bytes`
//!   (exact refusal, or the worksheet count),
//! - the public `Workbook::new` (exact refusal, or a digest of every decoded
//!   cell and sheet entry).
//!
//! Diffing the output of the before-leg and after-leg builds is the oracle.
//! The generator is deterministic (xorshift with a fixed seed per file).

use std::fmt::Write as _;
use std::io::Cursor;

use litchi_biff::Records;
use litchi_cfb::{OleFile, OleWriter};
use sha2::{Digest, Sha256};

type Record = (u16, Vec<u8>);

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000A;
const BOUND_SHEET: u16 = 0x0085;
const FORMULA: u16 = 0x0006;
const STRING: u16 = 0x0207;
const ARRAY: u16 = 0x0221;
const SHR_FMLA: u16 = 0x04BC;

fn hex(bytes: &[u8]) -> String {
    let mut text = String::new();
    for byte in bytes {
        let _ = write!(text, "{byte:02x}");
    }
    text
}

fn escape(text: &str) -> String {
    let mut out = String::new();
    for character in text.chars() {
        match character {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out
}

struct XorShift(u64);

impl XorShift {
    fn next(&mut self) -> u64 {
        let mut value = self.0;
        value ^= value << 13;
        value ^= value >> 7;
        value ^= value << 17;
        self.0 = value;
        value
    }

    fn below(&mut self, bound: usize) -> usize {
        if bound == 0 {
            0
        } else {
            (self.next() % bound as u64) as usize
        }
    }
}

fn is_cell(kind: u16) -> bool {
    matches!(kind, 0x0203 | 0x027E | 0x00FD | 0x0201 | 0x0205 | 0x0006 | 0x0204)
}

fn records(stream: &[u8]) -> Option<Vec<Record>> {
    let mut out = Vec::new();
    for record in Records::new(stream) {
        let record = record.ok()?;
        out.push((record.kind().get(), record.payload().to_vec()));
    }
    Some(out)
}

fn encode(records: &[Record]) -> Vec<u8> {
    let mut bytes = Vec::new();
    for (kind, payload) in records {
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&(payload.len() as u16).to_le_bytes());
        bytes.extend_from_slice(payload);
    }
    bytes
}

/// Repoints every BoundSheet8 at the n-th non-globals BOF, in order.
fn retarget(records: &mut [Record]) {
    let mut offset = 0_usize;
    let mut starts = Vec::new();
    let mut in_globals = true;
    for (kind, payload) in records.iter() {
        if *kind == BOF && !in_globals {
            starts.push(offset);
        }
        if *kind == EOF && in_globals {
            in_globals = false;
        }
        offset += 4 + payload.len();
    }
    let mut next = 0;
    for (kind, payload) in records.iter_mut() {
        if *kind == BOUND_SHEET && payload.len() >= 4 {
            if let Some(start) = starts.get(next) {
                payload[0..4].copy_from_slice(&(*start as u32).to_le_bytes());
            }
            next += 1;
        }
    }
}

fn package(stream: &[u8]) -> Option<Vec<u8>> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], stream).ok()?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).ok()?;
    Some(output.into_inner())
}

fn mutate(records: &mut Vec<Record>, random: &mut XorShift) -> &'static str {
    let globals_end = records.iter().position(|(kind, _)| *kind == EOF).unwrap_or(0);
    let cells = records
        .iter()
        .enumerate()
        .skip(globals_end + 1)
        .filter(|(_, (kind, payload))| is_cell(*kind) && payload.len() >= 6)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    let body_start = globals_end + 1;
    let body_len = records.len().saturating_sub(body_start);
    let pick = |random: &mut XorShift| cells.get(random.below(cells.len())).copied();
    let any = |random: &mut XorShift| {
        if body_len == 0 {
            None
        } else {
            let index = body_start + random.below(body_len);
            (records.get(index).is_some_and(|(kind, _)| *kind != BOF)).then_some(index)
        }
    };
    match random.below(13) {
        0 => {
            if let Some(index) = pick(random) {
                let copy = records[index].clone();
                records.insert(index + 1, copy);
            }
            "duplicate-in-place"
        },
        1 => {
            if let Some(index) = pick(random) {
                let copy = records[index].clone();
                let at = (index + 1 + random.below(8)).min(records.len().saturating_sub(1));
                records.insert(at, copy);
            }
            "duplicate-later"
        },
        2 => {
            if let (Some(from), Some(onto)) = (pick(random), pick(random)) {
                let position = records[onto].1[0..4].to_vec();
                records[from].1[0..4].copy_from_slice(&position);
            }
            "collide"
        },
        3 => {
            if let Some(index) = pick(random) {
                let column = 256 + random.below(4) as u16;
                records[index].1[2..4].copy_from_slice(&column.to_le_bytes());
            }
            "outside-grid"
        },
        4 | 5 => {
            if let Some(index) = any(random) {
                let payload = &mut records[index].1;
                if !payload.is_empty() {
                    let at = random.below(payload.len());
                    payload[at] ^= 1 << random.below(8);
                }
            }
            "bit-flip"
        },
        6 => {
            if let Some(index) = any(random) {
                let payload = &mut records[index].1;
                let cut = 1 + random.below(4);
                payload.truncate(payload.len().saturating_sub(cut));
            }
            "truncate"
        },
        7 => {
            if let Some(index) = any(random) {
                records.remove(index);
            }
            "drop"
        },
        8 => {
            if let Some(index) = any(random) {
                if index + 1 < records.len() {
                    records.swap(index, index + 1);
                }
            }
            "swap"
        },
        9 => {
            if let Some(index) = pick(random) {
                records.insert(index + 1, (STRING, vec![1, 0, 0, b'x']));
            }
            "stray-string"
        },
        10 => {
            if let Some(index) = pick(random) {
                records[index].1[4..6].copy_from_slice(&0x0fff_u16.to_le_bytes());
            }
            "xf-past-end"
        },
        11 => {
            if let Some(index) = records
                .iter()
                .enumerate()
                .skip(body_start)
                .find(|(_, (kind, _))| matches!(*kind, ARRAY | SHR_FMLA | FORMULA))
                .map(|(index, _)| index)
            {
                let copy = records[index].clone();
                records.insert(index + 1, copy);
            }
            "companion-repeat"
        },
        _ => {
            // Globals: flip one byte of one globals record after the BOF.
            if globals_end > 1 {
                let index = 1 + random.below(globals_end - 1);
                let payload = &mut records[index].1;
                if !payload.is_empty() {
                    let at = random.below(payload.len());
                    payload[at] ^= 1 << random.below(8);
                }
            }
            "globals-bit-flip"
        },
    }
}

fn cell_values_outcome(bytes: &[u8]) -> String {
    use litchi_xls::cell_values::Snapshot;
    match Snapshot::from_bytes(bytes.to_vec()) {
        Ok(snapshot) => {
            let mut hasher = Sha256::new();
            for worksheet in snapshot.worksheets() {
                hasher.update(worksheet.name().as_bytes());
                for cell in worksheet.cells() {
                    hasher.update(format!("{cell:?}").as_bytes());
                }
            }
            format!("ok worksheets={} cells={}", snapshot.worksheet_count(), hex(&hasher.finalize()))
        },
        Err(error) => format!("refused {error}"),
    }
}

fn comments_outcome(bytes: &[u8]) -> String {
    match litchi_xls::comments::Snapshot::from_bytes(bytes.to_vec()) {
        Ok(snapshot) => format!("ok worksheets={}", snapshot.worksheet_count()),
        Err(error) => format!("refused {error}"),
    }
}

fn visibility_outcome(bytes: &[u8]) -> String {
    match litchi_xls::sheet_visibility::Snapshot::from_bytes(bytes.to_vec()) {
        Ok(snapshot) => format!("ok worksheets={}", snapshot.worksheet_count()),
        Err(error) => format!("refused {error}"),
    }
}

fn reader_outcome(bytes: &[u8]) -> String {
    use litchi_core::sheet::Worksheet as _;
    match litchi_xls::Workbook::new(Cursor::new(bytes)) {
        Ok(workbook) => {
            let mut hasher = Sha256::new();
            for metadata in workbook.sheets() {
                hasher.update(format!("{metadata:?}").as_bytes());
                let Some(index) = metadata.parsed_worksheet_index() else {
                    continue;
                };
                let Ok(worksheet) = workbook.xls_worksheet(index) else {
                    continue;
                };
                let mut iterator = worksheet.cells();
                while let Some(Ok(cell)) = iterator.next() {
                    if let Some(full) = worksheet.get_cell(cell.row(), cell.column()) {
                        hasher.update(format!("{full:?}").as_bytes());
                    }
                }
            }
            format!("ok {}", hex(&hasher.finalize()))
        },
        Err(error) => format!("refused {error}"),
    }
}

fn main() {
    let rounds: usize = std::env::var("MUTATION_ROUNDS")
        .ok()
        .and_then(|value| value.parse().ok())
        .unwrap_or(40);
    for path in std::env::args().skip(1) {
        let Ok(bytes) = std::fs::read(&path) else {
            continue;
        };
        let Ok(mut ole) = OleFile::open(Cursor::new(bytes.as_slice())) else {
            continue;
        };
        let Ok(stream) = ole
            .open_stream(&["Workbook"])
            .or_else(|_| ole.open_stream(&["Book"]))
        else {
            continue;
        };
        let Some(original) = records(&stream) else {
            continue;
        };
        let mut seed = 0x0746_u64;
        for byte in path.as_bytes() {
            seed = seed.wrapping_mul(0x100_0000_01b3) ^ u64::from(*byte);
        }
        let mut random = XorShift(seed | 1);
        for round in 0..rounds {
            let mut mutated = original.clone();
            let kind = if round == 0 {
                "repackaged-unmodified"
            } else {
                mutate(&mut mutated, &mut random)
            };
            if round % 4 == 3 {
                let _ = mutate(&mut mutated, &mut random);
            }
            retarget(&mut mutated);
            let Some(package) = package(&encode(&mutated)) else {
                continue;
            };
            println!(
                "{{\"path\":\"{}\",\"round\":{round},\"mutation\":\"{kind}\",\"package_sha256\":\"{}\",\"cell_values\":\"{}\",\"comments\":\"{}\",\"visibility\":\"{}\",\"reader\":\"{}\"}}",
                escape(&path),
                hex(&Sha256::digest(&package)),
                escape(&cell_values_outcome(&package)),
                escape(&comments_outcome(&package)),
                escape(&visibility_outcome(&package)),
                escape(&reader_outcome(&package)),
            );
        }
    }
}
