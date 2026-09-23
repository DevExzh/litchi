//! 0541-style first-error matrix for the XLS `cell_values` edit owner
//! (change 0633).
//!
//! `Snapshot::from_bytes` frames the Workbook stream twice: once for the
//! private offset inventory (`parse_workbook_stream`) and once inside the
//! complete eager `Workbook::new`. The inventory runs first, so its refusals
//! shadow the eager parser's whenever a stream is malformed in a way both
//! owners can see. Any change that fuses the two passes has to keep that
//! shadowing exactly, and this matrix is the oracle for it.
//!
//! Each case rewrites one small valid package's Workbook stream with one
//! defect at one chosen position and prints the exact first typed refusal the
//! public entry point produces. Output is one JSON object per case on stdout,
//! sorted by case name. Scratch measurement code, not production code.

use std::io::Cursor;

use litchi_biff::Records;
use litchi_cfb::{OleFile, OleWriter};
use litchi_xls::Writer;
use litchi_xls::cell_values::Snapshot;

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000a;
const CODE_PAGE: u16 = 0x0042;
const BOUND_SHEET: u16 = 0x0085;
const FILE_PASS: u16 = 0x002f;
const SST: u16 = 0x00fc;
const XF: u16 = 0x00e0;
const FORMULA: u16 = 0x0006;
const STRING: u16 = 0x0207;
const NUMBER: u16 = 0x0203;
const LABEL_SST: u16 = 0x00fd;
const BOOL_ERR: u16 = 0x0205;

type Record = (u16, Vec<u8>);

/// A small package the `cell_values` editor admits: one worksheet with
/// `Number`, `LabelSst`, `BoolErr` and `Formula` cells and a live SST.
fn base_package() -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(sheet, 3, 2, 4.5).unwrap();
    writer.write_number(sheet, 4, 0, 1.0).unwrap();
    writer.write_number(sheet, 4, 1, 2.0).unwrap();
    writer.write_string(sheet, 6, 0, "alpha").unwrap();
    writer.write_string(sheet, 6, 1, "beta").unwrap();
    writer.write_string(sheet, 6, 2, "alpha").unwrap();
    writer.write_boolean(sheet, 7, 0, true).unwrap();
    writer.write_formula(sheet, 8, 0, "1+1").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

/// A two-worksheet package, so per-tab cases have a second tab to damage.
fn two_sheet_package() -> Vec<u8> {
    let mut writer = Writer::new();
    let first = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(first, 3, 2, 4.5).unwrap();
    writer.write_string(first, 6, 0, "alpha").unwrap();
    let second = writer.add_worksheet("Sheet2").unwrap();
    writer.write_number(second, 1, 1, 8.5).unwrap();
    writer.write_string(second, 2, 0, "beta").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn workbook_records(package: &[u8]) -> Vec<Record> {
    let mut ole = OleFile::open(Cursor::new(package.to_vec())).unwrap();
    let stream = ole.open_stream(&["Workbook"]).unwrap();
    let mut records = Vec::new();
    for record in Records::new(&stream) {
        let record = record.unwrap();
        records.push((record.kind().get(), record.payload().to_vec()));
    }
    records
}

fn encode(records: &[Record]) -> Vec<u8> {
    let mut bytes = Vec::new();
    for (kind, payload) in records {
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&u16::try_from(payload.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(payload);
    }
    bytes
}

/// Rebuilds a CFB around a Workbook stream, fixing up every `BoundSheet8`
/// position so the substream offsets stay consistent with the new framing.
fn repackage(records: &[Record], retarget_bound_sheets: bool) -> Vec<u8> {
    let mut records = records.to_vec();
    if retarget_bound_sheets {
        // Two passes: the first sizes the globals, the second writes the real
        // worksheet offsets. Record lengths never change between them.
        for _ in 0..2 {
            let mut offset = 0_usize;
            let mut globals_end = None;
            let mut worksheet_starts = Vec::new();
            let mut in_globals = true;
            for (kind, payload) in &records {
                if *kind == BOF && !in_globals {
                    worksheet_starts.push(offset);
                }
                if *kind == EOF && in_globals {
                    in_globals = false;
                    globals_end = Some(offset + 4 + payload.len());
                }
                offset += 4 + payload.len();
            }
            let _ = globals_end;
            let mut next = 0;
            for (kind, payload) in &mut records {
                if *kind == BOUND_SHEET
                    && payload.len() >= 4
                    && let Some(start) = worksheet_starts.get(next)
                {
                    payload[0..4].copy_from_slice(&u32::try_from(*start).unwrap().to_le_bytes());
                    next += 1;
                }
            }
        }
    }
    let stream = encode(&records);
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &stream).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn first_globals_index(records: &[Record], kind: u16) -> Option<usize> {
    let end = records.iter().position(|(k, _)| *k == EOF)?;
    records[..end].iter().position(|(k, _)| *k == kind)
}

/// The index of the first record of `kind` in the first worksheet substream.
fn first_worksheet_index(records: &[Record], kind: u16) -> Option<usize> {
    let globals_end = records.iter().position(|(k, _)| *k == EOF)?;
    records[globals_end + 1..]
        .iter()
        .position(|(k, _)| *k == kind)
        .map(|index| index + globals_end + 1)
}

fn escape(value: &str) -> String {
    let mut out = String::new();
    for character in value.chars() {
        match character {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            character if (character as u32) < 0x20 => {
                out.push_str(&format!("\\u{:04x}", character as u32));
            },
            character => out.push(character),
        }
    }
    out
}

fn main() {
    let mut rows: Vec<(String, String, String)> = Vec::new();

    let case = |name: &str, bytes: Vec<u8>, rows: &mut Vec<(String, String, String)>| {
        let (outcome, message) = match Snapshot::from_bytes(bytes) {
            Ok(snapshot) => (
                "ok".to_string(),
                format!("worksheets={}", snapshot.worksheet_count()),
            ),
            Err(error) => ("refused".to_string(), error.to_string()),
        };
        rows.push((name.to_string(), outcome, message));
    };

    let base = base_package();
    let records = workbook_records(&base);

    // --- control -----------------------------------------------------------
    case("00-control-unmodified", base.clone(), &mut rows);
    case(
        "01-control-repackaged",
        repackage(&records, false),
        &mut rows,
    );

    // --- globals defects, all reachable by the offset inventory first -------
    {
        let mut damaged = records.clone();
        damaged[0].0 = 0x003c;
        case("10-globals-bof-kind", repackage(&damaged, false), &mut rows);
    }
    {
        let mut damaged = records.clone();
        damaged[0].1.truncate(2);
        case(
            "11-globals-bof-payload-truncated",
            repackage(&damaged, false),
            &mut rows,
        );
    }
    {
        let mut damaged = records.clone();
        damaged[0].1[2..4].copy_from_slice(&0x0006_u16.to_le_bytes());
        case(
            "12-globals-bof-substream-kind",
            repackage(&damaged, false),
            &mut rows,
        );
    }
    {
        let mut damaged = records.clone();
        let eof = damaged.iter().position(|(k, _)| *k == EOF).unwrap();
        damaged.remove(eof);
        case("13-globals-no-eof", repackage(&damaged, true), &mut rows);
    }
    {
        let mut damaged = records.clone();
        let insert = first_globals_index(&damaged, BOUND_SHEET).unwrap();
        damaged.insert(insert, (FILE_PASS, vec![1, 0, 1, 0]));
        case("14-globals-file-pass", repackage(&damaged, true), &mut rows);
    }
    {
        let mut damaged = records.clone();
        let sst = first_globals_index(&damaged, SST).unwrap();
        let payload = damaged[sst].1.clone();
        damaged.insert(sst, (SST, payload));
        case(
            "15-globals-duplicate-sst",
            repackage(&damaged, true),
            &mut rows,
        );
    }
    {
        let mut damaged = records.clone();
        let sst = first_globals_index(&damaged, SST).unwrap();
        damaged[sst].1.truncate(4);
        case(
            "16-globals-truncated-sst",
            repackage(&damaged, true),
            &mut rows,
        );
    }
    {
        let mut damaged = records.clone();
        damaged.retain(|(kind, _)| *kind != XF);
        case("17-globals-no-xf", repackage(&damaged, true), &mut rows);
    }
    {
        let mut damaged = records.clone();
        if let Some(index) = first_globals_index(&damaged, CODE_PAGE) {
            damaged[index].1[0..2].copy_from_slice(&0xdead_u16.to_le_bytes());
        } else {
            let insert = first_globals_index(&damaged, BOUND_SHEET).unwrap();
            damaged.insert(insert, (CODE_PAGE, 0xdead_u16.to_le_bytes().to_vec()));
        }
        case(
            "18-globals-unknown-codepage",
            repackage(&damaged, true),
            &mut rows,
        );
    }

    // --- BoundSheet8 defects ------------------------------------------------
    {
        let mut damaged = records.clone();
        let bound = first_globals_index(&damaged, BOUND_SHEET).unwrap();
        let stream_len = encode(&damaged).len();
        damaged[bound].1[0..4]
            .copy_from_slice(&u32::try_from(stream_len + 64).unwrap().to_le_bytes());
        case(
            "20-boundsheet-position-past-end",
            repackage(&damaged, false),
            &mut rows,
        );
    }
    {
        let two = two_sheet_package();
        let mut damaged = workbook_records(&two);
        let bound = first_globals_index(&damaged, BOUND_SHEET).unwrap();
        let first = damaged[bound].1[0..4].to_vec();
        let second = damaged[bound + 1..]
            .iter()
            .position(|(k, _)| *k == BOUND_SHEET)
            .unwrap()
            + bound
            + 1;
        damaged[second].1[0..4].copy_from_slice(&first);
        case(
            "21-boundsheet-duplicate-position",
            repackage(&damaged, false),
            &mut rows,
        );
    }
    {
        let two = two_sheet_package();
        let mut damaged = workbook_records(&two);
        let bound = first_globals_index(&damaged, BOUND_SHEET).unwrap();
        let second = damaged[bound + 1..]
            .iter()
            .position(|(k, _)| *k == BOUND_SHEET)
            .unwrap()
            + bound
            + 1;
        let name = damaged[bound].1[6..].to_vec();
        let mut payload = damaged[second].1[0..6].to_vec();
        payload.extend_from_slice(&name);
        damaged[second].1 = payload;
        case(
            "22-boundsheet-duplicate-name",
            repackage(&damaged, true),
            &mut rows,
        );
    }

    // --- worksheet substream defects ---------------------------------------
    {
        let mut damaged = records.clone();
        let globals_end = damaged.iter().position(|(k, _)| *k == EOF).unwrap();
        damaged[globals_end + 1].1[2..4].copy_from_slice(&0x0005_u16.to_le_bytes());
        case(
            "30-worksheet-bof-substream-kind",
            repackage(&damaged, true),
            &mut rows,
        );
    }
    {
        let mut damaged = records.clone();
        let globals_end = damaged.iter().position(|(k, _)| *k == EOF).unwrap();
        damaged[globals_end + 1].0 = 0x003c;
        case(
            "31-worksheet-bof-kind",
            repackage(&damaged, true),
            &mut rows,
        );
    }
    {
        let mut damaged = records.clone();
        while damaged.last().map(|(kind, _)| *kind) == Some(EOF) {
            damaged.pop();
        }
        case("32-worksheet-no-eof", repackage(&damaged, true), &mut rows);
    }

    // --- cell record defects, inventory-visible -----------------------------
    {
        let mut damaged = records.clone();
        let number = first_worksheet_index(&damaged, NUMBER).unwrap();
        damaged[number].1.truncate(10);
        case(
            "40-number-payload-truncated",
            repackage(&damaged, true),
            &mut rows,
        );
    }
    {
        let mut damaged = records.clone();
        let number = first_worksheet_index(&damaged, NUMBER).unwrap();
        damaged[number].1[6..14].copy_from_slice(&f64::NAN.to_le_bytes());
        case("41-number-nan", repackage(&damaged, true), &mut rows);
    }
    {
        let mut damaged = records.clone();
        let number = first_worksheet_index(&damaged, NUMBER).unwrap();
        damaged[number].1[4..6].copy_from_slice(&0x0fff_u16.to_le_bytes());
        case(
            "42-cell-xf-index-past-end",
            repackage(&damaged, true),
            &mut rows,
        );
    }
    {
        let mut damaged = records.clone();
        let label = first_worksheet_index(&damaged, LABEL_SST).unwrap();
        damaged[label].1[6..10].copy_from_slice(&0x0000_7fff_u32.to_le_bytes());
        case(
            "43-labelsst-index-past-sst",
            repackage(&damaged, true),
            &mut rows,
        );
    }
    {
        let mut damaged = records.clone();
        let boolerr = first_worksheet_index(&damaged, BOOL_ERR).unwrap();
        damaged[boolerr].1[7] = 9;
        case("44-boolerr-bad-flag", repackage(&damaged, true), &mut rows);
    }
    {
        let mut damaged = records.clone();
        let formula = first_worksheet_index(&damaged, FORMULA).unwrap();
        damaged[formula].1[6..14].copy_from_slice(&[0, 0, 0, 0, 0, 0, 0xff, 0xff]);
        case(
            "45-formula-string-cache-without-string",
            repackage(&damaged, true),
            &mut rows,
        );
    }
    {
        let mut damaged = records.clone();
        let formula = first_worksheet_index(&damaged, FORMULA).unwrap();
        damaged[formula].1.truncate(20);
        case(
            "46-formula-payload-truncated",
            repackage(&damaged, true),
            &mut rows,
        );
    }
    {
        let mut damaged = records.clone();
        let formula = first_worksheet_index(&damaged, FORMULA).unwrap();
        damaged.insert(formula + 1, (STRING, vec![1, 0, 0, b'a']));
        case(
            "47-stray-string-after-numeric-formula",
            repackage(&damaged, true),
            &mut rows,
        );
    }

    // --- framing defects ----------------------------------------------------
    {
        let mut stream = encode(&records);
        stream.extend_from_slice(&[0x01, 0x02]);
        let mut writer = OleWriter::new();
        writer.create_stream(&["Workbook"], &stream).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        case("50-trailing-partial-header", output.into_inner(), &mut rows);
    }
    {
        let mut records = records.clone();
        let number = first_worksheet_index(&records, NUMBER).unwrap();
        let mut stream = encode(&records);
        // Declare a payload longer than the stream can hold.
        let mut offset = 0;
        for (index, (_, payload)) in records.iter().enumerate() {
            if index == number {
                break;
            }
            offset += 4 + payload.len();
        }
        stream[offset + 2..offset + 4].copy_from_slice(&0xfff0_u16.to_le_bytes());
        records[number].1.clear();
        let mut writer = OleWriter::new();
        writer.create_stream(&["Workbook"], &stream).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        case("51-record-payload-overruns", output.into_inner(), &mut rows);
    }

    // --- container defects --------------------------------------------------
    {
        let mut ole = OleFile::open(Cursor::new(base.clone())).unwrap();
        let stream = ole.open_stream(&["Workbook"]).unwrap();
        let mut writer = OleWriter::new();
        writer.create_stream(&["NotAWorkbook"], &stream).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        case("60-no-workbook-stream", output.into_inner(), &mut rows);
    }
    {
        let mut ole = OleFile::open(Cursor::new(base.clone())).unwrap();
        let stream = ole.open_stream(&["Workbook"]).unwrap();
        let mut writer = OleWriter::new();
        writer.create_stream(&["Book"], &stream).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        case("61-book-stream-alias", output.into_inner(), &mut rows);
    }

    rows.sort();
    for (name, outcome, message) in rows {
        println!(
            "{{\"case\":\"{}\",\"outcome\":\"{}\",\"message\":\"{}\"}}",
            escape(&name),
            escape(&outcome),
            escape(&message)
        );
    }
}
