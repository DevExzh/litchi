//! Differential tests: the complete reader against its validation-only mode.
//!
//! The validation-only mode is correct only if it is the same parser: for
//! every input, the same accept or the same first refusal, the same
//! workbook- and worksheet-level facts, and, at a kept position, the same
//! decoded cell. These tests check that over every `.xls` fixture in the
//! repository, over deterministic record-level mutations of real worksheet
//! substreams, and over synthetic worksheets aimed at the two cell questions
//! the occupancy map answers (duplicates and `Array` ownership).

use super::codec::{DecodeEveryCell, ValidateCells};
use super::model::{OpenOptions, Workbook};
use super::validation_only::{KeptCells, ValidationWorkbook};
use crate::compatibility::CompatibilityProfile;
use crate::records::{BoundSheetRecord, Encoding, SheetType};
use crate::worksheet::Worksheet;
use litchi_biff::Records;
use litchi_cfb::OleWriter;
use std::io::{Cursor, Read, Seek};
use std::path::{Path, PathBuf};

mod multi_defect_cases;

type Record = (u16, Vec<u8>);

const BOF: u16 = 0x0809;
const EOF: u16 = 0x000A;
const FORMULA: u16 = 0x0006;
const STRING: u16 = 0x0207;
const NUMBER: u16 = 0x0203;
const RK: u16 = 0x027E;
const LABEL_SST: u16 = 0x00FD;
const BLANK: u16 = 0x0201;
const BOOL_ERR: u16 = 0x0205;
const MUL_RK: u16 = 0x00BD;
const MUL_BLANK: u16 = 0x00BE;
const ARRAY: u16 = 0x0221;
const SHR_FMLA: u16 = 0x04BC;
const DIMENSIONS: u16 = 0x0200;

fn xls_fixtures() -> Vec<PathBuf> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let mut pending = vec![root];
    let mut found = Vec::new();
    while let Some(directory) = pending.pop() {
        let Ok(entries) = std::fs::read_dir(&directory) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                pending.push(path);
            } else if path
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"))
            {
                found.push(path);
            }
        }
    }
    found.sort();
    found
}

/// Every fact the two modes must agree on, as deterministic text: all of the
/// workbook's state except the CFB reader, the decoded cells and the
/// hash-indexed formatting lookup (whose ordered contents are compared).
fn facts<R: Read + Seek>(workbook: &Workbook<R>) -> String {
    let worksheets = workbook
        .worksheets
        .iter()
        .map(Worksheet::without_cells_for_tests)
        .collect::<Vec<_>>();
    let formatting = &workbook.formatting;
    [
        format!("{:#?}", (&workbook.worksheet_names, &workbook.sheets)),
        format!("{worksheets:#?}"),
        format!(
            "{:#?}",
            (
                &workbook.shared_strings,
                &workbook.shared_string_properties,
                workbook.shared_string_reference_count,
                &workbook.palette,
                &workbook.fonts,
                &workbook.biff_version,
                workbook.is_1904_date_system,
                &workbook.formula_context,
            )
        ),
        format!(
            "{:#?}",
            (
                &workbook.defined_names,
                &workbook.defined_name_records,
                formatting.number_formats(),
                formatting.extended_formats(),
                formatting.differential_formats(),
                formatting.date_system(),
                &workbook.protection,
                &workbook.calculation,
                &workbook.picture_compression,
                &workbook.vba_metadata,
                &workbook.environment,
                &workbook.book_ext,
            )
        ),
        format!(
            "{:#?}",
            (
                &workbook.style_extensions,
                &workbook.theme,
                &workbook.write_access,
                &workbook.table_styles,
                &workbook.shared_string_index,
                &workbook.workbook_view,
                &workbook.custom_views,
                &workbook.real_time_data,
                &workbook.mdx_metadata,
                &workbook.web_publications,
                &workbook.function_groups,
                &workbook.external_links,
            )
        ),
        format!(
            "{:#?}",
            (
                &workbook.pivot_caches,
                &workbook.pivot_cache_stream_ids,
                &workbook.xml_map,
                &workbook.tolerance,
                workbook.vba_project_storage(),
            )
        ),
    ]
    .join("\n")
}

/// The outcome both modes must agree on: the first refusal, or the facts.
fn outcome<R: Read + Seek>(result: Result<&Workbook<R>, &crate::Error>) -> String {
    match result {
        Ok(workbook) => format!("ok\n{}", facts(workbook)),
        Err(error) => format!("error: {error}\n{error:?}"),
    }
}

fn open_both(bytes: &[u8], kept: KeptCells) -> (String, String) {
    let full = Workbook::new(Cursor::new(bytes));
    let validated = Workbook::validation_only(Cursor::new(bytes), kept);
    (
        outcome(full.as_ref()),
        outcome(
            validated
                .as_ref()
                .map(ValidationWorkbook::as_workbook_for_tests),
        ),
    )
}

#[test]
fn facts_are_deterministic_across_two_complete_opens() {
    for path in xls_fixtures().into_iter().take(24) {
        let bytes = std::fs::read(&path).unwrap();
        let first = Workbook::new(Cursor::new(bytes.as_slice()));
        let second = Workbook::new(Cursor::new(bytes.as_slice()));
        assert_eq!(
            outcome(first.as_ref()),
            outcome(second.as_ref()),
            "{}",
            path.display()
        );
    }
}

/// Every `.xls` in the repository: the same acceptance or first refusal and
/// the same non-cell facts, keeping no cell.
#[test]
fn validation_only_open_agrees_with_the_complete_reader_on_every_fixture() {
    let fixtures = xls_fixtures();
    assert!(fixtures.len() >= 100, "fixture corpus went missing");
    let mut accepted = 0;
    let mut refused = 0;
    for path in &fixtures {
        let bytes = std::fs::read(path).unwrap();
        let (full, validated) = open_both(&bytes, KeptCells::none());
        assert_eq!(full, validated, "{}", path.display());
        if full.starts_with("ok") {
            accepted += 1;
        } else {
            refused += 1;
        }
    }
    // Both populations are exercised.
    assert!(
        accepted > 50 && refused > 0,
        "{accepted} accepted, {refused} refused"
    );
}

/// Up to `limit` positions of one worksheet's decoded cells, spread over the
/// whole sheet, plus every Formula cell among the first `limit`.
fn sample_positions(worksheet: &Worksheet, limit: usize) -> Vec<(u16, u16)> {
    let positions = worksheet.cell_positions_for_tests();
    let step = positions.len().div_ceil(limit.max(1)).max(1);
    let mut sample = positions
        .iter()
        .step_by(step)
        .chain(positions.last())
        .map(|&(row, column)| (u16::try_from(row).unwrap(), u16::try_from(column).unwrap()))
        .collect::<Vec<_>>();
    sample.extend(
        positions
            .iter()
            .filter(|&&(row, column)| {
                worksheet
                    .get_cell(row, column)
                    .is_some_and(crate::Cell::is_formula_record)
            })
            .take(limit)
            .map(|&(row, column)| (u16::try_from(row).unwrap(), u16::try_from(column).unwrap())),
    );
    sample.sort_unstable();
    sample.dedup();
    sample
}

/// At kept positions the validation-only open decodes the complete reader's
/// cell exactly, including shared-formula renderings and Array owners; a
/// kept vacant position reads back as vacant.
#[test]
fn kept_cells_are_the_complete_readers_cells() {
    let mut compared = 0;
    for path in xls_fixtures() {
        let bytes = std::fs::read(&path).unwrap();
        let Ok(full) = Workbook::new(Cursor::new(bytes.as_slice())) else {
            continue;
        };
        let mut kept = Vec::new();
        for (tab, sheet) in full.sheets().iter().enumerate() {
            let Some(index) = sheet.parsed_worksheet_index() else {
                continue;
            };
            let worksheet = full.xls_worksheet(index).unwrap();
            for (row, column) in sample_positions(worksheet, 48) {
                kept.push((tab, row, column));
            }
            // Vacant positions: past the last row, and beyond the grid.
            kept.push((tab, u16::MAX, 255));
            kept.push((tab, 0, 300));
        }
        let validated = Workbook::validation_only(
            Cursor::new(bytes.as_slice()),
            KeptCells::from_cells(kept.iter().copied()).unwrap(),
        )
        .unwrap();
        assert_eq!(
            facts(&full),
            facts(validated.as_workbook_for_tests()),
            "{}",
            path.display()
        );
        for &(tab, row, column) in &kept {
            let index = full.sheet(tab).unwrap().parsed_worksheet_index().unwrap();
            let expected = full
                .xls_worksheet(index)
                .unwrap()
                .get_cell(u32::from(row), u32::from(column));
            let actual = validated.kept_cell(index, row, column).unwrap();
            assert_eq!(
                format!("{actual:?}"),
                format!("{expected:?}"),
                "{} tab {tab} ({row}, {column})",
                path.display()
            );
            compared += 1;
        }
        // Only the kept positions were decoded.
        for (tab, sheet) in full.sheets().iter().enumerate() {
            if let Some(index) = sheet.parsed_worksheet_index() {
                let decoded = validated
                    .as_workbook_for_tests()
                    .xls_worksheet(index)
                    .unwrap()
                    .cell_count_for_tests();
                let kept_here = kept
                    .iter()
                    .filter(|(kept_tab, ..)| *kept_tab == tab)
                    .count();
                assert!(decoded <= kept_here, "{} tab {tab}", path.display());
            }
        }
    }
    assert!(compared > 1_000, "only {compared} kept cells were compared");
}

// ---------------------------------------------------------------------------
// Worksheet-level differential over mutated substreams
// ---------------------------------------------------------------------------

struct WorksheetInputs<'w> {
    encoding: Encoding,
    shared_strings: std::sync::Arc<Vec<String>>,
    properties: std::sync::Arc<Vec<Option<Box<crate::records::SharedStringProperties>>>>,
    formula_context: &'w crate::formula::FormulaContext,
    formatting: std::sync::Arc<crate::number_format::Formatting>,
}

fn inputs<R: Read + Seek>(workbook: &Workbook<R>) -> WorksheetInputs<'_> {
    WorksheetInputs {
        encoding: Encoding::from_codepage(1252).unwrap(),
        shared_strings: workbook.shared_strings_shared(),
        properties: workbook
            .shared_string_properties_shared()
            .unwrap_or_default(),
        formula_context: &workbook.formula_context,
        formatting: std::sync::Arc::clone(&workbook.formatting),
    }
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

/// Parses one worksheet substream (starting at its BOF) with both stores and
/// returns the two outcomes, plus the kept cells of each.
fn parse_both(
    inputs: &WorksheetInputs<'_>,
    records: &[Record],
    kept: &[(u16, u16)],
    profile: CompatibilityProfile,
) -> (String, String) {
    let stream = encode(records);
    let parse = |validate: bool| -> String {
        let mut framed = Records::new(&stream);
        let Some(bof) = framed.next() else {
            return "empty".into();
        };
        let bof = match bof {
            Ok(bof) => bof,
            Err(error) => return format!("error: {error}"),
        };
        if bof.kind().get() != BOF {
            return "no BOF".into();
        }
        let start = u64::try_from(bof.encoded().len()).unwrap();
        let result = if validate {
            Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records_with_compatibility(
                &mut framed,
                stream.len() as u64,
                0,
                start,
                &inputs.encoding,
                "Mutated",
                std::sync::Arc::clone(&inputs.shared_strings),
                std::sync::Arc::clone(&inputs.properties),
                Some(inputs.formula_context),
                std::sync::Arc::clone(&inputs.formatting),
                profile,
                &mut ValidateCells::new(kept),
            )
        } else {
            Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records_with_compatibility(
                &mut framed,
                stream.len() as u64,
                0,
                start,
                &inputs.encoding,
                "Mutated",
                std::sync::Arc::clone(&inputs.shared_strings),
                std::sync::Arc::clone(&inputs.properties),
                Some(inputs.formula_context),
                std::sync::Arc::clone(&inputs.formatting),
                profile,
                &mut DecodeEveryCell,
            )
        };
        match result {
            Ok(worksheet) => {
                let kept_cells = kept
                    .iter()
                    .map(|&(row, column)| {
                        format!(
                            "{:?}",
                            worksheet.get_cell(u32::from(row), u32::from(column))
                        )
                    })
                    .collect::<Vec<_>>();
                format!(
                    "ok\n{:#?}\n{kept_cells:#?}",
                    worksheet.without_cells_for_tests()
                )
            },
            Err(error) => format!("error: {error}\n{error:?}"),
        }
    };
    (parse(false), parse(true))
}

/// A small deterministic generator for mutation choices.
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
            usize::try_from(self.next() % u64::try_from(bound).unwrap()).unwrap()
        }
    }
}

fn is_cell(kind: u16) -> bool {
    matches!(
        kind,
        NUMBER | RK | LABEL_SST | BLANK | BOOL_ERR | FORMULA | 0x0204
    )
}

/// One deterministic record-level mutation of a worksheet substream.
fn mutate(records: &[Record], random: &mut XorShift) -> Vec<Record> {
    let mut records = records.to_vec();
    let cells = records
        .iter()
        .enumerate()
        .filter(|(_, (kind, payload))| is_cell(*kind) && payload.len() >= 6)
        .map(|(index, _)| index)
        .collect::<Vec<_>>();
    // Never mutate the BOF: its refusal is the caller's, not the store's.
    let body = 1..records.len().max(2);
    let pick_cell = |random: &mut XorShift| cells.get(random.below(cells.len())).copied();
    match random.below(12) {
        // A cell record repeated in place: a duplicate position.
        0 => {
            if let Some(index) = pick_cell(random) {
                let copy = records[index].clone();
                records.insert(index + 1, copy);
            }
        },
        // A cell record repeated further down the sheet.
        1 => {
            if let Some(index) = pick_cell(random) {
                let copy = records[index].clone();
                let at = index + 1 + random.below(records.len().saturating_sub(index + 1));
                records.insert(at.min(records.len().saturating_sub(1)), copy);
            }
        },
        // A cell moved onto another cell's position.
        2 => {
            if let (Some(from), Some(onto)) = (pick_cell(random), pick_cell(random)) {
                let position = records[onto].1[0..4].to_vec();
                records[from].1[0..4].copy_from_slice(&position);
            }
        },
        // A single-cell record moved outside the 256-column grid.
        3 => {
            if let Some(index) = pick_cell(random) {
                let column = 256 + u16::try_from(random.below(4)).unwrap();
                records[index].1[2..4].copy_from_slice(&column.to_le_bytes());
            }
        },
        // A byte flipped in any record after the BOF.
        4 | 5 => {
            let index = body.start + random.below(body.len());
            if let Some((_, payload)) = records.get_mut(index)
                && !payload.is_empty()
            {
                let at = random.below(payload.len());
                payload[at] ^= 1 << random.below(8);
            }
        },
        // A record truncated.
        6 => {
            let index = body.start + random.below(body.len());
            if let Some((_, payload)) = records.get_mut(index) {
                let cut = 1 + random.below(4);
                payload.truncate(payload.len().saturating_sub(cut));
            }
        },
        // A record dropped.
        7 => {
            let index = body.start + random.below(body.len());
            if index < records.len() && records.len() > 2 {
                records.remove(index);
            }
        },
        // Two adjacent records swapped.
        8 => {
            let index = body.start + random.below(body.len());
            if index + 1 < records.len() {
                records.swap(index, index + 1);
            }
        },
        // A stray String record after a cell.
        9 => {
            if let Some(index) = pick_cell(random) {
                records.insert(index + 1, (STRING, vec![1, 0, 0, b'x']));
            }
        },
        // A cell's XF index pushed past the workbook's resources.
        10 => {
            if let Some(index) = pick_cell(random) {
                records[index].1[4..6].copy_from_slice(&0x0fff_u16.to_le_bytes());
            }
        },
        // A companion (Array, ShrFmla) repeated, or a Formula repeated.
        _ => {
            let companion = records
                .iter()
                .position(|(kind, _)| matches!(*kind, ARRAY | SHR_FMLA | FORMULA));
            if let Some(index) = companion {
                let copy = records[index].clone();
                records.insert(index + 1, copy);
            }
        },
    }
    records
}

/// The records of one worksheet substream, from its BOF through its EOF.
fn worksheet_records(workbook_stream: &[u8], bound: &BoundSheetRecord) -> Option<Vec<Record>> {
    let start = usize::try_from(bound.position).ok()?;
    let mut records = Vec::new();
    for record in Records::new(workbook_stream.get(start..)?) {
        let record = record.ok()?;
        let kind = record.kind().get();
        records.push((kind, record.payload().to_vec()));
        if kind == EOF {
            return Some(records);
        }
    }
    None
}

fn workbook_stream(bytes: &[u8]) -> Option<Vec<u8>> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).ok()?;
    ole.open_stream(&["Workbook"])
        .or_else(|_| ole.open_stream(&["Book"]))
        .ok()
}

fn bound_sheets(stream: &[u8]) -> Vec<BoundSheetRecord> {
    let encoding = Encoding::from_codepage(1252).unwrap();
    let mut sheets = Vec::new();
    for record in Records::new(stream) {
        let Ok(record) = record else {
            break;
        };
        match record.kind().get() {
            0x0085 => {
                if let Ok(bound) = BoundSheetRecord::parse(record.payload(), &encoding) {
                    sheets.push(bound);
                }
            },
            EOF => break,
            _ => {},
        }
    }
    sheets
}

/// Deterministic mutations of every worksheet of a set of real fixtures,
/// parsed by both stores under both compatibility profiles: the same first
/// refusal, or the same facts and the same kept cells.
#[test]
fn worksheet_parse_agrees_with_the_complete_reader_under_mutation() {
    let chosen = [
        "54016.xls",
        "FormulaEvalTestData.xls",
        "WithCustomViews.xls",
        "SimpleWithFormula.xls",
        "SimpleMultiCell.xls",
        "44010-SingleChart.xls",
        "ConditionalFormattingSamples.xls",
        "45365-2.xls",
        "shared_formulas.xls",
        "SimpleWithComments.xls",
    ];
    let fixtures = xls_fixtures()
        .into_iter()
        .filter(|path| {
            path.file_name()
                .and_then(|name| name.to_str())
                .is_some_and(|name| chosen.contains(&name))
        })
        .collect::<Vec<_>>();
    assert!(fixtures.len() >= 6, "mutation fixtures went missing");
    let mut compared = 0;
    let mut refused = 0;
    let mut random = XorShift(0x0746_0746_0746_0746);
    for path in fixtures {
        let bytes = std::fs::read(&path).unwrap();
        let Ok(full) =
            Workbook::new_with_options(Cursor::new(bytes.as_slice()), OpenOptions::new())
        else {
            continue;
        };
        let Some(stream) = workbook_stream(&bytes) else {
            continue;
        };
        let inputs = inputs(&full);
        for bound in bound_sheets(&stream) {
            if bound.sheet_type != SheetType::WorkSheet {
                continue;
            }
            let Some(records) = worksheet_records(&stream, &bound) else {
                continue;
            };
            let kept_candidates = records
                .iter()
                .filter(|(kind, payload)| is_cell(*kind) && payload.len() >= 4)
                .map(|(_, payload)| {
                    (
                        u16::from_le_bytes([payload[0], payload[1]]),
                        u16::from_le_bytes([payload[2], payload[3]]),
                    )
                })
                .collect::<Vec<_>>();
            for round in 0..48 {
                let mutated = if round == 0 {
                    records.clone()
                } else {
                    let mut mutated = mutate(&records, &mut random);
                    if round % 3 == 0 {
                        mutated = mutate(&mutated, &mut random);
                    }
                    mutated
                };
                let mut kept = (0..3)
                    .filter_map(|_| {
                        kept_candidates
                            .get(random.below(kept_candidates.len()))
                            .copied()
                    })
                    .collect::<Vec<_>>();
                kept.sort_unstable();
                kept.dedup();
                for profile in [
                    CompatibilityProfile::Strict,
                    CompatibilityProfile::SharedFormulaFlagWithoutPtgExpV1,
                ] {
                    let (complete, validated) = parse_both(&inputs, &mutated, &kept, profile);
                    assert_eq!(
                        complete,
                        validated,
                        "{} sheet {:?} round {round}",
                        path.display(),
                        bound.name
                    );
                    compared += 1;
                    if complete.starts_with("error") {
                        refused += 1;
                    }
                }
            }
        }
    }
    assert!(
        compared > 500,
        "only {compared} mutated worksheets were compared"
    );
    assert!(
        refused > 50,
        "only {refused} mutated worksheets were refused"
    );
}

// ---------------------------------------------------------------------------
// Synthetic worksheets aimed at the two occupancy questions
// ---------------------------------------------------------------------------

fn cell_head(row: u16, column: u16, xf: u16) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&xf.to_le_bytes());
    payload
}

fn number(row: u16, column: u16, value: f64) -> Record {
    let mut payload = cell_head(row, column, 0);
    payload.extend_from_slice(&value.to_le_bytes());
    (NUMBER, payload)
}

/// A `Formula` whose tokens are one `PtgExp` to `anchor`; `string` makes its
/// cached value a pending `String`, `shared` sets `fShrFmla`.
fn ptg_exp_formula(
    row: u16,
    column: u16,
    anchor: (u16, u16),
    string: bool,
    shared: bool,
) -> Record {
    let mut payload = cell_head(row, column, 0);
    if string {
        payload.extend_from_slice(&[0, 0, 0, 0, 0, 0, 0xff, 0xff]);
    } else {
        payload.extend_from_slice(&1.5_f64.to_le_bytes());
    }
    payload.extend_from_slice(&(if shared { 0x0008_u16 } else { 0 }).to_le_bytes());
    payload.extend_from_slice(&0_u32.to_le_bytes());
    payload.extend_from_slice(&5_u16.to_le_bytes());
    payload.push(0x01);
    payload.extend_from_slice(&anchor.0.to_le_bytes());
    payload.extend_from_slice(&anchor.1.to_le_bytes());
    (FORMULA, payload)
}

/// A `Formula` with a numeric cache and a one-`PtgInt` body.
fn plain_formula(row: u16, column: u16) -> Record {
    let mut payload = cell_head(row, column, 0);
    payload.extend_from_slice(&2.5_f64.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&0_u32.to_le_bytes());
    payload.extend_from_slice(&3_u16.to_le_bytes());
    payload.extend_from_slice(&[0x1E, 7, 0]);
    (FORMULA, payload)
}

/// An `Array` owning `first..=last` rows of one column with a `PtgInt` body.
fn array(first_row: u16, last_row: u16, column: u8) -> Record {
    let mut payload = Vec::new();
    payload.extend_from_slice(&first_row.to_le_bytes());
    payload.extend_from_slice(&last_row.to_le_bytes());
    payload.push(column);
    payload.push(column);
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&0_u32.to_le_bytes());
    payload.extend_from_slice(&3_u16.to_le_bytes());
    payload.extend_from_slice(&[0x1E, 7, 0]);
    (ARRAY, payload)
}

fn dimensions(last_row: u32, last_column: u16) -> Record {
    let mut payload = Vec::new();
    payload.extend_from_slice(&0_u32.to_le_bytes());
    payload.extend_from_slice(&last_row.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&last_column.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    (DIMENSIONS, payload)
}

fn bof() -> Record {
    (BOF, vec![0, 6, 0x10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
}

/// Streams that reach every branch of the duplicate and `Array` checks the
/// occupancy map answers, including the two refusals only a trailing
/// unresolved string `Formula` can produce.
#[test]
fn synthetic_worksheets_agree_on_duplicate_and_array_ownership() {
    let bytes = {
        let mut writer = crate::Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_number(sheet, 0, 0, 1.0).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    };
    let full = Workbook::new(Cursor::new(bytes.as_slice())).unwrap();
    let inputs = inputs(&full);
    let anchor = (1_u16, 0_u16);
    let valid_array = vec![
        bof(),
        dimensions(8, 4),
        ptg_exp_formula(1, 0, anchor, false, false),
        array(1, 2, 0),
        ptg_exp_formula(2, 0, anchor, false, false),
        (EOF, Vec::new()),
    ];
    let cases: Vec<(&str, Vec<Record>)> = vec![
        ("valid-array", valid_array.clone()),
        ("duplicate-number", {
            vec![
                bof(),
                number(3, 3, 1.0),
                number(3, 3, 2.0),
                (EOF, Vec::new()),
            ]
        }),
        ("number-over-array-cell", {
            let mut records = valid_array.clone();
            records.insert(5, number(2, 0, 4.0));
            records
        }),
        ("array-cell-before-number", {
            vec![
                bof(),
                dimensions(8, 4),
                number(2, 0, 4.0),
                ptg_exp_formula(1, 0, anchor, false, false),
                array(1, 2, 0),
                ptg_exp_formula(2, 0, anchor, false, false),
                (EOF, Vec::new()),
            ]
        }),
        ("array-cell-missing", {
            let mut records = valid_array.clone();
            records.remove(4);
            records
        }),
        // The string Formula at (2, 0) never meets its String: with no EOF
        // the walk ends while it is pending, so the cell is never stored.
        ("array-cell-unmaterialized", {
            vec![
                bof(),
                dimensions(8, 4),
                ptg_exp_formula(1, 0, anchor, false, false),
                array(1, 2, 0),
                ptg_exp_formula(2, 0, anchor, true, false),
            ]
        }),
        // The same, over a Number already at (2, 0): the slot holds a
        // non-Formula cell and the owner cannot attach.
        ("array-cell-non-formula", {
            vec![
                bof(),
                dimensions(8, 4),
                number(2, 0, 4.0),
                ptg_exp_formula(1, 0, anchor, false, false),
                array(1, 2, 0),
                ptg_exp_formula(2, 0, anchor, true, false),
            ]
        }),
        ("array-cell-shared-flag", {
            let mut records = valid_array.clone();
            records[4] = ptg_exp_formula(2, 0, anchor, false, true);
            records
        }),
        ("formula-then-number-then-formula", {
            vec![
                bof(),
                plain_formula(5, 5),
                number(5, 5, 1.0),
                plain_formula(5, 5),
                (EOF, Vec::new()),
            ]
        }),
        ("orphan-ptg-exp", {
            vec![
                bof(),
                ptg_exp_formula(5, 5, (5, 5), false, false),
                (EOF, Vec::new()),
            ]
        }),
        ("outside-grid-duplicate", {
            vec![
                bof(),
                number(0, 300, 1.0),
                number(0, 300, 2.0),
                (EOF, Vec::new()),
            ]
        }),
        ("mul-rk-over-number", {
            let mut mul_rk = Vec::new();
            mul_rk.extend_from_slice(&4_u16.to_le_bytes());
            mul_rk.extend_from_slice(&1_u16.to_le_bytes());
            for value in [3_u32, 5] {
                mul_rk.extend_from_slice(&0_u16.to_le_bytes());
                mul_rk.extend_from_slice(&((value << 2) | 0x02).to_le_bytes());
            }
            mul_rk.extend_from_slice(&2_u16.to_le_bytes());
            let mut mul_blank = Vec::new();
            mul_blank.extend_from_slice(&4_u16.to_le_bytes());
            mul_blank.extend_from_slice(&2_u16.to_le_bytes());
            mul_blank.extend_from_slice(&0_u16.to_le_bytes());
            mul_blank.extend_from_slice(&0_u16.to_le_bytes());
            mul_blank.extend_from_slice(&3_u16.to_le_bytes());
            vec![
                bof(),
                number(4, 2, 9.0),
                (MUL_RK, mul_rk),
                (MUL_BLANK, mul_blank),
                (EOF, Vec::new()),
            ]
        }),
    ];
    let mut outcomes = Vec::new();
    for (name, records) in cases {
        for kept in [
            &[][..],
            &[(0, 300), (1, 0), (2, 0), (3, 3), (4, 2), (5, 5)][..],
        ] {
            let (complete, validated) =
                parse_both(&inputs, &records, kept, CompatibilityProfile::Strict);
            assert_eq!(complete, validated, "{name} kept {kept:?}");
            if kept.is_empty() {
                outcomes.push((name, complete.lines().next().unwrap_or("").to_string()));
            }
        }
    }
    // The cases reach the branches they are named for.
    let outcome = |name: &str| {
        outcomes
            .iter()
            .find(|(case, _)| *case == name)
            .map(|(_, outcome)| outcome.clone())
            .unwrap()
    };
    assert_eq!(outcome("valid-array"), "ok");
    assert_eq!(outcome("duplicate-number"), "ok");
    assert!(outcome("number-over-array-cell").contains("is not exactly one Formula/PtgExp"));
    assert!(outcome("array-cell-before-number").contains("is not exactly one Formula/PtgExp"));
    assert!(outcome("array-cell-missing").contains("has no Formula/PtgExp record"));
    assert!(outcome("array-cell-unmaterialized").contains("was not materialized"));
    assert!(outcome("array-cell-non-formula").contains("cannot attach to a non-Formula cell"));
    assert!(
        outcome("array-cell-shared-flag").contains("orphan PtgExp")
            || outcome("array-cell-shared-flag").contains("not exactly one")
    );
    assert_eq!(outcome("formula-then-number-then-formula"), "ok");
    assert!(outcome("orphan-ptg-exp").contains("Formula at (5, 5) contains an orphan PtgExp"));
    assert_eq!(outcome("outside-grid-duplicate"), "ok");
    assert_eq!(outcome("mul-rk-over-number"), "ok");
}

/// Complete packages built from a real fixture with one mutated worksheet:
/// the complete reader and the validation-only open agree on the package
/// outcome, including which tabs are published.
#[test]
fn package_outcome_agrees_under_worksheet_mutation() {
    let path = xls_fixtures()
        .into_iter()
        .find(|path| path.ends_with("WithCustomViews.xls"))
        .expect("fixture");
    let bytes = std::fs::read(&path).unwrap();
    let stream = workbook_stream(&bytes).unwrap();
    let bounds = bound_sheets(&stream);
    let mut random = XorShift(0x5eed_0746);
    let mut compared = 0;
    for round in 0..40 {
        let Some(bound) = bounds
            .iter()
            .filter(|bound| bound.sheet_type == SheetType::WorkSheet)
            .nth(round % 3)
        else {
            continue;
        };
        let Some(records) = worksheet_records(&stream, bound) else {
            continue;
        };
        let mutated = mutate(&records, &mut random);
        // Splice the mutated substream in and repoint later BoundSheet8
        // offsets by the length change.
        let start = usize::try_from(bound.position).unwrap();
        let old_length = encode(&records).len();
        let new_substream = encode(&mutated);
        let delta =
            i64::try_from(new_substream.len()).unwrap() - i64::try_from(old_length).unwrap();
        let mut rewritten = Vec::new();
        rewritten.extend_from_slice(&stream[..start]);
        rewritten.extend_from_slice(&new_substream);
        rewritten.extend_from_slice(&stream[start + old_length..]);
        let mut offset = 0;
        let mut globals = Vec::new();
        for record in Records::new(&rewritten) {
            let record = record.unwrap();
            globals.push((offset, record.kind().get(), record.payload().len()));
            offset += record.encoded().len();
            if record.kind().get() == EOF {
                break;
            }
        }
        for (record_offset, kind, _) in globals {
            if kind == 0x0085 {
                let at = record_offset + 4;
                let position = u32::from_le_bytes(rewritten[at..at + 4].try_into().unwrap());
                if usize::try_from(position).unwrap() > start {
                    let moved = i64::from(position) + delta;
                    rewritten[at..at + 4]
                        .copy_from_slice(&u32::try_from(moved).unwrap().to_le_bytes());
                }
            }
        }
        let mut writer = OleWriter::new();
        writer.create_stream(&["Workbook"], &rewritten).unwrap();
        let mut package = Cursor::new(Vec::new());
        writer.write_to(&mut package).unwrap();
        let package = package.into_inner();
        let (full, validated) = open_both(&package, KeptCells::none());
        assert_eq!(full, validated, "round {round}");
        compared += 1;
    }
    assert!(compared >= 30);
}

// ---------------------------------------------------------------------------
// Frozen worksheet-level first-error matrix, generated on the base
// ---------------------------------------------------------------------------

/// The first refusal of every multi-defect case, as the base commit
/// (`009d515bef`) produced it: `multi_defect_cases.rs` was compiled into a
/// temporary test on that commit and parsed through its only worksheet walk
/// (the complete reader's), three runs with identical output. The package walk
/// swallows these refusals (`parse_workbook` drops the sheet), so only a
/// worksheet-level matrix can see which one comes first.
const BASE_MULTI_DEFECT_MATRIX: &[(&str, &str)] = &[
    (
        "dup-then-bad-xf",
        "Invalid record 0x00E0: cell references out-of-range XF 4095",
    ),
    (
        "bad-xf-then-truncated",
        "Invalid record 0x00E0: cell references out-of-range XF 4095",
    ),
    (
        "truncated-then-bad-xf",
        "Invalid length: expected 14, found 10",
    ),
    (
        "reserved-xf-then-bad-xf",
        "Invalid record 0x00E0: cell references reserved style-XF slot 3",
    ),
    (
        "pending-string-formula-then-stray-string",
        "Invalid record 0x0203: String-valued Formula must be followed by a String record",
    ),
    (
        "stray-string-then-orphan",
        "Invalid record 0x0207: String record has no pending string-valued Formula",
    ),
    (
        "string-continuation-not-continue-then-bad-xf",
        "Invalid record 0x0203: String result continuation must be a Continue record",
    ),
    (
        "orphan-then-array-without-dimensions",
        "Invalid record 0x0006: Formula at (2, 2) contains an orphan PtgExp",
    ),
    (
        "array-without-dimensions-then-duplicate-member",
        "Invalid record 0x0221: Array formulas require worksheet Dimensions",
    ),
    (
        "array-duplicate-anchor-then-missing-member",
        "Invalid record 0x0221: Array range cell (4, 0) is not exactly one Formula/PtgExp to its anchor",
    ),
    (
        "array-duplicate-member-before-missing-member",
        "Invalid record 0x0221: Array range cell (5, 0) is not exactly one Formula/PtgExp to its anchor",
    ),
    (
        "second-array-duplicate-member",
        "Invalid record 0x0221: Array range cell (5, 1) is not exactly one Formula/PtgExp to its anchor",
    ),
    (
        "array-outside-dimensions-with-duplicates",
        "Invalid record 0x0221: Array range is outside worksheet Dimensions",
    ),
    (
        "shrfmla-without-formula-then-bad-xf",
        "Invalid record 0x04BC: Shared must immediately follow its Formula record",
    ),
    (
        "duplicate-anchor-claims-second-companion",
        "Invalid record 0x0221: Formula at (1, 1) already owns a Shared companion",
    ),
    (
        "mulblank-bad-xf-then-mulrk-bad-range",
        "Invalid record 0x00E0: cell references out-of-range XF 4095",
    ),
    (
        "mulrk-bad-range-then-bad-xf",
        "Invalid data: MulRk column range 2..=9 does not match 2 cells",
    ),
    (
        "mulrk-over-number-then-reserved-xf",
        "Invalid record 0x00E0: cell references reserved style-XF slot 3",
    ),
    (
        "unmaterialized-member-behind-orphan",
        "Invalid record 0x0006: Formula at (3, 3) contains an orphan PtgExp",
    ),
    (
        "unmaterialized-member",
        "Invalid record 0x0221: Array Formula cell was not materialized",
    ),
    (
        "non-formula-member-with-duplicates",
        "Invalid record 0x0221: Array owner cannot attach to a non-Formula cell",
    ),
    (
        "outside-grid-duplicates-then-bad-xf",
        "Invalid record 0x00E0: cell references out-of-range XF 4095",
    ),
    (
        "custom-view-end-without-begin-then-bad-xf",
        "Invalid record 0x01AB: UserSViewEnd without a matching UserSViewBegin",
    ),
    (
        "dval-cut-short-by-a-cell-then-bad-xf",
        "Invalid record 0x0203: DVAL must be followed immediately by its declared DV records",
    ),
    (
        "dval-cut-short-by-eof-after-duplicates",
        "Invalid record 0x000A: DVAL must be followed immediately by its declared DV records",
    ),
    ("valid-duplicates-shared-and-array", "ok"),
    ("no-eof-pending-string-formula-after-duplicates", "ok"),
];

/// Every multi-defect case refuses first with the base's exact refusal under
/// the public reader's store and under the validation-only store, keeping no
/// cell and keeping every cell the case names; accepted cases stay accepted.
#[test]
fn worksheet_first_error_matrix_matches_the_base_for_multi_defect_inputs() {
    let bytes = {
        let mut writer = crate::Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_number(sheet, 0, 0, 1.0).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    };
    let full = Workbook::new(Cursor::new(bytes.as_slice())).unwrap();
    let ok = full
        .xls_worksheet(0)
        .unwrap()
        .get_cell(0, 0)
        .unwrap()
        .xf_index();
    assert_eq!(ok, 15, "the base matrix was generated with cell XF 15");
    let inputs = inputs(&full);
    let cases = multi_defect_cases::cases(ok);
    assert_eq!(cases.len(), BASE_MULTI_DEFECT_MATRIX.len());
    for ((name, records), (expected_name, expected)) in cases.iter().zip(BASE_MULTI_DEFECT_MATRIX) {
        assert_eq!(name, expected_name);
        let stream = encode(records);
        let mut every_position = records
            .iter()
            .filter(|(kind, payload)| is_cell(*kind) && payload.len() >= 4)
            .map(|(_, payload)| {
                (
                    u16::from_le_bytes([payload[0], payload[1]]),
                    u16::from_le_bytes([payload[2], payload[3]]),
                )
            })
            .collect::<Vec<_>>();
        every_position.sort_unstable();
        every_position.dedup();
        let outcome = |result: crate::Result<Worksheet>| match result {
            Ok(_) => "ok".to_string(),
            Err(error) => error.to_string(),
        };
        let parse = |store: &mut dyn FnMut(&mut Records<'_>) -> crate::Result<Worksheet>| {
            let mut framed = Records::new(&stream);
            outcome(store(&mut framed))
        };
        let complete = parse(&mut |framed| {
            Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records_with_compatibility(
                framed,
                stream.len() as u64,
                0,
                0,
                &inputs.encoding,
                "Matrix",
                std::sync::Arc::clone(&inputs.shared_strings),
                std::sync::Arc::clone(&inputs.properties),
                Some(inputs.formula_context),
                std::sync::Arc::clone(&inputs.formatting),
                CompatibilityProfile::Strict,
                &mut DecodeEveryCell,
            )
        });
        for kept in [&[][..], every_position.as_slice()] {
            let validated = parse(&mut |framed| {
                Workbook::<Cursor<Vec<u8>>>::parse_worksheet_records_with_compatibility(
                    framed,
                    stream.len() as u64,
                    0,
                    0,
                    &inputs.encoding,
                    "Matrix",
                    std::sync::Arc::clone(&inputs.shared_strings),
                    std::sync::Arc::clone(&inputs.properties),
                    Some(inputs.formula_context),
                    std::sync::Arc::clone(&inputs.formatting),
                    CompatibilityProfile::Strict,
                    &mut ValidateCells::new(kept),
                )
            });
            assert_eq!(
                validated, *expected,
                "{name}: validation-only, kept {kept:?}"
            );
        }
        assert_eq!(complete, *expected, "{name}: complete reader");
    }
}

/// A position kept on one tab is not kept on another: the validation-only
/// open refuses to answer it there instead of reporting it as vacant.
#[test]
fn kept_cell_refuses_a_position_kept_on_a_different_tab() {
    let bytes = {
        let mut writer = crate::Writer::new();
        let first = writer.add_worksheet("First").unwrap();
        writer.write_number(first, 1, 1, 10.0).unwrap();
        let second = writer.add_worksheet("Second").unwrap();
        writer.write_number(second, 1, 1, 20.0).unwrap();
        writer.write_number(second, 2, 2, 30.0).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    };
    let full = Workbook::new(Cursor::new(bytes.as_slice())).unwrap();
    let first = full.sheet(0).unwrap().parsed_worksheet_index().unwrap();
    let second = full.sheet(1).unwrap().parsed_worksheet_index().unwrap();
    let validated = Workbook::validation_only(
        Cursor::new(bytes.as_slice()),
        KeptCells::from_cells([(1, 1, 1)]).unwrap(),
    )
    .unwrap();
    let kept = validated.kept_cell(second, 1, 1).unwrap().unwrap();
    assert_eq!(
        format!("{kept:?}"),
        format!(
            "{:?}",
            full.xls_worksheet(second).unwrap().get_cell(1, 1).unwrap()
        )
    );
    for (worksheet, row, column) in [(first, 1, 1), (first, 2, 2), (second, 2, 2)] {
        assert!(
            matches!(
                validated.kept_cell(worksheet, row, column),
                Err(crate::Error::UnsafeEdit(message)) if message.contains("not asked to keep")
            ),
            "worksheet {worksheet} ({row}, {column})"
        );
    }
    // Neither tab decoded anything it was not asked to keep.
    assert_eq!(
        validated
            .as_workbook_for_tests()
            .xls_worksheet(first)
            .unwrap()
            .cell_count_for_tests(),
        0
    );
    assert_eq!(
        validated
            .as_workbook_for_tests()
            .xls_worksheet(second)
            .unwrap()
            .cell_count_for_tests(),
        1
    );
}
