//! Drawing edits over workbooks whose drawing records are corrupted: the
//! drawing group (`MsoDrawingGroup`), the first worksheet's `MsoDrawing`
//! records and its `Obj` records. Every length, offset and count the edits
//! read from them is hostile, and each worksheet-feature and OLE-object
//! edit must return a typed error or succeed, never panic.

use litchi_cfb::{OleFile, OleWriter};
use litchi_xls::sheet_features::EquationKind;
use litchi_xls::sheet_features::edit::{
    Anchor, EmbeddedObject, EquationContent, EquationSpec, Picture,
};
use litchi_xls::{CheckState, FormControl, FtCblsData, ObjSubrecord, ObjectType, Package};
use litchi_xls::{EmbeddedObjectDraft, EmbeddedPayload};
use std::io::Cursor;
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::PathBuf;

const PNG: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0xe9, 0xfa, 0xdc, 0xd8, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
    0xae, 0x42, 0x60, 0x82,
];

/// The streams of a compound file, by path.
pub(crate) fn streams(bytes: &[u8]) -> Vec<(Vec<String>, Vec<u8>)> {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let parts: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = ole.open_stream(&parts).unwrap();
            (path, data)
        })
        .collect()
}

/// A compound file of `streams` with its Workbook stream replaced.
pub(crate) fn with_workbook(streams: &[(Vec<String>, Vec<u8>)], workbook: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for (path, data) in streams {
        let parts: Vec<&str> = path.iter().map(String::as_str).collect();
        for depth in 1..parts.len() {
            let _ = writer.create_storage(&parts[..depth]);
        }
        let data = if parts == ["Workbook"] {
            workbook
        } else {
            data
        };
        writer.create_stream(&parts, data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

/// The payload byte ranges of the globals' `MsoDrawingGroup` record and
/// the `Continue` records after it, and of the first worksheet's
/// `MsoDrawing` and `Obj` records.
pub(crate) fn drawing_ranges(
    workbook: &[u8],
) -> (
    Vec<(usize, usize)>,
    Vec<(usize, usize)>,
    Vec<(usize, usize)>,
) {
    let mut group = Vec::new();
    let mut drawing = Vec::new();
    let mut objects = Vec::new();
    let mut at = 0;
    let mut in_group = false;
    let mut substream = 0;
    while at + 4 <= workbook.len() {
        let kind = u16::from_le_bytes([workbook[at], workbook[at + 1]]);
        let length = usize::from(u16::from_le_bytes([workbook[at + 2], workbook[at + 3]]));
        let body = at + 4;
        if kind == 0x0809 {
            substream += 1;
        }
        if substream == 1 && (kind == 0x00eb || (in_group && kind == 0x003c)) {
            in_group = true;
            group.push((body, body + length));
        } else {
            in_group = false;
        }
        if substream == 2 && kind == 0x00ec {
            drawing.push((body, body + length));
        }
        if substream == 2 && kind == 0x005d {
            objects.push((body, body + length));
        }
        at = body + length;
    }
    (group, drawing, objects)
}

/// The byte ranges to corrupt and, as offsets into their concatenation,
/// the OfficeArt record headers of the drawing group and the worksheet
/// drawing.
pub(crate) fn targets(workbook: &[u8]) -> (Vec<(usize, usize)>, Vec<usize>) {
    let (group, drawing, objects) = drawing_ranges(workbook);
    let concat = |ranges: &[(usize, usize)]| -> Vec<u8> {
        ranges
            .iter()
            .flat_map(|(start, end)| workbook[*start..*end].iter().copied())
            .collect()
    };
    let group_bytes = concat(&group);
    let mut headers = record_headers(&group_bytes);
    headers.extend(
        record_headers(&concat(&drawing))
            .into_iter()
            .map(|at| at + group_bytes.len()),
    );
    let mut ranges = group;
    ranges.extend(drawing);
    ranges.extend(objects);
    (ranges, headers)
}

/// A deterministic xorshift generator.
pub(crate) struct Rng(u64);

impl Rng {
    pub(crate) fn new(seed: u64) -> Self {
        Self(if seed == 0 { 1 } else { seed })
    }

    pub(crate) fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    pub(crate) fn below(&mut self, bound: usize) -> usize {
        (self.next() % bound as u64) as usize
    }
}

/// The offsets, in concatenated OfficeArt records, of their record headers
/// (containers are walked into).
pub(crate) fn record_headers(group: &[u8]) -> Vec<usize> {
    fn walk(group: &[u8], start: usize, end: usize, out: &mut Vec<usize>) {
        let mut at = start;
        while at + 8 <= end && out.len() < 4_096 {
            out.push(at);
            let container = group[at] & 0x0f == 0x0f;
            let length = u32::from_le_bytes(group[at + 4..at + 8].try_into().unwrap()) as usize;
            let body_end = (at + 8).saturating_add(length).min(end);
            if container {
                walk(group, at + 8, body_end, out);
            }
            at = body_end;
        }
    }
    let mut out = Vec::new();
    walk(group, 0, group.len(), &mut out);
    out
}

/// `workbook` with one to four corruptions inside `ranges`: a random byte,
/// a flipped bit, or an extreme 32-bit value (the shape of a record length
/// or count), half of them aimed at an OfficeArt record header or the
/// fields after it.
pub(crate) fn mutate(
    workbook: &[u8],
    ranges: &[(usize, usize)],
    headers: &[usize],
    rng: &mut Rng,
) -> Vec<u8> {
    let mut out = workbook.to_vec();
    let positions: usize = ranges.iter().map(|(start, end)| end - start).sum();
    for _ in 0..=rng.below(4) {
        let mut index = if rng.below(2) == 0 && !headers.is_empty() {
            (headers[rng.below(headers.len())] + rng.below(48)).min(positions - 1)
        } else {
            rng.below(positions)
        };
        let (start, end) = ranges
            .iter()
            .copied()
            .find(|(start, end)| {
                let fits = index < end - start;
                if !fits {
                    index -= end - start;
                }
                fits
            })
            .unwrap();
        let at = start + index;
        match rng.below(3) {
            0 => out[at] = rng.next() as u8,
            1 => out[at] ^= 1 << rng.below(8),
            _ => {
                let value: u32 =
                    [0, 1, 0x7fff_ffff, 0xffff_ffff, 0xcafe_babe, 0x0001_0000][rng.below(6)];
                for (offset, byte) in value.to_le_bytes().into_iter().enumerate() {
                    if at + offset < end {
                        out[at + offset] = byte;
                    }
                }
            },
        }
    }
    out
}

fn anchor() -> Anchor {
    Anchor {
        first_row: 4,
        first_column: 1,
        last_row: 8,
        last_column: 3,
        behavior: litchi_xls::AnchorBehavior::MoveWithoutResize,
    }
}

fn equation_native() -> Vec<u8> {
    let mtef = [
        3u8, 1, 1, 3, 0, 0x0A, 0x01, 0x02, 0x88, 0x32, 0x00, 0x00, 0x00,
    ];
    let mut out = vec![28, 0, 0, 0, 2, 0, 0x6D, 0xC1];
    out.extend_from_slice(&u32::try_from(mtef.len()).unwrap().to_le_bytes());
    out.extend_from_slice(&[0; 16]);
    out.extend_from_slice(&mtef);
    out
}

fn payload() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["CONTENTS"], b"inert").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn checkbox() -> FormControl {
    FormControl::new(
        ObjectType::CheckBox,
        1,
        [ObjSubrecord::CheckBoxData(FtCblsData {
            state: CheckState::Checked,
            accelerator: 0,
            reserved: 0,
            flags: 0,
        })],
    )
}

/// Run every drawing edit on `package`, each in its own edit, and commit
/// the ones that succeed. Errors are fine; panics are not.
pub(crate) fn exercise(package: &Package) {
    type Step = fn(&mut litchi_xls::sheet_features::edit::Edit) -> bool;
    let steps: [Step; 7] = [
        |edit| {
            edit.add_picture(
                0,
                &Picture {
                    data: PNG.to_vec(),
                    anchor: anchor(),
                    name: None,
                    description: None,
                    rotation: 0.0,
                    flip_h: false,
                    flip_v: false,
                },
            )
            .is_ok()
        },
        |edit| {
            edit.add_equation(
                0,
                &EquationSpec {
                    kind: EquationKind::Equation3,
                    content: EquationContent {
                        native: equation_native(),
                        preview: PNG.to_vec(),
                        size_pt: (20.0, 14.0),
                    },
                    row: 3,
                    column: 2,
                },
            )
            .is_ok()
        },
        |edit| {
            edit.set_equation(
                0,
                0,
                &EquationContent {
                    native: equation_native(),
                    preview: PNG.to_vec(),
                    size_pt: (20.0, 14.0),
                },
            )
            .is_ok()
        },
        |edit| edit.remove_equation(0, 0).is_ok(),
        |edit| edit.remove_drawing(0, true, 0).is_ok(),
        |edit| {
            edit.add_embedded_object(
                0,
                &EmbeddedObject {
                    compound_file: payload(),
                    prog_id: "Package".to_owned(),
                    preview: PNG.to_vec(),
                    anchor: anchor(),
                },
            )
            .is_ok()
        },
        |edit| {
            let ids = edit
                .features(0, litchi_xls::sheet_features::FeatureSet::DRAWINGS)
                .map(|_| ())
                .is_ok();
            ids && (1..4).any(|id| edit.remove_embedded_object(0, id).is_ok())
        },
    ];
    for step in steps {
        let Ok(mut edit) = package.edit_sheet_features() else {
            break;
        };
        if step(&mut edit) {
            let _ = edit.commit();
        }
    }
    type OleStep = fn(&mut litchi_xls::ole_object::Edit) -> bool;
    let ole_steps: [OleStep; 3] = [
        |edit| {
            let ids = edit.objects(0).map(|objects| {
                objects
                    .iter()
                    .map(litchi_xls::OleObjectRecord::object_id)
                    .collect::<Vec<_>>()
            });
            ids.is_ok_and(|ids| ids.into_iter().any(|id| edit.remove_object(0, id).is_ok()))
        },
        |edit| edit.add_form_control(0, checkbox(), anchor()).is_ok(),
        |edit| {
            let Ok(payload) =
                EmbeddedPayload::new(payload(), litchi_xls::ole_object::Limits::default())
            else {
                return false;
            };
            edit.add_embedded_payload(0, EmbeddedObjectDraft::new(payload, anchor(), PNG.to_vec()))
                .is_ok()
        },
    ];
    for step in ole_steps {
        let Ok(mut edit) = package.edit_ole_objects() else {
            return;
        };
        if step(&mut edit) {
            let _ = edit.commit();
        }
    }
}

/// Run `cases` corruptions of `fixture`'s drawing records; returns the
/// panicking cases' seeds.
pub(crate) fn run(fixture: &[u8], seed: u64, cases: usize) -> Vec<u64> {
    let streams = streams(fixture);
    let workbook = &streams
        .iter()
        .find(|(path, _)| path == &["Workbook"])
        .unwrap()
        .1;
    let (ranges, headers) = targets(workbook);
    assert!(!headers.is_empty(), "the fixture has a drawing");
    let mut panics = Vec::new();
    for case in 0..cases {
        let case_seed = seed.wrapping_mul(0x9e37_79b9_7f4a_7c15) ^ case as u64;
        let mut rng = Rng::new(case_seed);
        let mutated = mutate(workbook, &ranges, &headers, &mut rng);
        let bytes = with_workbook(&streams, &mutated);
        let Ok(package) = Package::from_bytes(&bytes) else {
            continue;
        };
        if catch_unwind(AssertUnwindSafe(|| exercise(&package))).is_err() {
            panics.push(case_seed);
        }
    }
    panics
}

pub(crate) fn fixture(path: &str) -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(path),
    )
    .unwrap()
}

#[test]
fn corrupted_drawings_never_panic_an_edit() {
    for (path, seed) in [
        ("test-data/ole/xls/WithEmbeddedObjects.xls", 0x5eed_0001),
        (
            "test-data/microsoft-office/excel/excel-links-objects.xls",
            0x5eed_0002,
        ),
    ] {
        let panics = run(&fixture(path), seed, 200);
        assert!(panics.is_empty(), "{path}: panicking seeds {panics:x?}");
    }
}

/// Overwrite bytes at `offset` of the concatenation of `ranges`.
fn patch_concat(workbook: &mut [u8], ranges: &[(usize, usize)], offset: usize, bytes: &[u8]) {
    for (index, byte) in bytes.iter().enumerate() {
        let mut at = offset + index;
        for (start, end) in ranges {
            if at < end - start {
                workbook[start + at] = *byte;
                break;
            }
            at -= end - start;
        }
    }
}

/// `WithEmbeddedObjects.xls` with `change` applied to its Workbook stream.
fn crafted(change: impl FnOnce(&mut Vec<u8>)) -> Package {
    let streams = streams(&fixture("test-data/ole/xls/WithEmbeddedObjects.xls"));
    let mut workbook = streams
        .iter()
        .find(|(path, _)| path == &["Workbook"])
        .unwrap()
        .1
        .clone();
    change(&mut workbook);
    Package::from_bytes(&with_workbook(&streams, &workbook)).unwrap()
}

fn u32_le(bytes: &[u8], at: usize) -> usize {
    u32::from_le_bytes(bytes[at..at + 4].try_into().unwrap()) as usize
}

fn object_ids(package: &Package) -> Vec<u16> {
    package
        .ole_objects()
        .unwrap()
        .objects(0)
        .unwrap()
        .iter()
        .map(litchi_xls::OleObjectRecord::object_id)
        .collect()
}

/// Remove every embedded object through both editors: an error is fine, a
/// panic is not.
fn remove_each_object(package: &Package) {
    for id in object_ids(package) {
        let mut features = package.edit_sheet_features().unwrap();
        if features.remove_embedded_object(0, id).is_ok() {
            let _ = features.commit();
        }
        let mut ole = package.edit_ole_objects().unwrap();
        if ole.remove_object(0, id).is_ok() {
            let _ = ole.commit();
        }
    }
}

#[test]
fn a_preview_entry_that_is_not_an_fbse_is_refused_on_removal() {
    // The last blip store entry becomes an empty 0xF007 record followed by
    // a record holding the rest: the store is still exactly filled, and the
    // last shape's `pib` names the empty entry.
    let package = crafted(|workbook| {
        let (group, _, _) = drawing_ranges(workbook);
        let bytes: Vec<u8> = group
            .iter()
            .flat_map(|(start, end)| workbook[*start..*end].iter().copied())
            .collect();
        let mut at = 8;
        let mut store = None;
        while at + 8 <= bytes.len() {
            if u16::from_le_bytes([bytes[at + 2], bytes[at + 3]]) == 0xf001 {
                store = Some(at);
                break;
            }
            at += 8 + u32_le(&bytes, at + 4);
        }
        let store = store.unwrap();
        let store_end = store + 8 + u32_le(&bytes, store + 4);
        let mut entry = store + 8;
        let mut last = entry;
        while entry < store_end {
            last = entry;
            entry += 8 + u32_le(&bytes, entry + 4);
        }
        let length = u32_le(&bytes, last + 4);
        let mut patch = vec![0x02, 0x00, 0x07, 0xf0, 0, 0, 0, 0];
        patch.extend_from_slice(&[0x00, 0x00, 0x00, 0xf0]);
        patch.extend_from_slice(&u32::try_from(length - 8).unwrap().to_le_bytes());
        patch_concat(workbook, &group, last, &patch);
    });
    remove_each_object(&package);
    let ids = object_ids(&package);
    let mut features = package.edit_sheet_features().unwrap();
    assert!(features.remove_embedded_object(0, ids[1]).is_err());
}

#[test]
fn a_shape_that_runs_past_the_drawing_is_not_rewritten() {
    let package = crafted(|workbook| {
        let (_, drawing, _) = drawing_ranges(workbook);
        let bytes: Vec<u8> = drawing
            .iter()
            .flat_map(|(start, end)| workbook[*start..*end].iter().copied())
            .collect();
        // The patriarch group's children: the patriarch, then the shapes.
        let mut at = 24 + 8;
        let mut seen = 0;
        while at + 8 <= bytes.len() {
            if u16::from_le_bytes([bytes[at + 2], bytes[at + 3]]) == 0xf004 {
                seen += 1;
                if seen == 3 {
                    patch_concat(workbook, &drawing, at + 4, &0x7fff_fff0u32.to_le_bytes());
                    break;
                }
            }
            at += 8 + u32_le(&bytes, at + 4);
        }
        assert_eq!(seen, 3);
    });
    remove_each_object(&package);
    for id in object_ids(&package) {
        let mut features = package.edit_sheet_features().unwrap();
        assert!(features.remove_embedded_object(0, id).is_err());
    }
}
