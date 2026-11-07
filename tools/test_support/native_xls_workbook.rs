//! Native XLS test setup through the public CFB owner with exact offset relocation.
use litchi_biff::{Limits, Records};
use litchi_core::sheet::Cell as _;
use litchi_xls::Package;
use std::{collections::BTreeMap, io::Cursor, path::Path};
#[path = "cfb_directory_metadata.rs"]
pub(crate) mod cfb_metadata;
pub(crate) fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(file),
    )
    .unwrap()
}
pub(crate) fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|name| {
            let data = ole
                .open_stream(&name.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (name, data)
        })
        .collect()
}
pub(crate) fn workbook(p: &Package) -> Vec<u8> {
    streams(&p.to_bytes().unwrap())[&vec!["Workbook".into()]].clone()
}
pub(crate) fn frame(kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut r = kind.to_le_bytes().to_vec();
    r.extend(u16::try_from(payload.len()).unwrap().to_le_bytes());
    r.extend_from_slice(payload);
    r
}
pub(crate) fn records(bytes: &[u8]) -> Vec<(usize, u16, Vec<u8>)> {
    Records::new(bytes, Limits::default())
        .unwrap()
        .map(|r| {
            let r = r.unwrap();
            (r.offset(), r.kind().get(), r.payload().to_vec())
        })
        .collect()
}
/// The chosen fixtures use BoundSheet8, Index and ExtSST absolute offsets.
/// Keep every other source byte intact when installing the explicit record
/// derivative. This is raw test setup via the public CFB owner; the caller
/// remains responsible for placing records in the appropriate BIFF production.
pub(crate) fn publish_range(p: &Package, start: usize, end: usize, replacement: &[u8]) -> Package {
    use litchi_ole_common::object;
    let before = p.to_bytes().unwrap();
    let original = workbook(p);
    let mut output = original[..start].to_vec();
    output.extend_from_slice(replacement);
    output.extend_from_slice(&original[end..]);
    let delta = i64::try_from(replacement.len()).unwrap() - i64::try_from(end - start).unwrap();
    let mut checks = Vec::new();
    for (at, kind, payload) in records(&output) {
        let fields = match kind {
            0x0085 => vec![0],
            0x020b => (12..payload.len()).step_by(4).collect(),
            0x00ff => (2..payload.len()).step_by(8).collect(),
            _ => Vec::new(),
        };
        for field in fields {
            let old = u32::from_le_bytes(payload[field..field + 4].try_into().unwrap()) as usize;
            if old == 0 {
                continue;
            }
            assert!(
                old < start || old >= end,
                "a source offset targets replaced metadata"
            );
            assert!(old + 4 <= original.len());
            let new = if old >= end {
                usize::try_from(i64::try_from(old).unwrap() + delta).unwrap()
            } else {
                old
            };
            output[at + 4 + field..at + 8 + field]
                .copy_from_slice(&u32::try_from(new).unwrap().to_le_bytes());
            checks.push((old, new));
        }
    }
    for (old, new) in checks {
        assert_eq!(&output[new..new + 4], &original[old..old + 4]);
    }
    let source = object::Package::open(before.clone()).unwrap();
    let mut edit = source.edit().unwrap();
    edit.put_stream(&["Workbook".into()], output.clone())
        .unwrap();
    let c = edit.commit().unwrap();
    let q = Package::from_bytes(c.patch().after()).unwrap();
    let file = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
    q.save(file.path()).unwrap();
    let q = Package::open(file.path()).unwrap();
    assert_eq!(workbook(&q), output);
    let after = q.to_bytes().unwrap();
    assert_eq!(
        cfb_metadata::directory(&before),
        cfb_metadata::directory(&after)
    );
    let a = streams(&before);
    let b = streams(&after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, data) in a {
        if name != ["Workbook"] {
            assert!(b.get(&name) == Some(&data), "changed stream {name:?}");
        }
    }
    let original_cell = p
        .reader_ref()
        .unwrap()
        .xls_worksheet(0)
        .unwrap()
        .get_cell(0, 0)
        .unwrap()
        .value()
        .clone();
    assert_eq!(
        q.reader_ref()
            .unwrap()
            .xls_worksheet(0)
            .unwrap()
            .get_cell(0, 0)
            .unwrap()
            .value(),
        &original_cell
    );
    assert_eq!(
        q.reader_ref().unwrap().date_system(),
        p.reader_ref().unwrap().date_system()
    );
    assert_eq!(c.patch().apply(&source).unwrap().patch().after(), after);
    let target = object::Package::open(after).unwrap();
    assert_eq!(
        c.patch().inverse().apply(&target).unwrap().patch().after(),
        before
    );
    q
}
