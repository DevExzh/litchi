//! Native worksheet cell-watch operations; phonetic additions remain inert.
use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xlsb::{
    Package,
    cell_watches::{self, Alignment, Info, Reference, Type, Watch, Watches},
    raw::{Limits, Records, kind},
};
use std::collections::BTreeMap;
const SHEET: &str = "xl/worksheets/sheet1.bin";
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-metadata.xlsb"),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn changed(a: &[u8], b: &[u8]) -> Vec<String> {
    let a = members(a);
    let b = members(b);
    let names: std::collections::BTreeSet<_> = a.keys().chain(b.keys()).collect();
    names
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect()
}

fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}
fn without_owned(bytes: &[u8]) -> Vec<u8> {
    let mut it = Records::new(bytes, Limits::DEFAULT);
    let mut out = Vec::new();
    while let Some(r) = it.next() {
        let r = r.unwrap();
        if ![
            kind::BEGIN_CELL_WATCHES,
            kind::CELL_WATCH,
            kind::END_CELL_WATCHES,
            kind::PHONETIC_INFO,
        ]
        .contains(&r.kind())
        {
            out.extend_from_slice(&bytes[r.offset()..it.offset()]);
        }
    }
    out
}
fn caches(p: &Package) {
    let w = p.workbook().unwrap();
    let s = w.worksheet(0).unwrap();
    for (row, expected) in [
        (3, CellValue::Error("#VALUE!".into())),
        (4, CellValue::Error("#DIV/0!".into())),
        (6, CellValue::Float(30.)),
    ] {
        let CellValue::Formula {
            cached_value: Some(v),
            ..
        } = s.get_cell(row, 0).unwrap().value()
        else {
            panic!("formula cache")
        };
        assert_eq!(**v, expected);
    }
}
fn exact_except_owned(before: &[u8], after: &[u8]) {
    assert_eq!(changed(before, after), [SHEET]);
    assert_eq!(
        without_owned(&members(before)[SHEET]),
        without_owned(&members(after)[SHEET])
    );
}

#[test]
fn native_watch_reads_exact_a7_and_keeps_worksheet_error_caches() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let parts = members(&bytes);
    let edit = p.edit_cell_watches().unwrap();
    assert_eq!(
        edit.watches(0).unwrap().as_slice(),
        [Watch::new(6, 0).unwrap()]
    );
    assert_eq!(edit.phonetic(0).unwrap(), None);
    let snapshot = cell_watches::worksheet::read(&parts[SHEET], Limits::DEFAULT).unwrap();
    assert!(snapshot.has_collection());
    assert!(!snapshot.has_opaque_records());
    assert_eq!(snapshot.source_bytes(), parts[SHEET]);
    assert_eq!(snapshot.watches(), edit.watches(0).unwrap().as_slice());
    for i in 1..5 {
        assert!(edit.watches(i).unwrap().is_empty());
        assert_eq!(edit.phonetic(i).unwrap(), None);
    }
    caches(&p);
    assert_eq!(edit.commit().unwrap().package().to_bytes().unwrap(), bytes);
}

#[test]
fn native_watch_crud_preserves_every_unowned_record_and_failed_staging() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let mut e = p.edit_cell_watches().unwrap();
    let added = Watch::new(4, 0).unwrap();
    e.add(0, added).unwrap();
    assert!(matches!(
        e.add(0, added),
        Err(litchi_xlsb::package::PackageError::Unrecognized {
            typ: "BrtCellWatch collection",
            ..
        })
    ));
    assert_eq!(
        e.watches(0).unwrap().as_slice(),
        [Watch::new(6, 0).unwrap(), added]
    );
    let c = e.commit().unwrap();
    let after = save(c.package());
    let out = after.to_bytes().unwrap();
    exact_except_owned(&bytes, &out);
    caches(&after);
    assert_eq!(
        after
            .edit_cell_watches()
            .unwrap()
            .watches(0)
            .unwrap()
            .as_slice(),
        [Watch::new(6, 0).unwrap(), added]
    );
    assert!(
        members(
            &p.apply_cell_watches(c.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&out)
    );
    assert!(matches!(
        after.apply_cell_watches(c.patch()),
        Err(litchi_xlsb::package::PackageError::PatchConflict { .. })
    ));
    assert!(
        members(
            &after
                .apply_cell_watches(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&bytes)
    );
    let mut e = after.edit_cell_watches().unwrap();
    e.replace(
        0,
        Watches::new(vec![added, Watch::new(6, 0).unwrap()]).unwrap(),
    )
    .unwrap();
    assert!(e.add(99, Watch::new(0, 0).unwrap()).is_err());
    assert_eq!(
        e.watches(0).unwrap().as_slice(),
        [added, Watch::new(6, 0).unwrap()]
    );
    let c = e.commit().unwrap();
    let reordered = save(c.package());
    assert_eq!(
        reordered
            .edit_cell_watches()
            .unwrap()
            .watches(0)
            .unwrap()
            .as_slice(),
        [added, Watch::new(6, 0).unwrap()]
    );
    exact_except_owned(&out, &reordered.to_bytes().unwrap());
    let mut e = reordered.edit_cell_watches().unwrap();
    assert!(e.remove(0, Reference::new(4, 0).unwrap()).unwrap());
    assert!(!e.remove(0, Reference::new(4, 0).unwrap()).unwrap());
    let c = e.commit().unwrap();
    let restored = save(c.package());
    assert_eq!(members(&restored.to_bytes().unwrap()), members(&bytes));
    let mut e = restored.edit_cell_watches().unwrap();
    assert!(e.clear(0).unwrap());
    let c = e.commit().unwrap();
    let empty = save(c.package());
    assert!(
        empty
            .edit_cell_watches()
            .unwrap()
            .watches(0)
            .unwrap()
            .is_empty()
    );
    exact_except_owned(&bytes, &empty.to_bytes().unwrap());
    assert!(
        members(
            &empty
                .apply_cell_watches(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&bytes)
    );
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

#[test]
fn native_source_phonetic_default_crud_preserves_watches_text_and_saved_values() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let value = Info::new(0, Type::Hiragana, Alignment::Center);
    let mut e = p.edit_cell_watches().unwrap();
    e.set_phonetic(0, value).unwrap();
    let c = e.commit().unwrap();
    let after = save(c.package());
    let out = after.to_bytes().unwrap();
    exact_except_owned(&bytes, &out);
    caches(&after);
    assert_eq!(
        after.edit_cell_watches().unwrap().phonetic(0).unwrap(),
        Some(value)
    );
    assert_eq!(
        after
            .edit_cell_watches()
            .unwrap()
            .watches(0)
            .unwrap()
            .as_slice(),
        [Watch::new(6, 0).unwrap()]
    );
    assert!(
        members(
            &after
                .apply_cell_watches(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&bytes)
    );
    let mut e = after.edit_cell_watches().unwrap();
    e.set_phonetic(0, value).unwrap();
    assert_eq!(e.commit().unwrap().package().to_bytes().unwrap(), out);
    let changed_value = Info::new(0, Type::AsEntered, Alignment::Distribute);
    let mut e = after.edit_cell_watches().unwrap();
    e.set_phonetic(0, changed_value).unwrap();
    let c = e.commit().unwrap();
    let updated = save(c.package());
    assert_eq!(
        updated.edit_cell_watches().unwrap().phonetic(0).unwrap(),
        Some(changed_value)
    );
    exact_except_owned(&out, &updated.to_bytes().unwrap());
    let mut e = updated.edit_cell_watches().unwrap();
    assert!(e.clear_phonetic(0).unwrap());
    assert!(!e.clear_phonetic(0).unwrap());
    let c = e.commit().unwrap();
    let removed = save(c.package());
    assert_eq!(
        removed.edit_cell_watches().unwrap().phonetic(0).unwrap(),
        None
    );
    assert!(members(&removed.to_bytes().unwrap()) == members(&bytes));
    assert!(
        members(
            &removed
                .apply_cell_watches(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&updated.to_bytes().unwrap())
    );
    assert_eq!(p.to_bytes().unwrap(), bytes);
}
