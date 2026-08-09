use super::*;
use litchi_cfb::{OleFile, OleWriter};

fn package() -> Vec<u8> {
    let mut writer = crate::Writer::new();
    let sheet = writer.add_worksheet("Sheet1").unwrap();
    writer.write_number(sheet, 3, 2, 4.5).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut package =
        PackageEditor::open(output.into_inner(), Targets::default(), Limits::default()).unwrap();
    package
        .add_stream(vec!["Opaque".to_string()], b"untouched".to_vec())
        .unwrap();
    package.finish().unwrap()
}

fn signed_package() -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(package())).unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let opaque = ole.open_stream(&["Opaque"]).unwrap();
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer.create_stream(&["Opaque"], &opaque).unwrap();
    writer
        .create_stream(&["DigitalSignature"], b"signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn stream(bytes: &[u8], name: &str) -> Vec<u8> {
    OleFile::open(Cursor::new(bytes))
        .unwrap()
        .open_stream(&[name])
        .unwrap()
}

#[test]
fn edits_only_one_number_field_and_round_trips_patch() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(3, 2).unwrap();
    let worksheet = source.worksheet("Sheet1".into()).unwrap().unwrap();
    assert_eq!(worksheet.position(), 0);
    assert_eq!(worksheet.number(reference).unwrap().unwrap().value(), 4.5);

    let before_workbook = source.workbook_stream().to_vec();
    let mut edit = source.edit();
    edit.set_number(0usize.into(), reference, 9.25).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().changed_number_fields(), 1);
    assert_eq!(commit.diagnostics().touched_streams(), 1);
    let after_workbook = commit.snapshot().workbook_stream();
    assert_eq!(before_workbook.len(), after_workbook.len());
    assert_eq!(
        before_workbook
            .iter()
            .zip(after_workbook)
            .filter(|(left, right)| left != right)
            .count(),
        2
    );
    assert_eq!(stream(source.bytes(), "Opaque"), b"untouched");
    assert_eq!(stream(commit.snapshot().bytes(), "Opaque"), b"untouched");
    assert_eq!(
        source
            .worksheet("Sheet1".into())
            .unwrap()
            .unwrap()
            .number(reference)
            .unwrap()
            .unwrap()
            .value(),
        4.5
    );

    let applied = commit.patch().apply(&source).unwrap();
    let value = applied
        .worksheet("Sheet1".into())
        .unwrap()
        .unwrap()
        .number(reference)
        .unwrap()
        .unwrap()
        .value();
    assert_eq!(value, 9.25);
    assert!(commit.patch().apply(&applied).is_err());
    assert_eq!(
        commit.patch().inverse().apply(&applied).unwrap().bytes(),
        source.bytes()
    );
}

#[test]
fn rejected_and_noop_edits_are_failure_atomic() {
    let source = Snapshot::from_bytes(package()).unwrap();
    let reference = Reference::new(3, 2).unwrap();
    let mut edit = source.edit();
    assert!(
        edit.set_number("Sheet1".into(), reference, f64::NAN)
            .is_err()
    );
    assert!(
        edit.set_number("Sheet1".into(), reference, f64::INFINITY)
            .is_err()
    );
    assert!(edit.set_number("Sheet1".into(), reference, -0.0).is_err());
    assert!(
        edit.set_number("Sheet1".into(), reference, f64::MIN_POSITIVE / 2.0)
            .is_err()
    );
    assert!(
        edit.set_number("Sheet1".into(), Reference::new(9, 9).unwrap(), 1.0)
            .is_err()
    );
    edit.set_number("sHeEt1".into(), reference, 0.0).unwrap();
    edit.set_number("Sheet1".into(), reference, 4.5).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().bytes(), source.bytes());
    assert_eq!(commit.diagnostics(), Diagnostics::default());
}

#[test]
fn references_enforce_the_biff8_grid() {
    let reference = Reference::new(u32::from(u16::MAX), u32::from(u8::MAX)).unwrap();
    assert_eq!(reference.row(), u16::MAX);
    assert_eq!(reference.column(), u8::MAX);
    assert!(Reference::new(u32::from(u16::MAX) + 1, 0).is_err());
    assert!(Reference::new(0, u32::from(u8::MAX) + 1).is_err());
}

#[test]
fn signed_packages_are_refused_before_editing() {
    assert!(Snapshot::from_bytes(signed_package()).is_err());
}
