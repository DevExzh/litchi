//! OLE equations of worksheet drawings: create, read, update and delete,
//! on a fresh workbook and on an Excel-authored workbook with embedded
//! objects, read back through this reader, the OLE-object view and the
//! compound file.

use litchi_cfb::OleFile;
use litchi_xls::Package;
use litchi_xls::sheet_features::edit::{EquationContent, EquationSpec, Error};
use litchi_xls::sheet_features::{EquationKind, FeatureSet};
use std::collections::BTreeMap;
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(name: &str) -> Package {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/xls")
        .join(name);
    Package::from_bytes(&std::fs::read(path).unwrap()).unwrap()
}

/// Every stream of the package's compound file, by path.
fn streams(package: &Package) -> BTreeMap<Vec<String>, Vec<u8>> {
    let bytes = package.to_bytes().unwrap();
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let mut out = BTreeMap::new();
    for path in ole.list_streams() {
        let parts: Vec<&str> = path.iter().map(String::as_str).collect();
        let data = ole.open_stream(&parts).unwrap();
        out.insert(path, data);
    }
    out
}

/// The streams outside the Workbook stream and the given storages.
fn others(
    streams: &BTreeMap<Vec<String>, Vec<u8>>,
    storages: &[&str],
) -> BTreeMap<Vec<String>, Vec<u8>> {
    streams
        .iter()
        .filter(|(path, _)| {
            path[0] != "Workbook" && !storages.iter().any(|storage| path[0] == *storage)
        })
        .map(|(path, data)| (path.clone(), data.clone()))
        .collect()
}

/// The package reopened from its bytes, with its worksheets parsed.
fn reopen(package: &Package) -> Package {
    let bytes = package.to_bytes().unwrap();
    let package = Package::from_bytes(&bytes).unwrap();
    package.reader_ref().unwrap().xls_worksheet(0).unwrap();
    package
}

const PNG: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0xe9, 0xfa, 0xdc, 0xd8, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
    0xae, 0x42, 0x60, 0x82,
];

/// An Equation Editor 3.0 `Equation Native` stream holding the digit 2.
fn equation3_native() -> Vec<u8> {
    let mtef = [
        3u8, 1, 1, 3, 0, 0x0A, 0x01, 0x02, 0x88, 0x32, 0x00, 0x00, 0x00,
    ];
    let mut out = vec![28, 0, 0, 0, 2, 0, 0x6D, 0xC1];
    out.extend_from_slice(&u32::try_from(mtef.len()).unwrap().to_le_bytes());
    out.extend_from_slice(&[0; 16]);
    out.extend_from_slice(&mtef);
    out
}

fn equation3(row: u16, column: u16) -> EquationSpec {
    EquationSpec {
        kind: EquationKind::Equation3,
        content: EquationContent {
            native: equation3_native(),
            preview: PNG.to_vec(),
            size_pt: (20.0, 14.0),
        },
        row,
        column,
    }
}

#[test]
fn raw_equation_round_trips_through_the_drawing_and_the_compound_file() {
    let package = Package::new().unwrap();
    let mut edit = package.edit_sheet_features().unwrap();
    assert_eq!(edit.add_equation(0, &equation3(2, 1)).unwrap(), 0);
    let staged = edit.equations(0).unwrap();
    assert_eq!(staged.len(), 1);
    let package = reopen(&edit.commit().unwrap().into_package());

    let equations = package.equations(0).unwrap();
    assert_eq!(equations.len(), 1);
    let equation = &equations[0];
    assert_eq!(equation.prog_id, "Equation.3");
    assert_eq!(equation.kind(), EquationKind::Equation3);
    assert_eq!(
        equation.native.as_deref(),
        Some(equation3_native().as_slice())
    );
    assert_eq!(equation.preview.unwrap().0, "image/png");
    let anchor = equation.anchor.unwrap();
    assert_eq!(
        (anchor.top_left().column(), anchor.top_left().row()),
        (1, 2)
    );
    assert_eq!(
        anchor.behavior(),
        litchi_xls::AnchorBehavior::MoveWithoutResize
    );

    // The equation is not a picture.
    let features = package.sheet_features(0, FeatureSet::DRAWINGS).unwrap();
    assert!(features.drawings.is_empty());

    // The storage is the object Equation Editor writes.
    let streams = streams(&package);
    let storage = &equation.storage;
    assert!(storage.starts_with("MBD"));
    let comp_obj = &streams[&vec![storage.clone(), "\u{1}CompObj".to_owned()]];
    assert!(String::from_utf8_lossy(comp_obj).contains("Equation.3\0"));
    assert!(streams.contains_key(&vec![storage.clone(), "\u{1}Ole".to_owned()]));
    let bytes = package.to_bytes().unwrap();
    let ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let entry = ole
        .list_directory_entries(&[])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == *storage)
        .unwrap();
    assert!(
        entry.clsid.to_ascii_uppercase().contains("0002CE02"),
        "{}",
        entry.clsid
    );

    // The OLE-object view reads the Obj as an ordinary MBD embedding.
    let objects = package.ole_objects().unwrap();
    let records = objects.objects(0).unwrap();
    assert_eq!(records.len(), 1);
    assert_eq!(records[0].storage_name().as_deref(), Some(storage.as_str()));
}

#[test]
fn update_changes_only_the_selected_equation() {
    let package = Package::new().unwrap();
    let mut edit = package.edit_sheet_features().unwrap();
    edit.add_equation(0, &equation3(0, 0)).unwrap();
    edit.add_equation(0, &equation3(5, 3)).unwrap();
    let package = edit.commit().unwrap().into_package();
    let before = package.equations(0).unwrap();
    let before_streams = streams(&package);

    let mut native = equation3_native();
    native[37] = 0x33; // the digit 3
    let preview = PNG.to_vec();
    let mut edit = package.edit_sheet_features().unwrap();
    edit.set_equation(
        0,
        1,
        &EquationContent {
            native: native.clone(),
            preview,
            size_pt: (60.0, 30.0),
        },
    )
    .unwrap();
    let package = reopen(&edit.commit().unwrap().into_package());
    let after = package.equations(0).unwrap();
    assert_eq!(after[0], before[0]);
    assert_eq!(after[1].native.as_deref(), Some(native.as_slice()));
    assert_eq!(after[1].prog_id, before[1].prog_id);
    assert_eq!(after[1].storage, before[1].storage);
    let (old, new) = (before[1].anchor.unwrap(), after[1].anchor.unwrap());
    assert_eq!(old.top_left(), new.top_left());
    assert_ne!(old.bottom_right(), new.bottom_right());

    // Every other stream keeps its bytes.
    let after_streams = streams(&package);
    assert_eq!(
        others(&before_streams, &[&before[1].storage]),
        others(&after_streams, &[&before[1].storage])
    );
    let storage = |streams: &BTreeMap<Vec<String>, Vec<u8>>| {
        streams
            .iter()
            .filter(|(path, _)| path[0] == before[1].storage && path[1] != "Equation Native")
            .map(|(path, data)| (path.clone(), data.clone()))
            .collect::<Vec<_>>()
    };
    assert_eq!(storage(&before_streams), storage(&after_streams));
}

#[test]
fn update_refuses_another_mtef_version_and_missing_targets() {
    let package = Package::new().unwrap();
    let mut edit = package.edit_sheet_features().unwrap();
    edit.add_equation(0, &equation3(0, 0)).unwrap();
    let mut native = equation3_native();
    native[28] = 5;
    let content = EquationContent {
        native,
        preview: PNG.to_vec(),
        size_pt: (10.0, 10.0),
    };
    assert!(matches!(
        edit.set_equation(0, 0, &content),
        Err(Error::Refused(_))
    ));
    let content = EquationContent {
        native: equation3_native(),
        ..content
    };
    assert!(matches!(
        edit.set_equation(0, 1, &content),
        Err(Error::Refused(_))
    ));
    assert!(matches!(edit.remove_equation(0, 3), Err(Error::Refused(_))));
    let mut bad = equation3(0, 0);
    bad.content.native.truncate(20);
    assert!(matches!(edit.add_equation(0, &bad), Err(Error::Refused(_))));
    let mut bad = equation3(0, 0);
    bad.kind = EquationKind::MathType;
    assert!(matches!(edit.add_equation(0, &bad), Err(Error::Refused(_))));
    let mut bad = equation3(0, 0);
    bad.content.preview = b"GIF89a".to_vec();
    assert!(matches!(edit.add_equation(0, &bad), Err(Error::Refused(_))));
    let mut bad = equation3(0, 0);
    bad.content.size_pt = (f64::NAN, 1.0);
    assert!(matches!(edit.add_equation(0, &bad), Err(Error::Refused(_))));
}

#[test]
fn delete_removes_the_shape_obj_blip_and_storage() {
    let source = Package::new().unwrap();
    let source_streams = streams(&source);
    let mut edit = source.edit_sheet_features().unwrap();
    edit.add_equation(0, &equation3(0, 0)).unwrap();
    edit.add_equation(0, &equation3(4, 4)).unwrap();
    let package = edit.commit().unwrap().into_package();
    let equations = package.equations(0).unwrap();
    let with_two = streams(&package);

    let mut edit = package.edit_sheet_features().unwrap();
    edit.remove_equation(0, 0).unwrap();
    let package = reopen(&edit.commit().unwrap().into_package());
    let remaining = package.equations(0).unwrap();
    assert_eq!(remaining.len(), 1);
    assert_eq!(remaining[0].storage, equations[1].storage);
    let with_one = streams(&package);
    assert!(!with_one.keys().any(|path| path[0] == equations[0].storage));
    assert_eq!(
        others(&with_two, &[&equations[0].storage]),
        others(&with_one, &[])
    );
    assert_eq!(package.ole_objects().unwrap().objects(0).unwrap().len(), 1);

    // Removing the last one leaves the workbook as it was.
    let mut edit = package.edit_sheet_features().unwrap();
    edit.remove_equation(0, 0).unwrap();
    let package = reopen(&edit.commit().unwrap().into_package());
    assert!(package.equations(0).unwrap().is_empty());
    assert_eq!(streams(&package), source_streams);
}

#[test]
fn add_then_remove_in_one_edit_is_a_no_op() {
    let source = Package::new().unwrap();
    let mut edit = source.edit_sheet_features().unwrap();
    edit.add_equation(0, &equation3(1, 1)).unwrap();
    edit.remove_equation(0, 0).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(streams(commit.package()), streams(&source));
}

#[test]
fn excel_embedded_objects_stay_intact_beside_a_new_equation() {
    let package = fixture("WithEmbeddedObjects.xls");
    let before = streams(&package);
    let before_objects = package.ole_objects().unwrap().objects(0).unwrap().to_vec();
    let before_drawings = package.sheet_features(0, FeatureSet::DRAWINGS).unwrap();
    assert!(package.equations(0).unwrap().is_empty());

    let mut edit = package.edit_sheet_features().unwrap();
    edit.add_equation(0, &equation3(10, 2)).unwrap();
    let package = reopen(&edit.commit().unwrap().into_package());
    let equations = package.equations(0).unwrap();
    assert_eq!(equations.len(), 1);
    let after = streams(&package);
    assert_eq!(
        others(&before, &[]),
        others(&after, &[&equations[0].storage])
    );
    let objects = package.ole_objects().unwrap();
    let after_objects = objects.objects(0).unwrap();
    assert_eq!(
        &after_objects[..before_objects.len()],
        before_objects.as_slice()
    );
    assert_eq!(after_objects.len(), before_objects.len() + 1);
    let drawings = package.sheet_features(0, FeatureSet::DRAWINGS).unwrap();
    assert_eq!(drawings.drawings, before_drawings.drawings);

    let mut edit = package.edit_sheet_features().unwrap();
    edit.remove_equation(0, 0).unwrap();
    let package = reopen(&edit.commit().unwrap().into_package());
    assert_eq!(others(&streams(&package), &[]), others(&before, &[]));
    assert_eq!(
        package.ole_objects().unwrap().objects(0).unwrap(),
        before_objects.as_slice()
    );
}

#[cfg(feature = "equation")]
mod converted {
    use super::*;
    use litchi_equation::Limits;
    use litchi_xls::sheet_features::edit::EquationSource;

    #[test]
    fn latex_creates_a_mathtype_object_with_a_rendered_preview() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_sheet_features().unwrap();
        let (index, losses) = edit
            .add_equation_from(
                0,
                EquationSource::Latex(r"\frac{a}{b}"),
                1,
                1,
                &Limits::DEFAULT,
            )
            .unwrap();
        assert_eq!(index, 0);
        assert!(losses.is_empty());
        let package = reopen(&edit.commit().unwrap().into_package());
        let equation = &package.equations(0).unwrap()[0];
        assert_eq!(equation.prog_id, "Equation.DSMT4");
        assert_eq!(equation.kind(), EquationKind::MathType);
        assert_eq!(equation.latex(&Limits::DEFAULT).unwrap(), r"\frac{a}{b}");
        assert_eq!(equation.preview.unwrap().0, "image/png");
        let native = equation.native.as_deref().unwrap();
        assert_eq!(native[28], 5);
        let streams = streams(&package);
        let comp_obj = &streams[&vec![equation.storage.clone(), "\u{1}CompObj".to_owned()]];
        assert!(String::from_utf8_lossy(comp_obj).contains("Equation.DSMT4\0"));
    }

    #[test]
    fn omml_and_mtef_sources_convert() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_sheet_features().unwrap();
        edit.add_equation_from(
            0,
            EquationSource::Omml(
                "<m:oMath><m:sSup><m:e><m:r><m:t>e</m:t></m:r></m:e>\
                 <m:sup><m:r><m:t>x</m:t></m:r></m:sup></m:sSup></m:oMath>",
            ),
            0,
            0,
            &Limits::DEFAULT,
        )
        .unwrap();
        // Bare MTEF 3 gets its OLE header and keeps its kind.
        let native = equation3_native();
        edit.add_equation_from(
            0,
            EquationSource::Mtef(&native[28..]),
            3,
            0,
            &Limits::DEFAULT,
        )
        .unwrap();
        let package = reopen(&edit.commit().unwrap().into_package());
        let equations = package.equations(0).unwrap();
        assert_eq!(equations[0].latex(&Limits::DEFAULT).unwrap(), "e^x");
        assert_eq!(equations[1].prog_id, "Equation.3");
        assert_eq!(equations[1].latex(&Limits::DEFAULT).unwrap(), "2");
    }

    #[test]
    fn update_keeps_an_equation_editor_object_mtef3() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_sheet_features().unwrap();
        edit.add_equation(0, &equation3(0, 0)).unwrap();
        let losses = edit
            .set_equation_from(0, 0, EquationSource::Latex(r"\sqrt{x}"), &Limits::DEFAULT)
            .unwrap();
        assert!(losses.is_empty());
        let package = reopen(&edit.commit().unwrap().into_package());
        let equation = &package.equations(0).unwrap()[0];
        assert_eq!(equation.prog_id, "Equation.3");
        assert_eq!(equation.native.as_deref().unwrap()[28], 3);
        assert_eq!(equation.latex(&Limits::DEFAULT).unwrap(), r"\sqrt{x}");
    }

    #[test]
    fn hostile_input_is_a_typed_error() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_sheet_features().unwrap();
        let deep = "{".repeat(5_000) + &"}".repeat(5_000);
        assert!(matches!(
            edit.add_equation_from(0, EquationSource::Latex(&deep), 0, 0, &Limits::DEFAULT),
            Err(Error::Equation(_))
        ));
        assert!(matches!(
            edit.add_equation_from(
                0,
                EquationSource::Mtef(b"\x05garbage"),
                0,
                0,
                &Limits::DEFAULT
            ),
            Err(Error::Equation(_))
        ));
        assert!(edit.equations(0).unwrap().is_empty());
    }
}

#[test]
fn an_unreadable_equation_stream_is_reported_not_fatal() {
    let package = Package::new().unwrap();
    let mut edit = package.edit_sheet_features().unwrap();
    edit.add_equation(0, &equation3(0, 0)).unwrap();
    edit.add_equation(0, &equation3(3, 0)).unwrap();
    let package = edit.commit().unwrap().into_package();
    let broken = package.equations(0).unwrap()[0].storage.clone();
    // Rebuild the compound file with the first equation's `Equation Native`
    // turned into a storage.
    let mut writer = litchi_cfb::OleWriter::new();
    for (path, data) in streams(&package) {
        let mut parts: Vec<&str> = path.iter().map(String::as_str).collect();
        if parts == [broken.as_str(), "Equation Native"] {
            parts.push("data");
        }
        writer.create_stream(&parts, &data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let package = Package::from_bytes(&output.into_inner()).unwrap();
    let equations = package.equations(0).unwrap();
    assert_eq!(equations.len(), 2);
    assert!(equations[0].native.is_none());
    assert!(equations[0].native_error.is_some());
    assert_eq!(
        equations[1].native.as_deref(),
        Some(equation3_native().as_slice())
    );
    assert!(equations[1].native_error.is_none());
    let mut edit = package.edit_sheet_features().unwrap();
    assert_eq!(edit.equations(0).unwrap().len(), 2);
}
