//! Native Excel form controls are inert worksheet/part/VML metadata.
use litchi_xlsx::form_control::{
    Checked, ControlSelector, KnownOrUnknown, ObjectType, ScalarField, ScalarValue,
};
use litchi_xlsx::{Package, SourceBackedWorkbook};

fn bytes() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-form-controls.xlsx"),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let data = reader.read_member(&n).unwrap();
            (n, data)
        })
        .collect()
}

#[test]
fn native_form_control_owners_expose_all_eleven_stored_states_and_bindings() {
    let bytes = bytes();
    let package = Package::from_bytes(&bytes).unwrap();
    let book = package.workbook();
    let sheet = book.sheet("Sheet1").unwrap().unwrap();
    let controls = sheet.form_controls().unwrap();
    let source = SourceBackedWorkbook::from_reader(std::io::Cursor::new(bytes)).unwrap();
    let source_sheet = source.sheet("Sheet1").unwrap().unwrap();
    assert_eq!(controls, source_sheet.form_controls().unwrap());
    assert_eq!(controls.len(), 11);
    assert!(controls.diagnostics().is_empty());
    for (control, (name, id, kind)) in controls.iter().zip([
        ("Check Box 1", 1025, ObjectType::CheckBox),
        ("Drop Down 2", 1026, ObjectType::Drop),
        ("List Box 3", 1027, ObjectType::List),
        ("Spinner 4", 1028, ObjectType::Spin),
        ("Scroll Bar 5", 1029, ObjectType::Scroll),
        ("Group Box 6", 1030, ObjectType::GBox),
        ("Option Button 7", 1031, ObjectType::Radio),
        ("Label 9", 1033, ObjectType::Label),
        ("Button 10", 1034, ObjectType::Button),
        ("Check Box 11", 1035, ObjectType::CheckBox),
        ("Option Button 12", 1036, ObjectType::Radio),
    ]) {
        assert_eq!(control.name(), Some(name));
        assert_eq!(control.shape().shape_id(), id);
        assert_eq!(
            control
                .properties()
                .object_type()
                .and_then(KnownOrUnknown::known),
            Some(&kind)
        );
        assert!(
            control.diagnostics().iter().all(|d| d.code()
                == litchi_xlsx::form_control::FormControlDiagnosticCode::UnprovenMirror)
        );
    }
    let properties = |index| {
        controls
            .get(ControlSelector::position(index))
            .unwrap()
            .unwrap()
            .properties()
    };
    assert_eq!(
        properties(0).checked(),
        Some(&KnownOrUnknown::Known(Checked::Checked))
    );
    assert_eq!(
        properties(9).checked(),
        Some(&KnownOrUnknown::Known(Checked::Mixed))
    );
    assert_eq!(properties(6).checked(), None);
    assert!(properties(6).effective_first_button());
    assert_eq!(
        properties(10).checked(),
        Some(&KnownOrUnknown::Known(Checked::Checked))
    );
    for (index, link) in [
        (0, "$B$2"),
        (1, "$B$3"),
        (2, "$B$4"),
        (3, "$B$5"),
        (4, "$B$6"),
        (6, "$B$7"),
        (9, "$B$12"),
    ] {
        assert_eq!(properties(index).fmla_link().unwrap().as_str(), link);
    }
    for (index, selected) in [(1, 2), (2, 3)] {
        assert_eq!(properties(index).sel(), Some(selected));
        assert_eq!(
            properties(index).fmla_range().unwrap().as_str(),
            "$D$9:$D$11"
        );
        assert!(properties(index).item_list().is_none());
    }
    assert_eq!(
        (
            properties(3).val(),
            properties(3).min(),
            properties(3).max(),
            properties(3).inc(),
            properties(3).page()
        ),
        (Some(5), Some(1), Some(20), Some(2), Some(10))
    );
    assert_eq!(
        (
            properties(4).val(),
            properties(4).min(),
            properties(4).effective_min(),
            properties(4).max(),
            properties(4).inc(),
            properties(4).page()
        ),
        (Some(10), None, 0, Some(50), Some(2), Some(5))
    );
    assert!(properties(4).effective_horiz());
    assert_eq!(properties(7).fmla_txbx().unwrap().as_str(), "Sheet1!$A$1");
}

#[test]
fn native_checkbox_scalar_edit_preserves_cells_and_unrelated_controls() {
    let bytes = bytes();
    let before = members(&bytes);
    let source = SourceBackedWorkbook::from_reader(std::io::Cursor::new(bytes.clone())).unwrap();
    let mut edit = source.edit_form_controls("Sheet1").unwrap();
    edit.set_scalar(
        ControlSelector::position(0),
        ScalarField::Checked,
        Some(ScalarValue::Checked(Checked::Unchecked)),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("unchecked.xlsx");
    commit.package().save(&path).unwrap();
    let after_bytes = std::fs::read(&path).unwrap();
    let reopened = Package::from_bytes(&after_bytes).unwrap();
    let controls = reopened
        .workbook()
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .form_controls()
        .unwrap();
    assert_eq!(
        controls
            .get(ControlSelector::position(0))
            .unwrap()
            .unwrap()
            .properties()
            .checked(),
        Some(&KnownOrUnknown::Known(Checked::Unchecked))
    );
    let after = members(&after_bytes);
    let expected_properties = String::from_utf8(before["xl/ctrlProps/ctrlProp1.xml"].clone())
        .unwrap()
        .replacen("checked=\"Checked\"", "checked=\"Unchecked\"", 1);
    assert!(after["xl/ctrlProps/ctrlProp1.xml"] == expected_properties.as_bytes());
    let expected_vml = String::from_utf8(before["xl/drawings/vmlDrawing1.vml"].clone())
        .unwrap()
        .replacen("<x:Checked>1</x:Checked>", "<x:Checked>0</x:Checked>", 1);
    assert!(after["xl/drawings/vmlDrawing1.vml"] == expected_vml.as_bytes());
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (name, data) in &before {
        if name != "xl/ctrlProps/ctrlProp1.xml" && name != "xl/drawings/vmlDrawing1.vml" {
            assert!(after.get(name) == Some(data), "changed {name}");
        }
    }
    assert!(
        commit
            .package()
            .apply_form_controls(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(source.to_bytes().unwrap() == bytes);
}
