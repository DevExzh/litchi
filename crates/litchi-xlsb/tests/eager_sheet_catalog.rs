//! Eager XLSB tab catalog coverage (L-XLSB-R1).

#![allow(
    clippy::unwrap_used,
    reason = "the compact package fixtures and assertions intentionally panic on construction failures"
)]

use std::io::Cursor;

use litchi_core::Position;
use litchi_core::sheet::WorkbookTrait;
use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsb::Workbook;
use litchi_xlsb::builder::chart::{Anchor, Chart};
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::package::PackageError;
use litchi_xlsb::raw::{Encoder, Records, kind};
use litchi_xlsb::sheet::{Kind, State};

const DIALOG_SHEET: (&str, &str) = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
    "application/vnd.ms-excel.dialogsheet",
);
const MACRO_SHEET: (&str, &str) = (
    "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
    "application/vnd.ms-excel.macrosheet",
);
const INTL_MACRO_SHEET: (&str, &str) = (
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet",
    "application/vnd.ms-excel.intlmacrosheet",
);

fn every_kind_workbook() -> Vec<u8> {
    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .unwrap();
    let mut writer = Builder::new();
    writer.add_worksheet(builder::Worksheet::new("Data"));
    writer
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .unwrap();
    for name in ["Dialog", "Macro", "Intl"] {
        writer.add_worksheet(builder::Worksheet::new(name));
    }
    writer.add_worksheet(builder::Worksheet::new("Tail"));
    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build package")
        .write_to(&mut output)
        .unwrap();
    let mut bytes = output.into_inner();
    for (name, (relationship_type, content_type)) in [
        ("Dialog", DIALOG_SHEET),
        ("Macro", MACRO_SHEET),
        ("Intl", INTL_MACRO_SHEET),
    ] {
        bytes = retype_sheet(&bytes, name, relationship_type, content_type);
    }
    with_states(&bytes, &[(1, 2), (3, 1), (5, 1)])
}

fn bundle_entries(package: &OpcPackage) -> Vec<(String, String)> {
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let workbook_part = package.get_part(&workbook_uri).unwrap();
    Records::new(workbook_part.blob(), litchi_xlsb::raw::Limits::DEFAULT)
        .map(|record| record.expect("valid workbook record"))
        .filter(|record| record.kind() == kind::BUNDLE_SH)
        .map(|record| {
            let payload = record.payload();
            let rel_units = u32::from_le_bytes(payload[8..12].try_into().unwrap()) as usize;
            let rel_start = 12;
            let rel_end = rel_start + rel_units * 2;
            let rel_id = String::from_utf16(
                &payload[rel_start..rel_end]
                    .chunks_exact(2)
                    .map(|unit| u16::from_le_bytes([unit[0], unit[1]]))
                    .collect::<Vec<_>>(),
            )
            .unwrap();
            let name_units =
                u32::from_le_bytes(payload[rel_end..rel_end + 4].try_into().unwrap()) as usize;
            let name_start = rel_end + 4;
            let name_end = name_start + name_units * 2;
            let name = String::from_utf16(
                &payload[name_start..name_end]
                    .chunks_exact(2)
                    .map(|unit| u16::from_le_bytes([unit[0], unit[1]]))
                    .collect::<Vec<_>>(),
            )
            .unwrap();
            (name, rel_id)
        })
        .collect()
}

fn sheet_partname(package: &OpcPackage, position: usize) -> PackURI {
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let workbook_part = package.get_part(&workbook_uri).unwrap();
    let relationship_ids = bundle_entries(package);
    workbook_part
        .rels()
        .get(&relationship_ids[position].1)
        .unwrap()
        .target_partname()
        .unwrap()
}

fn retype_sheet(bytes: &[u8], name: &str, relationship_type: &str, content_type: &str) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let relationship_ids = bundle_entries(&package);
    let workbook_part = package.get_part(&workbook_uri).unwrap();
    let position = Records::new(workbook_part.blob(), litchi_xlsb::raw::Limits::DEFAULT)
        .filter_map(Result::ok)
        .filter(|record| record.kind() == kind::BUNDLE_SH)
        .enumerate()
        .find_map(|(position, record)| {
            let _ = record;
            (relationship_ids[position].0 == name).then_some(position)
        })
        .unwrap();
    let sheet_uri = sheet_partname(&package, position);
    let (r_id, target_ref) = package
        .get_part(&workbook_uri)
        .unwrap()
        .rels()
        .iter()
        .find_map(|relationship| {
            (relationship.target_partname().ok()? == sheet_uri).then(|| {
                (
                    relationship.r_id().to_owned(),
                    relationship.target_ref().to_owned(),
                )
            })
        })
        .unwrap();
    {
        let workbook = package.get_part_mut(&workbook_uri).unwrap();
        workbook.rels_mut().remove(&r_id).unwrap();
        workbook
            .rels_mut()
            .add_relationship(relationship_type.to_owned(), target_ref, r_id, false);
    }
    package
        .get_part_mut(&sheet_uri)
        .unwrap()
        .set_content_type(content_type.to_owned())
        .unwrap();
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

fn with_states(bytes: &[u8], states: &[(usize, u32)]) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let original = package.get_part(&workbook_uri).unwrap().blob().to_vec();
    let mut rewritten = Vec::new();
    {
        let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
        let mut position = 0;
        for record in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = record.unwrap();
            let mut payload = record.payload().to_vec();
            if record.kind() == kind::BUNDLE_SH {
                if let Some(&(_, state)) = states.iter().find(|(at, _)| *at == position) {
                    payload[0..4].copy_from_slice(&state.to_le_bytes());
                }
                position += 1;
            }
            writer.write_record(record.kind(), &payload).unwrap();
        }
    }
    package
        .get_part_mut(&workbook_uri)
        .unwrap()
        .set_blob(rewritten);
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

fn with_null_relationship(bytes: &[u8], position: usize, state: u32) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let original = package.get_part(&workbook_uri).unwrap().blob().to_vec();
    let mut rewritten = Vec::new();
    let mut current_position = 0;
    {
        let mut writer = Encoder::new(&mut rewritten, litchi_xlsb::raw::Limits::DEFAULT);
        for record in Records::new(&original, litchi_xlsb::raw::Limits::DEFAULT) {
            let record = record.unwrap();
            let mut payload = record.payload().to_vec();
            if record.kind() == kind::BUNDLE_SH {
                if current_position == position {
                    payload[0..4].copy_from_slice(&state.to_le_bytes());
                    let current_id = u32::from_le_bytes(payload[4..8].try_into().unwrap());
                    let strings_offset = if (1..=0xFFFF).contains(&current_id) {
                        8
                    } else {
                        12
                    };
                    let rel_units = u32::from_le_bytes(
                        payload[strings_offset..strings_offset + 4]
                            .try_into()
                            .unwrap(),
                    ) as usize;
                    let rel_end = strings_offset + 4 + rel_units * 2;
                    let mut null_payload = Vec::with_capacity(payload.len() - rel_units * 2);
                    null_payload.extend_from_slice(&payload[..strings_offset]);
                    null_payload.extend_from_slice(&u32::MAX.to_le_bytes());
                    null_payload.extend_from_slice(&payload[rel_end..]);
                    payload = null_payload;
                }
                current_position += 1;
            }
            writer.write_record(record.kind(), &payload).unwrap();
        }
    }
    package
        .get_part_mut(&workbook_uri)
        .unwrap()
        .set_blob(rewritten);
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

fn rewrite_target(
    bytes: &[u8],
    position: usize,
    f: impl FnOnce(&mut OpcPackage, &PackURI),
) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    let uri = sheet_partname(&package, position);
    f(&mut package, &uri);
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

#[test]
fn eager_catalog_lists_every_kind_and_preserves_worksheet_only_views() {
    let workbook = Workbook::new(Cursor::new(every_kind_workbook())).unwrap();
    let expected = [
        ("Data", Kind::Worksheet, State::Visible, Some(0)),
        ("Sales Chart", Kind::Chart, State::VeryHidden, None),
        ("Dialog", Kind::Dialog, State::Visible, None),
        ("Macro", Kind::Macro, State::Hidden, None),
        ("Intl", Kind::IntlMacro, State::Visible, None),
        ("Tail", Kind::Worksheet, State::Hidden, Some(1)),
    ];
    let tabs = workbook.sheets().collect::<Vec<_>>();
    assert_eq!(tabs.len(), expected.len());
    for (position, (tab, (name, kind, state, worksheet_index))) in
        tabs.iter().zip(expected).enumerate()
    {
        assert_eq!(tab.name(), name);
        assert_eq!(tab.position().get(), position);
        assert_eq!(tab.kind(), kind);
        assert_eq!(tab.state(), state);
        assert_eq!(tab.worksheet_position().map(Position::get), worksheet_index);
    }
    assert_eq!(workbook.worksheet_names(), ["Data", "Tail"]);
    assert!(workbook.sheet(Position::new(6)).unwrap().is_none());
    let by_name = workbook.sheet("sales chart").unwrap().unwrap();
    assert_eq!(by_name.position().get(), 1);
    assert_eq!(by_name.kind(), Kind::Chart);
    let by_position = workbook.sheet(2_usize).unwrap().unwrap();
    assert_eq!(by_position.name(), "Dialog");
    assert!(workbook.sheet("missing").unwrap().is_none());
    let debug = format!("{:?}", tabs[1]);
    assert!(!debug.contains("Sales Chart"));
}

#[test]
fn eager_catalog_retains_a_very_hidden_unbound_tab_without_a_worksheet_ordinal() {
    let bytes = with_null_relationship(&every_kind_workbook(), 3, 2);
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    let tabs = workbook.sheets().collect::<Vec<_>>();
    assert_eq!(tabs[3].name(), "Macro");
    assert_eq!(tabs[3].kind(), Kind::Unbound);
    assert_eq!(tabs[3].state(), State::VeryHidden);
    assert_eq!(tabs[3].worksheet_position(), None);
    assert_eq!(workbook.worksheet_names(), ["Data", "Tail"]);
    assert_eq!(
        workbook.sheet(Position::new(3)).unwrap().unwrap().kind(),
        Kind::Unbound
    );
    assert_eq!(
        workbook.sheet("macro").unwrap().unwrap().kind(),
        Kind::Unbound
    );
}

#[test]
fn eager_catalog_rejects_unbound_tabs_that_are_not_very_hidden() {
    for state in [0, 1] {
        let bytes = with_null_relationship(&every_kind_workbook(), 3, state);
        assert!(Workbook::new(Cursor::new(bytes)).is_err());
    }
}

#[test]
fn eager_catalog_refuses_missing_external_and_wrong_content_type_targets() {
    let bytes = every_kind_workbook();
    for position in 0..6 {
        let wrong_type = rewrite_target(&bytes, position, |package, uri| {
            package
                .get_part_mut(uri)
                .unwrap()
                .set_content_type("application/octet-stream".to_owned())
                .unwrap();
        });
        assert!(matches!(
            Workbook::new(Cursor::new(wrong_type)),
            Err(PackageError::InvalidContentType { .. })
        ));

        let missing = rewrite_target(&bytes, position, |package, uri| {
            assert!(package.remove_part(uri));
        });
        assert!(Workbook::new(Cursor::new(missing)).is_err());
    }

    let external = {
        let mut package =
            OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
        let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
        let relationship_ids = bundle_entries(&package);
        let relationship_id = relationship_ids[0].1.clone();
        let target = package
            .get_part(&workbook_uri)
            .unwrap()
            .rels()
            .get(&relationship_id)
            .unwrap()
            .target_ref()
            .to_owned();
        let workbook_part = package.get_part_mut(&workbook_uri).unwrap();
        workbook_part.rels_mut().remove(&relationship_id).unwrap();
        workbook_part.rels_mut().add_relationship(
            litchi_opc::constants::relationship_type::WORKSHEET.to_owned(),
            target,
            relationship_id,
            true,
        );
        let mut output = Vec::new();
        package.write_to(&mut output).unwrap();
        output
    };
    assert!(matches!(
        Workbook::new(Cursor::new(external)),
        Err(PackageError::UnsupportedFeature(_))
    ));
}
