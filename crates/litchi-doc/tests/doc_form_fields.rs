//! Tests for legacy form-field binary data (`NilPICFAndBinData`/`FFData`,
//! MS-DOC 2.9.158 and 2.9.78) against a real Word-produced document.

use litchi_cfb::OleFile;
use litchi_doc::parts::form_fields::{FormFieldData, NilPicfAndBinData};
use litchi_doc::{CheckBoxState, FormFieldDataKind, FormFieldTextKind, LegacyFormFieldKind};
use std::fs::File;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

const TRAVEL_FORM: &str = "test-data/poi/test-data/document/au.edu.utas.www___data_assets_word_doc_0003_154335_International-Travel-Approval-Request-Form.doc";
const CONCERT_FORM: &str = "test-data/poi/test-data/document/ca.kwsymphony.www_education_School_Concert_Seat_Booking_Form_2011-12.doc";

/// Offset and `lcb` of a checkbox `NilPICFAndBinData` in the travel form's
/// Data stream (`FFDataBits` 0x0065: checkbox with undefined state).
const CHECKBOX_OFFSET: u32 = 9558;
/// Offset of the document's lone dropdown `NilPICFAndBinData`. Its stored
/// `FFData` is invalid (wDef 0 with an empty item list).
const DROPDOWN_OFFSET: u32 = 12509;

#[test]
fn parses_ffdata_from_the_data_stream() {
    let path = fixture(TRAVEL_FORM);
    let mut ole = OleFile::open(File::open(&path).unwrap()).unwrap();
    let data_stream = ole.open_stream(&["Data"]).unwrap();

    let nil_picf = NilPicfAndBinData::parse_at(&data_stream, CHECKBOX_OFFSET).unwrap();
    let data = FormFieldData::parse(nil_picf.bin_data()).unwrap();
    assert_eq!(data.kind(), FormFieldDataKind::CheckBox);
    assert_eq!(data.checkbox_state(), Some(CheckBoxState::Undefined));
    assert_eq!(data.is_checked_by_default(), Some(false));
    assert_eq!(data.checkbox_size_half_points(), Some(20));
    assert!(data.name().starts_with("Check"));

    // Round-trip: the re-encoded bytes reproduce the stored bytes exactly.
    let lcb = u32::from_le_bytes(
        data_stream[CHECKBOX_OFFSET as usize..CHECKBOX_OFFSET as usize + 4]
            .try_into()
            .unwrap(),
    ) as usize;
    assert_eq!(
        nil_picf.to_bytes(),
        data_stream[CHECKBOX_OFFSET as usize..CHECKBOX_OFFSET as usize + lcb]
    );
    assert_eq!(
        data.to_bytes(),
        data_stream[CHECKBOX_OFFSET as usize + 68..CHECKBOX_OFFSET as usize + lcb]
    );

    // The lone dropdown's stored FFData is invalid per MS-DOC 2.9.78 (wDef 0
    // with an empty hsttbDropList) and MUST be ignored (2.9.158).
    let nil_picf = NilPicfAndBinData::parse_at(&data_stream, DROPDOWN_OFFSET).unwrap();
    assert!(FormFieldData::parse(nil_picf.bin_data()).is_err());
}

#[test]
fn exposes_form_data_through_the_document_api() {
    let mut package =
        litchi_doc::Package::from_reader(File::open(fixture(TRAVEL_FORM)).unwrap()).unwrap();
    let document = package.document().unwrap();

    let fields = document.legacy_form_fields().unwrap();
    assert_eq!(fields.len(), 60);

    let mut text = 0;
    let mut checkbox = 0;
    let mut dropdown = 0;
    for field in &fields {
        match field.kind() {
            LegacyFormFieldKind::Text => {
                text += 1;
                let data = field.form_data().expect("text FFData is well-formed");
                assert_eq!(data.kind(), FormFieldDataKind::Text);
                assert_eq!(data.text_kind(), Some(FormFieldTextKind::Regular));
                assert_eq!(data.max_length(), Some(0));
                assert_eq!(data.default_text(), Some(""));
                assert!(data.name().starts_with("Text"));
                // Everything stays inert: no macros or custom texts stored.
                assert_eq!(data.entry_macro(), "");
                assert_eq!(data.exit_macro(), "");
            },
            LegacyFormFieldKind::CheckBox => {
                checkbox += 1;
                let data = field.form_data().expect("checkbox FFData is well-formed");
                assert_eq!(data.kind(), FormFieldDataKind::CheckBox);
                assert_eq!(data.checkbox_state(), Some(CheckBoxState::Undefined));
                assert_eq!(data.is_checked_by_default(), Some(false));
                assert_eq!(data.checkbox_size_half_points(), Some(20));
                assert!(data.name().starts_with("Check"));
            },
            LegacyFormFieldKind::DropDown => {
                dropdown += 1;
                // The stored dropdown FFData is invalid (empty item list) and
                // is ignored per MS-DOC 2.9.158.
                assert!(field.form_data().is_none());
            },
        }
    }
    assert_eq!((text, checkbox, dropdown), (47, 12, 1));
}

#[test]
fn documents_without_form_fields_report_none() {
    // The seat booking form uses no legacy form-code fields at all.
    let mut package =
        litchi_doc::Package::from_reader(File::open(fixture(CONCERT_FORM)).unwrap()).unwrap();
    let document = package.document().unwrap();
    assert!(document.legacy_form_fields().unwrap().is_empty());
}
