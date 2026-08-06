//! Focused MS-DOC form-field wire and semantic tests.

use super::codec::{CB_HEADER, FF_DATA_VERSION, HEADER_LEN, STTB_EXTENDED, write_xstz};
use super::model::UNDEFINED_STATE;
use super::validation::{F_HAS_LISTBOX, F_ISIZE, F_RECALC, IRES_SHIFT, ITYPE_TXT_SHIFT};
use super::{
    CheckBoxState, FormFieldData, FormFieldDataKind, FormFieldTextKind, NilPicfAndBinData,
};

fn xstz(text: &str) -> Vec<u8> {
    let mut out = Vec::new();
    write_xstz(&mut out, text);
    out
}

fn nil_picf(bin_data: &[u8]) -> Vec<u8> {
    let lcb = (HEADER_LEN + bin_data.len()) as u32;
    let mut out = Vec::new();
    out.extend_from_slice(&lcb.to_le_bytes());
    out.extend_from_slice(&CB_HEADER.to_le_bytes());
    out.extend_from_slice(&[0; HEADER_LEN - 6]);
    out.extend_from_slice(bin_data);
    out
}

fn text_ff_data(bits: u16, name: &str, default: &str, format: &str) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&FF_DATA_VERSION.to_le_bytes());
    out.extend_from_slice(&bits.to_le_bytes());
    out.extend_from_slice(&10u16.to_le_bytes()); // cch
    out.extend_from_slice(&0u16.to_le_bytes()); // hps
    out.extend_from_slice(&xstz(name));
    out.extend_from_slice(&xstz(default));
    out.extend_from_slice(&xstz(format));
    out.extend_from_slice(&xstz("help"));
    out.extend_from_slice(&xstz("status"));
    out.extend_from_slice(&xstz("EntryMacro"));
    out.extend_from_slice(&xstz("ExitMacro"));
    out
}

fn checkbox_ff_data(state: u16, w_def: u16) -> Vec<u8> {
    let bits = 1 | (state << IRES_SHIFT) | F_ISIZE;
    let mut out = Vec::new();
    out.extend_from_slice(&FF_DATA_VERSION.to_le_bytes());
    out.extend_from_slice(&bits.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // cch MUST be 0
    out.extend_from_slice(&20u16.to_le_bytes()); // hps
    out.extend_from_slice(&xstz("Check1"));
    out.extend_from_slice(&w_def.to_le_bytes());
    for text in ["", "help", "status", "", ""] {
        out.extend_from_slice(&xstz(text));
    }
    out
}

fn dropdown_ff_data(selection: u16, w_def: u16, items: &[&str]) -> Vec<u8> {
    let bits = 2 | (selection << IRES_SHIFT) | F_HAS_LISTBOX;
    let mut out = Vec::new();
    out.extend_from_slice(&FF_DATA_VERSION.to_le_bytes());
    out.extend_from_slice(&bits.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    out.extend_from_slice(&xstz("Drop1"));
    out.extend_from_slice(&w_def.to_le_bytes());
    for text in ["", "", "", "", ""] {
        out.extend_from_slice(&xstz(text));
    }
    out.extend_from_slice(&STTB_EXTENDED.to_le_bytes());
    out.extend_from_slice(&(items.len() as u16).to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    for item in items {
        let units: Vec<u16> = item.encode_utf16().collect();
        out.extend_from_slice(&(units.len() as u16).to_le_bytes());
        for unit in units {
            out.extend_from_slice(&unit.to_le_bytes());
        }
    }
    out
}

#[test]
fn parses_text_field() {
    // iType text, iTypeTxt date, fRecalc.
    let bytes = text_ff_data(
        2 << ITYPE_TXT_SHIFT | F_RECALC,
        "Text1",
        "1.1.2000",
        "d.M.yyyy",
    );
    let data = FormFieldData::parse(&bytes).unwrap();
    assert_eq!(data.kind(), FormFieldDataKind::Text);
    assert_eq!(data.text_kind(), Some(FormFieldTextKind::Date));
    assert_eq!(data.max_length(), Some(10));
    assert_eq!(data.name(), "Text1");
    assert_eq!(data.default_text(), Some("1.1.2000"));
    assert_eq!(data.text_format(), "d.M.yyyy");
    assert_eq!(data.help_text(), "help");
    assert_eq!(data.status_text(), "status");
    assert_eq!(data.entry_macro(), "EntryMacro");
    assert_eq!(data.exit_macro(), "ExitMacro");
    assert!(data.is_marked_for_recalculation());
    assert!(!data.is_protected());
    assert_eq!(data.checkbox_state(), None);
    assert_eq!(data.selected_item_index(), None);
    assert!(data.dropdown_items().is_empty());
    assert_eq!(data.to_bytes(), bytes);
}

#[test]
fn parses_checkbox_field() {
    let bytes = checkbox_ff_data(1, 1);
    let data = FormFieldData::parse(&bytes).unwrap();
    assert_eq!(data.kind(), FormFieldDataKind::CheckBox);
    assert_eq!(data.checkbox_state(), Some(CheckBoxState::Checked));
    assert_eq!(data.is_checked_by_default(), Some(true));
    assert_eq!(data.checkbox_size_half_points(), Some(20));
    assert!(data.is_checkbox_auto_sized());
    assert_eq!(data.text_kind(), None);
    assert_eq!(data.default_text(), None);
    assert_eq!(data.to_bytes(), bytes);

    let undefined = FormFieldData::parse(&checkbox_ff_data(UNDEFINED_STATE.into(), 0)).unwrap();
    assert_eq!(undefined.checkbox_state(), Some(CheckBoxState::Undefined));
    assert_eq!(undefined.is_checked_by_default(), Some(false));
}

#[test]
fn parses_dropdown_field() {
    let bytes = dropdown_ff_data(2, 0, &["one", "two", "three"]);
    let data = FormFieldData::parse(&bytes).unwrap();
    assert_eq!(data.kind(), FormFieldDataKind::DropDown);
    assert!(data.has_list_box());
    assert_eq!(data.selected_item_index(), Some(2));
    assert_eq!(data.default_item_index(), Some(0));
    assert_eq!(data.dropdown_items(), &["one", "two", "three"]);
    assert_eq!(data.name(), "Drop1");
    assert_eq!(data.to_bytes(), bytes);

    let undefined =
        FormFieldData::parse(&dropdown_ff_data(UNDEFINED_STATE.into(), 1, &["a", "b"])).unwrap();
    assert_eq!(undefined.selected_item_index(), None);
    assert_eq!(undefined.default_item_index(), Some(1));
}

#[test]
fn rejects_invalid_text_kinds() {
    // Reserved iType (3).
    assert!(FormFieldData::parse(&text_ff_data(3, "n", "", "")).is_err());
    // iRes not 0 for a text box.
    assert!(FormFieldData::parse(&text_ff_data(1 << IRES_SHIFT, "n", "", "")).is_err());
    // Reserved iTypeTxt (6) for a text box.
    assert!(FormFieldData::parse(&text_ff_data(6 << ITYPE_TXT_SHIFT, "n", "", "")).is_err());
    // fHasListBox set on a text box.
    assert!(FormFieldData::parse(&text_ff_data(F_HAS_LISTBOX, "n", "", "")).is_err());
    // iSize set on a text box.
    assert!(FormFieldData::parse(&text_ff_data(F_ISIZE, "n", "", "")).is_err());
    // Non-empty default text for a current-date text box.
    assert!(FormFieldData::parse(&text_ff_data(3 << ITYPE_TXT_SHIFT, "n", "x", "")).is_err());
}

#[test]
fn rejects_invalid_checkbox_and_dropdown_states() {
    // iRes 2 is not a checkbox state.
    assert!(FormFieldData::parse(&checkbox_ff_data(2, 0)).is_err());
    // wDef 2 is not a checkbox state.
    assert!(FormFieldData::parse(&checkbox_ff_data(0, 2)).is_err());
    // hps below the checkbox range.
    let mut bad_hps = checkbox_ff_data(0, 0);
    bad_hps[8..10].copy_from_slice(&1u16.to_le_bytes());
    assert!(FormFieldData::parse(&bad_hps).is_err());
    // iRes 26 exceeds the undefined-selection marker.
    assert!(FormFieldData::parse(&dropdown_ff_data(26, 0, &["a"])).is_err());
    // wDef past the item list.
    assert!(FormFieldData::parse(&dropdown_ff_data(0, 3, &["a"])).is_err());
    // Missing fHasListBox on a drop-down.
    let mut no_listbox = dropdown_ff_data(0, 0, &["a"]);
    no_listbox[4..6].copy_from_slice(&0u16.to_le_bytes());
    assert!(FormFieldData::parse(&no_listbox).is_err());
    // More than 25 items.
    let items: Vec<&str> = vec!["x"; 26];
    assert!(FormFieldData::parse(&dropdown_ff_data(0, 0, &items)).is_err());
    // Non-zero cch for a non-text field.
    let mut bad_cch = checkbox_ff_data(0, 0);
    bad_cch[6..8].copy_from_slice(&5u16.to_le_bytes());
    assert!(FormFieldData::parse(&bad_cch).is_err());
}

#[test]
fn rejects_malformed_payloads() {
    let good = text_ff_data(0, "Text1", "", "");
    // Wrong version.
    let mut bad_version = good.clone();
    bad_version[0..4].copy_from_slice(&0u32.to_le_bytes());
    assert!(FormFieldData::parse(&bad_version).is_err());
    // Truncated.
    assert!(FormFieldData::parse(&good[..good.len() - 1]).is_err());
    assert!(FormFieldData::parse(&good[..4]).is_err());
    // Trailing bytes.
    let mut trailing = good.clone();
    trailing.push(0);
    assert!(FormFieldData::parse(&trailing).is_err());
    // xstzName beyond its 20-character cap.
    let long_name = text_ff_data(0, &"n".repeat(21), "", "");
    assert!(FormFieldData::parse(&long_name).is_err());
    // Non-zero Xstz terminator.
    let mut bad_terminator = good.clone();
    // xstzName "Text1": cch at 10, units at 12..22, terminator at 22.
    bad_terminator[22..24].copy_from_slice(&1u16.to_le_bytes());
    assert!(FormFieldData::parse(&bad_terminator).is_err());
    // Lone surrogate in an Xstz.
    let mut bad_utf16 = good.clone();
    bad_utf16[12..14].copy_from_slice(&0xD800u16.to_le_bytes());
    assert!(FormFieldData::parse(&bad_utf16).is_err());
    // Non-text field with a non-empty format string.
    let mut bad_format = checkbox_ff_data(0, 0);
    // name "Check1" ends at 10+2+12+2=26, wDef 26..28, format cch at 28.
    bad_format[28..30].copy_from_slice(&1u16.to_le_bytes());
    assert!(FormFieldData::parse(&bad_format).is_err());
}

#[test]
fn parses_nil_picf_and_bin_data() {
    let ff_data = text_ff_data(0, "Text1", "", "");
    let bytes = nil_picf(&ff_data);
    let parsed = NilPicfAndBinData::parse(&bytes).unwrap();
    assert_eq!(parsed.bin_data(), ff_data.as_slice());
    assert_eq!(parsed.to_bytes(), bytes);

    // Parsing from a longer buffer consumes exactly lcb bytes.
    let mut padded = bytes.clone();
    padded.extend_from_slice(&[0xAA; 16]);
    let parsed = NilPicfAndBinData::parse(&padded).unwrap();
    assert_eq!(parsed.bin_data(), ff_data.as_slice());

    // Parsing at an offset of a Data-stream-like buffer.
    let mut stream = vec![0u8; 7];
    stream.extend_from_slice(&bytes);
    let parsed = NilPicfAndBinData::parse_at(&stream, 7).unwrap();
    assert_eq!(parsed.bin_data(), ff_data.as_slice());
    let data = FormFieldData::parse_at(&stream, 7).unwrap();
    assert_eq!(data.name(), "Text1");
}

#[test]
fn rejects_malformed_nil_picf_and_bin_data() {
    let ff_data = text_ff_data(0, "Text1", "", "");
    let good = nil_picf(&ff_data);
    // Truncated header.
    assert!(NilPicfAndBinData::parse(&good[..HEADER_LEN - 1]).is_err());
    // Wrong cbHeader.
    let mut bad_header = good.clone();
    bad_header[4..6].copy_from_slice(&0x0042u16.to_le_bytes());
    assert!(NilPicfAndBinData::parse(&bad_header).is_err());
    // lcb smaller than the header.
    let mut small_lcb = good.clone();
    small_lcb[0..4].copy_from_slice(&10i32.to_le_bytes());
    assert!(NilPicfAndBinData::parse(&small_lcb).is_err());
    // lcb past the containing data.
    let mut huge_lcb = good.clone();
    huge_lcb[0..4].copy_from_slice(&1_000_000i32.to_le_bytes());
    assert!(NilPicfAndBinData::parse(&huge_lcb).is_err());
    // Negative lcb.
    let mut negative_lcb = good.clone();
    negative_lcb[0..4].copy_from_slice(&(-1i32).to_le_bytes());
    assert!(NilPicfAndBinData::parse(&negative_lcb).is_err());
    // Offset past the stream.
    assert!(NilPicfAndBinData::parse_at(&good, 1_000_000).is_err());
}
