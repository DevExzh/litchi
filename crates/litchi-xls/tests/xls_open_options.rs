//! Canonical option setters retain explicit policy and redact credentials.
use litchi_xls::{CompatibilityProfile, Leniency, OpenOptions};

#[test]
fn options_are_consuming_and_password_debug_is_redacted() {
    let defaults = OpenOptions::default();
    assert_eq!(defaults.get_password(), None);
    assert_eq!(defaults.get_leniency(), Leniency::Strict);
    assert_eq!(
        defaults.get_compatibility_profile(),
        CompatibilityProfile::Strict
    );
    let configured = defaults
        .password("private-workbook-password")
        .leniency(Leniency::TolerateFormattingDefects)
        .compatibility_profile(CompatibilityProfile::SharedFormulaFlagWithoutPtgExpV1);
    assert_eq!(configured.get_password(), Some("private-workbook-password"));
    assert_eq!(
        configured.get_leniency(),
        Leniency::TolerateFormattingDefects
    );
    assert_eq!(
        configured.get_compatibility_profile(),
        CompatibilityProfile::SharedFormulaFlagWithoutPtgExpV1
    );
    assert_eq!(defaults.get_password(), None);
    let debug = format!("{configured:?}");
    assert!(!debug.contains("private-workbook-password"));
    assert!(debug.contains("[REDACTED]"));
    assert_eq!(configured.password(None).get_password(), None);
}

#[test]
fn selected_source_profile_survives_other_option_setters() {
    let limits = litchi_xls::SourceLimits::builder()
        .max_input_bytes(1234)
        .unwrap()
        .max_entry_bytes(345)
        .unwrap()
        .max_aggregate_bytes(567)
        .unwrap()
        .build()
        .unwrap();
    let configured = OpenOptions::new()
        .source_limits(limits)
        .password("private-workbook-password")
        .leniency(Leniency::TolerateFormattingDefects);
    assert_eq!(configured.get_source_limits(), limits);
    assert_eq!(configured.password(None).get_source_limits(), limits);
}

#[test]
fn valid_workbook_opens_at_exact_physical_and_stream_boundaries() {
    use litchi_xls::{SourceLimits, Writer, raw::Reader};
    use std::io::Cursor;
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Data").unwrap();
    writer.write_number(sheet, 0, 0, 42.0).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let bytes = output.into_inner();
    let ole = litchi_cfb::OleFile::open(Cursor::new(&bytes)).unwrap();
    let sizes: Vec<_> = (0..ole.directory_entry_count())
        .filter_map(|index| ole.directory_entry_by_sid(u32::try_from(index).unwrap()))
        .filter(|entry| entry.entry_type == litchi_cfb::consts::STGTY_STREAM)
        .map(|entry| usize::try_from(entry.size).unwrap())
        .collect();
    let limits = SourceLimits::builder()
        .max_input_bytes(bytes.len())
        .unwrap()
        .max_entry_bytes(*sizes.iter().max().unwrap())
        .unwrap()
        .max_aggregate_bytes(sizes.iter().sum())
        .unwrap()
        .build()
        .unwrap();
    let options = OpenOptions::new().source_limits(limits);
    Reader::new_with_options(Cursor::new(&bytes), options).unwrap();
    Reader::from_ole_file_with_options(ole, options).unwrap();
    let ole = litchi_cfb::OleFile::open(Cursor::new(&bytes)).unwrap();
    assert!(
        Reader::from_ole_file_or_password_required(ole, options)
            .unwrap()
            .is_ok()
    );
}
