//! Canonical option setters retain explicit policy and redact credentials.
use litchi_xls::{CompatibilityProfile, Leniency, OpenOptions};

#[test]
fn options_are_consuming_and_password_debug_is_redacted() {
    let defaults = OpenOptions::default();
    assert_eq!(defaults.password_value(), None);
    assert_eq!(defaults.leniency_value(), Leniency::Strict);
    assert_eq!(
        defaults.compatibility_profile_value(),
        CompatibilityProfile::Strict
    );
    let configured = defaults
        .password("private-workbook-password")
        .leniency(Leniency::TolerateFormattingDefects)
        .compatibility_profile(CompatibilityProfile::SharedFormulaFlagWithoutPtgExpV1);
    assert_eq!(
        configured.password_value(),
        Some("private-workbook-password")
    );
    assert_eq!(
        configured.leniency_value(),
        Leniency::TolerateFormattingDefects
    );
    assert_eq!(
        configured.compatibility_profile_value(),
        CompatibilityProfile::SharedFormulaFlagWithoutPtgExpV1
    );
    assert_eq!(defaults.password_value(), None);
    let debug = format!("{configured:?}");
    assert!(!debug.contains("private-workbook-password"));
    assert!(debug.contains("[REDACTED]"));
    assert_eq!(configured.password(None).password_value(), None);
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
    assert_eq!(configured.source_limits_value(), limits);
    assert_eq!(configured.password(None).source_limits_value(), limits);
}

#[test]
fn valid_workbook_opens_at_exact_physical_and_stream_boundaries() {
    use litchi_xls::{Builder, SourceLimits, raw::Reader};
    use std::io::Cursor;
    let mut writer = Builder::new();
    let sheet = writer.add_worksheet("Data").unwrap();
    writer.write_number(sheet, 0, 0, 42.0).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    let bytes = output.into_inner();
    let ole = litchi_cfb::OleFile::open(Cursor::new(&bytes), litchi_cfb::OleFileLimits::default())
        .unwrap();
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
    let workbook = Reader::new_with_options(Cursor::new(&bytes), options).unwrap();
    assert_eq!(workbook.source_limits(), limits);
    Reader::from_ole_file_with_options(ole, options).unwrap();
    let ole = litchi_cfb::OleFile::open(Cursor::new(&bytes), litchi_cfb::OleFileLimits::default())
        .unwrap();
    assert!(
        Reader::from_ole_file_or_password_required(ole, options)
            .unwrap()
            .is_ok()
    );
}
