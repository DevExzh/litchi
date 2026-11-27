#![cfg(feature = "rtf")]

use litchi_core::{FileFormat, Metadata};
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::{Code, Conformance, CreateOptions, Flavor, create, create_with};

#[test]
fn fresh_rtf_adopts_native_empty_owner() {
    let file = create(FileFormat::Rtf).unwrap();
    assert_eq!(file.format(), Some(FileFormat::Rtf));
    assert_eq!(file.text().unwrap(), "");
    assert_eq!(file.document().unwrap().paragraph_count().unwrap(), 0);
    assert!(
        Capabilities::of(FileFormat::Rtf)
            .verb(Verb::Create)
            .is_supported()
    );
    // Publication is supported independently of the element-level CRUD adapter.
    assert!(
        Capabilities::of(FileFormat::Rtf)
            .verb(Verb::Save)
            .is_supported()
    );
}

#[test]
fn create_options_and_capability_gaps_agree() {
    for options in [
        CreateOptions::new().flavor(Flavor::Template),
        CreateOptions::new().conformance(Conformance::Strict),
        CreateOptions::new().metadata(Metadata::default()),
    ] {
        let error = create_with(FileFormat::Rtf, &options).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.format(), Some(FileFormat::Rtf));
    }
    let gaps = format!(
        "{:?}",
        Capabilities::of(FileFormat::Rtf).gaps().collect::<Vec<_>>()
    );
    for name in ["template", "strict-conformance", "initial-metadata"] {
        assert!(gaps.contains(name), "{gaps}");
    }
}

#[test]
fn builder_limit_keeps_native_typed_source() {
    let limits = litchi_rtf::SourceLimits::builder()
        .max_input_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let native = litchi_rtf::Builder::new()
        .source_limits(limits)
        .build()
        .unwrap_err();
    let error = litchi_ops::Error::from(native);
    assert_eq!(error.code(), Code::LimitExceeded);
    assert_eq!(error.format(), Some(FileFormat::Rtf));
    assert!(matches!(
        error.leaf(),
        litchi_ops::error::Source::RtfBuild(_)
    ));
}
