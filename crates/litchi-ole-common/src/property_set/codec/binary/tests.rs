//! Focused invariants for the binary Property Set facade.

use super::super::super::model::{SUMMARY_INFORMATION_FMTID, Section, Stream, Value};

#[test]
fn facade_round_trips_typed_property_values() {
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section.add(2, Value::I4(42)).expect("valid property value");

    let bytes = Stream::new(section)
        .to_bytes()
        .expect("serializable property set");
    let parsed = Stream::parse(&bytes).expect("parseable property set");

    assert_eq!(parsed.sections[0].property(2), Some(&Value::I4(42)));
}
