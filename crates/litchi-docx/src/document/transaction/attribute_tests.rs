//! Change 0764: a complex-field marker is read from its first `fldCharType`
//! at a cost bounded in the attributes of its start tag.

use super::*;

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// ` n00000=""` to ` n19999=""`, then 20,000 repeats of the last name: a tag
/// quick-xml's checked iterator reads in `O(n²)` when every item is read.
fn repeated_attributes() -> String {
    let mut attributes = String::new();
    for index in 0..20_000 {
        attributes.push_str(&format!(" n{index:05}=\"\""));
    }
    for _ in 0..20_000 {
        attributes.push_str(" n19999=\"\"");
    }
    attributes
}

fn marker(attributes: &str) -> Result<Option<ComplexFieldMarker>, Refusal> {
    let run = format!(r#"<w:r xmlns:w="{WORD}"><w:fldChar{attributes}/></w:r>"#);
    complex_field_marker(run.as_bytes())
}

#[test]
fn repeated_field_character_types_keep_their_first_value() {
    assert!(matches!(
        marker(r#" w:fldCharType="begin""#),
        Ok(Some(ComplexFieldMarker::Begin))
    ));
    assert!(matches!(
        marker(r#" w:fldCharType="separate" w:fldCharType="end""#),
        Ok(Some(ComplexFieldMarker::Separate))
    ));
    assert_eq!(
        marker(r#" w:dirty="1""#).err(),
        Some(Refusal::ComplexContent)
    );
    let names = repeated_attributes();
    assert!(matches!(
        marker(&format!(
            r#"{names} w:fldCharType="end" w:fldCharType="begin""#
        )),
        Ok(Some(ComplexFieldMarker::End))
    ));
    assert_eq!(marker(&names).err(), Some(Refusal::ComplexContent));
}

#[test]
fn a_malformed_field_character_type_before_a_well_formed_one_is_refused() {
    // The base refused this: quick-xml recorded the malformed first
    // occurrence and reported the second as a duplicate.
    assert_eq!(
        marker(r#" w:fldCharType=begin w:fldCharType="end""#).err(),
        Some(Refusal::ComplexContent)
    );
    // Any lexical attribute error before the type now refuses the marker too.
    assert_eq!(
        marker(r#" w:dirty=1 w:fldCharType="end""#).err(),
        Some(Refusal::ComplexContent)
    );
    // An error after the type is not reached, as before.
    assert!(matches!(
        marker(r#" w:fldCharType="end" w:dirty=1"#),
        Ok(Some(ComplexFieldMarker::End))
    ));
    // Repeated names before the type are skipped, as before.
    assert!(matches!(
        marker(r#" w:dirty="1" w:dirty="0" w:fldCharType="separate""#),
        Ok(Some(ComplexFieldMarker::Separate))
    ));
}
