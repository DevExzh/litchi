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
