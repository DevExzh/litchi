use super::*;

fn offset(source: &[u8], needle: &[u8]) -> u32 {
    u32::try_from(
        source
            .windows(needle.len())
            .position(|window| window == needle)
            .expect("test XML contains the requested element"),
    )
    .expect("test source offset fits in u32")
}

#[test]
fn active_offsets_decode_entity_escaped_mce_namespace_declarations() {
    let source = format!(
        r#"<p:connections xmlns:p="{CORE_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility&#x2F;2006" xmlns:u="urn:unsupported"><?before http://schemas.openxmlformats.org/markup-compatibility/2006?><mc:AlternateContent><mc:Choice Requires="u"><u:connection id="9"/></mc:Choice><mc:Fallback><p:connection id="1"/></mc:Fallback></mc:AlternateContent></p:connections>"#
    );
    let inactive = offset(source.as_bytes(), b"<u:connection id=\"9\"");
    let active = offset(source.as_bytes(), b"<p:connection id=\"1\"");

    assert_eq!(
        active_source_offsets(source.as_bytes(), &[inactive, active, inactive])
            .expect("MCE source offsets"),
        Some(vec![active])
    );
}

#[test]
fn active_offsets_require_an_mce_namespace_declaration() {
    let source =
        b"<root>http://schemas.openxmlformats.org/markup-compatibility/2006<child/></root>";
    let child = offset(source, b"<child/>");

    assert_eq!(
        active_source_offsets(source, &[child]).expect("plain source offsets"),
        Some(vec![child])
    );
}

#[test]
fn processing_instruction_offset_mapping_is_streaming_and_exact() {
    let source = b"<root><?one?><first/><?two data?><second/></root>";
    let mut mapped = vec![offset(source, b"<first/>"), offset(source, b"<second/>")];
    let filtered = strip_processing_instructions(source).expect("valid PI source");
    let expected = vec![
        offset(filtered.as_ref(), b"<first/>"),
        offset(filtered.as_ref(), b"<second/>"),
    ];

    processing_instruction_ranges(source, &mut mapped).expect("PI offsets map");
    assert_eq!(mapped, expected);
}
