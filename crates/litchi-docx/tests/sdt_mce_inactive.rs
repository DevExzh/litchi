use litchi_docx::content_control::{Inventory, Kind, Limits};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const W15: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const HASH: &str = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash";

#[test]
fn inactive_choice_does_not_enforce_extension_ignorable_ownership() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:w15="{W15}" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><w:sdtPr><w15:repeatingSection/></w:sdtPr></mc:Choice><mc:Fallback><w:sdtPr><w:id w:val="22"/></w:sdtPr></mc:Fallback></mc:AlternateContent></w:document>"#
    );

    let inventory = Inventory::parse(xml.as_bytes()).unwrap();
    assert_eq!(inventory.occurrences().len(), 1);
    assert_eq!(inventory.occurrences()[0].id(), Some(22));
    assert_eq!(inventory.occurrences()[0].control().kind(), Kind::RichText);
}

#[test]
fn inactive_fallback_does_not_enforce_extension_attribute_ignorable_ownership() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:w15="{W15}" xmlns:h="{HASH}"><mc:AlternateContent><mc:Choice Requires="w15"><w:sdtPr><w:id w:val="15"/></w:sdtPr></mc:Choice><mc:Fallback><w:sdtPr><w:dataBinding w:xpath="/inactive" w:storeItemID="{{11111111-1111-4111-8111-111111111111}}" h:storeItemChecksum="AAAAAA=="/></w:sdtPr></mc:Fallback></mc:AlternateContent></w:document>"#
    );

    let inventory = Inventory::parse(xml.as_bytes()).unwrap();
    assert_eq!(inventory.occurrences().len(), 1);
    assert_eq!(inventory.occurrences()[0].id(), Some(15));
    assert!(
        inventory.occurrences()[0]
            .control()
            .data_binding()
            .is_none()
    );
}

#[test]
fn selected_branch_still_requires_effective_ignorable_ownership() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:w15="{W15}"><mc:AlternateContent><mc:Choice Requires="w15"><w:sdtPr><w15:repeatingSection/></w:sdtPr></mc:Choice><mc:Fallback><w:sdtPr><w:id w:val="2"/></w:sdtPr></mc:Fallback></mc:AlternateContent></w:document>"#
    );

    assert!(Inventory::parse(xml.as_bytes()).is_err());
}

#[test]
fn inactive_branch_still_counts_toward_raw_mce_security_limits() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><w:sdtPr xmlns:x="urn:future" mc:Ignorable="x"><w:id w:val="1"/></w:sdtPr></mc:Choice><mc:Fallback><w:sdtPr><w:id w:val="2"/></w:sdtPr></mc:Fallback></mc:AlternateContent></w:document>"#
    );
    let mut limits = Limits::default();
    limits.max_metadata_bytes = 0;

    assert!(Inventory::parse_with_limits(xml.as_bytes(), &limits).is_err());
}
