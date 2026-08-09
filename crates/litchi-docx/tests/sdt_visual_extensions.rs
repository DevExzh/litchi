use litchi_docx::content_control::{
    Appearance, AuthoringView, Inventory, Kind, Limits, Lock, SdtColor, WebExtensionBinding,
    write_sdt_pr,
};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W15: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

#[test]
fn reads_typed_visual_and_inert_web_extension_metadata() {
    let source = format!(
        r#"<w:document xmlns:w="{W}" xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w15"><w:body><w:sdt><w:sdtPr><w:id w:val="7"/><w15:appearance w15:val="tags"/><w15:color w:val="a1B2c3"/><w15:webExtensionLinked w:val="0"/><w15:webExtensionCreated w:val="on"/></w:sdtPr></w:sdt></w:body></w:document>"#
    );

    let inventory = Inventory::parse(source.as_bytes()).unwrap();
    let control = inventory.occurrences()[0].control();
    assert_eq!(control.appearance(), Some(Appearance::Tags));
    assert_eq!(control.color(), Some(SdtColor::Rgb([0xA1, 0xB2, 0xC3])));
    assert_eq!(control.web_extension_linked(), Some(false));
    assert_eq!(control.web_extension_created(), Some(true));
    assert_eq!(
        control.web_extension_binding(),
        Some(WebExtensionBinding::Created(true))
    );
}

#[test]
fn authors_compact_visual_and_inert_web_extension_metadata() {
    let view = AuthoringView::new(7, Kind::RichText, Lock::Unlocked)
        .appearance(Some(Appearance::Hidden))
        .color(Some(SdtColor::Rgb([0xA1, 0xB2, 0xC3])))
        .web_extension_linked(Some(false))
        .web_extension_created(Some(true));

    let authored = write_sdt_pr(&view, &Limits::default()).unwrap();
    assert!(authored.namespace_requirements().word_2012());
    assert_eq!(
        authored.xml(),
        format!(
            r#"<w:sdtPr xmlns:mc="{MC}" xmlns:w15="{W15}" mc:Ignorable="w15"><w:id w:val="7"/><w15:appearance w15:val="hidden"/><w15:color w:val="A1B2C3"/><w15:webExtensionLinked w:val="0"/><w15:webExtensionCreated w:val="1"/><w:richText/></w:sdtPr>"#
        )
    );
    assert!(!authored.xml().contains('\n'));
}

#[test]
fn rejects_invalid_visual_extension_values() {
    for body in [
        r#"<w15:appearance w15:val="outline"/>"#,
        r#"<w15:appearance/>"#,
        r#"<w15:color w:val="beef"/>"#,
        r#"<w15:color w:val="nothex"/>"#,
    ] {
        let source = format!(
            r#"<w:document xmlns:w="{W}" xmlns:w15="{W15}" xmlns:mc="{MC}" mc:Ignorable="w15"><w:body><w:sdtPr>{body}</w:sdtPr></w:body></w:document>"#
        );
        assert!(
            Inventory::parse(source.as_bytes()).is_err(),
            "accepted {body}"
        );
    }
}
