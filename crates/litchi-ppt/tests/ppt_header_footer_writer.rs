use litchi_ppt::writer::Writer;
use litchi_ppt::{
    HeaderFooter, HeaderFooterOptions, HeaderFooterParent, HeaderFooterParentOrdinal,
    HeaderFooterScope, Package,
};
use std::io::Cursor;

fn header_footer(footer: &str) -> HeaderFooter {
    HeaderFooter {
        scope: HeaderFooterScope::PresentationSlides,
        options: HeaderFooterOptions {
            show_footer: true,
            ..HeaderFooterOptions::default()
        },
        user_date: None,
        header: None,
        footer: Some(footer.to_string()),
        placeholder_display: None,
    }
}

fn write(writer: &mut Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn writer_attaches_header_footers_to_every_supported_parent() {
    let mut writer = Writer::new();
    writer.add_slide().unwrap();
    writer.add_slide().unwrap();

    writer
        .set_presentation_header_footer(header_footer("全局页脚"))
        .unwrap();
    let mut notes = header_footer("Notes footer");
    notes.header = Some("Notes 标题".to_string());
    notes.options.show_header = true;
    writer.set_notes_and_handouts_header_footer(notes).unwrap();
    writer
        .set_main_master_header_footer(header_footer("Master footer"))
        .unwrap();
    writer
        .set_slide_header_footer(1, header_footer("Moved slide 页脚"))
        .unwrap();
    writer.move_slide(1, 0).unwrap();

    let mut package = Package::from_reader(Cursor::new(write(&mut writer))).unwrap();
    let values = package.presentation().unwrap().header_footers().unwrap();

    assert_eq!(values.entries().len(), 4);
    assert_eq!(
        values.entries()[0].scope,
        HeaderFooterScope::PresentationSlides
    );
    assert_eq!(values.entries()[0].footer.as_deref(), Some("全局页脚"));
    assert_eq!(
        values.entries()[1].scope,
        HeaderFooterScope::NotesAndHandouts
    );
    assert_eq!(values.entries()[1].header.as_deref(), Some("Notes 标题"));
    assert!(values.entries().iter().any(|entry| {
        entry.scope
            == HeaderFooterScope::Local {
                parent: HeaderFooterParent::MainMaster,
                parent_ordinal: HeaderFooterParentOrdinal::new(0),
            }
            && entry.footer.as_deref() == Some("Master footer")
    }));
    assert!(values.entries().iter().any(|entry| {
        entry.scope
            == HeaderFooterScope::Local {
                parent: HeaderFooterParent::Slide,
                parent_ordinal: HeaderFooterParentOrdinal::new(0),
            }
            && entry.footer.as_deref() == Some("Moved slide 页脚")
    }));
}

#[test]
fn rejected_local_header_does_not_replace_existing_slide_value() {
    let mut writer = Writer::new();
    writer.add_slide().unwrap();
    writer
        .set_slide_header_footer(0, header_footer("kept"))
        .unwrap();

    let mut invalid = header_footer("replacement");
    invalid.header = Some("not valid on a slide".to_string());
    assert!(writer.set_slide_header_footer(0, invalid).is_err());
    assert_eq!(
        writer
            .slide_header_footer(0)
            .unwrap()
            .unwrap()
            .footer
            .as_deref(),
        Some("kept")
    );
}
