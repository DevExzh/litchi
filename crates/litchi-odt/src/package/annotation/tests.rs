//! Focused scanner and mutation invariants for ODT annotations.

use super::*;

const NS: &str = "xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0' xmlns:text='urn:oasis:names:tc:opendocument:xmlns:text:1.0' xmlns:table='urn:oasis:names:tc:opendocument:xmlns:table:1.0' xmlns:draw='urn:oasis:names:tc:opendocument:xmlns:drawing:1.0' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:meta='urn:oasis:names:tc:opendocument:xmlns:meta:1.0'";

#[test]
fn scans_rich_nested_text_ranges_and_initials() {
    let xml = format!(
        "<office:document {NS}><office:body><office:text><text:p><office:annotation office:name='outer'><dc:creator>Ada</dc:creator><dc:date>2026-07-19T00:00:00Z</dc:date><meta:creator-initials>AL</meta:creator-initials><text:p>rich <office:annotation><text:p>nested</text:p></office:annotation></text:p></office:annotation>x<office:annotation-end office:name='outer'/></text:p></office:text></office:body></office:document>"
    );
    let items = annotations(&xml, AnnotationHost::Text).unwrap();
    assert_eq!(items.len(), 2);
    assert_eq!(items[0].annotation.initials().as_deref(), Some("AL"));
    assert!(items[0].anchor.end.is_some());
    assert_eq!(
        items[1].anchor.start,
        AnnotationPosition::AnnotationBody {
            annotation_index: 0
        }
    );
}

#[test]
fn scans_spreadsheet_cell_and_presentation_shape_anchors() {
    let ods = format!(
        "<office:document {NS}><office:body><office:spreadsheet><table:table><table:table-row><table:table-cell><office:annotation><text:p>cell</text:p></office:annotation></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document>"
    );
    let item = annotations(&ods, AnnotationHost::Spreadsheet)
        .unwrap()
        .remove(0);
    assert_eq!(
        item.anchor.start,
        AnnotationPosition::SpreadsheetCell {
            sheet_index: 0,
            row: 0,
            column: 0
        }
    );

    let odp = format!(
        "<office:document {NS}><office:body><office:presentation><draw:page><draw:frame draw:name='Title'><office:annotation><text:p>shape</text:p></office:annotation></draw:frame></draw:page></office:presentation></office:body></office:document>"
    );
    let item = annotations(&odp, AnnotationHost::Presentation)
        .unwrap()
        .remove(0);
    assert_eq!(
        item.anchor.start,
        AnnotationPosition::PresentationShape {
            page_index: 0,
            shape_name: "Title".to_string()
        }
    );
}

#[test]
fn rejects_crossing_and_duplicate_ranges() {
    let crossing = format!(
        "<office:document {NS}><office:body><office:text><text:p><office:annotation office:name='a'/><office:annotation office:name='b'/>x<office:annotation-end office:name='a'/><office:annotation-end office:name='b'/></text:p></office:text></office:body></office:document>"
    );
    assert!(annotations(&crossing, AnnotationHost::Text).is_err());
    let duplicate = format!(
        "<office:document {NS}><office:body><office:text><text:p><office:annotation office:name='a'/><office:annotation office:name='a'/></text:p></office:text></office:body></office:document>"
    );
    assert!(annotations(&duplicate, AnnotationHost::Text).is_err());
}

#[test]
fn generated_text_mutations_preserve_unknown_xml() {
    let xml = format!(
        "<office:document {NS} xmlns:v='urn:vendor'><office:body><office:text><text:p><v:keep key='1'/>text</text:p></office:text></office:body></office:document>"
    );
    let mut annotation = Annotation::new("review");
    annotation.set_name(Some("r1"));
    annotation.set_creator(Some("Ada"));
    let anchor = AnnotationAnchor::range(
        AnnotationPosition::TextParagraph { paragraph_index: 0 },
        AnnotationPosition::TextParagraph { paragraph_index: 0 },
    );
    let (updated, index) = add_xml(&xml, AnnotationHost::Text, &anchor, &annotation).unwrap();
    assert_eq!(index, 0);
    assert!(updated.contains("<v:keep key='1'/>") && updated.contains("office:annotation-end"));
    let replaced = replace_xml(&updated, AnnotationHost::Text, 0, &annotation).unwrap();
    let removed = remove_xml(&replaced, AnnotationHost::Text, 0).unwrap();
    assert!(removed.contains("<v:keep key='1'/>") && !removed.contains("office:annotation"));
}
