use litchi_odt::Builder;
use litchi_odt::style::paragraph::breaks::{Break, Breaks, PageNumber, Style, parse, set_xml};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const F: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const T: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
fn wrap(x: &str) -> String {
    format!(r#"<o:styles xmlns:o="{O}" xmlns:s="{S}" xmlns:f="{F}" xmlns:t="{T}">{x}</o:styles>"#)
}

fn full_properties() -> &'static str {
    r#"<s:style s:name="P1" s:family="paragraph"><s:paragraph-properties f:break-before="page" f:break-after="column" s:page-number="7" t:number-lines="false" t:line-number="0" f:keep-with-next="always" f:line-height="115%"/></s:style>"#
}

#[test]
fn parses_all_break_attributes() {
    let set = parse(&wrap(full_properties())).unwrap();
    let style = set.get("P1").unwrap();
    let p = style.properties.as_ref().unwrap();
    assert_eq!(p.break_before, Some(Break::Page));
    assert_eq!(p.break_after, Some(Break::Column));
    assert_eq!(p.page_number, Some(PageNumber::Number(7)));
    assert_eq!(p.number_lines, Some(false));
    assert_eq!(p.line_number, Some(0));
}

#[test]
fn ignores_sibling_owned_attributes_and_foreign_families() {
    let xml = wrap(
        r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:line-height="200%" f:break-before="auto"/></s:style><s:style s:name="T" s:family="table"><s:paragraph-properties f:break-before="page"/></s:style>"#,
    );
    let set = parse(&xml).unwrap();
    assert_eq!(set.styles.len(), 1);
    let p = set.get("P").unwrap().properties.as_ref().unwrap();
    assert_eq!(p.break_before, Some(Break::Auto));
    assert!(p.break_after.is_none());
    assert!(set.get("T").is_none());
}

#[test]
fn round_trip_serialize_reparse() {
    let set = parse(&wrap(full_properties())).unwrap();
    let style = set.get("P1").unwrap();
    let fragment = style.to_xml_fragment().unwrap();
    let reparsed = parse(&wrap(&fragment)).unwrap();
    assert_eq!(reparsed.get("P1"), Some(style));
    let default = Style::default_style(Some(Breaks {
        page_number: Some(PageNumber::Auto),
        number_lines: Some(true),
        ..Default::default()
    }));
    let fragment = default.to_xml_fragment().unwrap();
    let reparsed = parse(&wrap(&fragment)).unwrap();
    assert_eq!(reparsed.default_style(), Some(&default));
}

#[test]
fn rejects_malformed_values_and_duplicates() {
    let bad = [
        // Unknown break token.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:break-before="always"/></s:style>"#,
        ),
        // Page number zero is not positive.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:page-number="0"/></s:style>"#,
        ),
        // Page number must be auto or an integer.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties s:page-number="first"/></s:style>"#,
        ),
        // Negative line number is not a nonNegativeInteger.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties t:line-number="-1"/></s:style>"#,
        ),
        // Malformed boolean.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties t:number-lines="yes"/></s:style>"#,
        ),
        // Duplicate owned attribute (same expanded name).
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties f:break-after="page" f:break-after="auto"/></s:style>"#,
        ),
        // Duplicate properties element.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties/><s:paragraph-properties/></s:style>"#,
        ),
        // Duplicate style identity.
        wrap(
            r#"<s:style s:name="P" s:family="paragraph"><s:paragraph-properties/></s:style><s:style s:name="P" s:family="paragraph"><s:paragraph-properties/></s:style>"#,
        ),
        // Broken identity: default style with a name.
        wrap(
            r#"<s:default-style s:name="P" s:family="paragraph"><s:paragraph-properties/></s:default-style>"#,
        ),
    ];
    for xml in bad {
        assert!(parse(&xml).is_err(), "accepted {xml}");
    }
}

#[test]
fn mutation_replaces_inserts_and_strips_owned_attributes() {
    let xml = wrap(full_properties());
    // Replace: owned attributes are rewritten, sibling attributes survive.
    let changed = Style::named(
        "P1",
        Some(Breaks {
            break_before: Some(Break::Column),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = set_xml(&xml, &changed).unwrap();
    assert!(updated.contains(r#"fo:break-before="column""#));
    assert!(!updated.contains("break-after"));
    assert!(!updated.contains("page-number"));
    assert!(updated.contains(r#"f:keep-with-next="always""#));
    assert!(updated.contains(r#"f:line-height="115%""#));
    let reparsed = parse(&updated).unwrap();
    assert_eq!(
        reparsed.get("P1").unwrap().properties.as_ref().unwrap(),
        changed.properties.as_ref().unwrap()
    );
    // Strip: None removes only this module's attributes.
    let stripped = Style::named("P1", None).unwrap();
    let updated = set_xml(&xml, &stripped).unwrap();
    assert!(!updated.contains("break-before"));
    assert!(updated.contains(r#"f:keep-with-next="always""#));
    let reparsed = parse(&updated).unwrap();
    let p = reparsed.get("P1").unwrap().properties.as_ref().unwrap();
    assert_eq!(p, &Breaks::default());
    // Missing target is an error.
    assert!(set_xml(&xml, &changed_named("Other")).is_err());
    // Insert into a style without a properties element.
    fn changed_named(name: &str) -> Style {
        Style::named(
            name,
            Some(Breaks {
                break_after: Some(Break::Page),
                ..Default::default()
            }),
        )
        .unwrap()
    }
    let bare = wrap(r#"<s:style s:name="B" s:family="paragraph"/>"#);
    let updated = set_xml(&bare, &changed_named("B")).unwrap();
    let reparsed = parse(&updated).unwrap();
    assert_eq!(
        reparsed
            .get("B")
            .unwrap()
            .properties
            .as_ref()
            .unwrap()
            .break_after,
        Some(Break::Page)
    );
    // Insert into a style with other children but no properties element.
    let with_text =
        wrap(r#"<s:style s:name="C" s:family="paragraph"><s:text-properties/></s:style>"#);
    let updated = set_xml(&with_text, &changed_named("C")).unwrap();
    assert!(updated.contains("<s:text-properties/>"));
    let reparsed = parse(&updated).unwrap();
    assert!(reparsed.get("C").unwrap().properties.is_some());
    // Nothing to strip and no properties element: unchanged.
    let unchanged = set_xml(&bare, &Style::named("B", None).unwrap()).unwrap();
    assert_eq!(unchanged, bare);
}

#[test]
fn mutation_under_prefix_aliases() {
    // Aliased prefixes: owned attributes resolve by namespace, not by prefix.
    let xml = format!(
        r#"<a:styles xmlns:a="{O}" xmlns:b="{S}" xmlns:c="{F}" xmlns:d="{T}"><b:style b:name="P" b:family="paragraph"><b:paragraph-properties c:break-before="page" d:line-number="3" c:keep-together="always"/></b:style></a:styles>"#
    );
    let changed = Style::named(
        "P",
        Some(Breaks {
            break_before: Some(Break::Auto),
            ..Default::default()
        }),
    )
    .unwrap();
    let updated = set_xml(&xml, &changed).unwrap();
    assert!(updated.contains(r#"fo:break-before="auto""#));
    assert!(!updated.contains("line-number"));
    assert!(updated.contains(r#"c:keep-together="always""#));
    let reparsed = parse(&updated).unwrap();
    assert_eq!(
        reparsed.get("P").unwrap().properties.as_ref().unwrap(),
        changed.properties.as_ref().unwrap()
    );
}

#[test]
fn builder_package_round_trip() {
    let style = Style::named(
        "Body",
        Some(Breaks {
            break_before: Some(Break::Page),
            page_number: Some(PageNumber::Auto),
            number_lines: Some(false),
            line_number: Some(0),
            ..Default::default()
        }),
    )
    .unwrap();
    let mut builder = Builder::new();
    builder.add_paragraph_break_style(style.clone()).unwrap();
    assert!(
        builder
            .add_paragraph_break_style(Style::named("Body", None).unwrap())
            .is_err()
    );
    builder.add_paragraph("x").unwrap();
    let package = litchi_odt::generic::Package::from_bytes(builder.build().unwrap()).unwrap();
    assert_eq!(
        package.paragraph_style_breaks().unwrap().get("Body"),
        Some(&style)
    );
}

#[test]
fn parses_real_odf_fixtures() {
    // Line-numbering suppression and automatic page numbers from LibreOffice.
    let xml = include_str!("../../../test-data/odf/odt/note-tracked-changes.fodt");
    let set = parse(xml).unwrap();
    assert!(!set.styles.is_empty());
    let properties: Vec<_> = set
        .styles
        .iter()
        .filter_map(|x| x.properties.as_ref())
        .collect();
    assert!(
        properties
            .iter()
            .any(|p| p.number_lines == Some(false) && p.line_number == Some(0))
    );
    assert!(
        properties
            .iter()
            .any(|p| p.page_number == Some(PageNumber::Auto))
    );
    let flat = litchi_odt::generic::FlatDocument::from_reader(Cursor::new(xml)).unwrap();
    assert_eq!(
        flat.paragraph_style_breaks().unwrap().styles.len(),
        set.styles.len()
    );
    // Explicit page breaks before paragraphs.
    let breaks = include_str!(
        "../../../test-data/libreoffice-core/sw/qa/extras/pagelinespacing/data/pageColumns.fodt"
    );
    let set = parse(breaks).unwrap();
    assert!(
        set.styles
            .iter()
            .filter_map(|x| x.properties.as_ref())
            .any(|p| p.break_before == Some(Break::Page))
    );
}
