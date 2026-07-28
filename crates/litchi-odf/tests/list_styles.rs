use litchi_odf::{
    BulletRelativeSize, FlatOpenDocument, ListLevelBulletStyle, ListLevelImageSource,
    ListLevelKind, ListLevelStyle, ListStyle, OdfOutlineNumberFormat, OdfOutlinePositiveInteger,
    parse_list_styles,
};
use std::io::Cursor;
const O: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const S: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const T: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const X: &str = "http://www.w3.org/1999/xlink";
fn wrap(x: &str) -> String {
    format!(r#"<o:styles xmlns:o="{O}" xmlns:s="{S}" xmlns:t="{T}" xmlns:x="{X}">{x}</o:styles>"#)
}

#[test]
fn parses_number_bullet_and_image_levels() {
    let x = wrap(concat!(
        r#"<t:list-style s:name="L1" s:display-name="List 1" t:consecutive-numbering="true">"#,
        r#"<t:list-level-style-number t:level="1" t:style-name="Num" s:num-format="a" s:num-letter-sync="true" s:num-prefix="(" s:num-suffix=")" t:display-levels="2" t:start-value="3"><s:list-level-properties t:list-level-position-and-space-mode="label-alignment"><s:list-level-label-alignment t:label-followed-by="space"/></s:list-level-properties><s:text-properties s:font-name="f"/></t:list-level-style-number>"#,
        r#"<t:list-level-style-bullet t:level="2" t:bullet-char="•" t:bullet-relative-size="75%"/>"#,
        r#"<t:list-level-style-image t:level="3" x:type="simple" x:href="Pictures/bullet.png" x:show="embed" x:actuate="onLoad"/>"#,
        r#"</t:list-style>"#,
    ));
    let set = parse_list_styles(&x).unwrap();
    let style = set.get("L1").unwrap();
    assert_eq!(style.display_name.as_deref(), Some("List 1"));
    assert_eq!(style.consecutive_numbering, Some(true));
    assert_eq!(style.levels.len(), 3);

    let ListLevelKind::Number(number) = &style.level(1).unwrap().kind else {
        panic!("level 1 must be numbered");
    };
    assert_eq!(number.format.as_ref().unwrap().as_str(), "a");
    assert_eq!(number.letter_sync, Some(true));
    assert_eq!(number.prefix.as_deref(), Some("("));
    assert_eq!(number.suffix.as_deref(), Some(")"));
    assert_eq!(number.display_levels.as_ref().unwrap().as_str(), "2");
    assert_eq!(number.start_value.as_ref().unwrap().as_str(), "3");
    assert_eq!(style.level(1).unwrap().style_name.as_deref(), Some("Num"));

    let ListLevelKind::Bullet(bullet) = &style.level(2).unwrap().kind else {
        panic!("level 2 must be a bullet");
    };
    assert_eq!(bullet.bullet_char, '•');
    assert_eq!(bullet.relative_size.as_ref().unwrap().as_str(), "75%");

    let ListLevelKind::Image(ListLevelImageSource::Linked(href)) = &style.level(3).unwrap().kind
    else {
        panic!("level 3 must be a linked image");
    };
    assert_eq!(href, "Pictures/bullet.png");
}

#[test]
fn parses_embedded_binary_image_level() {
    let x = wrap(
        r#"<t:list-style s:name="Img"><t:list-level-style-image t:level="1"><o:binary-data>aGVsbG8=</o:binary-data></t:list-level-style-image></t:list-style>"#,
    );
    let set = parse_list_styles(&x).unwrap();
    let ListLevelKind::Image(ListLevelImageSource::Embedded(data)) =
        &set.get("Img").unwrap().level(1).unwrap().kind
    else {
        panic!("level 1 must be an embedded image");
    };
    assert_eq!(data, "aGVsbG8=");
}

#[test]
fn round_trip_fragment_reparses_identically() {
    let mut style = ListStyle::new("RT").unwrap();
    style.consecutive_numbering = Some(false);
    style.levels.push(ListLevelStyle {
        level: 1,
        style_name: None,
        kind: ListLevelKind::Number(litchi_odf::ListLevelNumberStyle {
            format: Some(OdfOutlineNumberFormat::new("I").unwrap()),
            prefix: None,
            suffix: Some(".".to_string()),
            letter_sync: None,
            display_levels: None,
            start_value: Some(OdfOutlinePositiveInteger::new("4").unwrap()),
        }),
    });
    style.levels.push(ListLevelStyle {
        level: 2,
        style_name: Some("Bullet_20_Symbols".to_string()),
        kind: ListLevelKind::Bullet(ListLevelBulletStyle {
            bullet_char: '⚑',
            relative_size: Some(BulletRelativeSize::new("110%").unwrap()),
            prefix: None,
            suffix: None,
        }),
    });
    style.levels.push(ListLevelStyle {
        level: 3,
        style_name: None,
        kind: ListLevelKind::Image(ListLevelImageSource::Embedded("AAAA".to_string())),
    });
    let fragment = style.to_xml_fragment().unwrap();
    assert!(fragment.contains(r#"text:consecutive-numbering="false""#));
    assert!(fragment.contains(r#"text:bullet-char="⚑""#));
    assert!(fragment.contains("<office:binary-data>AAAA</office:binary-data>"));

    let reparsed = parse_list_styles(&wrap(&fragment)).unwrap();
    let round = reparsed.get("RT").unwrap();
    assert_eq!(round, &style);
}

#[test]
fn linked_image_round_trip() {
    let mut style = ListStyle::new("Link").unwrap();
    style.levels.push(ListLevelStyle {
        level: 1,
        style_name: None,
        kind: ListLevelKind::Image(ListLevelImageSource::Linked("images/dot.png".to_string())),
    });
    let fragment = style.to_xml_fragment().unwrap();
    assert!(fragment.contains(r#"xlink:href="images/dot.png""#));
    let reparsed = parse_list_styles(&wrap(&fragment)).unwrap();
    assert_eq!(reparsed.get("Link").unwrap(), &style);
}

#[test]
fn parses_flat_odt_fixture() {
    let bytes = include_bytes!("../../../test-data/odf/odt/note-ordinary-numbering.fodt");
    let flat = FlatOpenDocument::from_reader(Cursor::new(bytes)).unwrap();
    let set = flat.list_styles().unwrap();
    let l1 = set.get("L1").expect("fixture declares list style L1");
    let ListLevelKind::Number(number) = &l1.level(1).unwrap().kind else {
        panic!("L1 level 1 must be numbered");
    };
    assert_eq!(number.format.as_ref().unwrap().as_str(), "1");
    assert_eq!(number.suffix.as_deref(), Some("."));
    let l2 = set.get("L2").expect("fixture declares list style L2");
    let ListLevelKind::Bullet(bullet) = &l2.level(1).unwrap().kind else {
        panic!("L2 level 1 must be a bullet");
    };
    assert_eq!(bullet.bullet_char, '•');
}

#[test]
fn parses_libreoffice_list_style_fixture() {
    let bytes = include_bytes!(
        "../../../test-data/libreoffice-core/xmloff/qa/unit/data/differentListStylesInOneList.fodt"
    );
    let flat = FlatOpenDocument::from_reader(Cursor::new(bytes)).unwrap();
    let set = flat.list_styles().unwrap();
    let one = set.get("ListStyleOne").expect("fixture declares ListStyleOne");
    assert!(!one.levels.is_empty());
    assert!(
        one.levels
            .iter()
            .all(|level| matches!(level.kind, ListLevelKind::Number(_)))
    );
}

#[test]
fn rejects_invalid_declarations() {
    // Missing text:bullet-char.
    let x = wrap(r#"<t:list-style s:name="B"><t:list-level-style-bullet t:level="1"/></t:list-style>"#);
    assert!(parse_list_styles(&x).is_err());
    // Duplicate level.
    let x = wrap(
        r#"<t:list-style s:name="D"><t:list-level-style-number t:level="1" s:num-format="1"/><t:list-level-style-number t:level="1" s:num-format="1"/></t:list-style>"#,
    );
    assert!(parse_list_styles(&x).is_err());
    // Level zero.
    let x = wrap(
        r#"<t:list-style s:name="Z"><t:list-level-style-number t:level="0" s:num-format="1"/></t:list-style>"#,
    );
    assert!(parse_list_styles(&x).is_err());
    // num-letter-sync without an alphabetic format.
    let x = wrap(
        r#"<t:list-style s:name="S"><t:list-level-style-number t:level="1" s:num-format="1" s:num-letter-sync="true"/></t:list-style>"#,
    );
    assert!(parse_list_styles(&x).is_err());
    // Image without href or binary data.
    let x = wrap(r#"<t:list-style s:name="I"><t:list-level-style-image t:level="1"/></t:list-style>"#);
    assert!(parse_list_styles(&x).is_err());
    // Image combining href and binary data.
    let x = wrap(
        r#"<t:list-style s:name="C"><t:list-level-style-image t:level="1" x:href="a.png"><o:binary-data>AAAA</o:binary-data></t:list-level-style-image></t:list-style>"#,
    );
    assert!(parse_list_styles(&x).is_err());
    // Unknown attribute in an ODF namespace.
    let x = wrap(
        r#"<t:list-style s:name="U"><t:list-level-style-bullet t:level="1" t:bullet-char="-" t:bogus="1"/></t:list-style>"#,
    );
    assert!(parse_list_styles(&x).is_err());
    // Missing style:name.
    let x = wrap(r#"<t:list-style><t:list-level-style-number t:level="1"/></t:list-style>"#);
    assert!(parse_list_styles(&x).is_err());
    // Multi-character bullet.
    let x = wrap(
        r#"<t:list-style s:name="M"><t:list-level-style-bullet t:level="1" t:bullet-char="--"/></t:list-style>"#,
    );
    assert!(parse_list_styles(&x).is_err());
}

#[test]
fn rejects_invalid_newtypes() {
    assert!(BulletRelativeSize::new("75").is_err());
    assert!(BulletRelativeSize::new("-75%").is_err());
    assert!(BulletRelativeSize::new("7a5%").is_err());
    assert!(ListLevelBulletStyle::new('\n').is_err());
    assert!(ListStyle::new("").is_err());
}
