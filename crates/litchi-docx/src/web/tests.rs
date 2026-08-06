use super::*;
use crate::color::Theme;
use litchi_opc::packuri::PackURI;
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, Part};

fn parse_settings(xml: &[u8]) -> Result<Settings> {
    parse(xml).map(|(settings, _)| settings)
}

fn read_settings_part(part: &dyn Part) -> Result<Settings> {
    read(part).map(|(settings, _)| settings)
}

fn contains_xml(xml: &[u8], needle: &str) -> bool {
    xml.windows(needle.len())
        .any(|window| window == needle.as_bytes())
}

fn id(value: i64) -> Id {
    Id::new(value).unwrap()
}

fn make_div(value: i64) -> Div {
    Div::new(id(value))
}

fn package(conformance: Conformance) -> OpcPackage {
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};

    let mut package = OpcPackage::new();
    let document = PackURI::new("/word/document.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
            document,
            ct::WML_DOCUMENT_MAIN.to_owned(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#.to_vec(),
        )));
    let relationship = match conformance {
        Conformance::Transitional => rt::OFFICE_DOCUMENT,
        Conformance::Strict => STRICT_OFFICE_DOCUMENT_RELATIONSHIP,
    };
    package.relate_to("word/document.xml", relationship);
    package
}

fn add_raw_web(package: &mut OpcPackage, xml: &[u8], conformance: Conformance) -> PackURI {
    let name = PackURI::new("/word/webSettings.xml").unwrap();
    package.add_part(Box::new(BlobPart::new(
        name.clone(),
        CONTENT_TYPE.to_owned(),
        xml.to_vec(),
    )));
    package
        .get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .relate_to("webSettings.xml", conformance.relationship());
    name
}

#[test]
fn semantic_and_positional_div_crud_is_checked_and_atomic() {
    let mut settings = Settings::default();
    let mut first = make_div(1);
    first.set_body_div(true);
    settings.add(first).unwrap();
    assert_eq!(settings.get(id(1)).unwrap().unwrap().id(), id(1));
    assert_eq!(settings.get(0).unwrap().unwrap().id(), id(1));

    let before = settings.clone();
    assert!(settings.add(make_div(1)).is_err());
    assert_eq!(settings, before);
    assert!(settings.move_to(id(1), 1).is_err());
    assert_eq!(settings, before);

    let mut replacement = make_div(1);
    replacement.set_block_quote(true);
    let old = settings.put(replacement).unwrap().unwrap();
    assert_eq!(old.is_body_div(), Some(true));
    assert_eq!(
        settings.get(id(1)).unwrap().unwrap().is_block_quote(),
        Some(true)
    );

    settings.add(make_div(2)).unwrap();
    settings.move_to(id(2), 0).unwrap();
    assert_eq!(settings.get(0).unwrap().unwrap().id(), id(2));
    assert_eq!(settings.remove(id(1)).unwrap().unwrap().id(), id(1));
    assert!(settings.remove(id(3)).unwrap().is_none());
    assert_eq!(settings.remove(0).unwrap().unwrap().id(), id(2));
    assert!(settings.divs().is_none());
}

#[test]
fn numeric_div_selectors_reject_missing_positions_without_mutation() {
    let mut settings = Settings::default();
    assert!(settings.get(id(99)).unwrap().is_none());
    assert!(settings.remove(id(99)).unwrap().is_none());
    assert!(settings.get(0usize).is_err());
    assert!(settings.remove(0usize).is_err());

    settings.add(make_div(1)).unwrap();
    let before = settings.clone();
    assert!(settings.get(1usize).is_err());
    assert!(settings.remove(1usize).is_err());
    assert_eq!(settings, before);

    let mut parent = make_div(10);
    assert!(parent.child(id(99)).unwrap().is_none());
    assert!(parent.remove_child(id(99)).unwrap().is_none());
    assert!(parent.child(0usize).is_err());
    assert!(parent.remove_child(0usize).is_err());

    parent.add_child(make_div(11)).unwrap();
    assert_eq!(parent.child(0usize).unwrap().unwrap().id(), id(11));
    let before = parent.clone();
    assert!(parent.child(1usize).is_err());
    assert!(parent.remove_child(1usize).is_err());
    assert_eq!(parent, before);
}

#[test]
fn package_graph_crud_round_trips_and_is_idempotent() {
    let mut package = package(Conformance::Transitional);
    assert!(load(&package).unwrap().is_none());
    assert!(!remove(&mut package).unwrap());

    let mut settings = Settings::default();
    settings.set_encoding("utf-8").unwrap().set_allow_png(true);
    assert!(put(&mut package, settings.clone(), Conformance::Transitional).unwrap());
    assert_eq!(
        load(&package).unwrap(),
        Some((settings.clone(), Conformance::Transitional))
    );
    assert!(!put(&mut package, settings.clone(), Conformance::Transitional).unwrap());

    settings.set_allow_png(false);
    assert!(put(&mut package, settings, Conformance::Transitional).unwrap());
    assert!(remove(&mut package).unwrap());
    assert!(load(&package).unwrap().is_none());
    assert!(!remove(&mut package).unwrap());
}

#[test]
fn semantic_noop_preserves_noncanonical_bytes_and_signatures() {
    use litchi_opc::constants::relationship_type as rt;

    let mut package = package(Conformance::Transitional);
    let source = br#"<?xml version="1.0"?>
          <w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:encoding w:val="utf-8"></w:encoding>
            <w:allowPNG w:val="1" />
          </w:webSettings>"#;
    let name = add_raw_web(&mut package, source, Conformance::Transitional);
    package.rels_mut().add_relationship(
        rt::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        "_xmlsignatures/origin.sigs".to_owned(),
        "rSignature".to_owned(),
        false,
    );
    let (settings, conformance) = load(&package).unwrap().unwrap();
    assert!(package.is_signed());

    assert!(put(&mut package, settings.clone(), Conformance::Strict).is_err());
    assert_eq!(package.get_part(&name).unwrap().blob(), source);
    assert!(package.is_signed());
    assert!(!put(&mut package, settings, conformance).unwrap());
    assert_eq!(package.get_part(&name).unwrap().blob(), source);
    assert!(package.is_signed());
}

#[test]
fn strict_graph_and_mce_round_trip() {
    let mut package = package(Conformance::Strict);
    let mut settings = Settings::default();
    settings.set_target_screen_size(Screen::Pixels1920x1200);
    assert!(put(&mut package, settings.clone(), Conformance::Strict).unwrap());
    assert_eq!(
        load(&package).unwrap(),
        Some((settings, Conformance::Strict))
    );

    let xml = br#"<w:webSettings xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:x="urn:unsupported" mc:Ignorable="x">
            <x:ignored/><w:allowPNG/>
        </w:webSettings>"#;
    let (parsed, conformance) = parse(xml).unwrap();
    assert_eq!(conformance, Conformance::Strict);
    assert_eq!(parsed.allow_png(), Some(true));
}

#[test]
fn graph_failures_are_atomic_and_shared_parts_are_rejected() {
    let mut duplicate = package(Conformance::Transitional);
    let document = PackURI::new("/word/document.xml").unwrap();
    {
        let part = duplicate.get_part_mut(&document).unwrap();
        part.rels_mut().add_relationship(
            Conformance::Transitional.relationship().to_owned(),
            "webSettings.xml".to_owned(),
            "rWeb1".to_owned(),
            false,
        );
        part.rels_mut().add_relationship(
            Conformance::Transitional.relationship().to_owned(),
            "webSettings2.xml".to_owned(),
            "rWeb2".to_owned(),
            false,
        );
    }
    let parts = duplicate.part_count();
    let relationships = duplicate.get_part(&document).unwrap().rels().iter().count();
    assert!(
        put(
            &mut duplicate,
            Settings::default(),
            Conformance::Transitional
        )
        .is_err()
    );
    assert_eq!(duplicate.part_count(), parts);
    assert_eq!(
        duplicate.get_part(&document).unwrap().rels().iter().count(),
        relationships
    );

    let mut shared = package(Conformance::Transitional);
    assert!(put(&mut shared, Settings::default(), Conformance::Transitional).unwrap());
    let name = PackURI::new("/word/webSettings.xml").unwrap();
    let bytes = shared.get_part(&name).unwrap().blob().to_vec();
    let mut other = BlobPart::new(
        PackURI::new("/word/other.xml").unwrap(),
        "application/xml".to_owned(),
        b"<other/>".to_vec(),
    );
    other.rels_mut().add_relationship(
        "urn:shared".to_owned(),
        "webSettings.xml".to_owned(),
        "rShared".to_owned(),
        false,
    );
    shared.add_part(Box::new(other));
    assert!(remove(&mut shared).is_err());
    assert_eq!(shared.get_part(&name).unwrap().blob(), bytes);
    assert!(load(&shared).unwrap().is_some());
}

#[test]
fn adversarial_xml_never_unwinds() {
    for xml in [
            b"".as_slice(),
            b"<w:webSettings".as_slice(),
            b"<webSettings/>".as_slice(),
            b"\xFF\xFE".as_slice(),
            br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:divs><w:div w:id="same"/><w:div w:id="same"/></w:divs></w:webSettings>"#.as_slice(),
        ] {
            let result = std::panic::catch_unwind(|| parse(xml));
            assert!(result.is_ok());
            assert!(result.unwrap().is_err());
        }
    let oversized = vec![b' '; MAX_XML_BYTES + 1];
    assert!(parse(&oversized).is_err());
}

#[test]
fn source_checked_web_edits_preserve_opaque_markup_and_round_trip_inverse() {
    let source = br#"<?xml version="1.0"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:future">
  <x:future w:val="untouched"/>
  <w:frameset>
    <w:sz w:val="1*"/><x:insideFrameset/><w:frameLayout w:val="rows"/>
  </w:frameset>
  <w:divs>
    <w:div w:id="7">
      <w:blockQuote w:val="1"/><w:bodyDiv w:val="0"/>
      <w:marLeft w:val="0"/><w:marRight w:val="0"/><w:marTop w:val="0"/><w:marBottom w:val="0"/>
      <w:divBdr><x:opaqueBorder/><w:top w:val="single" w:color="FF0000"/></w:divBdr>
    </w:div>
  </w:divs>
  <w:targetScreenSz w:val="800x600"/>
  <x:trailing>keep me</x:trailing>
</w:webSettings>"#;
    let snapshot = Snapshot::from_xml(source.to_vec()).unwrap();
    assert_eq!(
        snapshot.settings().target_screen_size(),
        Some(Screen::Pixels800x600)
    );

    let mut edit = snapshot.edit();
    edit.set_target_screen_size(Some(Screen::Pixels1920x1200));
    edit.set_frameset_layout(Some(Layout::Columns)).unwrap();
    let border = Border::new("double").unwrap();
    edit.set_div_border(id(7), BorderSide::Top, Some(border))
        .unwrap();
    let commit = edit.commit().unwrap();
    let updated = commit.snapshot();
    assert_ne!(updated.xml_bytes(), source);
    for needle in [
        "x:future",
        "x:insideFrameset",
        "x:opaqueBorder",
        "x:trailing",
        "targetScreenSz w:val=\"1920x1200\"",
        "frameLayout w:val=\"cols\"",
        "top w:val=\"double\"",
    ] {
        assert!(
            contains_xml(updated.xml_bytes(), needle),
            "missing {needle}"
        );
    }

    let restored = commit.patch().inverse().apply(updated).unwrap();
    assert_eq!(restored.xml_bytes(), source);
    assert_eq!(restored.settings(), snapshot.settings());

    let mut no_op = snapshot.edit();
    no_op.set_target_screen_size(Some(Screen::Pixels800x600));
    let no_op = no_op.commit().unwrap();
    assert_eq!(no_op.snapshot().xml_bytes(), source);
    assert_eq!(no_op.patch().apply(&snapshot).unwrap().xml_bytes(), source);

    let mut cancelled = snapshot.edit();
    cancelled
        .set_frameset_layout(Some(Layout::Columns))
        .unwrap();
    cancelled.set_frameset_layout(Some(Layout::Rows)).unwrap();
    let cancelled = cancelled.commit().unwrap();
    assert_eq!(cancelled.snapshot().xml_bytes(), source);

    let stale_source = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:targetScreenSz w:val="1024x768"/></w:webSettings>"#;
    let stale = Snapshot::from_xml(stale_source.to_vec()).unwrap();
    assert!(commit.patch().apply(&stale).is_err());
}

#[test]
fn package_web_patch_is_graph_checked_and_failure_atomic() {
    let source = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:targetScreenSz w:val="800x600"/></w:webSettings>"#;
    let mut package = package(Conformance::Transitional);
    let target = add_raw_web(&mut package, source, Conformance::Transitional);
    let snapshot = load_snapshot(&package).unwrap().unwrap();
    let mut edit = snapshot.edit();
    edit.set_target_screen_size(Some(Screen::Pixels1024x768));
    let commit = edit.commit().unwrap();
    let before = package.get_part(&target).unwrap().blob().to_vec();
    assert!(apply_patch(&mut package, commit.patch()).is_ok());
    assert_ne!(package.get_part(&target).unwrap().blob(), before);
    assert!(contains_xml(
        package.get_part(&target).unwrap().blob(),
        "targetScreenSz w:val=\"1024x768\""
    ));

    let mut stale_edit = snapshot.edit();
    stale_edit.set_target_screen_size(Some(Screen::Pixels1152x900));
    let stale_patch = stale_edit.commit().unwrap().into_patch();
    let unchanged = package.get_part(&target).unwrap().blob().to_vec();
    assert!(apply_patch(&mut package, &stale_patch).is_err());
    assert_eq!(package.get_part(&target).unwrap().blob(), unchanged);

    package
        .get_part_mut(&target)
        .unwrap()
        .set_blob(b"<broken/>".to_vec());
    let unchanged_broken = package.get_part(&target).unwrap().blob().to_vec();
    let inverse = stale_patch.inverse();
    assert!(apply_patch(&mut package, &inverse).is_err());
    assert_eq!(package.get_part(&target).unwrap().blob(), unchanged_broken);
}

#[test]
fn parser_budget_covers_nested_readers() {
    let mut xml = String::from(
        r#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset>"#,
    );
    xml.try_reserve(MAX_XML_EVENTS * 8).unwrap();
    for _ in 0..=MAX_XML_EVENTS {
        xml.push_str("<!--x-->");
    }
    xml.push_str("</w:frameset></w:webSettings>");

    let result = std::panic::catch_unwind(|| Settings::parse_xml(xml.as_bytes()));
    assert!(result.is_ok());
    assert!(result.unwrap().is_err());
}

#[test]
fn mce_preprocessing_respects_web_settings_xml_bound() {
    let prefix = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">"#;
    let suffix = b"</w:webSettings>";
    let mut xml = Vec::with_capacity(MAX_XML_BYTES + suffix.len() + 1);
    xml.extend_from_slice(prefix);
    xml.resize(MAX_XML_BYTES + 1, b' ');
    xml.extend_from_slice(suffix);

    assert!(matches!(
        parse(&xml),
        Err(Error::Mce(litchi_ooxml_common::mce::Error::LimitExceeded(
            _
        )))
    ));
}

#[test]
fn rejects_mismatched_leaf_end_tags() {
    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    for xml in [
        format!(r#"<w:webSettings xmlns:w="{W}"><w:allowPNG></w:encoding></w:webSettings>"#),
        format!(
            r#"<w:webSettings xmlns:w="{W}"><w:frameset><w:sz w:val="1*"></w:name></w:frameset></w:webSettings>"#
        ),
        format!(
            r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1"><w:marLeft w:val="0"></w:marRight><w:marRight w:val="0"/><w:marTop w:val="0"/><w:marBottom w:val="0"/></w:div></w:divs></w:webSettings>"#
        ),
    ] {
        assert!(parse_settings(xml.as_bytes()).is_err(), "accepted {xml}");
    }
}

#[test]
fn parses_all_scalar_web_settings_with_strict_namespaces() {
    let xml = br#"<s:webSettings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:false="urn:not-wordprocessingml">
            <s:encoding s:val="utf-8"/>
            <s:optimizeForBrowser s:val="on"/>
            <s:allowPNG/>
            <s:doNotRelyOnCSS s:val="0"/>
            <s:doNotSaveAsSingleFile s:val="1"/>
            <s:doNotOrganizeInFolder s:val="false"/>
            <s:doNotUseLongFileNames s:val="true"/>
            <s:pixelsPerInch s:val=" 1023 "/>
            <s:targetScreenSz s:val="1920x1200"/>
            <s:saveSmartTagsAsXml s:val="on"/>
            <false:saveSmartTagsAsXml false:val="off"/>
        </s:webSettings>"#;

    let settings = parse_settings(xml).unwrap();
    assert_eq!(settings.encoding(), Some("utf-8"));
    assert_eq!(settings.optimize_for_browser(), Some(true));
    assert_eq!(settings.rely_on_vml(), None);
    assert_eq!(settings.allow_png(), Some(true));
    assert_eq!(settings.do_not_rely_on_css(), Some(false));
    assert_eq!(settings.do_not_save_as_single_file(), Some(true));
    assert_eq!(settings.do_not_organize_in_folder(), Some(false));
    assert_eq!(settings.do_not_use_long_file_names(), Some(true));
    assert_eq!(settings.pixels_per_inch(), Some(1023));
    assert_eq!(settings.target_screen_size(), Some(Screen::Pixels1920x1200));
    assert_eq!(settings.save_smart_tags_as_xml(), Some(true));
}

#[test]
fn rejects_invalid_or_duplicate_scalar_web_settings() {
    let missing_value = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pixelsPerInch/></w:webSettings>"#;
    assert!(parse_settings(missing_value).is_err());

    let invalid_on_off = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:saveSmartTagsAsXml w:val="maybe"/></w:webSettings>"#;
    assert!(parse_settings(invalid_on_off).is_err());

    let duplicate = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:allowPNG/><w:allowPNG/></w:webSettings>"#;
    assert!(parse_settings(duplicate).is_err());

    let invalid_screen = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:targetScreenSz w:val="1366x768"/></w:webSettings>"#;
    assert!(parse_settings(invalid_screen).is_err());

    let excessive_pixels = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pixelsPerInch w:val="1024"/></w:webSettings>"#;
    assert!(parse_settings(excessive_pixels).is_err());

    let strict_rely = br#"<w:webSettings xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:relyOnVML/></w:webSettings>"#;
    assert!(parse_settings(strict_rely).is_err());

    let out_of_order = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:allowPNG/><w:encoding w:val="utf-8"/></w:webSettings>"#;
    assert!(parse_settings(out_of_order).is_err());

    let nested_scalar = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:allowPNG><w:doNotRelyOnCSS/></w:allowPNG></w:webSettings>"#;
    assert!(parse_settings(nested_scalar).is_err());

    let scalar_text = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:encoding w:val="utf-8">unexpected</w:encoding></w:webSettings>"#;
    assert!(parse_settings(scalar_text).is_err());
}

#[test]
fn parses_recursive_framesets_and_all_frame_properties() {
    let xml = br#"<s:webSettings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"
            xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships"
            xmlns:false="urn:not-wordprocessingml">
          <s:frameset>
            <s:sz s:val="2*"/>
            <s:framesetSplitbar>
              <s:w s:val="90"/>
              <s:color s:val="auto" s:themeColor="accent2" s:themeTint="7f" s:themeShade="00"/>
              <s:noBorder s:val="off"/>
              <s:flatBorders/>
            </s:framesetSplitbar>
            <s:frameLayout s:val="cols"/>
            <s:frame>
              <s:sz s:val="50%"/>
              <s:name s:val="navigation"/>
              <s:sourceFileName rel:id="rId7"/>
              <s:marW s:val="18446744073709551615"/>
              <s:marH s:val="24"/>
              <s:scrollbar s:val="auto"/>
              <s:noResizeAllowed/>
              <s:linkedToFile s:val="false"/>
              <s:futureExtension><s:nested/></s:futureExtension>
            </s:frame>
            <s:frameset>
              <s:frameLayout s:val="none"/>
              <s:frame><s:name s:val="content"/></s:frame>
            </s:frameset>
            <false:frame><false:name false:val="ignored"/></false:frame>
          </s:frameset>
        </s:webSettings>"#;

    let settings = parse_settings(xml).unwrap();
    let frameset = settings.frameset().unwrap();
    assert_eq!(frameset.size(), Some("2*"));
    assert_eq!(frameset.layout(), Some(Layout::Columns));
    let split_bar = frameset.split_bar().unwrap();
    assert_eq!(split_bar.width_twips(), Some(90));
    assert_eq!(split_bar.no_border(), Some(false));
    assert_eq!(split_bar.flat_borders(), Some(true));
    let color = split_bar.color().unwrap();
    assert_eq!(color.value(), "auto");
    assert_eq!(color.theme_color(), Some(Theme::Accent2));
    assert_eq!(color.theme_tint(), Some(0x7f));
    assert_eq!(color.theme_shade(), Some(0));
    assert_eq!(frameset.children().len(), 2);

    let Child::Frame(frame) = &frameset.children()[0] else {
        panic!("first frameset child must be a frame");
    };
    assert_eq!(frame.size(), Some("50%"));
    assert_eq!(frame.name(), Some("navigation"));
    assert_eq!(frame.rel(), Some("rId7"));
    assert_eq!(frame.margin_width(), Some(u64::MAX));
    assert_eq!(frame.margin_height(), Some(24));
    assert_eq!(frame.scrollbar(), Some(Scrollbar::Auto));
    assert_eq!(frame.no_resize_allowed(), Some(true));
    assert_eq!(frame.linked_to_file(), Some(false));

    let Child::Frameset(nested) = &frameset.children()[1] else {
        panic!("second frameset child must be nested");
    };
    assert_eq!(nested.layout(), Some(Layout::None));
    let Child::Frame(frame) = &nested.children()[0] else {
        panic!("nested child must be a frame");
    };
    assert_eq!(frame.name(), Some("content"));
}

#[test]
fn validates_frame_values_and_source_relationships() {
    let invalid_layout = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frameLayout w:val="diagonal"/></w:frameset></w:webSettings>"#;
    assert!(parse_settings(invalid_layout).is_err());

    let overflowing_pixels = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frame><w:marW w:val="18446744073709551616"/></w:frame></w:frameset></w:webSettings>"#;
    assert!(parse_settings(overflowing_pixels).is_err());

    let child_in_leaf = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frame><w:name w:val="bad"><w:frame/></w:name></w:frame></w:frameset></w:webSettings>"#;
    assert!(parse_settings(child_in_leaf).is_err());

    let duplicate = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset><w:frame><w:name w:val="one"/><w:name w:val="two"/></w:frame></w:frameset></w:webSettings>"#;
    assert!(parse_settings(duplicate).is_err());

    let xml = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:frameset><w:frame><w:sourceFileName r:id="rId1"/></w:frame></w:frameset></w:webSettings>"#;
    let mut part = BlobPart::new(
        PackURI::new("/word/webSettings.xml").unwrap(),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml".into(),
        xml.to_vec(),
    );
    assert!(read_settings_part(&part).is_err());
    part.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame".into(),
        "https://example.test/frame.html".into(),
        "rId1".into(),
        true,
    );
    assert!(read_settings_part(&part).is_ok());
}

#[test]
fn parses_recursive_html_divisions_and_border_properties() {
    let xml = br#"<s:webSettings xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:false="urn:not-wordprocessingml">
          <s:divs>
            <s:div s:id="1785730240">
              <s:blockQuote/>
              <s:bodyDiv s:val="off"/>
              <s:marLeft s:val=" -1234567890 "/>
              <s:marRight s:val="+42"/>
              <s:marTop s:val="0"/>
              <s:marBottom s:val="700"/>
              <s:divBdr>
                <s:top s:val="single" s:color="A0b1C2" s:themeColor="text2" s:themeTint="10" s:themeShade="ff" s:sz="18446744073709551615" s:space="6" s:shadow="on" s:frame="0"/>
                <s:left s:val="zigZagStitch"/>
              </s:divBdr>
              <s:divsChild>
                <s:div s:id="1785730241"><s:bodyDiv/><s:marLeft s:val="0"/><s:marRight s:val="0"/><s:marTop s:val="0"/><s:marBottom s:val="0"/></s:div>
              </s:divsChild>
              <s:divsChild><s:div s:id="1785730242"><s:marLeft s:val="1"/><s:marRight s:val="2"/><s:marTop s:val="3"/><s:marBottom s:val="4"/></s:div></s:divsChild>
            </s:div>
            <s:div s:id="1785730243"><s:marLeft s:val="0"/><s:marRight s:val="0"/><s:marTop s:val="0"/><s:marBottom s:val="0"/></s:div>
            <false:div false:id="ignored"/>
          </s:divs>
        </s:webSettings>"#;

    let settings = parse_settings(xml).unwrap();
    let divs = settings.divs().unwrap();
    assert_eq!(divs.len(), 2);
    let div = &divs[0];
    assert_eq!(div.id(), id(1785730240));
    assert_eq!(div.is_block_quote(), Some(true));
    assert_eq!(div.is_body_div(), Some(false));
    assert_eq!(div.left(), Twips::new(-1234567890));
    assert_eq!(div.right(), Twips::new(42));
    assert_eq!(div.top(), Twips::new(0));
    assert_eq!(div.bottom(), Twips::new(700));
    assert_eq!(div.children().len(), 2);
    assert_eq!(div.children()[0].id(), id(1785730241));
    assert_eq!(div.children()[0].is_body_div(), Some(true));
    assert_eq!(div.children()[1].id(), id(1785730242));

    let borders = div.borders().unwrap();
    let top = borders.top().unwrap();
    assert_eq!(top.style(), "single");
    assert_eq!(top.color(), Some("A0b1C2"));
    assert_eq!(top.theme_color(), Some(Theme::Text2));
    assert_eq!(top.theme_tint(), Some(0x10));
    assert_eq!(top.theme_shade(), Some(0xff));
    assert_eq!(top.size_eighth_points(), Some(u64::MAX));
    assert_eq!(top.space_points(), Some(6));
    assert_eq!(top.shadow(), Some(true));
    assert_eq!(top.frame(), Some(false));
    assert_eq!(borders.left().unwrap().style(), "zigZagStitch");
    assert!(borders.bottom().is_none());
    assert!(borders.right().is_none());
}

#[test]
fn validates_html_division_structure_and_values() {
    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let margins = r#"<w:marLeft w:val="0"/><w:marRight w:val="0"/><w:marTop w:val="0"/><w:marBottom w:val="0"/>"#;
    let missing_id = format!(
        r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div>{margins}</w:div></w:divs></w:webSettings>"#
    );
    assert!(parse_settings(missing_id.as_bytes()).is_err());

    for invalid_id in ["0", "-0", "not-a-number", "9223372036854775808"] {
        let xml = format!(
            r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="{invalid_id}">{margins}</w:div></w:divs></w:webSettings>"#
        );
        assert!(parse_settings(xml.as_bytes()).is_err());
    }

    let missing_margin = format!(
        r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1"><w:marLeft w:val="0"/><w:marRight w:val="0"/><w:marTop w:val="0"/></w:div></w:divs></w:webSettings>"#
    );
    assert!(parse_settings(missing_margin.as_bytes()).is_err());

    let invalid_margin = format!(
        r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1"><w:marLeft w:val="1.5"/><w:marRight w:val="0"/><w:marTop w:val="0"/><w:marBottom w:val="0"/></w:div></w:divs></w:webSettings>"#
    );
    assert!(parse_settings(invalid_margin.as_bytes()).is_err());

    let invalid_color = format!(
        r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1">{margins}<w:divBdr><w:left w:val="single" w:color="xyz"/></w:divBdr></w:div></w:divs></w:webSettings>"#
    );
    assert!(parse_settings(invalid_color.as_bytes()).is_err());

    let empty_child_container = format!(
        r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1">{margins}<w:divsChild/></w:div></w:divs></w:webSettings>"#
    );
    assert!(parse_settings(empty_child_container.as_bytes()).is_err());

    let empty_divs = format!(r#"<w:webSettings xmlns:w="{W}"><w:divs/></w:webSettings>"#);
    assert!(parse_settings(empty_divs.as_bytes()).is_err());

    let out_of_order = format!(
        r#"<w:webSettings xmlns:w="{W}"><w:divs><w:div w:id="1"><w:marRight w:val="0"/><w:marLeft w:val="0"/><w:marTop w:val="0"/><w:marBottom w:val="0"/></w:div></w:divs></w:webSettings>"#
    );
    assert!(parse_settings(out_of_order.as_bytes()).is_err());
}

#[test]
fn serializes_every_modeled_web_setting_for_round_trip() {
    let xml = br#"<w:webSettings
          xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <w:frameset>
            <w:sz w:val="2* &amp; 1*"/>
            <w:framesetSplitbar>
              <w:w w:val="18446744073709551615"/>
              <w:color w:val="A0b1C2" w:themeColor="accent4" w:themeTint="0a" w:themeShade="FF"/>
              <w:noBorder/>
              <w:flatBorders w:val="false"/>
            </w:framesetSplitbar>
            <w:frameLayout w:val="cols"/>
            <w:frame>
              <w:sz w:val="50%"/>
              <w:name w:val="main &amp; detail"/>
              <w:sourceFileName r:id="rId7"/>
              <w:marW w:val="42"/>
              <w:marH w:val="24"/>
              <w:scrollbar w:val="auto"/>
              <w:noResizeAllowed w:val="off"/>
              <w:linkedToFile/>
            </w:frame>
            <w:frameset><w:frameLayout w:val="none"/></w:frameset>
          </w:frameset>
          <w:divs>
            <w:div w:id="1">
              <w:blockQuote/>
              <w:bodyDiv w:val="0"/>
              <w:marLeft w:val="-1234567890"/>
              <w:marRight w:val="+42"/>
              <w:marTop w:val="0"/>
              <w:marBottom w:val="700"/>
              <w:divBdr>
                <w:top w:val="single" w:color="auto" w:themeColor="text2" w:themeTint="10" w:themeShade="ff" w:sz="18446744073709551615" w:space="6" w:shadow="on" w:frame="0"/>
                <w:left w:val="zigZagStitch"/>
              </w:divBdr>
              <w:divsChild><w:div w:id="2"><w:bodyDiv/><w:marLeft w:val="0"/><w:marRight w:val="0"/><w:marTop w:val="0"/><w:marBottom w:val="0"/></w:div></w:divsChild>
            </w:div>
          </w:divs>
          <w:encoding w:val="utf-8"/>
          <w:optimizeForBrowser/>
          <w:relyOnVML w:val="false"/>
          <w:allowPNG/>
          <w:doNotRelyOnCSS w:val="off"/>
          <w:doNotSaveAsSingleFile/>
          <w:doNotOrganizeInFolder w:val="0"/>
          <w:doNotUseLongFileNames/>
          <w:pixelsPerInch w:val="1023"/>
          <w:targetScreenSz w:val="1920x1200"/>
          <w:saveSmartTagsAsXml w:val="false"/>
        </w:webSettings>"#;

    let settings = parse_settings(xml).unwrap();
    let serialized = settings.xml(Conformance::Transitional).unwrap();
    let reparsed = parse_settings(&serialized).unwrap();

    assert_eq!(reparsed, settings);
    assert!(contains_xml(&serialized, "main &amp; detail"));
    assert!(contains_xml(&serialized, "w:themeTint=\"0A\""));
    assert!(contains_xml(&serialized, "w:themeShade=\"FF\""));
    assert!(contains_xml(&serialized, "<w:blockQuote w:val=\"1\"/>"));
    assert!(contains_xml(&serialized, "<w:bodyDiv w:val=\"0\"/>"));
    assert!(contains_xml(&serialized, "<w:bodyDiv w:val=\"1\"/>"));

    let mut strict_settings = settings.clone();
    strict_settings.clear_rely_on_vml();
    let strict = strict_settings.xml(Conformance::Strict).unwrap();
    assert!(contains_xml(&strict, "<w:blockQuote w:val=\"1\"/>"));
    assert!(contains_xml(&strict, "<w:bodyDiv w:val=\"0\"/>"));
    assert!(contains_xml(&strict, "<w:bodyDiv w:val=\"1\"/>"));
}

#[test]
fn edits_and_clears_every_scalar_web_setting() {
    let mut settings = Settings::default();
    settings
        .set_encoding("utf-8")
        .unwrap()
        .set_optimize_for_browser(true)
        .set_rely_on_vml(false)
        .set_allow_png(true)
        .set_do_not_rely_on_css(false)
        .set_do_not_save_as_single_file(true)
        .set_do_not_organize_in_folder(false)
        .set_do_not_use_long_file_names(true);
    settings
        .set_pixels_per_inch(96)
        .unwrap()
        .set_target_screen_size(Screen::Pixels1800x1440)
        .set_save_smart_tags_as_xml(false);

    let serialized = settings.xml(Conformance::Transitional).unwrap();
    let reparsed = parse_settings(&serialized).unwrap();
    assert_eq!(reparsed, settings);
    assert_eq!(reparsed.encoding(), Some("utf-8"));
    assert_eq!(reparsed.pixels_per_inch(), Some(96));
    assert_eq!(reparsed.target_screen_size(), Some(Screen::Pixels1800x1440));

    let previous_pixels = settings.pixels_per_inch().unwrap();
    assert!(settings.set_pixels_per_inch(1024).is_err());
    assert_eq!(settings.pixels_per_inch(), Some(previous_pixels));

    assert!(settings.xml(Conformance::Strict).is_err());

    settings
        .clear_encoding()
        .clear_optimize_for_browser()
        .clear_rely_on_vml()
        .clear_allow_png()
        .clear_do_not_rely_on_css()
        .clear_do_not_save_as_single_file()
        .clear_do_not_organize_in_folder()
        .clear_do_not_use_long_file_names()
        .clear_pixels_per_inch()
        .clear_target_screen_size()
        .clear_save_smart_tags_as_xml();
    assert_eq!(settings, Settings::default());
    assert_eq!(
        parse_settings(&settings.xml(Conformance::Transitional).unwrap()).unwrap(),
        Settings::default()
    );
}

#[test]
fn builds_and_edits_recursive_framesets_for_round_trip() {
    let mut color = Color::new("A0b1C2").unwrap();
    color
        .set_theme_color(Theme::Accent4)
        .set_theme_tint(0x0a)
        .set_theme_shade(0xff);

    let mut split_bar = SplitBar::default();
    split_bar
        .set_width_twips(u64::MAX)
        .set_color(color)
        .set_no_border(true)
        .set_flat_borders(false);

    let mut frameset = Frameset::default();
    frameset
        .set_size("2* & 1*")
        .unwrap()
        .set_split_bar(split_bar)
        .set_layout(Layout::Columns);
    let frame = frameset.add_frame().unwrap();
    frame.set_size("50%").unwrap();
    frame.set_name("main & detail").unwrap();
    frame
        .set_rel("rId7")
        .unwrap()
        .set_margin_width(42)
        .set_margin_height(24)
        .set_scrollbar(Scrollbar::Auto)
        .set_no_resize_allowed(false)
        .set_linked_to_file(true);
    let nested = frameset.add_frameset().unwrap();
    nested.set_size("1*").unwrap().set_layout(Layout::None);
    nested.add_frame().unwrap().set_name("nested").unwrap();

    let mut settings = Settings::default();
    settings.set_frameset(frameset);
    let serialized = settings.xml(Conformance::Transitional).unwrap();
    let reparsed = parse_settings(&serialized).unwrap();
    assert_eq!(reparsed, settings);
    assert!(contains_xml(&serialized, "main &amp; detail"));

    let frameset = settings.frameset_mut().unwrap();
    assert_eq!(frameset.children().len(), 2);
    assert!(matches!(frameset.children()[0], Child::Frame(_)));
    assert!(matches!(frameset.children()[1], Child::Frameset(_)));
    frameset
        .clear_size()
        .clear_split_bar()
        .clear_layout()
        .clear_children();
    assert_eq!(frameset, &Frameset::default());
    settings.clear_frameset();
    assert!(settings.frameset().is_none());
}

#[test]
fn validates_mutable_frameset_colors_without_losing_prior_value() {
    assert!(Color::new("12345").is_err());
    let mut color = Color::new("auto").unwrap();
    assert!(color.set_value("GG0000").is_err());
    assert_eq!(color.value(), "auto");
    color.set_value("00ffAA").unwrap();
    assert_eq!(color.value(), "00ffAA");
    color
        .set_theme_color(Theme::Text1)
        .set_theme_tint(1)
        .set_theme_shade(2)
        .clear_theme_color()
        .clear_theme_tint()
        .clear_theme_shade();
    assert_eq!(color.theme_color(), None);
    assert_eq!(color.theme_tint(), None);
    assert_eq!(color.theme_shade(), None);
}

#[test]
fn builds_and_edits_recursive_html_divisions_for_round_trip() {
    let mut top = Border::new("single").unwrap();
    top.set_color("A0b1C2")
        .unwrap()
        .set_theme_color(Theme::Text2)
        .set_theme_tint(0x10)
        .set_theme_shade(0xff)
        .set_size_eighth_points(u64::MAX)
        .set_space_points(6)
        .set_shadow(true)
        .set_frame(false);
    let mut borders = Borders::default();
    borders
        .set_top(top)
        .set_left(Border::new("zigZagStitch").unwrap())
        .set_bottom(Border::new("double").unwrap())
        .set_right(Border::new("nil").unwrap());

    let mut div = make_div(1);
    div.set_block_quote(true)
        .set_body_div(false)
        .set_left(-1_234_567_890)
        .set_right(42)
        .set_top(0)
        .set_bottom(700)
        .set_borders(borders);
    let mut grandchild = make_div(3);
    grandchild.set_block_quote(false);
    let mut child = make_div(2);
    child.set_body_div(true);
    child.add_child(grandchild).unwrap();
    div.add_child(child).unwrap();

    let mut settings = Settings::default();
    settings.set_divs(vec![div]).unwrap();
    let serialized = settings.xml(Conformance::Transitional).unwrap();
    let reparsed = parse_settings(&serialized).unwrap();
    assert_eq!(reparsed, settings);
    assert!(contains_xml(&serialized, "w:id=\"1\""));
    assert_eq!(
        reparsed.divs().unwrap()[0].left(),
        Twips::new(-1_234_567_890)
    );
    let left = serialized
        .windows(b"<w:marLeft".len())
        .position(|window| window == b"<w:marLeft")
        .unwrap();
    let right = serialized
        .windows(b"<w:marRight".len())
        .position(|window| window == b"<w:marRight")
        .unwrap();
    let top = serialized
        .windows(b"<w:marTop".len())
        .position(|window| window == b"<w:marTop")
        .unwrap();
    let bottom = serialized
        .windows(b"<w:marBottom".len())
        .position(|window| window == b"<w:marBottom")
        .unwrap();
    assert!(left < right && right < top && top < bottom);

    settings.add(make_div(4)).unwrap();
    assert_eq!(settings.divs().unwrap().len(), 2);
    assert_eq!(settings.get(id(4)).unwrap().unwrap().id(), id(4));
    settings.move_to(id(4), 0).unwrap();
    let mut first = settings.remove(id(1)).unwrap().unwrap();
    first
        .clear_block_quote()
        .clear_body_div()
        .set_left(0)
        .set_right(0)
        .set_top(0)
        .set_bottom(0)
        .clear_children();
    let borders = first.borders_mut().unwrap();
    borders
        .clear_top()
        .clear_left()
        .clear_bottom()
        .clear_right();
    first.clear_borders();
    settings.add(first).unwrap();
    settings.clear_divs();
    assert!(settings.divs().is_none());
}

#[test]
fn validates_mutable_html_division_values_atomically() {
    assert!(Id::new(0).is_err());
    assert!(Id::parse("word-id").is_err());
    assert!(Twips::parse("1.5").is_err());
    let mut div = make_div(1);
    div.set_left(-42);
    assert_eq!(div.left(), Twips::new(-42));

    let mut border = Border::new("single").unwrap();
    border.set_color("auto").unwrap();
    assert!(border.set_color("xyz").is_err());
    assert_eq!(border.color(), Some("auto"));
    border
        .set_style("double")
        .unwrap()
        .clear_color()
        .set_theme_color(Theme::Accent1)
        .set_theme_tint(1)
        .set_theme_shade(2)
        .set_size_eighth_points(8)
        .set_space_points(1)
        .set_shadow(false)
        .set_frame(true)
        .clear_theme_color()
        .clear_theme_tint()
        .clear_theme_shade()
        .clear_size_eighth_points()
        .clear_space_points()
        .clear_shadow()
        .clear_frame();
    assert_eq!(border.style(), "double");
    assert_eq!(border.color(), None);
    assert_eq!(border.theme_color(), None);
    assert_eq!(border.theme_tint(), None);
    assert_eq!(border.theme_shade(), None);
    assert_eq!(border.size_eighth_points(), None);
    assert_eq!(border.space_points(), None);
    assert_eq!(border.shadow(), None);
    assert_eq!(border.frame(), None);
}

#[test]
fn serialization_rejects_empty_division_containers() {
    let xml = br#"<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:frameset/><w:divs/></w:webSettings>"#;
    assert!(parse_settings(xml).is_err());
    assert!(Settings::default().set_divs(Vec::new()).is_err());
}

#[test]
fn serialization_rejects_excessive_recursive_nesting() {
    let mut frameset = Frameset::default();
    for _ in 0..=MAX_FRAMESET_NESTING {
        frameset = Frameset {
            children: vec![Child::Frameset(frameset)],
            ..Frameset::default()
        };
    }
    let settings = Settings {
        frameset: Some(frameset),
        ..Settings::default()
    };
    assert!(settings.xml(Conformance::Transitional).is_err());

    let mut div = make_div(1);
    for value in 2..=(MAX_FRAMESET_NESTING as i64 + 3) {
        let mut parent = make_div(value);
        parent.children.push(div);
        div = parent;
    }
    let settings = Settings {
        divs: Some(vec![div]),
        ..Settings::default()
    };
    assert!(settings.xml(Conformance::Transitional).is_err());
}
