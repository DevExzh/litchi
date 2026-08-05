use crate::error::Result;

use super::codec::{CORE, REL, STRICT, STRICT_REL, invalid};
use super::model::{MAX_MEASURE_BYTES, MAX_RELATIONSHIP_ID_BYTES};
use super::*;
use litchi_opc::{OpcPackage, PackURI};

const START: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#;

fn parse(body: &str) -> Result<Option<Setup>> {
    parse_worksheet_page_setup(format!("{START}{body}</worksheet>").as_bytes())
}

fn parse_fixture(path: &str) -> Result<Setup> {
    let package = OpcPackage::open(path).unwrap();
    let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    parse_worksheet_page_setup(package.get_part(&uri).unwrap().blob())?
        .ok_or_else(|| invalid("fixture has no pageSetup"))
}

#[test]
fn closed_tokens_have_one_from_str_and_display_mapping() {
    for (token, value) in [
        ("default", Orientation::Default),
        ("portrait", Orientation::Portrait),
        ("landscape", Orientation::Landscape),
    ] {
        assert_eq!(token.parse(), Ok(value));
        assert_eq!(value.to_string(), token);
    }
    for (token, value) in [
        ("downThenOver", Order::DownThenOver),
        ("overThenDown", Order::OverThenDown),
    ] {
        assert_eq!(token.parse(), Ok(value));
        assert_eq!(value.to_string(), token);
    }
    for (token, value) in [
        ("none", Comments::None),
        ("asDisplayed", Comments::AsDisplayed),
        ("atEnd", Comments::AtEnd),
    ] {
        assert_eq!(token.parse(), Ok(value));
        assert_eq!(value.to_string(), token);
    }
    for (token, value) in [
        ("displayed", ErrorMode::Displayed),
        ("blank", ErrorMode::Blank),
        ("dash", ErrorMode::Dash),
        ("NA", ErrorMode::NotAvailable),
    ] {
        assert_eq!(token.parse(), Ok(value));
        assert_eq!(value.to_string(), token);
    }
    for (token, value) in [
        ("mm", Unit::Millimeter),
        ("cm", Unit::Centimeter),
        ("in", Unit::Inch),
        ("pt", Unit::Point),
        ("pc", Unit::Pica),
        ("pi", Unit::PicaAlternative),
    ] {
        assert_eq!(token.parse(), Ok(value));
        assert_eq!(value.to_string(), token);
    }

    assert!("Landscape".parse::<Orientation>().is_err());
    assert!("na".parse::<ErrorMode>().is_err());
    assert!("px".parse::<Unit>().is_err());
}

#[test]
fn parses_every_attribute_without_erasing_lexical_values() {
    let setup = parse(r#"<pageSetup paperSize="9" paperWidth="21cm" paperHeight="297mm" scale="125" firstPageNumber="3" fitToWidth="2" fitToHeight="4" pageOrder="overThenDown" orientation="landscape" usePrinterDefaults="1" blackAndWhite="true" draft="1" cellComments="atEnd" useFirstPageNumber="true" errors="NA" horizontalDpi="1200" verticalDpi="600" copies="2" r:id="rId7"/>"#).unwrap().unwrap();
    assert_eq!(
        setup,
        Setup {
            paper: Some(Paper::A4),
            paper_width: Some(Measure::new("21", Unit::Centimeter).unwrap()),
            paper_height: Some(Measure::new("297", Unit::Millimeter).unwrap()),
            scale: Some(Scale::new(125).unwrap()),
            first_page: Some(FirstPage::new(3).unwrap()),
            fit_to_width: Some(Fit::new(2).unwrap()),
            fit_to_height: Some(Fit::new(4).unwrap()),
            order: Some(Order::OverThenDown),
            orientation: Some(Orientation::Landscape),
            use_printer_defaults: Some(true),
            black_and_white: Some(true),
            draft: Some(true),
            comments: Some(Comments::AtEnd),
            use_first_page: Some(true),
            errors: Some(ErrorMode::NotAvailable),
            horizontal_dpi: Some(Dpi::new(1_200).unwrap()),
            vertical_dpi: Some(Dpi::new(600).unwrap()),
            copies: Some(Copies::new(2).unwrap()),
        }
    );
    let document = format!(r#"{START}<pageSetup r:id="rId7"/></worksheet>"#);
    assert_eq!(
        parse_worksheet_page_setup_relationship_id(document.as_bytes())
            .unwrap()
            .as_ref()
            .map(RelId::as_str),
        Some("rId7")
    );
}

#[test]
fn preserves_absence_and_resolves_the_printer_default_only_on_request() {
    let absent = parse("<pageSetup/>").unwrap().unwrap();
    assert_eq!(absent, Setup::default());
    assert!(absent.uses_printer_defaults());

    let explicit_false = parse(r#"<pageSetup usePrinterDefaults="0"/>"#)
        .unwrap()
        .unwrap();
    assert_eq!(explicit_false.use_printer_defaults, Some(false));
    assert!(!explicit_false.uses_printer_defaults());
    assert_ne!(absent, explicit_false);
    assert!(parse("").unwrap().is_none());
}

#[test]
fn checked_numeric_types_cover_excel_boundaries() {
    assert_eq!(Scale::new(0), Ok(Scale::AUTO));
    assert_eq!(Scale::new(10).unwrap().get(), 10);
    assert_eq!(Scale::new(400).unwrap().get(), 400);
    assert!(Scale::new(9).is_err());
    assert!(Scale::new(401).is_err());
    assert_eq!(Fit::new(0), Ok(Fit::NONE));
    assert_eq!(Fit::new(Fit::MAX).unwrap().get(), Fit::MAX);
    assert!(Fit::new(Fit::MAX + 1).is_err());

    for value in [1, 47, 50, 118, 256, 2_147_483_647, Paper::MAX] {
        assert_eq!(Paper::new(value).unwrap().get(), value, "paper {value}");
    }
    assert_eq!(Paper::try_from(9), Ok(Paper::A4));
    assert!(Paper::try_from(256).unwrap().is_custom());
    assert!(!Paper::A4.is_custom());

    for value in [1, 9, 401, u32::MAX] {
        assert!(Scale::try_from(value).is_err(), "scale {value}");
    }
    for value in [32_768, u32::MAX] {
        assert!(Fit::try_from(value).is_err(), "fit {value}");
    }
    for value in [0, 48, 49, 119, 255] {
        assert!(Paper::try_from(value).is_err(), "paper {value}");
    }

    assert_eq!(Copies::new(1), Ok(Copies::ONE));
    assert_eq!(Copies::new(Copies::MAX).unwrap().get(), Copies::MAX);
    assert!(Copies::new(0).is_err());
    assert!(Copies::try_from(32_768).is_err());

    assert_eq!(Dpi::new(1).unwrap().get(), 1);
    assert_eq!(Dpi::new(u32::MAX).unwrap().get(), u32::MAX);
    assert!(Dpi::new(0).is_err());
}

#[test]
fn signed_first_page_round_trips_through_the_unsigned_wire_domain() {
    let minimum = FirstPage::new(-32_767).unwrap();
    assert_eq!(minimum.wire(), 4_294_934_529);
    assert_eq!(FirstPage::from_wire(minimum.wire()), Ok(minimum));
    assert_eq!(FirstPage::new(-1).unwrap().wire(), u32::MAX);
    assert_eq!(FirstPage::from_wire(u32::MAX).unwrap().get(), -1);
    assert!(FirstPage::new(0).is_err());
    assert!(FirstPage::new(-32_768).is_err());
    assert!(FirstPage::new(32_768).is_err());
    assert!(FirstPage::from_wire(0).is_err());
    assert!(FirstPage::from_wire(32_768).is_err());
    assert!(FirstPage::from_wire(4_294_934_528).is_err());
    assert!(parse(r#"<pageSetup firstPageNumber="0"/>"#).is_err());

    let parsed = parse(r#"<pageSetup firstPageNumber="4294934529"/>"#)
        .unwrap()
        .unwrap();
    assert_eq!(parsed.first_page, Some(minimum));
}

#[test]
fn exact_measures_reject_non_schema_numbers_and_oversized_lexicals() {
    let exact = Measure::new("00.50", Unit::Centimeter).unwrap();
    assert_eq!(exact.decimal(), "00.50");
    assert_eq!(exact.unit(), Unit::Centimeter);
    assert_eq!(exact.to_string(), "00.50cm");
    assert_eq!(
        Measure::new("0", Unit::Millimeter).unwrap().to_string(),
        "0mm"
    );

    for decimal in ["+1", "1e2", ".5", "1.", "-1", ""] {
        assert!(
            Measure::new(decimal, Unit::Millimeter).is_err(),
            "decimal {decimal}"
        );
    }

    let maximum = "1".repeat(MAX_MEASURE_BYTES - Unit::Millimeter.as_str().len());
    assert!(Measure::new(maximum, Unit::Millimeter).is_ok());
    let oversized = "1".repeat(MAX_MEASURE_BYTES - Unit::Millimeter.as_str().len() + 1);
    assert!(Measure::new(oversized.as_str(), Unit::Millimeter).is_err());
    assert!(Measure::new(oversized.clone(), Unit::Millimeter).is_err());
    assert!(parse(&format!(r#"<pageSetup paperWidth="{oversized}mm"/>"#)).is_err());

    let parsed = parse(r#"<pageSetup paperWidth="00.50cm" paperHeight="0mm"/>"#)
        .unwrap()
        .unwrap();
    assert_eq!(parsed.paper_width, Some(exact));
    assert_eq!(
        parsed.paper_height,
        Some(Measure::new("0", Unit::Millimeter).unwrap())
    );
}

#[test]
fn relationship_ids_validate_before_borrowed_or_owned_storage() {
    assert_eq!(RelId::new("rId7").unwrap().as_str(), "rId7");
    assert_eq!(
        RelId::new(String::from("printerSettings1"))
            .unwrap()
            .as_str(),
        "printerSettings1"
    );

    let oversized = "r".repeat(MAX_RELATIONSHIP_ID_BYTES + 1);
    assert!(RelId::new(oversized.as_str()).is_err());
    assert!(RelId::new(oversized).is_err());
}

#[test]
fn relationship_projection_is_independent_from_nonconforming_settings() {
    let xml = format!(r#"{START}<pageSetup horizontalDpi="0" r:id="rIdPrinter"/></worksheet>"#);
    assert!(parse_worksheet_page_setup(xml.as_bytes()).is_err());
    assert_eq!(
        parse_worksheet_page_setup_relationship_id(xml.as_bytes())
            .unwrap()
            .unwrap()
            .as_str(),
        "rIdPrinter"
    );
}

#[test]
fn parser_rejects_every_out_of_domain_numeric_family() {
    let automatic = parse(r#"<pageSetup scale="0" fitToWidth="0" fitToHeight="32767"/>"#)
        .unwrap()
        .unwrap();
    assert!(automatic.scale.unwrap().is_auto());
    assert!(automatic.fit_to_width.unwrap().is_unbounded());
    assert_eq!(automatic.fit_to_height.unwrap().get(), 32_767);

    assert!(parse(r#"<pageSetup scale="9"/>"#).is_err());
    assert!(parse(r#"<pageSetup scale="401"/>"#).is_err());
    assert!(parse(r#"<pageSetup fitToWidth="32768"/>"#).is_err());
    assert!(parse(r#"<pageSetup paperSize="0"/>"#).is_err());
    assert!(parse(r#"<pageSetup paperSize="48"/>"#).is_err());
    assert!(parse(r#"<pageSetup paperSize="119"/>"#).is_err());
    assert!(parse(r#"<pageSetup horizontalDpi="0"/>"#).is_err());
    assert!(parse(r#"<pageSetup verticalDpi="0"/>"#).is_err());
    assert!(parse(r#"<pageSetup copies="0"/>"#).is_err());
    assert!(parse(r#"<pageSetup copies="32768"/>"#).is_err());
}

#[test]
fn rejects_bad_enums_measures_and_content() {
    assert!(parse(r#"<pageSetup orientation="sideways"/>"#).is_err());
    assert!(parse(r#"<pageSetup paperWidth="+1mm"/>"#).is_err());
    assert!(parse(r#"<pageSetup paperWidth="1e2mm"/>"#).is_err());
    assert!(parse(r#"<pageSetup paperWidth=".5mm"/>"#).is_err());
    assert!(parse(r#"<pageSetup paperWidth="1.mm"/>"#).is_err());
    assert!(parse(r#"<pageSetup errors="na"/>"#).is_err());
    assert!(parse(r#"<pageSetup r:id="1bad"/>"#).is_err());
    assert!(parse(r#"<pageSetup><x/></pageSetup>"#).is_err());
}

#[test]
fn page_setup_attributes_use_expanded_names_strictly() {
    let local_default = parse(&format!(
        r#"<pageSetup xmlns="{}" xmlns:unused="urn:unused" orientation="landscape"/>"#,
        String::from_utf8_lossy(CORE)
    ))
    .unwrap()
    .unwrap();
    assert_eq!(local_default.orientation, Some(Orientation::Landscape));

    let aliased_relationship = format!(
        r#"<worksheet xmlns="{}"><pageSetup xmlns:rels="{}" rels:id="rIdAlias"/></worksheet>"#,
        String::from_utf8_lossy(CORE),
        String::from_utf8_lossy(REL)
    );
    assert_eq!(
        parse_worksheet_page_setup_relationship_id(aliased_relationship.as_bytes())
            .unwrap()
            .unwrap()
            .as_str(),
        "rIdAlias"
    );
    assert!(
            parse(r#"<pageSetup xmlns:rels="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" rels:id="rId2"/>"#)
                .is_err()
        );

    let undeclared_relationship = format!(
        r#"<worksheet xmlns="{}"><pageSetup r:id="rId1"/></worksheet>"#,
        String::from_utf8_lossy(CORE)
    );
    assert!(parse_worksheet_page_setup(undeclared_relationship.as_bytes()).is_err());
    assert!(parse(r#"<pageSetup xmlns:v="urn:vendor" v:mode="x"/>"#).is_err());
    assert!(parse(r#"<pageSetup id="rId1"/>"#).is_err());
    assert!(parse(r#"<pageSetup r:target="rId1"/>"#).is_err());
    assert!(
        parse(&format!(
            r#"<pageSetup xmlns:s="{}" s:orientation="landscape"/>"#,
            String::from_utf8_lossy(CORE)
        ))
        .is_err()
    );

    let mismatched_relationship = format!(
        r#"<worksheet xmlns="{}" xmlns:r="{}"><pageSetup r:id="rId1"/></worksheet>"#,
        String::from_utf8_lossy(CORE),
        String::from_utf8_lossy(STRICT_REL)
    );
    assert!(
        parse_worksheet_page_setup_relationship_id(mismatched_relationship.as_bytes()).is_err()
    );

    let mismatched_element = format!(
        r#"<worksheet xmlns="{}" xmlns:s="{}" xmlns:r="{}"><s:pageSetup r:id="rId1"/></worksheet>"#,
        String::from_utf8_lossy(CORE),
        String::from_utf8_lossy(STRICT),
        String::from_utf8_lossy(STRICT_REL)
    );
    assert!(parse_worksheet_page_setup(mismatched_element.as_bytes()).is_err());

    let strict = format!(
        r#"<s:worksheet xmlns:s="{}" xmlns:r="{}"><s:pageSetup r:id="rId9"/></s:worksheet>"#,
        String::from_utf8_lossy(STRICT),
        String::from_utf8_lossy(STRICT_REL)
    );
    assert_eq!(
        parse_worksheet_page_setup_relationship_id(strict.as_bytes())
            .unwrap()
            .unwrap()
            .as_str(),
        "rId9"
    );

    let strict_with_transitional_relationship = format!(
        r#"<s:worksheet xmlns:s="{}" xmlns:r="{}"><s:pageSetup r:id="rId9"/></s:worksheet>"#,
        String::from_utf8_lossy(STRICT),
        String::from_utf8_lossy(REL)
    );
    assert!(
        parse_worksheet_page_setup_relationship_id(
            strict_with_transitional_relationship.as_bytes()
        )
        .is_err()
    );
}

#[test]
fn loads_poi_resolution_and_relationship_fixture() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/45540_classic_Header.xlsx"
    );
    let setup = parse_fixture(path).unwrap();
    assert_eq!(setup.orientation, Some(Orientation::Portrait));
    assert_eq!(setup.horizontal_dpi, Some(Dpi::new(1_200).unwrap()));
    assert_eq!(setup.vertical_dpi, Some(Dpi::new(1_200).unwrap()));
    let package = OpcPackage::open(path).unwrap();
    let uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    assert_eq!(
        parse_worksheet_page_setup_relationship_id(package.get_part(&uri).unwrap().blob())
            .unwrap()
            .as_ref()
            .map(RelId::as_str),
        Some("rId1")
    );
}

#[test]
fn rejects_a_libreoffice_fixture_with_nonconforming_zero_dpi() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf136721_letter_sized_paper.xlsx"
    );
    let error = parse_fixture(path).unwrap_err();
    assert!(error.to_string().contains("DPI"));
}
