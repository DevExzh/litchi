//! Regression tests for the conditional-formatting model and codec.

use super::codec::capture_conditional_formatting;
use super::*;

use crate::color::Rgb;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, TargetMode};

use std::mem::size_of;

fn fixture(relative: &str, sheet: u32) -> (Vec<Differential>, Vec<Formatting>) {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let package = litchi_opc::phys_pkg::OwnedPhysPkgReader::open(root.join(relative))
        .unwrap_or_else(|error| panic!("failed to open {relative}: {error}"));
    let styles_uri = PackURI::new("/xl/styles.xml").expect("valid styles URI");
    let styles_bytes = package
        .blob_for(&styles_uri)
        .unwrap_or_else(|error| panic!("failed to read styles in {relative}: {error}"));
    let differential_formats = parse_differential_formats(&styles_bytes)
        .unwrap_or_else(|error| panic!("failed to parse styles in {relative}: {error}"));
    let sheet_uri =
        PackURI::new(format!("/xl/worksheets/sheet{sheet}.xml")).expect("valid sheet URI");
    let sheet_bytes = package
        .blob_for(&sheet_uri)
        .unwrap_or_else(|error| panic!("failed to read sheet1 in {relative}: {error}"));
    let values = parse_conditional_formattings(&sheet_bytes, differential_formats.len())
        .unwrap_or_else(|error| {
            panic!("failed to parse conditional formatting in {relative}: {error}")
        });
    (differential_formats, values)
}

#[test]
fn rewrite_readback_is_reused_and_public_noop_remains_exact() {
    let source = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting><hyperlinks/></worksheet>"#;
    let before = parse_conditional_formattings(source, 0).unwrap();

    package::reset_readback_parse_count();
    let (rewritten, readback) =
        package::replace_conditional_formattings_with_readback(source, &[], 0).unwrap();
    assert!(readback.is_empty());
    assert_eq!(
        parse_conditional_formattings(&rewritten, 0).unwrap(),
        readback
    );
    assert_eq!(package::readback_parse_count(), 1);

    package::reset_readback_parse_count();
    assert_eq!(
        replace_conditional_formattings(source, &before, 0).unwrap(),
        source
    );
    assert_eq!(package::readback_parse_count(), 0);

    package::reset_readback_parse_count();
    assert_eq!(
        replace_conditional_formattings(source, &[], 0).unwrap(),
        rewritten
    );
    assert_eq!(package::readback_parse_count(), 1);
}

fn owned_package_with_conditional_formatting() -> OpcPackage {
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/workbook.xml").unwrap(),
            ct::SML_SHEET_MAIN.to_owned(),
            br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rIdSheet"/></sheets></workbook>"#.to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/worksheets/sheet1.xml").unwrap(),
            ct::SML_WORKSHEET.to_owned(),
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting></worksheet>"#.to_vec(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/styles.xml").unwrap(),
            ct::SML_STYLES.to_owned(),
            br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dxfs count="0"/></styleSheet>"#.to_vec(),
        )))
        .unwrap();
    let workbook = package
        .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
        .unwrap();
    workbook
        .rels_mut()
        .try_add_relationship(
            rt::WORKSHEET.to_owned(),
            "worksheets/sheet1.xml".to_owned(),
            "rIdSheet".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    workbook
        .rels_mut()
        .try_add_relationship(
            rt::STYLES.to_owned(),
            "styles.xml".to_owned(),
            "rIdStyles".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package
}

#[test]
fn source_edit_commit_reuses_one_rewrite_readback_and_noop_reuses_none() {
    let package = owned_package_with_conditional_formatting();
    let before = Snapshot::load(&package, "Sheet1").unwrap();
    let mut edit = SourceEdit::from_snapshot_for_test(before);
    assert!(edit.clear());

    package::reset_readback_parse_count();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert!(commit.snapshot().collections().is_empty());
    assert_eq!(package::readback_parse_count(), 1);

    let package = owned_package_with_conditional_formatting();
    let before = Snapshot::load(&package, "Sheet1").unwrap();
    let edit = SourceEdit::from_snapshot_for_test(before);
    package::reset_readback_parse_count();
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(package::readback_parse_count(), 0);
}

#[test]
fn parses_core_payloads_and_differential_reference() {
    let xml=br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><conditionalFormatting sqref="A1:A3 C1"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color theme="2" tint="0.5"/></colorScale></cfRule><cfRule type="cellIs" priority="2" dxfId="0" operator="between" stopIfTrue="1"><formula>1</formula><formula><![CDATA[2]]></formula></cfRule></conditionalFormatting></worksheet>"#;
    let values = parse_conditional_formattings(xml, 1).unwrap();
    assert_eq!(values[0].ranges.len(), 2);
    assert_eq!(values[0].rules.len(), 2);
    assert_eq!(values[0].rules[1].formulas.as_slice(), ["1", "2"]);
    assert!(values[0].rules[1].stop_if_true);
    let encoded=br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>$B$3&gt;5</formula></cfRule></conditionalFormatting></worksheet>"#;
    let captured = capture_conditional_formatting(encoded).unwrap();
    assert!(
        String::from_utf8_lossy(&captured[0].bytes).contains("&gt;"),
        "first capture was {:?}",
        String::from_utf8_lossy(&captured[0].bytes)
    );
    let values = parse_conditional_formattings(encoded, 0).unwrap();
    assert_eq!(values[0].rules[0].formulas.as_slice(), ["$B$3>5"]);
    let extension=br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="futureThing"><x14:futurePayload/></x14:cfRule><xm:sqref>A1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></worksheet>"#;
    assert!(parse_conditional_formattings(extension, 0).is_err());
}

#[test]
fn fixed_domains_are_compact_exact_and_strict() {
    assert_eq!(size_of::<Kind>(), 1);
    assert_eq!(size_of::<Operator>(), 1);
    assert_eq!(size_of::<ValueKind>(), 1);
    assert_eq!(size_of::<Period>(), 1);
    assert_eq!(size_of::<Direction>(), 1);
    assert_eq!(size_of::<Axis>(), 1);
    assert_eq!(size_of::<ColorRole>(), 1);
    assert_eq!(size_of::<IconSet>(), 1);
    assert_eq!(size_of::<IconSet14>(), 1);

    assert_eq!("cellIs".parse::<Kind>(), Ok(Kind::CellIs));
    assert_eq!("between".parse::<Operator>(), Ok(Operator::Between));
    assert_eq!("last7Days".parse::<Period>(), Ok(Period::Last7Days));
    assert_eq!(
        "rightToLeft".parse::<Direction>(),
        Ok(Direction::RightToLeft)
    );
    assert_eq!("middle".parse::<Axis>(), Ok(Axis::Middle));
    assert_eq!(
        "negativeFillColor".parse::<ColorRole>(),
        Ok(ColorRole::NegativeFill)
    );
    assert_eq!("5Rating".parse::<IconSet>(), Ok(IconSet::FiveRating));
    assert_eq!("3Stars".parse::<IconSet14>(), Ok(IconSet14::ThreeStars));
    assert!("futureThing".parse::<Kind>().is_err());
    assert!("3Stars".parse::<IconSet>().is_err());
    assert!("FF00Aa10".parse::<Rgb>().is_ok());
    assert!("00AA10".parse::<Rgb>().is_err());
}

#[test]
fn office_2010_icon_sets_are_distinct_from_core_sets() {
    let x14 = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="iconSet"><x14:iconSet iconSet="3Stars"><x14:cfvo type="autoMin"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="autoMax"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>A1:A3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></worksheet>"#;
    let values = parse_conditional_formattings(x14, 0).unwrap();
    let Some(Payload::IconSet14(icons)) = &values[0].rules[0].payload else {
        panic!("expected a typed Office 2010 icon set")
    };
    assert_eq!(icons.set, IconSet14::ThreeStars);
    assert_eq!(icons.thresholds.len(), 3);
    assert_eq!(icons.thresholds[0].kind, ValueKind::AutoMin);
    assert_eq!(icons.thresholds[0].formula.as_deref(), Some("0"));

    let core = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><conditionalFormatting sqref="A1:A3"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Stars"><cfvo type="min"/><cfvo type="percent" val="33"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting></worksheet>"#;
    assert!(parse_conditional_formattings(core, 0).is_err());

    let core_auto = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><conditionalFormatting sqref="A1:A3"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="autoMin"/><cfvo type="percent" val="33"/><cfvo type="max"/></iconSet></cfRule></conditionalFormatting></worksheet>"#;
    assert!(parse_conditional_formattings(core_auto, 0).is_err());
}

#[test]
fn rejects_malformed_and_dangerous_content() {
    let duplicate=br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"/></conditionalFormatting><conditionalFormatting sqref="B1"><cfRule type="expression" priority="1"/></conditionalFormatting></worksheet>"#;
    assert!(parse_conditional_formattings(duplicate, 0).is_err());
    let dangling=br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1" dxfId="2"/></conditionalFormatting></worksheet>"#;
    assert!(parse_conditional_formattings(dangling, 0).is_err());
    assert!(parse_conditional_formattings(br#"<!DOCTYPE x><worksheet/>"#, 0).is_err());
}

#[test]
fn parses_poi_and_libreoffice_conditional_formatting_fixtures() {
    let fixtures = [
        (
            "test-data/ooxml/xlsx/NewStyleConditionalFormattings.xlsx",
            1,
        ),
        ("test-data/ooxml/xlsx/test_conditional_formatting.xlsx", 1),
        (
            "test-data/ooxml/xlsx/conditional_formatting_multiple_ranges.xlsx",
            1,
        ),
        ("test-data/ooxml/xlsx/conditional_fmt_origin.xlsx", 1),
        ("test-data/ooxml/xlsx/conditional_fmt_checkpriority.xlsx", 1),
        (
            "test-data/poi/test-data/spreadsheet/ConditionalFormattingSamples.xlsx",
            2,
        ),
        (
            "test-data/poi/test-data/spreadsheet/61060-conditional-number-formatting.xlsx",
            1,
        ),
    ];
    let mut parsed = Vec::new();
    for (relative, sheet) in fixtures {
        let value = fixture(relative, sheet);
        assert!(
            !value.1.is_empty(),
            "no conditional formatting in {relative}"
        );
        parsed.push(value);
    }

    assert!(
        parsed[0]
            .1
            .iter()
            .flat_map(|value| &value.rules)
            .any(|rule| matches!(
                rule.payload,
                Some(
                    Payload::ColorScale(_)
                        | Payload::DataBar(_)
                        | Payload::IconSet(_)
                        | Payload::IconSet14(_)
                )
            ))
    );
    let formulas: Vec<_> = parsed[1]
        .1
        .iter()
        .flat_map(|value| &value.rules)
        .flat_map(|rule| &rule.formulas)
        .map(String::as_str)
        .collect();
    assert!(
        formulas.iter().any(|formula| formula.contains("$B$3>5")),
        "opaque formulas were {formulas:?}"
    );
    assert!(parsed[2].1.iter().any(|value| value.ranges.len() >= 3));
    assert!(
        parsed[2]
            .1
            .iter()
            .flat_map(|value| &value.rules)
            .flat_map(|rule| &rule.formulas)
            .any(|formula| formula.contains("B2"))
    );
    assert!(
        parsed[3]
            .1
            .iter()
            .flat_map(|value| &value.rules)
            .flat_map(|rule| &rule.formulas)
            .any(|formula| formula == "NOT(ISERROR(SEARCH(\"BAC\",B1)))")
    );
    let mut priorities: Vec<_> = parsed[4]
        .1
        .iter()
        .flat_map(|value| &value.rules)
        .filter_map(|rule| rule.priority)
        .collect();
    priorities.sort_unstable();
    assert!(priorities.windows(2).all(|pair| pair[0] < pair[1]));
    assert!(!parsed[5].0.is_empty());
    assert!(
        parsed[6]
            .0
            .iter()
            .any(|format| format.number_format.is_some())
    );
}
