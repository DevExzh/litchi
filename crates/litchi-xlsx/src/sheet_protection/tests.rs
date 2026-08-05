use super::codec::{CORE, PROTECTED_RANGES_EXTENSION_URI, STRICT, X14, XM};
use super::*;

const START: &str =
    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#;

fn parse(body: &str) -> Result<Metadata> {
    parse_protection(format!("{START}{body}</worksheet>").as_bytes())
}

#[test]
fn parses_legacy_verifier_and_schema_defaults() {
    let metadata =
        parse(r#"<sheetProtection password="CC3D" sheet="1" objects="true" scenarios="1"/>"#)
            .unwrap();
    let protection = metadata.sheet_protection().unwrap();
    assert_eq!(
        protection.verifier(),
        Some(&ProtectionPasswordVerifier::Legacy(0xCC3D))
    );
    assert!(protection.sheet_locked());
    assert!(protection.format_cells_locked());
    assert!(!protection.select_locked_cells_locked());
}

#[test]
fn parses_core_strong_ranges_and_column_shorthand() {
    let metadata = parse(r#"<sheetProtection sheet="1"/><protectedRanges><protectedRange algorithmName="SHA-512" hashValue="AQI=" saltValue="AwQ=" spinCount="100000" sqref="A6 C:C" name="editable" securityDescriptor="D:test"/></protectedRanges>"#).unwrap();
    let range = metadata.protected_ranges().next().unwrap();
    assert_eq!(range.name(), "editable");
    assert_eq!(
        range.sqref().references()[0].kind(),
        ProtectionRangeReferenceKind::Cells {
            start_row: 6,
            start_column: 1,
            end_row: 6,
            end_column: 1
        }
    );
    assert_eq!(
        range.sqref().references()[1].kind(),
        ProtectionRangeReferenceKind::Columns {
            start_column: 3,
            end_column: 3
        }
    );
    let ProtectionPasswordVerifier::Strong(verifier) = range.verifier().unwrap() else {
        panic!("expected strong verifier")
    };
    assert_eq!(verifier.algorithm_name(), "SHA-512");
    assert_eq!(verifier.hash_value(), &[1, 2]);
    assert_eq!(verifier.spin_count(), 100_000);
}

#[test]
fn parses_x14_extension_range() {
    let xml = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" mc:Ignorable="x14 xm"><extLst><ext uri="{FC87AEE6-9EDD-4A0A-B7FB-166176984837}"><x14:protectedRanges><x14:protectedRange name="Range1" password="1234"><xm:sqref>$B$2:$C$4</xm:sqref></x14:protectedRange></x14:protectedRanges></ext></extLst></worksheet>"#;
    let metadata = parse_protection(xml).unwrap();
    let collection = &metadata.protected_range_collections()[0];
    assert_eq!(collection.source(), ProtectedRangeSource::Office2010);
    assert_eq!(
        collection.ranges()[0].sqref().references()[0].kind(),
        ProtectionRangeReferenceKind::Cells {
            start_row: 2,
            start_column: 2,
            end_row: 4,
            end_column: 3
        }
    );
}

#[test]
fn rejects_incomplete_or_conflicting_verifiers() {
    assert!(parse(r#"<sheetProtection algorithmName="SHA-512" hashValue="AQI="/>"#).is_err());
    assert!(parse(r#"<protectedRanges><protectedRange name="bad" sqref="A1" password="1234" algorithmName="SHA-512" hashValue="AQI=" saltValue="AwQ=" spinCount="1"/></protectedRanges>"#).is_err());
    assert!(
        parse(r#"<protectedRanges><protectedRange name="bad" sqref="XFE1"/></protectedRanges>"#)
            .is_err()
    );
}

#[test]
fn deterministic_writer_round_trips_core_x14_and_strict_metadata() {
    let source = format!(
        r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="{}" xmlns:xm="{}" mc:Ignorable="x14 xm"><sheetProtection algorithmName="SHA-512" hashValue="AQI=" saltValue="AwQ=" spinCount="0" sheet="1" formatCells="0"/><protectedRanges><protectedRange name="A&amp;B" sqref="A1  C:C" password="00af" securityDescriptor="D:&quot;test&quot;"/></protectedRanges><extLst><ext uri="{}"><x14:protectedRanges><x14:protectedRange name="X"><xm:sqref>$B$2:$C$4</xm:sqref></x14:protectedRange></x14:protectedRanges></ext></extLst></worksheet>"#,
        std::str::from_utf8(STRICT).unwrap(),
        std::str::from_utf8(X14).unwrap(),
        std::str::from_utf8(XM).unwrap(),
        PROTECTED_RANGES_EXTENSION_URI
    );
    let metadata = parse_protection(source.as_bytes()).unwrap();
    let fragment = write_protection(&metadata, Conformance::Strict).unwrap();
    assert!(fragment.contains("password=\"00AF\""));
    assert!(fragment.contains("sqref=\"A1 C:C\""));
    assert!(fragment.contains("<x14:protectedRanges"));
    let wrapped = format!(
        r#"<worksheet xmlns="{}">{fragment}</worksheet>"#,
        std::str::from_utf8(STRICT).unwrap()
    );
    let reparsed = parse_protection(wrapped.as_bytes()).unwrap();
    assert_eq!(
        write_protection(&reparsed, Conformance::Strict).unwrap(),
        fragment
    );
}

#[test]
fn rejects_spoofed_unknown_out_of_order_and_noncanonical_metadata() {
    let invalid = [
        r#"<sheetProtection xmlns:f="urn:fake" f:sheet="1"/>"#,
        r#"<protectedRanges><protectedRange name="R" sqref="A1"/></protectedRanges><sheetProtection/>"#,
        r#"<sheetProtection/><sheetData/>"#,
        r#"<protectedRanges><unknown/></protectedRanges>"#,
        r#"<protectedRanges><protectedRange name="R" sqref="A1" algorithmName="SHA-512" hashValue="" saltValue="AQ==" spinCount="1"/></protectedRanges>"#,
        r#"<protectedRanges><protectedRange name="R" sqref="A1" algorithmName="SHA-512" hashValue="AQI=" saltValue="AQ=" spinCount="1"/></protectedRanges>"#,
        r#"<f:sheetProtection xmlns:f="urn:fake"/>"#,
    ];
    for body in invalid {
        assert!(parse(body).is_err(), "accepted {body}");
    }

    let ignored = format!(
        r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:test" mc:Ignorable="x"><sheetProtection x:future="1"/></worksheet>"#,
        std::str::from_utf8(CORE).unwrap()
    );
    assert!(parse_protection(ignored.as_bytes()).is_ok());
    let preserved = format!(
        r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:test" mc:Ignorable="x" mc:PreserveAttributes="x:*"><sheetProtection x:future="1"/></worksheet>"#,
        std::str::from_utf8(CORE).unwrap()
    );
    assert!(parse_protection(preserved.as_bytes()).is_err());
}
