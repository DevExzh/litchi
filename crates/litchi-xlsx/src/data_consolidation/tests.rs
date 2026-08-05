//! Focused regression tests for the data-consolidation facade and codec.

use super::*;

const T: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const S: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn worksheet(namespace: &str, body: &str) -> Vec<u8> {
    format!(r#"<worksheet xmlns="{namespace}" xmlns:r="{R}"><sheetData/>{body}</worksheet>"#)
        .into_bytes()
}

#[test]
fn parses_every_function_and_effective_defaults() {
    let values = [
        "average",
        "count",
        "countNums",
        "max",
        "min",
        "product",
        "stdDev",
        "stdDevp",
        "sum",
        "var",
        "varp",
    ];
    for value in values {
        let xml = worksheet(T, &format!(r#"<dataConsolidate function="{value}"/>"#));
        assert_eq!(
            parse_worksheet_data_consolidation(&xml)
                .unwrap()
                .unwrap()
                .function()
                .as_str(),
            value
        );
    }
    let value = parse_worksheet_data_consolidation(&worksheet(T, "<dataConsolidate/>"))
        .unwrap()
        .unwrap();
    assert_eq!(value.function(), Function::Sum);
    assert!(!value.left_labels() && !value.start_labels() && !value.top_labels() && !value.link());
    assert!(value.data_references().is_none());
}

#[test]
fn canonical_writer_round_trips_range_name_relationships_and_flags() {
    let references = References::new(vec![
        Reference::range(
            "Sales & West",
            RangeReference::new("$A$1:XFD1048576").unwrap(),
        )
        .unwrap(),
        Reference::named("Workbook_Name")
            .unwrap()
            .with_relationship_id("rId7")
            .unwrap(),
    ])
    .unwrap();
    let value = DataConsolidation::new(Function::CountNumbers, Some(references))
        .with_left_labels(true)
        .with_start_labels(true)
        .with_top_labels(true)
        .with_link(true);
    let fragment = write_worksheet_data_consolidation(&value, Conformance::Transitional).unwrap();
    assert_eq!(
        fragment,
        format!(
            r#"<dataConsolidate xmlns="{T}" xmlns:r="{R}" function="countNums" leftLabels="1" startLabels="1" topLabels="1" link="1"><dataRefs count="2"><dataRef ref="$A$1:XFD1048576" sheet="Sales &amp; West"/><dataRef name="Workbook_Name" r:id="rId7"/></dataRefs></dataConsolidate>"#
        )
    );
    let parsed = parse_worksheet_data_consolidation(&worksheet(T, &fragment))
        .unwrap()
        .unwrap();
    assert_eq!(parsed, value);
    assert_eq!(
        write_worksheet_data_consolidation(&parsed, Conformance::Transitional).unwrap(),
        fragment
    );
}

#[test]
fn supports_strict_mce_preservation_and_exact_schema_position() {
    let strict = format!(
        r#"<worksheet xmlns="{S}" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><sheetData/><dataConsolidate link="true"><dataRefs><dataRef name="N" r:id="rId1"/></dataRefs></dataConsolidate><phoneticPr fontId="0"/></worksheet>"#
    );
    let value = parse_worksheet_data_consolidation(strict.as_bytes())
        .unwrap()
        .unwrap();
    assert!(value.link());
    assert_eq!(value.data_references().unwrap().declared_count(), None);

    let mce = format!(
        r#"<worksheet xmlns="{T}" xmlns:mc="{MC}" xmlns:x="urn:test" mc:Ignorable="x" mc:PreserveAttributes="x:*" mc:PreserveElements="x:keep"><sheetData/><x:wrapper mc:Ignorable="x" mc:ProcessContent="x:wrapper"><dataConsolidate><dataRefs count="1"><dataRef sheet="S" ref="A1:B2"/></dataRefs></dataConsolidate></x:wrapper></worksheet>"#
    );
    assert_eq!(
        parse_worksheet_data_consolidation(mce.as_bytes())
            .unwrap()
            .unwrap()
            .data_references()
            .unwrap()
            .references()
            .len(),
        1
    );

    assert!(
        parse_worksheet_data_consolidation(&worksheet(
            T,
            "<phoneticPr fontId=\"0\"/><dataConsolidate/>",
        ))
        .is_err()
    );
    assert!(
        parse_worksheet_data_consolidation(&worksheet(
            T,
            "<dataConsolidate/><sortState ref=\"A1\"/>",
        ))
        .is_err()
    );
    assert!(
        parse_worksheet_data_consolidation(&worksheet(
            T,
            "<extLst><ext uri=\"u\"><dataConsolidate/></ext></extLst>",
        ))
        .unwrap()
        .is_none()
    );
}

#[test]
fn rejects_malformed_counts_choices_spoofing_duplicates_unknowns_and_bounds() {
    let invalid_bodies = [
        r#"<dataConsolidate function="median"/>"#,
        r#"<dataConsolidate link="maybe"/>"#,
        r#"<dataConsolidate function="sum" function="count"/>"#,
        r#"<dataConsolidate bogus="1"/>"#,
        r#"<dataConsolidate><dataRefs count="2"><dataRef name="N"/></dataRefs></dataConsolidate>"#,
        r#"<dataConsolidate><dataRefs count="65537"/></dataConsolidate>"#,
        r#"<dataConsolidate><dataRefs><dataRef/></dataRefs></dataConsolidate>"#,
        r#"<dataConsolidate><dataRefs><dataRef name="N" sheet="S" ref="A1"/></dataRefs></dataConsolidate>"#,
        r#"<dataConsolidate><dataRefs><dataRef sheet="S"/></dataRefs></dataConsolidate>"#,
        r#"<dataConsolidate><dataRefs><dataRef sheet="S" ref="XFE1"/></dataRefs></dataConsolidate>"#,
        r#"<dataConsolidate><dataRefs><dataRef name="N"><child/></dataRef></dataRefs></dataConsolidate>"#,
        r#"<dataConsolidate><unknown/></dataConsolidate>"#,
        r#"<dataConsolidate><dataRefs/><dataRefs/></dataConsolidate>"#,
        r#"<x:dataConsolidate xmlns:x="urn:fake"/>"#,
        r#"<dataConsolidate xmlns:f="urn:fake" f:link="1"/>"#,
        r#"<dataConsolidate><dataRefs><dataRef xmlns:f="urn:fake" name="N" f:id="rId1"/></dataRefs></dataConsolidate>"#,
    ];
    for body in invalid_bodies {
        assert!(
            parse_worksheet_data_consolidation(&worksheet(T, body)).is_err(),
            "accepted {body}"
        );
    }
}
