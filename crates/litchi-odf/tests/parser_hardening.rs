//! Fuzz-style hardening for the ODS and ODT parsers.
//!
//! Truncated, bit-flipped, and hand-crafted malformed documents must always
//! produce a typed `Result` — `Ok` or `Error::InvalidFormat` — and never a
//! panic. The sweeps exercise every truncation point and single-byte mutation
//! of feature-rich seed documents; the targeted cases assert typed errors for
//! specific malformations (caps, nesting rules, spoofed content).

use litchi_odf::{FlatSpreadsheet, FlatTextDocument};

/// Feature-rich flat ODS seed: cells, formulas, annotations, named
/// expressions, conditional formats, and sparkline groups.
const ODS_SEED: &str = r##"<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet" office:version="1.3"><office:body><office:spreadsheet><table:calculation-settings table:case-sensitive="false"/><table:table table:name="Data"><table:table-row><table:table-cell office:value-type="float" office:value="7" table:formula="of:=1+6"><text:p>7</text:p><office:annotation><text:p>note</text:p></office:annotation></table:table-cell><table:table-cell table:number-columns-repeated="3" office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row><table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="string"><text:p>tail</text:p></table:table-cell></table:table-row><calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="Data.A1:Data.D1"><calcext:condition calcext:apply-style-name="Good" calcext:value="cell-content()&gt;5" calcext:base-cell-address="Data.A1"/><calcext:color-scale><calcext:color-scale-entry calcext:value="0" calcext:type="minimum" calcext:color="#ff0000"/><calcext:color-scale-entry calcext:value="1" calcext:type="maximum" calcext:color="#00ff00"/></calcext:color-scale></calcext:conditional-format></calcext:conditional-formats><calcext:sparkline-groups><calcext:sparkline-group calcext:type="line" calcext:markers="true"><calcext:sparklines><calcext:sparkline calcext:cell-address="Data.E1" calcext:data-range="Data.A1:Data.D1"/></calcext:sparklines></calcext:sparkline-group></calcext:sparkline-groups></table:table><table:named-expressions><table:named-range table:name="Total" table:base-cell-address="$Data.$A$1" table:cell-range-address="$Data.$A$1:$Data.$D$1"/></table:named-expressions></office:spreadsheet></office:body></office:document>"##;

/// Feature-rich flat ODT seed: page sequence, forms, tracked changes,
/// sections, ruby, dynamic fields, notes, and a table.
const ODT_SEED: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" office:mimetype="application/vnd.oasis.opendocument.text" office:version="1.3"><office:body><office:text><text:page-sequence><text:page text:master-page-name="Standard"/></text:page-sequence><text:tracked-changes><text:changed-region text:id="c1"><text:insertion><office:change-info><dc:date xmlns:dc="http://purl.org/dc/elements/1.1/">2024-01-01T00:00:00</dc:date></office:change-info><text:p>added</text:p></text:insertion></text:changed-region></text:tracked-changes><office:forms><form:form form:name="f"><form:text form:name="Name" xml:id="name_field" form:current-value="Ada"/></form:form></office:forms><text:h text:outline-level="1">Title</text:h><text:p>body <text:bookmark text:name="mark"/> <text:note text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>fn</text:p></text:note-body></text:note></text:p><text:section text:name="s1" text:protected="true" text:protection-key="aGk="><text:p>secret</text:p></text:section><text:p><text:ruby><text:ruby-base>漢</text:ruby-base><text:rt>kan</text:rt></text:ruby> <text:conditional-text text:condition="of:=1=1" text:string-value-if-true="y" text:string-value-if-false="n">y</text:conditional-text></text:p><table:table table:name="T"><table:table-row><table:table-cell office:value-type="string"><text:p>c</text:p></table:table-cell></table:table-row></table:table></office:text></office:body></office:document>"#;

/// Exercise every eager and lazy ODS parse path, returning the first typed
/// error encountered. A panic anywhere fails the test.
fn exercise_ods(bytes: Vec<u8>) -> Option<litchi_core::Error> {
    let mut flat = match FlatSpreadsheet::from_bytes(bytes) {
        Ok(flat) => flat,
        Err(error) => return Some(error),
    };
    let spreadsheet = flat.spreadsheet_mut();
    for result in [
        spreadsheet.sheets().map(|_| ()),
        spreadsheet.to_csv().map(|_| ()),
    ] {
        if let Err(error) = result {
            return Some(error);
        }
    }
    None
}

/// Exercise the ODT parse paths, returning the first typed error.
fn exercise_odt(bytes: Vec<u8>) -> Option<litchi_core::Error> {
    let mut flat = match FlatTextDocument::from_bytes(bytes) {
        Ok(flat) => flat,
        Err(error) => return Some(error),
    };
    let document = flat.document_mut();
    for result in [
        document.text().map(|_| ()),
        document.paragraphs().map(|_| ()),
        document.tables().map(|_| ()),
        document.sections().map(|_| ()),
        document.forms().map(|_| ()),
        document.page_sequence().map(|_| ()),
        document.tracked_changes().map(|_| ()),
        document.dynamic_text_fields().map(|_| ()),
        document.text_indexes().map(|_| ()),
        document.master_pages().map(|_| ()),
        document.ruby_annotations().map(|_| ()),
    ] {
        if let Err(error) = result {
            return Some(error);
        }
    }
    None
}

fn ods_table(inner: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="S">{inner}</table:table></office:spreadsheet></office:body></office:document>"#
    )
    .into_bytes()
}

fn assert_ods_typed_error(bytes: Vec<u8>, case: &str) {
    let Some(error) = exercise_ods(bytes) else {
        panic!("case '{case}' unexpectedly parsed");
    };
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(_)),
        "case '{case}' produced a non-typed error: {error:?}"
    );
}

fn assert_odt_typed_error(bytes: Vec<u8>, case: &str) {
    let Some(error) = exercise_odt(bytes) else {
        panic!("case '{case}' unexpectedly parsed");
    };
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(_)),
        "case '{case}' produced a non-typed error: {error:?}"
    );
}

#[test]
fn ods_truncation_sweep_never_panics() {
    let bytes = ODS_SEED.as_bytes();
    for end in 0..bytes.len() {
        exercise_ods(bytes[..end].to_vec());
    }
    exercise_ods(bytes.to_vec());
}

#[test]
fn ods_byte_mutation_sweep_never_panics() {
    let bytes = ODS_SEED.as_bytes();
    for position in 0..bytes.len() {
        let mut mutated = bytes.to_vec();
        mutated[position] ^= 0x01;
        exercise_ods(mutated);
    }
}

#[test]
fn odt_truncation_sweep_never_panics() {
    let bytes = ODT_SEED.as_bytes();
    for end in 0..bytes.len() {
        exercise_odt(bytes[..end].to_vec());
    }
    exercise_odt(bytes.to_vec());
}

#[test]
fn odt_byte_mutation_sweep_never_panics() {
    let bytes = ODT_SEED.as_bytes();
    for position in 0..bytes.len() {
        let mut mutated = bytes.to_vec();
        mutated[position] ^= 0x01;
        exercise_odt(mutated);
    }
}

#[test]
fn ods_malformed_inputs_yield_typed_errors() {
    // Mismatched closing tag inside the calcext container.
    assert_ods_typed_error(
        ods_table(
            r#"<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="S.A1"><calcext:condition calcext:apply-style-name="A" calcext:value="x"/></calcext:conditional-formats></calcext:conditional-format>"#,
        ),
        "mismatched calcext close",
    );
    // Unterminated sparkline container at EOF.
    assert_ods_typed_error(
        ods_table(
            r#"<calcext:sparkline-groups><calcext:sparkline-group><calcext:sparklines><calcext:sparkline calcext:cell-address="S.B1" calcext:data-range="S.A1"/></calcext:sparklines>"#,
        ),
        "unterminated sparkline-groups",
    );
    // DOCTYPE with a custom entity definition: quick-xml never expands
    // entities, so using the entity is the error path.
    assert_ods_typed_error(
        r#"<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE office [<!ENTITY x "y">]><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="S"><table:table-row><table:table-cell office:value-type="string"><text:p>&x;</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document>"#
            .as_bytes()
            .to_vec(),
        "DOCTYPE entity use",
    );
    // Undefined entity references are rejected.
    assert_ods_typed_error(
        ods_table(
            r#"<table:table-row><table:table-cell office:value-type="string"><text:p>&undefined;</text:p></table:table-cell></table:table-row>"#,
        ),
        "undefined entity",
    );
    // Duplicate table names are rejected (they break sheet association).
    assert_ods_typed_error(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet" office:version="1.3"><office:body><office:spreadsheet><table:table table:name="S"/><table:table table:name="S"/></office:spreadsheet></office:body></office:document>"#
            .as_bytes()
            .to_vec(),
        "duplicate table name",
    );
    // Row repetition beyond the expansion safety limit.
    assert_ods_typed_error(
        ods_table(
            r#"<table:table-row table:number-rows-repeated="999999999999"><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row>"#,
        ),
        "row repetition overflow",
    );
    // A calcext:value beyond the attribute safety limit.
    assert_ods_typed_error(
        ods_table(&format!(
            r#"<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="S.A1"><calcext:condition calcext:apply-style-name="A" calcext:value="{}"/></calcext:conditional-format></calcext:conditional-formats>"#,
            "x".repeat(70 * 1024)
        )),
        "oversized calcext:value",
    );
    // More rules than the per-format safety limit.
    assert_ods_typed_error(
        ods_table(&format!(
            r#"<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="S.A1">{}</calcext:conditional-format></calcext:conditional-formats>"#,
            r#"<calcext:condition calcext:apply-style-name="A" calcext:value="x"/>"#.repeat(2_000)
        )),
        "rule count overflow",
    );
    // office:dde-source must not have children.
    assert_ods_typed_error(
        ods_table(
            r#"<office:dde-source office:dde-application="a" office:dde-topic="t" office:dde-item="i"><table:table-row/></office:dde-source>"#,
        ),
        "dde-source with children",
    );
    // A second what-if scenario in one sheet.
    assert_ods_typed_error(
        ods_table(
            r#"<table:scenario table:scenario-ranges=".A1:.B2" table:is-active="true"/><table:scenario table:scenario-ranges=".C1:.D2" table:is-active="false"/>"#,
        ),
        "duplicate scenarios",
    );
}

#[test]
fn ods_misplaced_but_wellformed_inputs_do_not_panic() {
    // Extension containers at the wrong depth are ignored, not panicked on.
    exercise_ods(ods_table(
        r#"<table:table-row><table:table-cell><calcext:sparkline-groups/></table:table-cell></table:table-row>"#,
    ));
    exercise_ods(ods_table(
        r#"<table:table-row><calcext:conditional-formats/></table:table-row>"#,
    ));
    // Covered cells and text outside rows are tolerated or typed errors.
    exercise_ods(ods_table(r#"<table:covered-table-cell/>"#));
    exercise_ods(ods_table("stray text"));
    // Invalid UTF-8 anywhere is a typed error, never a panic.
    let mut bytes = ods_table("<table:table-row/>");
    bytes.insert(bytes.len() / 2, 0xFF);
    exercise_ods(bytes);
    // A UTF-16 document is rejected without a panic.
    let utf16: Vec<u8> = ODS_SEED
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    exercise_ods(utf16);
}

#[test]
fn odt_malformed_inputs_yield_typed_errors() {
    let odt_body = |inner: &str| -> Vec<u8> {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" office:mimetype="application/vnd.oasis.opendocument.text" office:version="1.3"><office:body><office:text>{inner}</office:text></office:body></office:document>"#
        )
        .into_bytes()
    };

    // Unterminated section at EOF.
    assert_odt_typed_error(
        odt_body(r#"<text:section text:name="s"><text:p>open"#),
        "unterminated section",
    );
    // Page sequence nested inside a paragraph instead of office:text.
    assert_odt_typed_error(
        odt_body(
            r#"<text:p><text:page-sequence><text:page text:master-page-name="A"/></text:page-sequence></text:p>"#,
        ),
        "misplaced page-sequence",
    );
    // Page sequence with an attribute.
    assert_odt_typed_error(
        odt_body(
            r#"<text:page-sequence text:style-name="x"><text:page text:master-page-name="A"/></text:page-sequence>"#,
        ),
        "page-sequence attributes",
    );
    // Duplicate control xml:id inside one form.
    assert_odt_typed_error(
        odt_body(
            r#"<office:forms><form:form form:name="f"><form:text form:name="a" xml:id="dup"/><form:button form:name="b" xml:id="dup"/></form:form></office:forms>"#,
        ),
        "duplicate control id",
    );
    // A non-Boolean text:protected value.
    assert_odt_typed_error(
        odt_body(r#"<text:section text:name="s" text:protected="maybe"><text:p>x</text:p></text:section>"#),
        "invalid section boolean",
    );
    // A protection key without text:protected is a write-time validation
    // error, but the read path tolerates it — assert it never panics.
    exercise_odt(odt_body(
        r#"<text:section text:name="s" text:protection-key="aGk="><text:p>x</text:p></text:section>"#,
    ));
    // Mismatched section closing tag.
    assert_odt_typed_error(
        odt_body(r#"<text:section text:name="s"><text:p>x</text:p></text:p>"#),
        "mismatched section close",
    );
}

#[test]
fn odt_misplaced_but_wellformed_inputs_do_not_panic() {
    let odt_body = |inner: &str| -> Vec<u8> {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:mimetype="application/vnd.oasis.opendocument.text" office:version="1.3"><office:body><office:text>{inner}</office:text></office:body></office:document>"#
        )
        .into_bytes()
    };
    // Deeply nested valid content stays within depth limits.
    let nested = "<text:span>".repeat(200) + &"</text:span>".repeat(200);
    exercise_odt(odt_body(&nested));
    // Unknown foreign elements inside text are ignored or typed errors.
    exercise_odt(odt_body(r#"<text:p><foo:bar xmlns:foo="urn:example"/>text</text:p>"#));
    // An empty document body.
    exercise_odt(odt_body(""));
    // Invalid UTF-8 is a typed error, never a panic.
    let mut bytes = odt_body("<text:p>x</text:p>");
    bytes.insert(bytes.len() / 2, 0xFE);
    exercise_odt(bytes);
}
