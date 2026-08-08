use litchi_ods::FlatSpreadsheet;

const ODS_SEED: &str = r##"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-row><table:table-cell office:value-type="float" office:value="7" table:formula="of:=1+6"><text:p>7</text:p><office:annotation><text:p>note</text:p></office:annotation></table:table-cell><table:table-cell table:number-columns-repeated="3" office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row><table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="string"><text:p>tail</text:p></table:table-cell></table:table-row><calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="Data.A1:Data.D1"><calcext:condition calcext:apply-style-name="Good" calcext:value="cell-content()&gt;5" calcext:base-cell-address="Data.A1"/><calcext:color-scale><calcext:color-scale-entry calcext:value="0" calcext:type="minimum" calcext:color="#ff0000"/><calcext:color-scale-entry calcext:value="1" calcext:type="maximum" calcext:color="#00ff00"/></calcext:color-scale></calcext:conditional-format></calcext:conditional-formats><calcext:sparkline-groups><calcext:sparkline-group calcext:type="line" calcext:markers="true"><calcext:sparklines><calcext:sparkline calcext:cell-address="Data.E1" calcext:data-range="Data.A1:Data.D1"/></calcext:sparklines></calcext:sparkline-group></calcext:sparkline-groups></table:table><table:named-expressions><table:named-range table:name="Total" table:base-cell-address="$Data.$A$1" table:cell-range-address="$Data.$A$1:$Data.$D$1"/></table:named-expressions></office:spreadsheet></office:body></office:document>"##;

fn table(inner: &str) -> Vec<u8> {
    format!(r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet"><office:body><office:spreadsheet><table:table table:name="S">{inner}</table:table></office:spreadsheet></office:body></office:document>"#).into_bytes()
}

fn exercise(bytes: Vec<u8>) -> Option<litchi_core::Error> {
    match FlatSpreadsheet::from_bytes(bytes) {
        Ok(flat) => {
            let _ = flat.sheets();
            let _ = flat.dde();
            let _ = flat.scenarios();
            None
        },
        Err(error) => Some(error),
    }
}

fn typed_error(bytes: Vec<u8>, case: &str) {
    let error = exercise(bytes).unwrap_or_else(|| panic!("{case} unexpectedly parsed"));
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(_)),
        "{case}: {error:?}"
    );
}

#[test]
fn truncation_and_single_byte_mutation_sweeps_never_panic() {
    let bytes = ODS_SEED.as_bytes();
    for end in 0..bytes.len() {
        exercise(bytes[..end].to_vec());
    }
    exercise(bytes.to_vec());
    for position in 0..bytes.len() {
        let mut mutated = bytes.to_vec();
        mutated[position] ^= 1;
        exercise(mutated);
    }
}

#[test]
fn malformed_flat_ods_inputs_return_typed_errors() {
    typed_error(
        table(
            "<calcext:conditional-formats><calcext:conditional-format></calcext:conditional-formats></calcext:conditional-format>",
        ),
        "mismatched close",
    );
    typed_error(
        table("<calcext:sparkline-groups><calcext:sparkline-group>"),
        "unterminated sparkline group",
    );
    typed_error(br#"<?xml version="1.0"?><!DOCTYPE office [<!ENTITY x "y">]><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet"><office:body><office:spreadsheet/></office:body></office:document>"#.to_vec(), "DOCTYPE");
    typed_error(
        table(
            "<table:table-row><table:table-cell><text:p>&undefined;</text:p></table:table-cell></table:table-row>",
        ),
        "undefined entity",
    );
    typed_error(
        table(
            "<table:table-row table:number-rows-repeated=\"999999999999\"><table:table-cell/></table:table-row>",
        ),
        "row repetition overflow",
    );
    typed_error(
        table(&format!(
            "<calcext:conditional-formats><calcext:conditional-format calcext:target-range-address=\"S.A1\"><calcext:condition calcext:value=\"{}\"/></calcext:conditional-format></calcext:conditional-formats>",
            "x".repeat(70 * 1024)
        )),
        "oversized attribute",
    );
    typed_error(
        table(&format!(
            "<calcext:conditional-formats><calcext:conditional-format>{}</calcext:conditional-format></calcext:conditional-formats>",
            "<calcext:condition/>".repeat(2_000)
        )),
        "conditional rule overflow",
    );
    typed_error(
        table(
            "<office:dde-source office:dde-application=\"a\" office:dde-topic=\"t\" office:dde-item=\"i\"><table:table-row/></office:dde-source>",
        ),
        "DDE source child",
    );
    typed_error(
        table(
            "<table:scenario table:scenario-ranges=\".A1:.B2\" table:is-active=\"true\"/><table:scenario table:scenario-ranges=\".C1:.D2\" table:is-active=\"false\"/>",
        ),
        "duplicate scenarios",
    );
}

#[test]
fn misplaced_well_formed_and_encoding_inputs_never_panic() {
    exercise(table(
        "<table:table-row><table:table-cell><calcext:sparkline-groups/></table:table-cell></table:table-row>",
    ));
    exercise(table(
        "<table:table-row><calcext:conditional-formats/></table:table-row>",
    ));
    exercise(table("<table:covered-table-cell/>"));
    exercise(table("stray text"));
    let mut invalid_utf8 = table("<table:table-row/>");
    invalid_utf8.insert(invalid_utf8.len() / 2, 0xff);
    typed_error(invalid_utf8, "invalid UTF-8");
    let utf16 = ODS_SEED.encode_utf16().flat_map(u16::to_le_bytes).collect();
    typed_error(utf16, "UTF-16");
}
