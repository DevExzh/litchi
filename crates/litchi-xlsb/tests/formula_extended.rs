#![allow(
    clippy::pedantic,
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure extraction and exact fixture comparisons"
)]

use litchi_xlsb::formula::{
    ExternalTableReference, Parser, TableColumns, TableDataType, TableNamedColumns, TableReference,
    TableRowType, Token,
};

fn utf16(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

#[test]
fn parses_resident_and_nonresident_ptg_list() {
    let resident = [
        0x18, 0x19, 0x02, 0x00, 0x1A, 0x00, 0x07, 0x00, 0x00, 0x00, 0x01, 0x00, 0x03, 0x00,
    ];
    assert_eq!(
        Parser::new(&resident).parse().unwrap(),
        vec![Token::TableReference(TableReference {
            sheet_index: 2,
            row_type: Some(TableRowType::DataAndHeaders),
            columns: Some(TableColumns::Range { first: 1, last: 3 }),
            square_bracket_space: false,
            comma_space: false,
            data_type: TableDataType::Reference,
            invalid: false,
            list_index: Some(7),
            external: None,
        })]
    );

    let token = [0x18, 0x19, 0x04, 0x00, 0x00, 0x20, 0, 0, 0, 0, 0, 0, 0, 0];
    let mut extra = vec![1, 0x06, 0x00, 5, 0];
    extra.extend(utf16("Sales"));
    extra.extend([0, 0, 2]);
    for (not_last, name) in [(1, "From"), (0, "To")] {
        extra.push(not_last);
        extra.extend([2, 0]);
        extra.extend((name.encode_utf16().count() as u32).to_le_bytes());
        extra.extend(utf16(name));
    }
    let parsed = Parser::with_extra(&token, &extra).parse().unwrap();
    assert_eq!(
        parsed,
        vec![Token::TableReference(TableReference {
            sheet_index: 4,
            row_type: None,
            columns: None,
            square_bracket_space: false,
            comma_space: false,
            data_type: TableDataType::Reference,
            invalid: false,
            list_index: None,
            external: Some(ExternalTableReference {
                table: "Sales".into(),
                row_type: TableRowType::DataAndHeaders,
                columns: TableNamedColumns::Range {
                    first: "From".into(),
                    last: "To".into()
                },
            }),
        })]
    );
    let (written_token, written_extra) = parsed[0].to_extended_binary().unwrap();
    assert_eq!(
        Parser::with_extra(&written_token, &written_extra)
            .parse()
            .unwrap(),
        parsed
    );
}

#[test]
fn parses_ptg_sx_name_and_rejects_reserved_extended_fields() {
    assert_eq!(
        Parser::new(&[0x18, 0x1D, 5, 0, 0, 0]).parse().unwrap(),
        vec![Token::PivotName(5)]
    );
    let pivot = Token::PivotName(5);
    let (token, extra) = pivot.to_extended_binary().unwrap();
    assert!(extra.is_empty());
    assert_eq!(Parser::new(&token).parse().unwrap(), vec![pivot]);
    for malformed in [
        vec![0x18, 0x20],
        vec![0x18, 0x19, 0, 0, 0, 0x40, 1, 0, 0, 0, 0, 0, 0, 0],
        vec![0x18, 0x19, 0, 0, 3, 0, 1, 0, 0, 0, 0, 0, 0, 0],
    ] {
        assert!(Parser::new(&malformed).parse().is_err());
    }
}
