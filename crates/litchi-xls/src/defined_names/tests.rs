use crate::formula::FormulaContext;

use super::*;

fn lbl(flags: u16, itab: u16, name: &str, unicode: bool, formula: &[u8]) -> Vec<u8> {
    let units: Vec<u16> = name.encode_utf16().collect();
    let cch = if unicode { units.len() } else { name.len() };
    let mut data = Vec::new();
    data.extend_from_slice(&flags.to_le_bytes());
    data.push(0);
    data.push(cch as u8);
    data.extend_from_slice(&(formula.len() as u16).to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&itab.to_le_bytes());
    data.extend_from_slice(&[0; 4]);
    data.push(u8::from(unicode));
    if unicode {
        for unit in units {
            data.extend_from_slice(&unit.to_le_bytes());
        }
    } else {
        data.extend_from_slice(name.as_bytes());
    }
    data.extend_from_slice(formula);
    data
}

#[test]
fn parses_compressed_unicode_hidden_and_scoped_names() {
    let ordinary =
        DefinedNameSlot::parse(&lbl(FLAG_HIDDEN, 2, "Rate", false, &[0x1e, 2, 0]), 1).unwrap();
    assert_eq!(ordinary.name, "Rate");
    assert!(ordinary.hidden);
    assert_eq!(ordinary.itab, 2);
    assert_eq!(ordinary.formula_tokens, [0x1e, 2, 0]);

    let unicode = DefinedNameSlot::parse(&lbl(0, 0, "税率", true, &[0x1e, 3, 0]), 2).unwrap();
    assert_eq!(unicode.name, "税率");
}

#[test]
fn parses_every_built_in_name() {
    for code in 0u8..=0x0d {
        let slot = DefinedNameSlot::parse(
            &lbl(
                FLAG_BUILT_IN,
                1,
                &char::from(code).to_string(),
                false,
                &[0x1e, 1, 0],
            ),
            u32::from(code) + 1,
        )
        .unwrap();
        assert!(matches!(slot.kind, DefinedNameKind::BuiltIn(_)));
    }
}

#[test]
fn macro_slots_keep_indices_but_have_no_symbol_or_public_value() {
    let slot = DefinedNameSlot::parse(&lbl(FLAG_PROCEDURE, 0, "Macro", false, &[]), 7).unwrap();
    assert!(slot.formula_symbol().is_none());
    assert!(
        slot.into_public(1, &FormulaContext::default())
            .unwrap()
            .is_macro()
    );
}

#[test]
fn rejects_malformed_names_and_scope() {
    assert!(DefinedNameSlot::parse(&[0; 14], 1).is_err());
    let mut truncated = lbl(0, 0, "Name", true, &[0x1e, 1, 0]);
    truncated.truncate(17);
    assert!(DefinedNameSlot::parse(&truncated, 1).is_err());
    assert!(DefinedNameSlot::parse(&lbl(FLAG_BUILT_IN, 0, "x", false, &[]), 1).is_err());
    let invalid_scope = DefinedNameSlot::parse(&lbl(0, 3, "Name", false, &[]), 1).unwrap();
    assert!(
        invalid_scope
            .into_public(2, &FormulaContext::default())
            .is_err()
    );
}

#[test]
fn name_comment_requires_exact_header_and_matching_preceding_name() {
    let mut slot = DefinedNameSlot::parse(&lbl(0, 0, "Rate", false, &[]), 1).unwrap();
    let mut comment = Vec::new();
    comment.extend_from_slice(&NAME_CMT_RECORD_TYPE.to_le_bytes());
    comment.extend_from_slice(&[0; 10]);
    comment.extend_from_slice(&5u16.to_le_bytes());
    comment.extend_from_slice(&1u16.to_le_bytes());
    comment.extend_from_slice(&[0, b'O', b't', b'h', b'e', b'r', 0, b'X']);
    assert!(slot.attach_comment(&comment).is_err());
    comment[16] = 2;
    assert!(slot.attach_comment(&comment).is_err());
}
