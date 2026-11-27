//! Focused BIFF8 Array owner and payload tests.

use super::super::{Cell, Range};
use super::codec::parse_payload;
use super::{Limits, Owner};

fn payload(tokens: &[u8]) -> Vec<u8> {
    let mut data = vec![0, 0, 1, 0, 1, 2, 1, 0, 0xaa, 0xbb, 0xcc, 0xdd];
    data.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
    data.extend_from_slice(tokens);
    data
}

#[test]
fn parses_iterates_and_round_trips_source_fields() {
    let source = payload(&[0x1e, 7, 0]);
    let owner = parse_payload(&source, Limits::default()).unwrap();
    assert_eq!(owner.range(), Range::try_new(0, 1, 1, 2).unwrap());
    assert_eq!(owner.anchor(), Cell::new(0, 1));
    assert!(owner.always_calculate());
    assert_eq!(owner.unused(), [0xaa, 0xbb, 0xcc, 0xdd]);
    assert_eq!(owner.cell_count(), 4);
    assert_eq!(owner.cells().count(), 4);
    assert_eq!(owner.to_payload().unwrap(), source);
}

#[test]
fn authored_owner_is_canonical_and_inert() {
    let range = Range::try_new(2, 3, 3, 4).unwrap();
    let owner = Owner::from_compiled(range, vec![0x1e, 2, 0]).unwrap();
    assert_eq!(owner.reserved(), 0);
    assert_eq!(owner.unused(), [0; 4]);
    assert!(owner.extra().is_empty());
    assert_eq!(owner.anchor_tokens(), [0x01, 2, 0, 3, 0]);
}

#[test]
fn validates_and_preserves_ordered_ptg_array_extra() {
    let mut source = payload(&[0x40, 1, 2, 3, 4, 5, 6, 7]);
    source.extend_from_slice(&[0, 0, 0]);
    source.extend_from_slice(&[0, 1, 2, 3, 4, 5, 6, 7, 8]);
    let owner = parse_payload(&source, Limits::default()).unwrap();
    assert_eq!(owner.extra(), &[0, 0, 0, 0, 1, 2, 3, 4, 5, 6, 7, 8]);
    assert_eq!(owner.to_payload().unwrap(), source);
}

#[test]
fn rejects_forbidden_unsafe_truncated_and_bounded_inputs() {
    assert!(parse_payload(&payload(&[0x01, 0, 0, 0, 0]), Limits::default()).is_err());
    assert!(
        Owner::from_compiled(Range::try_new(0, 0, 0, 0).unwrap(), vec![0x23, 0, 0, 0, 0]).is_err()
    );
    assert!(parse_payload(&payload(&[0x1f, 0]), Limits::default()).is_err());

    let limits = Limits::default().with_max_cells(1).unwrap();
    assert!(parse_payload(&payload(&[0x1e, 1, 0]), limits).is_err());
}

#[test]
fn validates_rpn_root_stack_types_and_fixed_function_arity() {
    let range = Range::try_new(0, 0, 0, 0).unwrap();
    assert!(Owner::from_compiled(range, vec![0x1e, 1, 0, 0x1e, 2, 0, 0x03]).is_ok());
    assert!(Owner::from_compiled(range, vec![0x03]).is_err());
    assert!(Owner::from_compiled(range, vec![0x1e, 1, 0, 0x1e, 2, 0]).is_err());

    let mut abs = vec![0x1e, 1, 0];
    for _ in 0..8 {
        abs.extend_from_slice(&[0x41, 24, 0]);
    }
    assert!(Owner::from_compiled(range, abs.clone()).is_ok());
    abs.extend_from_slice(&[0x41, 24, 0]);
    assert!(Owner::from_compiled(range, abs).is_err());
}

#[test]
fn enforces_normative_operand_pressure_and_configured_caps() {
    let range = Range::try_new(0, 0, 0, 0).unwrap();
    let mut tokens = Vec::new();
    for value in 0..41u16 {
        tokens.push(0x1e);
        tokens.extend_from_slice(&value.to_le_bytes());
    }
    tokens.extend(std::iter::repeat_n(0x03, 40));
    assert!(Owner::from_compiled(range, tokens).is_err());

    let limits = Limits::default().with_max_operands(1).unwrap();
    assert!(
        Owner::from_compiled_with_limits(range, vec![0x1e, 1, 0, 0x1e, 2, 0, 0x03], limits,)
            .is_err()
    );
    assert!(Limits::default().with_max_nesting_depth(9).is_err());
    assert!(Limits::default().with_max_operands(41).is_err());
}

#[test]
fn compressed_strings_use_actual_expression_size_and_configured_limit() {
    let range = Range::try_new(0, 0, 0, 0).unwrap();
    let mut tokens = Vec::new();
    for _ in 0..40 {
        tokens.extend_from_slice(&[0x17, 41, 0]);
        tokens.extend(std::iter::repeat_n(b'a', 41));
    }
    tokens.extend(std::iter::repeat_n(0x08, 39));
    assert_eq!(tokens.len(), 1_799);
    assert!(Owner::from_compiled(range, tokens).is_err());

    let limits = Limits::default().with_max_string_utf16_units(1).unwrap();
    assert!(
        Owner::from_compiled_with_limits(range, vec![0x17, 2, 0, b'a', b'b'], limits,).is_err()
    );
}

#[test]
fn rejects_invalid_xnum_encodings_in_ptg_and_extra_arrays() {
    let range = Range::try_new(0, 0, 0, 0).unwrap();
    for bits in [f64::NAN.to_bits(), 1, (-0.0f64).to_bits()] {
        let mut tokens = vec![0x1f];
        tokens.extend_from_slice(&bits.to_le_bytes());
        assert!(Owner::from_compiled(range, tokens).is_err());

        let mut source = payload(&[0x40, 0, 0, 0, 0, 0, 0, 0]);
        source.extend_from_slice(&[0, 0, 0, 1]);
        source.extend_from_slice(&bits.to_le_bytes());
        assert!(parse_payload(&source, Limits::default()).is_err());
    }
}

#[test]
fn configurable_limits_have_finite_structural_ceilings() {
    assert!(
        Limits::new(
            usize::MAX,
            usize::MAX,
            usize::MAX,
            usize::MAX,
            usize::MAX,
            usize::MAX,
            usize::MAX,
        )
        .is_err()
    );
    assert!(Limits::default().with_max_tokens(1_801).is_err());
    assert!(Limits::default().with_max_operator_depth(257).is_err());
    assert!(Limits::default().with_max_string_utf16_units(256).is_err());
}

#[test]
fn generic_ptg_exp_requires_exact_opcode_and_zero_high_column_byte() {
    use super::super::validation::is_ptg_exp;

    assert!(is_ptg_exp(&[0x01, 1, 0, 2, 0]));
    assert!(!is_ptg_exp(&[0x81, 1, 0, 2, 0]));
    assert!(!is_ptg_exp(&[0x01, 1, 0, 2, 1]));
}

#[test]
fn rejects_reference_class_root_malformed_attributes_and_orphan_memory() {
    let range = Range::try_new(0, 0, 0, 0).unwrap();
    assert!(Owner::from_compiled(range, vec![0x24, 0, 0, 0, 0]).is_err());

    assert!(Owner::from_compiled(range, vec![0x19, 0x40, 7, 1, 0x1e, 1, 0]).is_err());
    assert!(
        Owner::from_compiled(
            range,
            vec![
                0x1e, 1, 0, // condition
                0x19, 0x02, 99, 0, // invalid AttrIf branch offset
                0x1e, 2, 0, // branch
                0x19, 0x08, 0, 0, // AttrGoto
            ],
        )
        .is_err()
    );
    assert!(
        Owner::from_compiled(
            range,
            vec![0x29, 3, 0, 0x1e, 1, 0], // orphan PtgMemFunc + scalar
        )
        .is_err()
    );
}

#[test]
fn memory_ptg_result_uses_its_classified_value_type() {
    let range = Range::try_new(0, 0, 0, 0).unwrap();
    let expression = [
        11, 0, // memory cce: two PtgRef plus PtgRange
        0x24, 0, 0, 0, 0, // reference-class A1
        0x24, 0, 0, 1, 0,    // reference-class B1
        0x11, // PtgRange
    ];

    let mut value_class = vec![0x49];
    value_class.extend_from_slice(&expression);
    assert!(Owner::from_compiled(range, value_class).is_ok());

    let mut reference_class = vec![0x29];
    reference_class.extend_from_slice(&expression);
    assert!(Owner::from_compiled(range, reference_class).is_err());
}
