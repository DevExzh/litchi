use super::{
    Axis, AxisType, Color, Colors, EmptyCells, Error, Formula, FormulaKind, FrtState, Group,
    Groups, Limits, Location, Options, Sparkline, SparklineType, encode_block, parse_block,
};

fn black() -> Color {
    Color::rgb(0, 0, 0, 255, 0)
}

fn one_group() -> Groups {
    let item = Sparkline::new(Location::new(0, 0).unwrap(), None);
    Groups::new(vec![
        Group::new(SparklineType::Line, Colors::uniform(black()), vec![item]).unwrap(),
    ])
    .unwrap()
}

fn hard_coded_fixture() -> Vec<u8> {
    let mut bytes = vec![0xA2, 0x08, 0x00, 0x91, 0x08, 0x62];
    // BrtBeginSparklineGroup: empty FRT formula header and option word.
    bytes.extend_from_slice(&[0, 0, 0, 0, 0, 0]);
    // Eight explicit opaque-black BrtColor values in normative slot order.
    for _ in 0..8 {
        bytes.extend_from_slice(&[0x05, 0, 0, 0, 0, 0, 0, 0xFF]);
    }
    bytes.extend_from_slice(&0.0f64.to_le_bytes()); // dManualMax
    bytes.extend_from_slice(&0.0f64.to_le_bytes()); // dManualMin
    bytes.extend_from_slice(&1.0f64.to_le_bytes()); // dLineWeight
    bytes.extend_from_slice(&0u32.to_le_bytes()); // line
    bytes.extend_from_slice(&[0xA0, 0x08, 0x00, 0x93, 0x08, 0x20]);
    bytes.extend_from_slice(&2u32.to_le_bytes()); // FRTHeader: sqref only
    bytes.extend_from_slice(&1u32.to_le_bytes()); // one FRTSqref
    bytes.extend_from_slice(&2u32.to_le_bytes()); // fDoAdjust
    bytes.extend_from_slice(&1u32.to_le_bytes()); // one UncheckedRfX
    bytes.extend_from_slice(&[0; 16]); // A1:A1
    bytes.extend_from_slice(&[
        0xA1, 0x08, 0x00, // BrtEndSparklines
        0x92, 0x08, 0x00, // BrtEndSparklineGroup
        0xA3, 0x08, 0x00, // BrtEndSparklineGroups
    ]);
    bytes
}

fn group_payload(bytes: &[u8]) -> Vec<u8> {
    crate::raw::Records::new(bytes)
        .map(|record| record.unwrap())
        .find(|record| record.kind() == crate::raw::kind::BEGIN_SPARKLINE_GROUP)
        .unwrap()
        .payload()
        .to_vec()
}

fn area_token(rows: (u32, u32), columns: (u16, u16), ixti: u16) -> Vec<u8> {
    let mut token = vec![0x3B];
    token.extend_from_slice(&ixti.to_le_bytes());
    token.extend_from_slice(&rows.0.to_le_bytes());
    token.extend_from_slice(&rows.1.to_le_bytes());
    token.extend_from_slice(&columns.0.to_le_bytes());
    token.extend_from_slice(&columns.1.to_le_bytes());
    token
}

#[test]
fn hard_coded_spec_fixture_roundtrips_exactly() {
    let fixture = hard_coded_fixture();
    let (groups, consumed) = parse_block(&fixture, Limits::DEFAULT).unwrap();
    assert_eq!(consumed, fixture.len());
    assert_eq!(groups.len(), 1);
    assert_eq!(groups.as_slice()[0].sparklines()[0].location().row(), 0);
    assert_eq!(encode_block(&groups, Limits::DEFAULT).unwrap(), fixture);
}

#[test]
fn semantic_options_colors_state_and_formulas_roundtrip() {
    let formula =
        Formula::with_limits(vec![0x23, 1, 0, 0, 0], vec![0xAA], Limits::DEFAULT).unwrap();
    let state = FrtState::new(true, true, true, true);
    let item = Sparkline::new(
        Location::with_state(4, 7, state).unwrap(),
        Some(formula.clone()),
    );
    let group = Group::new(
        SparklineType::Stacked,
        Colors::uniform(Color::theme(11, -12).unwrap()),
        vec![item],
    )
    .unwrap()
    .with_empty_cells(EmptyCells::Span)
    .with_options(Options::MARKERS | Options::NEGATIVE | Options::RIGHT_TO_LEFT)
    .unwrap()
    .with_axes(Axis::custom(-2.5).unwrap(), Axis::group())
    .with_line_weight(1.25)
    .unwrap()
    .with_date_axis(formula);
    let groups = Groups::new(vec![group]).unwrap();
    let encoded = encode_block(&groups, Limits::DEFAULT).unwrap();
    let (parsed, consumed) = parse_block(&encoded, Limits::DEFAULT).unwrap();
    assert_eq!(consumed, encoded.len());
    assert_eq!(parsed, groups);
    assert_eq!(
        parsed.as_slice()[0].date_formula().unwrap().kind(),
        FormulaKind::Name
    );
}

#[test]
fn parser_returns_consumed_block_without_claiming_following_records() {
    let fixture = hard_coded_fixture();
    let mut worksheet_tail = fixture.clone();
    worksheet_tail.extend_from_slice(&[0x82, 0x01, 0x00]);
    let (_, consumed) = parse_block(&worksheet_tail, Limits::DEFAULT).unwrap();
    assert_eq!(consumed, fixture.len());
}

#[test]
fn strict_grammar_rejects_unknown_internal_records_and_nonempty_delimiters() {
    let mut unknown = hard_coded_fixture();
    let position = unknown.len() - 9;
    unknown.splice(position..position, [0x01, 0x00]);
    assert!(matches!(
        parse_block(&unknown, Limits::DEFAULT),
        Err(Error::Record { .. })
    ));

    let mut nonempty = hard_coded_fixture();
    nonempty[2] = 1;
    nonempty.insert(3, 0xCC);
    assert!(matches!(
        parse_block(&nonempty, Limits::DEFAULT),
        Err(Error::Delimiter { .. })
    ));
}

#[test]
fn formulas_reject_extra_tokens_reserved_class_and_non_vectors() {
    assert!(Formula::new(vec![0x23, 1, 0, 0, 0, 0], vec![]).is_err());
    assert!(Formula::new(vec![0x03, 1, 0, 0, 0], vec![]).is_err());

    let mut area = vec![0x3B, 0, 0];
    area.extend_from_slice(&0u32.to_le_bytes());
    area.extend_from_slice(&1u32.to_le_bytes());
    area.extend_from_slice(&0u16.to_le_bytes());
    area.extend_from_slice(&1u16.to_le_bytes());
    assert!(Formula::new(area, vec![]).is_err());
}

#[test]
fn all_allowed_formula_tokens_have_structural_accessors() {
    let name = Formula::new(vec![0x23, 9, 0, 0, 0], vec![]).unwrap();
    assert_eq!(name.kind(), FormulaKind::Name);
    assert_eq!(name.name_index(), Some(9));
    assert_eq!(name.ixti(), None);

    let external = Formula::new(vec![0x39, 0x34, 0x12, 7, 0, 0, 0], vec![]).unwrap();
    assert_eq!(external.kind(), FormulaKind::ExternalName);
    assert_eq!(external.ixti(), Some(0x1234));
    assert_eq!(external.name_index(), Some(7));

    let mut reference = vec![0x3A, 3, 0];
    reference.extend_from_slice(&42u32.to_le_bytes());
    reference.extend_from_slice(&(0xC000u16 | 11).to_le_bytes());
    let reference = Formula::new(reference, vec![]).unwrap();
    assert_eq!(reference.kind(), FormulaKind::Reference3d);
    assert_eq!(reference.ixti(), Some(3));
    let cell = reference.reference().unwrap();
    assert_eq!((cell.row(), cell.column()), (42, 11));
    assert!(cell.row_relative());
    assert!(cell.column_relative());

    let area = Formula::new(area_token((5, 5), (2, 8), 4), vec![]).unwrap();
    assert_eq!(area.kind(), FormulaKind::Area3d);
    assert_eq!(area.ixti(), Some(4));
    let range = area.area().unwrap();
    assert_eq!(
        (
            range.row_first(),
            range.row_last(),
            range.column_first(),
            range.column_last()
        ),
        (5, 5, 2, 8)
    );
}

#[test]
fn area3d_relative_masks_follow_rgce_column_bit_layout() {
    // A one-row source requires equal fRwRel (0x4000); differing fColRel is allowed.
    assert!(Formula::new(area_token((1, 1), (0x0001, 0x8002), 0), vec![]).is_ok());
    assert!(Formula::new(area_token((1, 1), (0x0001, 0x4002), 0), vec![]).is_err());

    // A one-column source requires equal fColRel (0x8000); differing fRwRel is allowed.
    assert!(Formula::new(area_token((1, 2), (0x0001, 0x4001), 0), vec![]).is_ok());
    assert!(Formula::new(area_token((1, 2), (0x0001, 0x8001), 0), vec![]).is_err());
}

#[test]
fn colors_axes_and_line_weight_enforce_wire_domains() {
    assert!(Color::from_raw([0; 8]).is_err());
    assert!(Color::from_raw([0x04, 0, 0, 0, 0, 0, 0, 0]).is_err());
    assert!(Color::theme(12, 0).is_err());
    assert!(Color::palette(0x52, 0).is_err());
    assert!(Axis::new(AxisType::Individual, 1.0).is_err());
    assert!(Axis::custom(f64::NAN).is_err());
    let group = Group::new(
        SparklineType::Line,
        Colors::uniform(black()),
        vec![Sparkline::new(Location::new(0, 0).unwrap(), None)],
    )
    .unwrap();
    assert!(group.clone().with_line_weight(-0.0).is_err());
    assert!(group.with_line_weight(1584.000_1).is_err());
}

#[test]
fn palette_rgb_theme_and_all_distinct_color_slots_roundtrip_exactly() {
    let colors = Colors::new(
        Color::palette(1, -1).unwrap(),
        Color::rgb(2, 3, 4, 5, 6),
        Color::theme(2, 7).unwrap(),
        Color::palette(8, 9).unwrap(),
        Color::rgb(10, 11, 12, 13, 14),
        Color::theme(5, 15).unwrap(),
        Color::palette(16, 17).unwrap(),
        Color::from_raw([0x05, 0xEE, 18, 0, 19, 20, 21, 22]).unwrap(),
    );
    let group = Group::new(
        SparklineType::Column,
        colors,
        vec![Sparkline::new(Location::new(0, 0).unwrap(), None)],
    )
    .unwrap();
    let groups = Groups::new(vec![group]).unwrap();
    let encoded = encode_block(&groups, Limits::DEFAULT).unwrap();
    let (parsed, _) = parse_block(&encoded, Limits::DEFAULT).unwrap();
    assert_eq!(parsed.as_slice()[0].colors(), colors);
    assert_eq!(parsed.as_slice()[0].colors().low().raw()[1], 0xEE);
}

#[test]
fn every_option_and_shared_enum_roundtrips_through_exact_wire_flags() {
    let options = [
        (Options::MARKERS, 0x0008),
        (Options::HIGH, 0x0010),
        (Options::LOW, 0x0020),
        (Options::FIRST, 0x0040),
        (Options::LAST, 0x0080),
        (Options::NEGATIVE, 0x0100),
        (Options::AXIS, 0x0200),
        (Options::DISPLAY_HIDDEN, 0x0400),
        (Options::RIGHT_TO_LEFT, 0x8000),
    ];
    for (option, wire) in options {
        let group = Group::new(
            SparklineType::Line,
            Colors::uniform(black()),
            vec![Sparkline::new(Location::new(0, 0).unwrap(), None)],
        )
        .unwrap()
        .with_options(option)
        .unwrap();
        let encoded = encode_block(&Groups::new(vec![group]).unwrap(), Limits::DEFAULT).unwrap();
        let payload = group_payload(&encoded);
        assert_ne!(u16::from_le_bytes([payload[4], payload[5]]) & wire, 0);
    }

    for kind in [
        SparklineType::Line,
        SparklineType::Column,
        SparklineType::Stacked,
    ] {
        for empty in [EmptyCells::Zero, EmptyCells::Gap, EmptyCells::Span] {
            for axis in [AxisType::Individual, AxisType::Group, AxisType::Custom] {
                let bound = match axis {
                    AxisType::Individual => Axis::individual(),
                    AxisType::Group => Axis::group(),
                    AxisType::Custom => Axis::custom(2.0).unwrap(),
                };
                let group = Group::new(
                    kind,
                    Colors::uniform(black()),
                    vec![Sparkline::new(Location::new(0, 0).unwrap(), None)],
                )
                .unwrap()
                .with_empty_cells(empty)
                .with_axes(bound, bound);
                let groups = Groups::new(vec![group]).unwrap();
                let encoded = encode_block(&groups, Limits::DEFAULT).unwrap();
                let (parsed, _) = parse_block(&encoded, Limits::DEFAULT).unwrap();
                let parsed = &parsed.as_slice()[0];
                assert_eq!(parsed.kind(), kind);
                assert_eq!(parsed.empty_cells(), empty);
                assert_eq!(parsed.minimum().kind(), axis);
                assert_eq!(parsed.maximum().kind(), axis);
            }
        }
    }
}

#[test]
fn unknown_semantic_option_bits_are_rejected() {
    let group = Group::new(
        SparklineType::Line,
        Colors::uniform(black()),
        vec![Sparkline::new(Location::new(0, 0).unwrap(), None)],
    )
    .unwrap();
    assert!(
        group
            .with_options(Options::from_bits_retain(0x8000))
            .is_err()
    );
}

#[test]
fn date_axis_flag_and_formula_presence_are_independent_wire_state() {
    let flag_only = Group::new(
        SparklineType::Line,
        Colors::uniform(black()),
        vec![Sparkline::new(Location::new(0, 0).unwrap(), None)],
    )
    .unwrap()
    .with_date_axis_enabled(true);
    let formula_only = Group::new(
        SparklineType::Line,
        Colors::uniform(black()),
        vec![Sparkline::new(Location::new(0, 1).unwrap(), None)],
    )
    .unwrap()
    .with_date_formula(Some(Formula::name(1).unwrap()));
    let groups = Groups::new(vec![flag_only, formula_only]).unwrap();
    let encoded = encode_block(&groups, Limits::DEFAULT).unwrap();
    let (parsed, _) = parse_block(&encoded, Limits::DEFAULT).unwrap();
    assert!(parsed.as_slice()[0].date_axis());
    assert!(parsed.as_slice()[0].date_formula().is_none());
    assert!(!parsed.as_slice()[1].date_axis());
    assert!(parsed.as_slice()[1].date_formula().is_some());
}

#[test]
fn worksheet_grid_is_strict_and_duplicate_destinations_are_preserved() {
    assert!(Location::new(litchi_sheet::ROWS, 0).is_err());
    assert!(Location::new(0, litchi_sheet::COLUMNS).is_err());
    let location = Location::new(2, 3).unwrap();
    let first = Group::new(
        SparklineType::Line,
        Colors::uniform(black()),
        vec![Sparkline::new(location, None)],
    )
    .unwrap();
    let second = Group::new(
        SparklineType::Column,
        Colors::uniform(black()),
        vec![Sparkline::new(location, None)],
    )
    .unwrap();
    let groups = Groups::new(vec![first, second]).unwrap();
    let encoded = encode_block(&groups, Limits::DEFAULT).unwrap();
    let (parsed, _) = parse_block(&encoded, Limits::DEFAULT).unwrap();
    assert_eq!(parsed, groups);
}

#[test]
fn exact_and_plus_one_collection_limits_are_enforced() {
    let groups = one_group();
    let exact = Limits::DEFAULT
        .with_groups(1)
        .unwrap()
        .with_per_group(1)
        .unwrap()
        .with_total(1)
        .unwrap();
    let encoded = encode_block(&groups, exact).unwrap();
    parse_block(&encoded, exact).unwrap();

    let location = Location::new(0, 1).unwrap();
    let extra = Sparkline::new(location, None);
    let two = Groups::new(vec![
        Group::new(
            SparklineType::Line,
            Colors::uniform(black()),
            vec![Sparkline::new(Location::new(0, 0).unwrap(), None), extra],
        )
        .unwrap(),
    ])
    .unwrap();
    let per_group_only = exact.with_total(2).unwrap();
    assert!(matches!(
        encode_block(&two, per_group_only),
        Err(Error::Limit {
            resource: "sparklines per group",
            actual: 2,
            maximum: 1
        })
    ));

    let aggregate_only = exact.with_per_group(2).unwrap();
    let encoded_two = encode_block(&two, Limits::DEFAULT).unwrap();
    assert!(matches!(
        encode_block(&two, aggregate_only),
        Err(Error::Limit {
            resource: "total sparklines",
            actual: 2,
            maximum: 1
        })
    ));
    assert!(matches!(
        parse_block(&encoded_two, aggregate_only),
        Err(Error::Limit {
            resource: "total sparklines",
            actual: 2,
            maximum: 1
        })
    ));
}

#[test]
fn parser_enforces_group_and_per_group_limits_before_growth() {
    let first = Group::new(
        SparklineType::Line,
        Colors::uniform(black()),
        vec![Sparkline::new(Location::new(0, 0).unwrap(), None)],
    )
    .unwrap();
    let second = Group::new(
        SparklineType::Column,
        Colors::uniform(black()),
        vec![Sparkline::new(Location::new(0, 1).unwrap(), None)],
    )
    .unwrap();
    let encoded_groups =
        encode_block(&Groups::new(vec![first, second]).unwrap(), Limits::DEFAULT).unwrap();
    let one_group = Limits::DEFAULT.with_groups(1).unwrap();
    assert!(matches!(
        parse_block(&encoded_groups, one_group),
        Err(Error::Limit {
            resource: "groups",
            actual: 2,
            maximum: 1
        })
    ));

    let group = Group::new(
        SparklineType::Line,
        Colors::uniform(black()),
        vec![
            Sparkline::new(Location::new(0, 0).unwrap(), None),
            Sparkline::new(Location::new(0, 1).unwrap(), None),
        ],
    )
    .unwrap();
    let encoded_items = encode_block(&Groups::new(vec![group]).unwrap(), Limits::DEFAULT).unwrap();
    let one_item = Limits::DEFAULT
        .with_per_group(1)
        .unwrap()
        .with_total(2)
        .unwrap();
    assert!(matches!(
        parse_block(&encoded_items, one_item),
        Err(Error::Limit {
            resource: "sparklines per group",
            actual: 2,
            maximum: 1
        })
    ));
}

#[test]
fn exact_and_plus_one_formula_limits_are_enforced() {
    let exact = Limits::DEFAULT
        .with_formula_tokens(5)
        .unwrap()
        .with_formula_ancillary(1)
        .unwrap();
    assert!(Formula::with_limits(vec![0x23, 1, 0, 0, 0], vec![1], exact).is_ok());
    assert!(matches!(
        Formula::with_limits(vec![0x23, 1, 0, 0, 0], vec![1, 2], exact),
        Err(Error::Limit {
            resource: "formula ancillary bytes",
            actual: 2,
            maximum: 1
        })
    ));
    assert!(matches!(
        Formula::with_limits(
            vec![0x23, 1, 0, 0, 0],
            vec![],
            exact.with_formula_tokens(4).unwrap()
        ),
        Err(Error::Limit {
            resource: "formula token bytes",
            actual: 5,
            maximum: 4
        })
    ));
}

#[test]
fn exact_and_plus_one_block_limits_are_enforced() {
    let groups = one_group();
    let encoded = encode_block(&groups, Limits::DEFAULT).unwrap();
    let exact = Limits::DEFAULT.with_block_bytes(encoded.len()).unwrap();
    assert_eq!(encode_block(&groups, exact).unwrap(), encoded);
    parse_block(&encoded, exact).unwrap();

    let short = Limits::DEFAULT.with_block_bytes(encoded.len() - 1).unwrap();
    assert!(matches!(
        encode_block(&groups, short),
        Err(Error::Limit {
            resource: "block bytes",
            ..
        })
    ));
    assert!(matches!(
        parse_block(&encoded, short),
        Err(Error::Limit {
            resource: "block bytes",
            ..
        })
    ));
}

#[test]
fn exact_and_plus_one_record_payload_limits_are_enforced() {
    let groups = one_group();
    let encoded = encode_block(&groups, Limits::DEFAULT).unwrap();
    let exact = Limits::DEFAULT.with_record_bytes(98).unwrap();
    assert_eq!(encode_block(&groups, exact).unwrap(), encoded);
    parse_block(&encoded, exact).unwrap();

    let short = Limits::DEFAULT.with_record_bytes(97).unwrap();
    assert!(matches!(
        encode_block(&groups, short),
        Err(Error::Limit {
            resource: "record payload bytes",
            actual: 98,
            maximum: 97
        })
    ));
    assert!(matches!(parse_block(&encoded, short), Err(Error::Wire(_))));
}

#[test]
fn multi_item_block_budget_is_checked_at_each_internal_record() {
    let group = Group::new(
        SparklineType::Line,
        Colors::uniform(black()),
        vec![
            Sparkline::new(Location::new(0, 0).unwrap(), None),
            Sparkline::new(Location::new(0, 1).unwrap(), None),
        ],
    )
    .unwrap();
    let groups = Groups::new(vec![group]).unwrap();
    let encoded = encode_block(&groups, Limits::DEFAULT).unwrap();
    let exact = Limits::DEFAULT.with_block_bytes(encoded.len()).unwrap();
    parse_block(&encoded, exact).unwrap();

    let mut records = crate::raw::Records::new(&encoded);
    let mut seen = 0;
    let mut second_end = 0;
    while let Some(record) = records.next() {
        if record.unwrap().kind() == crate::raw::kind::SPARKLINE {
            seen += 1;
            if seen == 2 {
                second_end = records.offset();
                break;
            }
        }
    }
    assert!(second_end > 1);
    let too_short = Limits::DEFAULT.with_block_bytes(second_end - 1).unwrap();
    assert!(matches!(
        parse_block(&encoded, too_short),
        Err(Error::Limit {
            resource: "block bytes",
            actual,
            maximum
        }) if actual == second_end && maximum == second_end - 1
    ));
}

#[test]
fn begin_groups_header_obeys_one_two_exact_and_plus_one_block_limits() {
    let header = [0xA2, 0x08, 0x00];
    for maximum in [1, 2] {
        let limits = Limits::DEFAULT.with_block_bytes(maximum).unwrap();
        assert!(matches!(
            parse_block(&header, limits),
            Err(Error::Limit {
                resource: "block bytes",
                actual: 3,
                maximum: found
            }) if found == maximum
        ));
    }
    for maximum in [3, 4] {
        let limits = Limits::DEFAULT.with_block_bytes(maximum).unwrap();
        assert!(matches!(
            parse_block(&header, limits),
            Err(Error::Record { .. })
        ));
    }
}

#[test]
fn malformed_declared_formula_lengths_are_rejected() {
    let item = Sparkline::new(
        Location::new(0, 0).unwrap(),
        Some(Formula::with_limits(vec![0x23, 1, 0, 0, 0], vec![0xAA], Limits::DEFAULT).unwrap()),
    );
    let group = Group::new(SparklineType::Line, Colors::uniform(black()), vec![item]).unwrap();
    let encoded = encode_block(&Groups::new(vec![group]).unwrap(), Limits::DEFAULT).unwrap();
    let payload_offset = crate::raw::Records::new(&encoded)
        .map(|record| record.unwrap())
        .find(|record| record.kind() == crate::raw::kind::SPARKLINE)
        .map(|record| record.payload().as_ptr() as usize - encoded.as_ptr() as usize)
        .unwrap();

    let mut bad_cce = encoded.clone();
    bad_cce[payload_offset + 40..payload_offset + 44].copy_from_slice(&6u32.to_le_bytes());
    assert!(parse_block(&bad_cce, Limits::DEFAULT).is_err());

    let mut bad_cb = encoded;
    bad_cb[payload_offset + 44..payload_offset + 48].copy_from_slice(&2u32.to_le_bytes());
    assert!(parse_block(&bad_cb, Limits::DEFAULT).is_err());
}

#[test]
fn configured_limits_are_nonzero_and_formula_wire_bounded() {
    assert_eq!(Limits::DEFAULT.groups(), 230);
    assert_eq!(Limits::DEFAULT.per_group(), 230);
    assert_eq!(Limits::DEFAULT.total(), 52_900);
    assert_eq!(Limits::DEFAULT.formula_ancillary(), 64 * 1024);
    assert_eq!(Limits::DEFAULT.record_bytes(), 1024 * 1024);
    assert_eq!(Limits::DEFAULT.block_bytes(), 8 * 1024 * 1024);
    assert_eq!(Limits::DEFAULT.worksheet_bytes(), 512 * 1024 * 1024);
    assert!(Limits::DEFAULT.with_groups(0).is_err());
    assert!(Limits::DEFAULT.with_formula_tokens(16_385).is_err());
}

#[test]
fn worksheet_source_limit_is_nonzero_and_coherent_with_block_limit() {
    assert!(Limits::DEFAULT.with_worksheet_bytes(0).is_err());
    assert!(
        Limits::DEFAULT
            .with_worksheet_bytes(Limits::DEFAULT.block_bytes() - 1)
            .is_err()
    );
    assert!(
        Limits::DEFAULT
            .with_block_bytes(Limits::DEFAULT.worksheet_bytes() + 1)
            .is_err()
    );

    let exact = Limits::new(1, 1, 1, 1, 1, 1, 1, 1).unwrap();
    assert_eq!(exact.block_bytes(), 1);
    assert_eq!(exact.worksheet_bytes(), 1);

    let configured = Limits::DEFAULT
        .with_block_bytes(64)
        .unwrap()
        .with_worksheet_bytes(64)
        .unwrap();
    assert_eq!(configured.worksheet_bytes(), configured.block_bytes());
}
