//! Strict BIFF12 sparkline block codec.

use crate::raw::{self, Cursor, Record, Records, Writer, kind};
use litchi_sheet::sparkline::{AxisType, EmptyCells, SparklineType};

use super::model::{
    Axis, Color, Colors, Error, Formula, FrtState, Group, Groups, Limits, Location, Options,
    Result, Sparkline, allocation, check_limit, value,
};

const GROUP_FIXED_BYTES: usize = 2 + 8 * 8 + 3 * 8 + 4;
const LOCATION_HEADER_BYTES: usize = 4 + 4 + 4 + 4 + 16;

/// Parse one block beginning at `BrtBeginSparklineGroups` and return the
/// number of source bytes consumed through `BrtEndSparklineGroups`.
pub(crate) fn parse_block(input: &[u8], limits: Limits) -> Result<(Groups, usize)> {
    limits.validate()?;
    let raw_limits = raw::Limits::new(limits.record_bytes(), 1);
    let mut records = Records::with_limits(input, raw_limits);

    let begin = next(&mut records, "BrtBeginSparklineGroups")?;
    enforce_block_bytes(records.offset(), limits)?;
    expect_kind(
        begin,
        kind::BEGIN_SPARKLINE_GROUPS,
        "BrtBeginSparklineGroups",
    )?;
    empty(begin, "BrtBeginSparklineGroups")?;

    let mut groups = Vec::new();
    let mut total_items = 0usize;
    loop {
        let record = next(
            &mut records,
            "BrtBeginSparklineGroup or BrtEndSparklineGroups",
        )?;
        enforce_block_bytes(records.offset(), limits)?;
        if record.kind() == kind::END_SPARKLINE_GROUPS {
            empty(record, "BrtEndSparklineGroups")?;
            break;
        }
        expect_kind(
            record,
            kind::BEGIN_SPARKLINE_GROUP,
            "BrtBeginSparklineGroup",
        )?;
        check_limit("groups", groups.len().saturating_add(1), limits.groups())?;
        groups
            .try_reserve(1)
            .map_err(|_| allocation("group collection"))?;
        let group = parse_group(record.payload(), &mut records, limits, total_items)?;
        total_items = total_items
            .checked_add(group.sparklines().len())
            .ok_or(Error::Limit {
                resource: "total sparklines",
                actual: usize::MAX,
                maximum: limits.total(),
            })?;
        groups.push(group);
        enforce_block_bytes(records.offset(), limits)?;
    }
    let consumed = records.offset();
    enforce_block_bytes(consumed, limits)?;
    Ok((Groups::with_limits(groups, limits)?, consumed))
}

/// Validate and buffer one complete sparkline block before exposing bytes.
pub(crate) fn encode_block(value: &Groups, limits: Limits) -> Result<Vec<u8>> {
    value.validate(limits)?;
    let encoded_len = encoded_block_len(value)?;
    check_limit("block bytes", encoded_len, limits.block_bytes())?;

    let mut output = Vec::new();
    output
        .try_reserve_exact(encoded_len)
        .map_err(|_| allocation("encoded block"))?;
    let mut writer = Writer::with_limits(output, raw::Limits::new(limits.record_bytes(), 1));
    writer.write_record(kind::BEGIN_SPARKLINE_GROUPS, &[])?;
    for group in value.as_slice() {
        let payload = encode_group_payload(group, limits)?;
        writer.write_record(kind::BEGIN_SPARKLINE_GROUP, &payload)?;
        writer.write_record(kind::BEGIN_SPARKLINES, &[])?;
        for sparkline in group.sparklines() {
            let payload = encode_sparkline_payload(sparkline, limits)?;
            writer.write_record(kind::SPARKLINE, &payload)?;
        }
        writer.write_record(kind::END_SPARKLINES, &[])?;
        writer.write_record(kind::END_SPARKLINE_GROUP, &[])?;
    }
    writer.write_record(kind::END_SPARKLINE_GROUPS, &[])?;
    let output = writer.finish();
    debug_assert_eq!(output.len(), encoded_len);
    Ok(output)
}

fn parse_group(
    payload: &[u8],
    records: &mut Records<'_>,
    limits: Limits,
    prior_items: usize,
) -> Result<Group> {
    let (date_formula, header_bytes) = parse_formula_header(payload, limits)?;
    let mut cursor = Cursor::new(
        payload
            .get(header_bytes..)
            .ok_or_else(|| value("BrtBeginSparklineGroup", "FRTHeader length overflow"))?,
        "BrtBeginSparklineGroup",
    );
    let flags = cursor.read_u16()?;
    let date_axis = flags & 0x0001 != 0;
    let empty_cells = match (flags >> 1) & 0x03 {
        0 => EmptyCells::Zero,
        1 => EmptyCells::Gap,
        2 => EmptyCells::Span,
        other => {
            return Err(value(
                "BrtBeginSparklineGroup",
                format!("fShowEmptyCellAsZero {other} is reserved"),
            ));
        },
    };
    let mut options = Options::empty();
    for (wire, semantic) in [
        (0x0008, Options::MARKERS),
        (0x0010, Options::HIGH),
        (0x0020, Options::LOW),
        (0x0040, Options::FIRST),
        (0x0080, Options::LAST),
        (0x0100, Options::NEGATIVE),
        (0x0200, Options::AXIS),
        (0x0400, Options::DISPLAY_HIDDEN),
        (0x8000, Options::RIGHT_TO_LEFT),
    ] {
        if flags & wire != 0 {
            options.insert(semantic);
        }
    }
    let individual_max = flags & 0x0800 != 0;
    let individual_min = flags & 0x1000 != 0;
    let group_max = flags & 0x2000 != 0;
    let group_min = flags & 0x4000 != 0;
    if individual_max && group_max {
        return Err(value(
            "BrtBeginSparklineGroup",
            "individual and group maximum flags are mutually exclusive",
        ));
    }
    if individual_min && group_min {
        return Err(value(
            "BrtBeginSparklineGroup",
            "individual and group minimum flags are mutually exclusive",
        ));
    }

    let colors = Colors::new(
        read_color(&mut cursor)?,
        read_color(&mut cursor)?,
        read_color(&mut cursor)?,
        read_color(&mut cursor)?,
        read_color(&mut cursor)?,
        read_color(&mut cursor)?,
        read_color(&mut cursor)?,
        read_color(&mut cursor)?,
    );
    let manual_max = cursor.read_f64()?;
    let manual_min = cursor.read_f64()?;
    let line_weight = cursor.read_f64()?;
    let kind = match cursor.read_u32()? {
        0 => SparklineType::Line,
        1 => SparklineType::Column,
        2 => SparklineType::Stacked,
        other => {
            return Err(value(
                "BrtBeginSparklineGroup",
                format!("isltype {other} is outside 0..=2"),
            ));
        },
    };
    cursor.finish()?;

    let maximum = Axis::new(axis_kind(individual_max, group_max), manual_max)?;
    let minimum = Axis::new(axis_kind(individual_min, group_min), manual_min)?;

    let begin = next(records, "BrtBeginSparklines")?;
    enforce_block_bytes(records.offset(), limits)?;
    expect_kind(begin, kind::BEGIN_SPARKLINES, "BrtBeginSparklines")?;
    empty(begin, "BrtBeginSparklines")?;
    let mut items = Vec::new();
    loop {
        let record = next(records, "BrtSparkline or BrtEndSparklines")?;
        enforce_block_bytes(records.offset(), limits)?;
        if record.kind() == kind::END_SPARKLINES {
            empty(record, "BrtEndSparklines")?;
            break;
        }
        expect_kind(record, kind::SPARKLINE, "BrtSparkline")?;
        check_limit(
            "sparklines per group",
            items.len().saturating_add(1),
            limits.per_group(),
        )?;
        let aggregate = prior_items
            .checked_add(items.len())
            .and_then(|count| count.checked_add(1))
            .ok_or(Error::Limit {
                resource: "total sparklines",
                actual: usize::MAX,
                maximum: limits.total(),
            })?;
        check_limit("total sparklines", aggregate, limits.total())?;
        items
            .try_reserve(1)
            .map_err(|_| allocation("sparkline collection"))?;
        items.push(parse_sparkline(record.payload(), limits)?);
    }
    let end = next(records, "BrtEndSparklineGroup")?;
    enforce_block_bytes(records.offset(), limits)?;
    expect_kind(end, kind::END_SPARKLINE_GROUP, "BrtEndSparklineGroup")?;
    empty(end, "BrtEndSparklineGroup")?;

    Group::from_wire(
        kind,
        empty_cells,
        options,
        colors,
        minimum,
        maximum,
        line_weight,
        date_axis,
        date_formula,
        items,
        limits,
    )
}

fn parse_sparkline(payload: &[u8], limits: Limits) -> Result<Sparkline> {
    let mut cursor = Cursor::new(payload, "BrtSparkline");
    let flags = cursor.read_u32()?;
    if !matches!(flags, 0x02 | 0x06) {
        return Err(value(
            "BrtSparkline",
            format!("FRTHeader flags 0x{flags:08X} must be 0x02 or 0x06"),
        ));
    }
    if cursor.read_u32()? != 1 {
        return Err(value(
            "BrtSparkline",
            "FRTSqrefs must contain exactly one FRTSqref",
        ));
    }
    let state_flags = cursor.read_u32()?;
    if state_flags & 0x02 == 0 || state_flags & !0x0001_000f != 0 {
        return Err(value(
            "BrtSparkline",
            format!("invalid FRTSqref state flags 0x{state_flags:08X}"),
        ));
    }
    if cursor.read_i32()? != 1 {
        return Err(value(
            "BrtSparkline",
            "FRTSqref must contain exactly one UncheckedRfX",
        ));
    }
    let row_first = cursor.read_u32()?;
    let row_last = cursor.read_u32()?;
    let column_first = cursor.read_u32()?;
    let column_last = cursor.read_u32()?;
    if row_first != row_last || column_first != column_last {
        return Err(value(
            "BrtSparkline",
            "destination FRTSqref must identify one cell",
        ));
    }
    let location = Location::with_state(row_first, column_first, FrtState::from_wire(state_flags))?;
    let formula = if flags & 0x04 != 0 {
        Some(parse_formula_from_cursor(&mut cursor, limits)?)
    } else {
        None
    };
    cursor.finish()?;
    Ok(Sparkline::new(location, formula))
}

fn parse_formula_header(payload: &[u8], limits: Limits) -> Result<(Option<Formula>, usize)> {
    let mut cursor = Cursor::new(payload, "BrtBeginSparklineGroup");
    let flags = cursor.read_u32()?;
    if !matches!(flags, 0x00 | 0x04) {
        return Err(value(
            "BrtBeginSparklineGroup",
            format!("FRTHeader flags 0x{flags:08X} must be zero or 0x04"),
        ));
    }
    let formula = if flags == 0x04 {
        Some(parse_formula_from_cursor(&mut cursor, limits)?)
    } else {
        None
    };
    let consumed = payload.len().saturating_sub(cursor.remaining());
    Ok((formula, consumed))
}

fn parse_formula_from_cursor(cursor: &mut Cursor<'_>, limits: Limits) -> Result<Formula> {
    if cursor.read_u32()? != 1 {
        return Err(value("FRTHeader", "formula count must be exactly one"));
    }
    let formula_flags = cursor.read_u32()?;
    if formula_flags != 0x02 {
        return Err(value(
            "FRTFormula",
            format!("reserved flags must equal 0x00000002, found 0x{formula_flags:08X}"),
        ));
    }
    let cce = usize::try_from(cursor.read_u32()?)
        .map_err(|_| value("FRTFormula", "cce does not fit usize"))?;
    let cb = usize::try_from(cursor.read_u32()?)
        .map_err(|_| value("FRTFormula", "cb does not fit usize"))?;
    check_limit("formula token bytes", cce, limits.formula_tokens())?;
    check_limit("formula ancillary bytes", cb, limits.formula_ancillary())?;
    if cce == 0 || cce > 16_384 {
        return Err(value("FRTFormula", "cce must be in 1..=16384"));
    }
    let rgce = cursor.read_bytes(cce)?;
    let rgcb = cursor.read_bytes(cb)?;
    Formula::from_slices(rgce, rgcb, limits)
}

fn encode_group_payload(group: &Group, limits: Limits) -> Result<Vec<u8>> {
    let formula_bytes = group
        .date_formula()
        .map(encoded_formula_len)
        .transpose()?
        .unwrap_or(0);
    let payload_len = 4usize
        .checked_add(
            if group.date_formula().is_some() {
                4usize.checked_add(formula_bytes)
            } else {
                Some(0)
            }
            .ok_or_else(|| value("BrtBeginSparklineGroup", "formula length overflow"))?,
        )
        .and_then(|size| size.checked_add(GROUP_FIXED_BYTES))
        .ok_or_else(|| value("BrtBeginSparklineGroup", "payload length overflow"))?;
    check_limit("record payload bytes", payload_len, limits.record_bytes())?;
    let mut payload = Vec::new();
    payload
        .try_reserve_exact(payload_len)
        .map_err(|_| allocation("group payload"))?;
    payload.extend_from_slice(
        &if group.date_formula().is_some() {
            0x04u32
        } else {
            0u32
        }
        .to_le_bytes(),
    );
    if let Some(formula) = group.date_formula() {
        write_formula(&mut payload, formula)?;
    }
    let mut flags = match group.empty_cells() {
        EmptyCells::Zero => 0,
        EmptyCells::Gap => 1 << 1,
        EmptyCells::Span => 2 << 1,
    };
    if group.date_axis() {
        flags |= 0x0001;
    }
    for (semantic, wire) in [
        (Options::MARKERS, 0x0008),
        (Options::HIGH, 0x0010),
        (Options::LOW, 0x0020),
        (Options::FIRST, 0x0040),
        (Options::LAST, 0x0080),
        (Options::NEGATIVE, 0x0100),
        (Options::AXIS, 0x0200),
        (Options::DISPLAY_HIDDEN, 0x0400),
        (Options::RIGHT_TO_LEFT, 0x8000),
    ] {
        if group.options().contains(semantic) {
            flags |= wire;
        }
    }
    flags |= axis_flags(group.maximum().kind(), 0x0800, 0x2000);
    flags |= axis_flags(group.minimum().kind(), 0x1000, 0x4000);
    payload.extend_from_slice(&flags.to_le_bytes());
    for color in group.colors().ordered() {
        payload.extend_from_slice(&color.raw());
    }
    payload.extend_from_slice(&group.maximum().manual().to_le_bytes());
    payload.extend_from_slice(&group.minimum().manual().to_le_bytes());
    payload.extend_from_slice(&group.line_weight().to_le_bytes());
    payload.extend_from_slice(
        &match group.kind() {
            SparklineType::Line => 0u32,
            SparklineType::Column => 1u32,
            SparklineType::Stacked => 2u32,
        }
        .to_le_bytes(),
    );
    debug_assert_eq!(payload.len(), payload_len);
    Ok(payload)
}

fn encode_sparkline_payload(sparkline: &Sparkline, limits: Limits) -> Result<Vec<u8>> {
    let formula_bytes = sparkline
        .formula()
        .map(encoded_formula_len)
        .transpose()?
        .unwrap_or(0);
    let payload_len = LOCATION_HEADER_BYTES
        .checked_add(
            if sparkline.formula().is_some() {
                4usize.checked_add(formula_bytes)
            } else {
                Some(0)
            }
            .ok_or_else(|| value("BrtSparkline", "formula length overflow"))?,
        )
        .ok_or_else(|| value("BrtSparkline", "payload length overflow"))?;
    check_limit("record payload bytes", payload_len, limits.record_bytes())?;
    let mut payload = Vec::new();
    payload
        .try_reserve_exact(payload_len)
        .map_err(|_| allocation("sparkline payload"))?;
    payload.extend_from_slice(
        &if sparkline.formula().is_some() {
            0x06u32
        } else {
            0x02u32
        }
        .to_le_bytes(),
    );
    payload.extend_from_slice(&1u32.to_le_bytes());
    payload.extend_from_slice(&sparkline.location().state().wire().to_le_bytes());
    payload.extend_from_slice(&1u32.to_le_bytes());
    let location = sparkline.location();
    for coordinate in [
        location.row(),
        location.row(),
        location.column(),
        location.column(),
    ] {
        payload.extend_from_slice(&coordinate.to_le_bytes());
    }
    if let Some(formula) = sparkline.formula() {
        write_formula(&mut payload, formula)?;
    }
    debug_assert_eq!(payload.len(), payload_len);
    Ok(payload)
}

fn write_formula(output: &mut Vec<u8>, formula: &Formula) -> Result<()> {
    output.extend_from_slice(&1u32.to_le_bytes());
    output.extend_from_slice(&2u32.to_le_bytes());
    output.extend_from_slice(
        &u32::try_from(formula.tokens().len())
            .map_err(|_| value("FRTFormula", "cce overflows u32"))?
            .to_le_bytes(),
    );
    output.extend_from_slice(
        &u32::try_from(formula.ancillary().len())
            .map_err(|_| value("FRTFormula", "cb overflows u32"))?
            .to_le_bytes(),
    );
    output.extend_from_slice(formula.tokens());
    output.extend_from_slice(formula.ancillary());
    Ok(())
}

fn encoded_block_len(groups: &Groups) -> Result<usize> {
    let mut total = record_len(0)?;
    for group in groups.as_slice() {
        let formula = group
            .date_formula()
            .map(encoded_formula_len)
            .transpose()?
            .unwrap_or(0);
        let group_payload = 4usize
            .checked_add(
                if group.date_formula().is_some() {
                    4usize.checked_add(formula)
                } else {
                    Some(0)
                }
                .ok_or_else(|| value("sparkline block", "formula length overflow"))?,
            )
            .and_then(|size| size.checked_add(GROUP_FIXED_BYTES))
            .ok_or_else(|| value("sparkline block", "group length overflow"))?;
        total = checked_add(total, record_len(group_payload)?)?;
        total = checked_add(total, record_len(0)?)?;
        for item in group.sparklines() {
            let formula = item
                .formula()
                .map(encoded_formula_len)
                .transpose()?
                .unwrap_or(0);
            let payload = LOCATION_HEADER_BYTES
                .checked_add(
                    if item.formula().is_some() {
                        4usize.checked_add(formula)
                    } else {
                        Some(0)
                    }
                    .ok_or_else(|| value("sparkline block", "formula length overflow"))?,
                )
                .ok_or_else(|| value("sparkline block", "sparkline length overflow"))?;
            total = checked_add(total, record_len(payload)?)?;
        }
        total = checked_add(total, record_len(0)?)?;
        total = checked_add(total, record_len(0)?)?;
    }
    checked_add(total, record_len(0)?)
}

fn encoded_formula_len(formula: &Formula) -> Result<usize> {
    12usize
        .checked_add(formula.tokens().len())
        .and_then(|size| size.checked_add(formula.ancillary().len()))
        .ok_or_else(|| value("FRTFormula", "encoded length overflow"))
}

fn record_len(payload: usize) -> Result<usize> {
    let len_bytes = if payload < 1 << 7 {
        1
    } else if payload < 1 << 14 {
        2
    } else if payload < 1 << 21 {
        3
    } else {
        4
    };
    payload
        .checked_add(2 + len_bytes)
        .ok_or_else(|| value("sparkline block", "record length overflow"))
}

fn checked_add(left: usize, right: usize) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| value("sparkline block", "encoded length overflow"))
}

fn next<'a>(records: &mut Records<'a>, expected: &'static str) -> Result<Record<'a>> {
    match records.next() {
        Some(Ok(record)) => Ok(record),
        Some(Err(error)) => Err(error.into()),
        None => Err(Error::Record {
            expected,
            found: "end of input".to_string(),
        }),
    }
}

fn expect_kind(record: Record<'_>, expected: raw::Kind, name: &'static str) -> Result<()> {
    if record.kind() == expected {
        Ok(())
    } else {
        Err(Error::Record {
            expected: name,
            found: record.kind().to_string(),
        })
    }
}

fn empty(record: Record<'_>, name: &'static str) -> Result<()> {
    if record.is_empty() {
        Ok(())
    } else {
        Err(Error::Delimiter {
            record: name,
            length: record.len(),
        })
    }
}

fn enforce_block_bytes(actual: usize, limits: Limits) -> Result<()> {
    check_limit("block bytes", actual, limits.block_bytes())
}

fn read_color(cursor: &mut Cursor<'_>) -> Result<Color> {
    let bytes = cursor.read_bytes(8)?;
    let raw = <[u8; 8]>::try_from(bytes)
        .map_err(|_| value("BrtColor", "color payload must contain exactly eight bytes"))?;
    Color::from_raw(raw)
}

const fn axis_kind(individual: bool, group: bool) -> AxisType {
    if individual {
        AxisType::Individual
    } else if group {
        AxisType::Group
    } else {
        AxisType::Custom
    }
}

const fn axis_flags(kind: AxisType, individual: u16, group: u16) -> u16 {
    match kind {
        AxisType::Individual => individual,
        AxisType::Group => group,
        AxisType::Custom => 0,
    }
}
