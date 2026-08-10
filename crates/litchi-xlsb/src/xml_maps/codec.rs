#![allow(
    clippy::expect_used,
    reason = "legacy module confines extraction after an immediately preceding structural invariant check to this codec boundary"
)]

//! BIFF12 parser, canonical encoder, and linear source patcher.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use super::model::{
    CellReference, ColumnBinding, Limits, MappedTable, SingleCellBinding, XPath, XmlDataType,
};
use super::validation;
use crate::package::error::{Error, Result};
use crate::raw::{Cursor, Kind, Limits as RawLimits, Records, Writer, kind as rt};

const LIST_SINGLE_CELL: u32 = 1 << 1;
const XML_CAN_BE_SINGLE: u32 = 1 << 1;
const NO_DXF: u32 = u32::MAX;

#[derive(Debug, Clone, PartialEq, Eq)]
struct ColumnSlot {
    column_id: u32,
    xml_span: Option<(usize, usize)>,
    end_column_offset: usize,
    opaque_xml: bool,
    ignored_xml_flags: u32,
}

/// Exact source bytes and parsed ordinary-table XML bindings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableBindingsSource {
    source: Arc<[u8]>,
    value: MappedTable,
    slots: Vec<ColumnSlot>,
}

impl TableBindingsSource {
    #[must_use]
    pub const fn value(&self) -> &MappedTable {
        &self.value
    }

    #[must_use]
    pub fn source(&self) -> &[u8] {
        &self.source
    }
}

/// Exact source bytes and parsed single-cell bindings.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SingleCellsSource {
    source: Arc<[u8]>,
    value: Vec<SingleCellBinding>,
    connection_ids: Vec<u32>,
    has_opaque: bool,
}

impl SingleCellsSource {
    #[must_use]
    pub fn value(&self) -> &[SingleCellBinding] {
        &self.value
    }

    #[must_use]
    pub fn source(&self) -> &[u8] {
        &self.source
    }

    /// Inert `dwConnID` values aligned with [`Self::value`].
    #[allow(
        dead_code,
        reason = "retained for BIFF12 codec completeness and staged host integration"
    )]
    pub(crate) fn connection_ids(&self) -> &[u32] {
        &self.connection_ids
    }
}

#[derive(Clone, Copy)]
struct SeenRecord<'a> {
    kind: Kind,
    payload: &'a [u8],
    start: usize,
    end: usize,
}

fn records(data: &[u8], limits: Limits) -> Result<Vec<SeenRecord<'_>>> {
    if data.len() > limits.max_part_bytes {
        return Err(Error::InvalidLength {
            expected: limits.max_part_bytes,
            found: data.len(),
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve(data.len().min(limits.max_records).min(256))
        .map_err(|source| Error::Allocation {
            resource: "XML map BIFF12 record index",
            source,
        })?;
    let mut iterator = Records::with_limits(
        data,
        RawLimits::new(limits.max_part_bytes, limits.max_xpath_units),
    );
    while let Some(record) = iterator.next() {
        let record = record?;
        if output.len() >= limits.max_records {
            return Err(Error::InvalidLength {
                expected: limits.max_records,
                found: output.len().saturating_add(1),
            });
        }
        if output.len() == output.capacity() {
            let remaining = limits.max_records.saturating_sub(output.len());
            output
                .try_reserve(remaining.min(output.capacity().max(256)))
                .map_err(|source| Error::Allocation {
                    resource: "XML map BIFF12 record index",
                    source,
                })?;
        }
        output.push(SeenRecord {
            kind: record.kind(),
            payload: record.payload(),
            start: record.offset(),
            end: iterator.offset(),
        });
    }
    Ok(output)
}

fn count_opaque(
    record: SeenRecord<'_>,
    count: &mut usize,
    bytes: &mut usize,
    limits: Limits,
) -> Result<()> {
    *count = count.checked_add(1).ok_or(Error::CapacityOverflow {
        resource: "opaque XML map records",
    })?;
    *bytes = bytes
        .checked_add(record.payload.len())
        .ok_or(Error::CapacityOverflow {
            resource: "opaque XML map record payloads",
        })?;
    if *count > limits.max_opaque_records {
        return Err(Error::InvalidLength {
            expected: limits.max_opaque_records,
            found: *count,
        });
    }
    if *bytes > limits.max_opaque_bytes {
        return Err(Error::InvalidLength {
            expected: limits.max_opaque_bytes,
            found: *bytes,
        });
    }
    Ok(())
}

/// Parse XML column-property sidecars without changing the public Table model.
pub fn parse_table_bindings(data: &[u8], limits: Limits) -> Result<TableBindingsSource> {
    let (value, slots) = parse_table_bindings_impl(data, limits)?;
    Ok(TableBindingsSource {
        source: Arc::from(data),
        value,
        slots,
    })
}

/// Parse only semantic table bindings without retaining a second copy of the
/// complete BIFF part. Package readers use this while their OPC source already
/// owns the bytes; source-owning public parsers remain available for patching.
pub(crate) fn parse_table_bindings_value(data: &[u8], limits: Limits) -> Result<MappedTable> {
    parse_table_bindings_impl(data, limits).map(|(value, _)| value)
}

fn parse_table_bindings_impl(
    data: &[u8],
    limits: Limits,
) -> Result<(MappedTable, Vec<ColumnSlot>)> {
    let records = records(data, limits)?;
    let mut opaque_records = 0usize;
    let mut opaque_bytes = 0usize;
    let mut start = None;
    let mut wrapper_depth = 0usize;
    for (index, record) in records.iter().copied().enumerate() {
        if wrapper_depth != 0 {
            count_opaque(record, &mut opaque_records, &mut opaque_bytes, limits)?;
            if matches!(record.kind, rt::FRT_BEGIN | rt::AC_BEGIN) {
                wrapper_depth = wrapper_depth.saturating_add(1);
            } else if matches!(record.kind, rt::FRT_END | rt::AC_END) {
                wrapper_depth -= 1;
            }
            continue;
        }
        if matches!(record.kind, rt::FRT_BEGIN | rt::AC_BEGIN) {
            count_opaque(record, &mut opaque_records, &mut opaque_bytes, limits)?;
            wrapper_depth = 1;
        } else if record.kind == rt::BEGIN_LIST {
            start = Some(index);
            break;
        } else if known_table_structure(record.kind) {
            return Err(validation::invalid(
                "Table XML bindings",
                "known table record precedes BrtBeginList",
            ));
        } else {
            count_opaque(record, &mut opaque_records, &mut opaque_bytes, limits)?;
        }
    }
    if wrapper_depth != 0 {
        return Err(Error::UnexpectedEndOfStream(
            "future-record wrapper before BrtBeginList".to_string(),
        ));
    }
    let start = start.ok_or_else(|| Error::UnexpectedEndOfStream("BrtBeginList".to_string()))?;
    let (table_id, table_type) = parse_table_header(records[start].payload)?;
    let mut bindings = Vec::new();
    let mut slots = Vec::new();
    let mut current: Option<ColumnSlot> = None;
    let mut xml_start = None;
    let mut xml_binding = None;
    let mut in_columns = false;
    let mut saw_columns = false;
    let mut declared_columns = 0u32;
    let mut found_columns = 0u32;
    let mut saw_end_list = false;

    for record in records.iter().copied().skip(start + 1) {
        if saw_end_list {
            if known_table_structure(record.kind) {
                return Err(validation::invalid(
                    "Table XML bindings",
                    "known table record follows BrtEndList",
                ));
            }
            count_opaque(record, &mut opaque_records, &mut opaque_bytes, limits)?;
            continue;
        }
        match record.kind {
            rt::BEGIN_LIST_COLS
                if !in_columns && !saw_columns && current.is_none() && xml_start.is_none() =>
            {
                let mut cursor = Cursor::new(record.payload, "BrtBeginListCols");
                declared_columns = cursor.read_u32()?;
                cursor.finish()?;
                in_columns = true;
                saw_columns = true;
            },
            rt::BEGIN_LIST_COL if in_columns && current.is_none() => {
                found_columns = found_columns
                    .checked_add(1)
                    .ok_or(Error::CapacityOverflow {
                        resource: "table column count",
                    })?;
                current = Some(ColumnSlot {
                    column_id: parse_column_id(record.payload)?,
                    xml_span: None,
                    end_column_offset: 0,
                    opaque_xml: false,
                    ignored_xml_flags: 0,
                });
            },
            rt::BEGIN_LIST_XML_CPR
                if in_columns
                    && current.is_some()
                    && xml_start.is_none()
                    && current.as_ref().expect("checked").xml_span.is_none() =>
            {
                let column_id = current.as_ref().expect("checked").column_id;
                let (binding, ignored_flags) =
                    parse_column_binding_with_flags(column_id, record.payload, limits)?;
                current.as_mut().expect("checked").ignored_xml_flags = ignored_flags;
                xml_binding = Some(binding);
                xml_start = Some(record.start);
            },
            rt::END_LIST_XML_CPR if current.is_some() && xml_start.is_some() => {
                let slot = current.as_mut().expect("checked");
                slot.xml_span = Some((xml_start.take().expect("checked"), record.end));
                bindings.push(xml_binding.take().expect("checked"));
            },
            rt::END_LIST_COL if current.is_some() && xml_start.is_none() => {
                let mut slot = current.take().expect("checked");
                slot.end_column_offset = record.start;
                slots.push(slot);
            },
            rt::END_LIST_COLS if in_columns && current.is_none() && xml_start.is_none() => {
                if found_columns != declared_columns {
                    return Err(validation::invalid(
                        "BrtBeginListCols",
                        format!("declared {declared_columns} columns, found {found_columns}"),
                    ));
                }
                in_columns = false;
            },
            rt::END_LIST => {
                if !saw_columns || in_columns || current.is_some() || xml_start.is_some() {
                    return Err(validation::invalid(
                        "Table XML bindings",
                        "unclosed collection",
                    ));
                }
                saw_end_list = true;
            },
            _ if xml_start.is_some() => {
                current.as_mut().expect("checked").opaque_xml = true;
                count_opaque(record, &mut opaque_records, &mut opaque_bytes, limits)?;
            },
            _ if known_table_structure(record.kind) => {
                return Err(validation::invalid(
                    "Table XML bindings",
                    format!("misplaced known record {}", record.kind.get()),
                ));
            },
            _ => {
                count_opaque(record, &mut opaque_records, &mut opaque_bytes, limits)?;
            },
        }
    }
    if !saw_end_list {
        return Err(Error::UnexpectedEndOfStream("BrtEndList".to_string()));
    }
    if !bindings.is_empty() && table_type != 2 {
        return Err(validation::invalid(
            "mapped table type",
            format!("expected LTXML, found {table_type}"),
        ));
    }
    let value = MappedTable::new_with_limits(table_id, bindings, limits)?;
    Ok((value, slots))
}

fn known_table_structure(kind: Kind) -> bool {
    matches!(
        kind,
        rt::BEGIN_LIST
            | rt::END_LIST
            | rt::BEGIN_LIST_COLS
            | rt::END_LIST_COLS
            | rt::BEGIN_LIST_COL
            | rt::END_LIST_COL
            | rt::BEGIN_LIST_XML_CPR
            | rt::END_LIST_XML_CPR
    )
}

fn parse_table_header(payload: &[u8]) -> Result<(u32, u32)> {
    let mut cursor = Cursor::new(payload, "BrtBeginList");
    cursor.skip(16)?;
    let table_type = cursor.read_u32()?;
    if !matches!(table_type, 0 | 2 | 3) {
        return Err(validation::invalid(
            "table type",
            format!("unknown value {table_type}"),
        ));
    }
    let id = cursor.read_u32()?;
    validation::list_id(id, "table ID")?;
    Ok((id, table_type))
}

fn parse_column_id(payload: &[u8]) -> Result<u32> {
    let mut cursor = Cursor::new(payload, "BrtBeginListCol");
    let id = cursor.read_u32()?;
    validation::nonzero_id(id, "table column ID")?;
    Ok(id)
}

fn parse_column_binding_with_flags(
    column_id: u32,
    payload: &[u8],
    limits: Limits,
) -> Result<(ColumnBinding, u32)> {
    let mut cursor = Cursor::with_limits(
        payload,
        "BrtBeginListXmlCPr",
        RawLimits::new(limits.max_part_bytes, limits.max_xpath_units),
    );
    let map_id = cursor.read_u32()?;
    let flags = cursor.read_u32()?;
    let data_type = XmlDataType::new(cursor.read_u32()?)?;
    let xpath = XPath::new_with_limits(cursor.read_wide_string()?, limits)?;
    cursor.finish()?;
    let binding = ColumnBinding::new(
        column_id,
        map_id,
        data_type,
        xpath,
        flags & XML_CAN_BE_SINGLE != 0,
    )?;
    Ok((binding, flags & !XML_CAN_BE_SINGLE))
}

/// Encode one complete `BrtBeginListXmlCPr`/`BrtEndListXmlCPr` pair.
pub fn serialize_column_binding(binding: &ColumnBinding, limits: Limits) -> Result<Vec<u8>> {
    serialize_column_binding_with_flags(binding, 0, limits)
}

fn serialize_column_binding_with_flags(
    binding: &ColumnBinding,
    ignored_flags: u32,
    limits: Limits,
) -> Result<Vec<u8>> {
    validation::xpath(binding.xpath().as_str(), limits.max_xpath_units)?;
    let units = binding.xpath().as_str().encode_utf16().count();
    let payload_len = 16usize
        .checked_add(units.checked_mul(2).ok_or(Error::CapacityOverflow {
            resource: "XML column binding payload",
        })?)
        .ok_or(Error::CapacityOverflow {
            resource: "XML column binding payload",
        })?;
    let mut payload = Vec::new();
    payload
        .try_reserve_exact(payload_len)
        .map_err(|source| Error::Allocation {
            resource: "XML column binding payload",
            source,
        })?;
    {
        let mut writer = Writer::with_limits(
            &mut payload,
            RawLimits::new(limits.max_part_bytes, limits.max_xpath_units),
        );
        writer.write_u32(binding.map_id())?;
        writer.write_u32(
            ignored_flags
                | if binding.can_be_single() {
                    XML_CAN_BE_SINGLE
                } else {
                    0
                },
        )?;
        writer.write_u32(binding.data_type().get())?;
        writer.write_wide_string(binding.xpath().as_str())?;
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(payload.len().saturating_add(8))
        .map_err(|source| Error::Allocation {
            resource: "XML column binding records",
            source,
        })?;
    let mut writer = Writer::with_limits(
        &mut output,
        RawLimits::new(limits.max_part_bytes, limits.max_xpath_units),
    );
    writer.write_record(rt::BEGIN_LIST_XML_CPR, &payload)?;
    writer.write_record(rt::END_LIST_XML_CPR, &[])?;
    enforce_output_limit(output, limits)
}

/// Encode the mapped-column sidecars in semantic column order.
pub fn serialize_table_bindings(value: &MappedTable, limits: Limits) -> Result<Vec<u8>> {
    validation::mapped_table(value, limits)?;
    let mut output = Vec::new();
    for binding in value.columns() {
        let encoded = serialize_column_binding(binding, limits)?;
        output
            .try_reserve(encoded.len())
            .map_err(|source| Error::Allocation {
                resource: "mapped table bindings",
                source,
            })?;
        output.extend_from_slice(&encoded);
        if output.len() > limits.max_part_bytes {
            return Err(Error::InvalidLength {
                expected: limits.max_part_bytes,
                found: output.len(),
            });
        }
    }
    Ok(output)
}

/// Patch supported table bindings while copying every unrelated source byte.
pub fn patch_table_bindings(
    source: &TableBindingsSource,
    value: &MappedTable,
    limits: Limits,
) -> Result<Vec<u8>> {
    validation::mapped_table(value, limits)?;
    if source.value == *value {
        return clone_bytes(&source.source, "mapped table no-op patch");
    }
    if source.value.table_id() != value.table_id() {
        return Err(validation::invalid(
            "mapped table patch",
            "table ID changed",
        ));
    }
    let slot_ids: HashSet<u32> = source.slots.iter().map(|slot| slot.column_id).collect();
    if let Some(binding) = value
        .columns()
        .iter()
        .find(|binding| !slot_ids.contains(&binding.column_id()))
    {
        return Err(validation::invalid(
            "mapped table patch",
            format!("column {} is absent from the source", binding.column_id()),
        ));
    }
    let target: HashMap<u32, &ColumnBinding> = value
        .columns()
        .iter()
        .map(|binding| (binding.column_id(), binding))
        .collect();
    let old: HashMap<u32, &ColumnBinding> = source
        .value
        .columns()
        .iter()
        .map(|binding| (binding.column_id(), binding))
        .collect();
    let mut edits = Vec::new();
    for slot in &source.slots {
        let before = old.get(&slot.column_id).copied();
        let after = target.get(&slot.column_id).copied();
        if before == after {
            continue;
        }
        if slot.opaque_xml {
            return Err(Error::UnsupportedFeature(format!(
                "cannot losslessly edit opaque BrtBeginListXmlCPr for column {}",
                slot.column_id
            )));
        }
        let replacement = match after {
            Some(binding) => {
                serialize_column_binding_with_flags(binding, slot.ignored_xml_flags, limits)?
            },
            None => Vec::new(),
        };
        let (start, end) = slot
            .xml_span
            .unwrap_or((slot.end_column_offset, slot.end_column_offset));
        edits.push((start, end, replacement));
    }
    apply_edits(&source.source, edits, limits, "mapped table patch")
}

/// Parse and patch a complete ordinary Table part in one bounded operation.
pub fn apply_table_bindings(
    table_part: &[u8],
    value: &MappedTable,
    limits: Limits,
) -> Result<Vec<u8>> {
    let source = parse_table_bindings(table_part, limits)?;
    patch_table_bindings(&source, value, limits)
}

/// Parse the known minimum Single Cell Tables grammar while retaining or
/// refusing edits to bounded opaque/FRT content.
pub fn parse_single_cells(data: &[u8], limits: Limits) -> Result<SingleCellsSource> {
    let (value, connection_ids, has_opaque) = parse_single_cells_impl(data, limits)?;
    Ok(SingleCellsSource {
        source: Arc::from(data),
        value,
        connection_ids,
        has_opaque,
    })
}

/// Parse semantic single-cell bindings and aligned inert connection IDs
/// without retaining a second copy of the complete BIFF part.
pub(crate) fn parse_single_cells_value_with_connection_ids(
    data: &[u8],
    limits: Limits,
) -> Result<(Vec<SingleCellBinding>, Vec<u32>)> {
    parse_single_cells_impl(data, limits).map(|(value, connection_ids, _)| (value, connection_ids))
}

fn parse_single_cells_impl(
    data: &[u8],
    limits: Limits,
) -> Result<(Vec<SingleCellBinding>, Vec<u32>, bool)> {
    let all = records(data, limits)?;
    let mut significant = Vec::new();
    significant
        .try_reserve(all.len().min(256))
        .map_err(|source| Error::Allocation {
            resource: "Single Cell Tables semantic record index",
            source,
        })?;
    let mut has_opaque = false;
    let mut wrapper_depth = 0usize;
    let mut opaque_records = 0usize;
    let mut opaque_bytes = 0usize;
    let known: HashSet<Kind> = [
        rt::BEGIN_SINGLE_CELLS,
        rt::END_SINGLE_CELLS,
        rt::BEGIN_LIST,
        rt::END_LIST,
        rt::BEGIN_LIST_COLS,
        rt::END_LIST_COLS,
        rt::BEGIN_LIST_COL,
        rt::END_LIST_COL,
        rt::BEGIN_LIST_XML_CPR,
        rt::END_LIST_XML_CPR,
    ]
    .into_iter()
    .collect();
    for record in all {
        if wrapper_depth != 0 {
            has_opaque = true;
            count_opaque(record, &mut opaque_records, &mut opaque_bytes, limits)?;
            if matches!(record.kind, rt::FRT_BEGIN | rt::AC_BEGIN) {
                wrapper_depth = wrapper_depth.saturating_add(1);
            } else if matches!(record.kind, rt::FRT_END | rt::AC_END) {
                wrapper_depth -= 1;
            }
            continue;
        }
        if matches!(record.kind, rt::FRT_BEGIN | rt::AC_BEGIN) {
            has_opaque = true;
            count_opaque(record, &mut opaque_records, &mut opaque_bytes, limits)?;
            wrapper_depth = 1;
        } else if known.contains(&record.kind) {
            significant.push(record);
        } else {
            has_opaque = true;
            count_opaque(record, &mut opaque_records, &mut opaque_bytes, limits)?;
        }
    }
    if wrapper_depth != 0 {
        return Err(Error::UnexpectedEndOfStream(
            "future-record wrapper in Single Cell Tables".to_string(),
        ));
    }
    let mut position = 0usize;
    expect_kind(&significant, &mut position, rt::BEGIN_SINGLE_CELLS, true)?;
    let mut values = Vec::new();
    let mut connection_ids = Vec::new();
    while significant
        .get(position)
        .is_some_and(|record| record.kind == rt::BEGIN_LIST)
    {
        if values.len() >= limits.max_bindings {
            return Err(Error::InvalidLength {
                expected: limits.max_bindings,
                found: values.len().saturating_add(1),
            });
        }
        let list = expect_kind(&significant, &mut position, rt::BEGIN_LIST, false)?;
        let (table_id, cell, connection_id, list_noncanonical) = parse_single_list(list.payload)?;
        has_opaque |= list_noncanonical;
        let cols = expect_kind(&significant, &mut position, rt::BEGIN_LIST_COLS, false)?;
        exact_u32(cols.payload, "BrtBeginListCols", 1)?;
        let col = expect_kind(&significant, &mut position, rt::BEGIN_LIST_COL, false)?;
        let (column_id, column_noncanonical) = parse_single_column(col.payload)?;
        has_opaque |= column_noncanonical;
        let xml = expect_kind(&significant, &mut position, rt::BEGIN_LIST_XML_CPR, false)?;
        let (binding, ignored_xml_flags) =
            parse_column_binding_with_flags(column_id, xml.payload, limits)?;
        has_opaque |= ignored_xml_flags != 0;
        if !binding.can_be_single() {
            return Err(validation::invalid(
                "single-cell XML binding",
                "fCanBeSingle is zero",
            ));
        }
        expect_kind(&significant, &mut position, rt::END_LIST_XML_CPR, true)?;
        expect_kind(&significant, &mut position, rt::END_LIST_COL, true)?;
        expect_kind(&significant, &mut position, rt::END_LIST_COLS, true)?;
        expect_kind(&significant, &mut position, rt::END_LIST, true)?;
        values.push(SingleCellBinding::new(
            table_id,
            column_id,
            cell,
            binding.map_id(),
            binding.data_type(),
            binding.xpath().clone(),
        )?);
        connection_ids.push(connection_id);
    }
    expect_kind(&significant, &mut position, rt::END_SINGLE_CELLS, true)?;
    if position != significant.len() {
        return Err(validation::invalid(
            "Single Cell Tables",
            "records follow BrtEndSingleCells",
        ));
    }
    validation::single_cells(&values, limits)?;
    Ok((values, connection_ids, has_opaque))
}

fn expect_kind<'a>(
    records: &'a [SeenRecord<'a>],
    position: &mut usize,
    kind: Kind,
    empty: bool,
) -> Result<SeenRecord<'a>> {
    let record = records
        .get(*position)
        .copied()
        .ok_or_else(|| Error::UnexpectedEndOfStream(format!("record {}", kind.get())))?;
    if record.kind != kind {
        return Err(Error::UnexpectedRecord {
            expected: kind.get(),
            found: record.kind.get(),
        });
    }
    if empty && !record.payload.is_empty() {
        return Err(Error::InvalidLength {
            expected: 0,
            found: record.payload.len(),
        });
    }
    *position += 1;
    Ok(record)
}

fn parse_single_list(payload: &[u8]) -> Result<(u32, CellReference, u32, bool)> {
    let mut cursor = Cursor::new(payload, "single-cell BrtBeginList");
    let first_row = cursor.read_u32()?;
    let last_row = cursor.read_u32()?;
    let first_column = cursor.read_u32()?;
    let last_column = cursor.read_u32()?;
    if first_row != last_row || first_column != last_column {
        return Err(validation::invalid(
            "single-cell BrtBeginList",
            "rfxList does not occupy exactly one cell",
        ));
    }
    exact_cursor_u32(&mut cursor, "table type", 2)?;
    let table_id = cursor.read_u32()?;
    validation::list_id(table_id, "single-cell table ID")?;
    exact_cursor_u32(&mut cursor, "header row count", 0)?;
    exact_cursor_u32(&mut cursor, "totals row count", 0)?;
    let flags = cursor.read_u32()?;
    if flags & LIST_SINGLE_CELL == 0 {
        return Err(validation::invalid(
            "single-cell flags",
            "fSingleCell is zero",
        ));
    }
    for _ in 0..6 {
        exact_cursor_u32(&mut cursor, "single-cell DXF ID", NO_DXF)?;
    }
    let connection_id = cursor.read_u32()?;
    for _ in 0..6 {
        if cursor.read_nullable_wide_string()?.is_some() {
            return Err(validation::invalid(
                "single-cell BrtBeginList",
                "non-NULL string",
            ));
        }
    }
    cursor.finish()?;
    Ok((
        table_id,
        CellReference::new(first_row, first_column)?,
        connection_id,
        flags != LIST_SINGLE_CELL || connection_id != 0,
    ))
}

fn parse_single_column(payload: &[u8]) -> Result<(u32, bool)> {
    let mut cursor = Cursor::new(payload, "single-cell BrtBeginListCol");
    let column_id = cursor.read_u32()?;
    validation::nonzero_id(column_id, "single-cell column ID")?;
    let totals_row_function = cursor.read_u32()?;
    if totals_row_function > 9 {
        return Err(validation::invalid(
            "totals-row function",
            format!("unknown value {totals_row_function}"),
        ));
    }
    for _ in 0..3 {
        exact_cursor_u32(&mut cursor, "single-cell column DXF ID", NO_DXF)?;
    }
    exact_cursor_u32(&mut cursor, "query-table field ID", 0)?;
    for context in ["name", "caption"] {
        if cursor.read_nullable_wide_string()?.is_some() {
            return Err(validation::invalid(
                "single-cell BrtBeginListCol",
                format!("non-NULL {context}"),
            ));
        }
    }
    let total = cursor.read_nullable_wide_string()?;
    if total
        .as_deref()
        .is_some_and(|value| value.encode_utf16().count() > 8_189)
    {
        return Err(validation::invalid(
            "single-cell BrtBeginListCol",
            "total label exceeds 8189 UTF-16 units",
        ));
    }
    if totals_row_function == 9 && total.is_some() {
        return Err(validation::invalid(
            "single-cell BrtBeginListCol",
            "custom totals-row function has a non-NULL total label",
        ));
    }
    for context in ["header style", "insert-row style", "totals style"] {
        if cursor.read_nullable_wide_string()?.is_some() {
            return Err(validation::invalid(
                "single-cell BrtBeginListCol",
                format!("non-NULL {context}"),
            ));
        }
    }
    cursor.finish()?;
    Ok((column_id, totals_row_function != 0 || total.is_some()))
}

fn exact_cursor_u32(cursor: &mut Cursor<'_>, context: &'static str, expected: u32) -> Result<()> {
    let found = cursor.read_u32()?;
    if found != expected {
        Err(validation::invalid(
            context,
            format!("expected 0x{expected:08X}, found 0x{found:08X}"),
        ))
    } else {
        Ok(())
    }
}

fn exact_u32(payload: &[u8], context: &'static str, expected: u32) -> Result<()> {
    let mut cursor = Cursor::new(payload, context);
    exact_cursor_u32(&mut cursor, context, expected)?;
    cursor.finish()?;
    Ok(())
}

/// Serialize a canonical Single Cell Tables part.
pub fn serialize_single_cells(values: &[SingleCellBinding], limits: Limits) -> Result<Vec<u8>> {
    validation::single_cells(values, limits)?;
    let mut output = Vec::new();
    output
        .try_reserve(values.len().saturating_mul(192).saturating_add(8))
        .map_err(|source| Error::Allocation {
            resource: "Single Cell Tables part",
            source,
        })?;
    let raw_limits = RawLimits::new(limits.max_part_bytes, limits.max_xpath_units);
    let mut writer = Writer::with_limits(&mut output, raw_limits);
    writer.write_record(rt::BEGIN_SINGLE_CELLS, &[])?;
    for value in values {
        let list = single_list_payload(value, limits)?;
        writer.write_record(rt::BEGIN_LIST, &list)?;
        writer.write_record(rt::BEGIN_LIST_COLS, &1u32.to_le_bytes())?;
        let column = single_column_payload(value, limits)?;
        writer.write_record(rt::BEGIN_LIST_COL, &column)?;
        let xml = column_binding_payload(value.column_binding(), limits)?;
        writer.write_record(rt::BEGIN_LIST_XML_CPR, &xml)?;
        writer.write_record(rt::END_LIST_XML_CPR, &[])?;
        writer.write_record(rt::END_LIST_COL, &[])?;
        writer.write_record(rt::END_LIST_COLS, &[])?;
        writer.write_record(rt::END_LIST, &[])?;
    }
    writer.write_record(rt::END_SINGLE_CELLS, &[])?;
    drop(writer);
    enforce_output_limit(output, limits)
}

fn single_list_payload(value: &SingleCellBinding, limits: Limits) -> Result<Vec<u8>> {
    let mut payload = checked_vec(92, "single-cell BrtBeginList")?;
    let mut writer = Writer::with_limits(
        &mut payload,
        RawLimits::new(limits.max_part_bytes, limits.max_xpath_units),
    );
    writer.write_u32(value.cell().row())?;
    writer.write_u32(value.cell().row())?;
    writer.write_u32(value.cell().column())?;
    writer.write_u32(value.cell().column())?;
    writer.write_u32(2)?;
    writer.write_u32(value.table_id())?;
    writer.write_u32(0)?;
    writer.write_u32(0)?;
    writer.write_u32(LIST_SINGLE_CELL)?;
    for _ in 0..6 {
        writer.write_u32(NO_DXF)?;
    }
    writer.write_u32(0)?;
    for _ in 0..6 {
        writer.write_u32(u32::MAX)?;
    }
    Ok(payload)
}

fn single_column_payload(value: &SingleCellBinding, limits: Limits) -> Result<Vec<u8>> {
    let mut payload = checked_vec(48, "single-cell BrtBeginListCol")?;
    let mut writer = Writer::with_limits(
        &mut payload,
        RawLimits::new(limits.max_part_bytes, limits.max_xpath_units),
    );
    writer.write_u32(value.column_id())?;
    writer.write_u32(0)?;
    for _ in 0..3 {
        writer.write_u32(NO_DXF)?;
    }
    writer.write_u32(0)?;
    for _ in 0..6 {
        writer.write_u32(u32::MAX)?;
    }
    Ok(payload)
}

fn column_binding_payload(binding: &ColumnBinding, limits: Limits) -> Result<Vec<u8>> {
    let records = serialize_column_binding(binding, limits)?;
    let mut parsed = Records::new(&records);
    let first = parsed
        .next()
        .transpose()?
        .ok_or_else(|| Error::UnexpectedEndOfStream("BrtBeginListXmlCPr".to_string()))?;
    clone_bytes(first.payload(), "XML column binding payload")
}

/// Patch an existing source or refuse to destroy opaque/FRT records.
pub fn patch_single_cells(
    source: &SingleCellsSource,
    values: &[SingleCellBinding],
    limits: Limits,
) -> Result<Vec<u8>> {
    validation::single_cells(values, limits)?;
    if source.value == values {
        return clone_bytes(&source.source, "single-cell no-op patch");
    }
    if source.has_opaque {
        return Err(Error::UnsupportedFeature(
            "cannot losslessly edit Single Cell Tables containing opaque or FRT records"
                .to_string(),
        ));
    }
    serialize_single_cells(values, limits)
}

fn checked_vec(capacity: usize, resource: &'static str) -> Result<Vec<u8>> {
    let mut value = Vec::new();
    value
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation { resource, source })?;
    Ok(value)
}

fn clone_bytes(data: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = checked_vec(data.len(), resource)?;
    output.extend_from_slice(data);
    Ok(output)
}

fn enforce_output_limit(output: Vec<u8>, limits: Limits) -> Result<Vec<u8>> {
    if output.len() > limits.max_part_bytes {
        Err(Error::InvalidLength {
            expected: limits.max_part_bytes,
            found: output.len(),
        })
    } else {
        Ok(output)
    }
}

fn apply_edits(
    source: &[u8],
    mut edits: Vec<(usize, usize, Vec<u8>)>,
    limits: Limits,
    resource: &'static str,
) -> Result<Vec<u8>> {
    edits.sort_by_key(|edit| edit.0);
    let mut planned = source.len();
    let mut previous = 0usize;
    for (start, end, replacement) in &edits {
        if *start < previous || *start > *end || *end > source.len() {
            return Err(validation::invalid(
                resource,
                "overlapping or invalid source edits",
            ));
        }
        planned = planned
            .checked_sub(end - start)
            .and_then(|value| value.checked_add(replacement.len()))
            .ok_or(Error::CapacityOverflow { resource })?;
        previous = *end;
    }
    if planned > limits.max_part_bytes {
        return Err(Error::InvalidLength {
            expected: limits.max_part_bytes,
            found: planned,
        });
    }
    let mut output = checked_vec(planned, resource)?;
    let mut cursor = 0usize;
    for (start, end, replacement) in edits {
        output.extend_from_slice(&source[cursor..start]);
        output.extend_from_slice(&replacement);
        cursor = end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}
