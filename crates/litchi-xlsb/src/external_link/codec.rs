//! BIFF12 parsing for XLSB external-link streams.
//!
//! This module validates the checked-in `BrtSupBook` record grammar and
//! constructs the inert semantic model from `model.rs`.

use super::model::validate_defined_name;
use super::*;

/// Parse one complete XLSB External Link part stream.
pub fn parse_external_link(data: &[u8]) -> Result<Parsed> {
    if data.len() > MAX_LINK_PART_BYTES {
        return Err(Error::InvalidLength {
            expected: MAX_LINK_PART_BYTES,
            found: data.len(),
        });
    }
    let limits = crate::raw::Limits::new(MAX_LINK_PART_BYTES, MAX_WIDE_STRING_UNITS);
    let mut link_type = None;
    let mut target_key = String::new();
    let mut target_detail = String::new();
    let mut sheet_names = Vec::new();
    let mut workbook_entries = Vec::new();
    let mut dde_entries = Vec::new();
    let mut ole_entries = Vec::new();
    let mut saw_sup_tabs = false;
    // 0 = outside a name, 1 = expect formula, 2 = expect bits,
    // 3 = expect end/value start, 4 = inside a cached matrix.
    let mut sup_name_state = 0u8;
    let mut current_name = None;
    let mut current_formula = None;
    let mut current_bits = None;
    let mut current_cache = None;
    let mut cache_dimensions = None;
    let mut cache_values = Vec::new();
    let mut saw_end = false;

    for record in crate::raw::Records::with_limits(data, limits) {
        let record = record?;
        if saw_end {
            return Err(invalid("external link has records after BrtEndSupBook"));
        }
        if link_type.is_none() && record.kind() != crate::raw::kind::BEGIN_SUP_BOOK {
            return Err(invalid("external link does not start with BrtBeginSupBook"));
        }
        match record.kind() {
            crate::raw::kind::BEGIN_SUP_BOOK => {
                if link_type.is_some() || record.payload().len() < 10 {
                    return Err(invalid("invalid BrtBeginSupBook framing"));
                }
                let mut cursor =
                    crate::raw::Cursor::with_limits(record.payload(), "BrtBeginSupBook", limits);
                let kind = cursor.read_u16()?;
                let first = cursor.read_wide_string()?;
                let second = if kind == EXTERNAL_REFERENCE_WORKBOOK {
                    cursor.read_nullable_wide_string()?
                } else {
                    Some(cursor.read_wide_string()?)
                };
                cursor.finish()?;
                if kind > EXTERNAL_REFERENCE_OLE || first.is_empty() {
                    return Err(invalid("invalid BrtBeginSupBook payload"));
                }
                if kind == EXTERNAL_REFERENCE_WORKBOOK && second.is_some() {
                    return Err(invalid(
                        "external workbook BrtBeginSupBook string2 is not NULL",
                    ));
                }
                link_type = Some(kind);
                target_key = first;
                target_detail = second.unwrap_or_default();
            },
            crate::raw::kind::SUP_TABS => {
                if link_type != Some(EXTERNAL_REFERENCE_WORKBOOK)
                    || saw_sup_tabs
                    || sup_name_state != 0
                {
                    return Err(invalid("unexpected BrtSupTabs"));
                }
                sheet_names = parse_external_sheet_names(record.payload(), limits)?;
                saw_sup_tabs = true;
            },
            crate::raw::kind::SUP_NAME_START => {
                let kind =
                    link_type.ok_or_else(|| invalid("BrtSupNameStart precedes BrtBeginSupBook"))?;
                if sup_name_state != 0 || (kind == EXTERNAL_REFERENCE_WORKBOOK && !saw_sup_tabs) {
                    return Err(invalid("unexpected BrtSupNameStart"));
                }
                let mut cursor =
                    crate::raw::Cursor::with_limits(record.payload(), "BrtSupNameStart", limits);
                let name = cursor.read_wide_string()?;
                cursor.finish()?;
                validate_defined_name(&name)?;
                current_name = Some(name);
                sup_name_state = if kind == EXTERNAL_REFERENCE_WORKBOOK {
                    1
                } else {
                    2
                };
            },
            crate::raw::kind::SUP_NAME_FORMULA => {
                if link_type != Some(EXTERNAL_REFERENCE_WORKBOOK) || sup_name_state != 1 {
                    return Err(invalid("unexpected BrtSupNameFmla"));
                }
                let mut cursor =
                    crate::raw::Cursor::with_limits(record.payload(), "BrtSupNameFmla", limits);
                if record.payload().len() < 4 {
                    return Err(Error::InvalidLength {
                        expected: 4,
                        found: record.payload().len(),
                    });
                }
                let formula_len = usize::try_from(cursor.read_u32()?)
                    .map_err(|_| invalid("BrtSupNameFmla size overflow"))?;
                let formula = cursor.read_bytes(formula_len)?.to_vec();
                cursor.finish()?;
                current_formula = if formula.is_empty() {
                    None
                } else {
                    Some(NameFormula::from_tokens(formula)?)
                };
                sup_name_state = 2;
            },
            crate::raw::kind::SUP_NAME_BITS => {
                if sup_name_state != 2 || record.payload().len() != 7 {
                    return Err(invalid("unexpected BrtSupNameBits"));
                }
                let mut bits = [0u8; 7];
                bits.copy_from_slice(record.payload());
                validate_external_name_bits(
                    link_type.expect("external link kind is present"),
                    &bits,
                )?;
                current_bits = Some(bits);
                sup_name_state = 3;
            },
            crate::raw::kind::SUP_NAME_VALUE_START => {
                if !matches!(
                    link_type,
                    Some(EXTERNAL_REFERENCE_DDE | EXTERNAL_REFERENCE_OLE)
                ) || sup_name_state != 3
                    || record.payload().len() != 8
                    || current_cache.is_some()
                {
                    return Err(invalid("unexpected BrtSupNameValueStart"));
                }
                let mut cursor = crate::raw::Cursor::with_limits(
                    record.payload(),
                    "BrtSupNameValueStart",
                    limits,
                );
                let rows = cursor.read_u32()?;
                let columns = cursor.read_u32()?;
                cursor.finish()?;
                let count = usize::try_from(rows)
                    .ok()
                    .and_then(|rows| {
                        usize::try_from(columns)
                            .ok()
                            .and_then(|columns| rows.checked_mul(columns))
                    })
                    .ok_or_else(|| invalid("external cached-value dimensions overflow"))?;
                if count > MAX_XLSB_EXTERNAL_CACHED_VALUES {
                    return Err(Error::InvalidLength {
                        expected: MAX_XLSB_EXTERNAL_CACHED_VALUES,
                        found: count,
                    });
                }
                cache_values.clear();
                cache_values
                    .try_reserve(count)
                    .map_err(|source| Error::Allocation {
                        resource: "external cached values",
                        source,
                    })?;
                cache_dimensions = Some((rows, columns, count));
                sup_name_state = 4;
            },
            crate::raw::kind::SUP_NAME_NIL
            | crate::raw::kind::SUP_NAME_NUM
            | crate::raw::kind::SUP_NAME_BOOL
            | crate::raw::kind::SUP_NAME_ERROR
            | crate::raw::kind::SUP_NAME_STRING => {
                let Some((_, _, count)) = cache_dimensions else {
                    return Err(invalid("cached external value occurs outside its matrix"));
                };
                if sup_name_state != 4 || cache_values.len() >= count {
                    return Err(invalid("too many or misplaced cached external values"));
                }
                cache_values.push(parse_external_cached_value(
                    record.kind(),
                    record.payload(),
                    limits,
                )?);
            },
            crate::raw::kind::SUP_NAME_VALUE_END => {
                let Some((rows, columns, count)) = cache_dimensions.take() else {
                    return Err(invalid("unexpected BrtSupNameValueEnd"));
                };
                if sup_name_state != 4
                    || !record.payload().is_empty()
                    || cache_values.len() != count
                {
                    return Err(invalid("invalid cached external value matrix"));
                }
                current_cache = Some(ValueMatrix::new(
                    rows,
                    columns,
                    std::mem::take(&mut cache_values),
                )?);
                sup_name_state = 3;
            },
            crate::raw::kind::SUP_NAME_END => {
                if sup_name_state != 3 || !record.payload().is_empty() {
                    return Err(invalid("invalid BrtSupNameEnd"));
                }
                let kind = link_type.expect("external link kind is present");
                let name = current_name
                    .take()
                    .ok_or_else(|| invalid("external name block has no name"))?;
                let bits = current_bits
                    .take()
                    .ok_or_else(|| invalid("external name block has no properties"))?;
                match kind {
                    EXTERNAL_REFERENCE_WORKBOOK => {
                        let scope = u32::from_le_bytes([bits[2], bits[3], bits[4], bits[5]]);
                        let mut entry = DefinedName::new(name)?
                            .with_built_in(bits[0] & EXTERNAL_NAME_BUILT_IN != 0);
                        if scope != 0 {
                            entry = entry
                                .with_sheet_scope(u16::try_from(scope - 1).map_err(|_| {
                                    invalid("external defined-name scope overflow")
                                })?);
                        }
                        if let Some(formula) = current_formula.take() {
                            entry = entry.with_formula(formula);
                        }
                        if workbook_entries.len() >= MAX_COLLECTION_ITEMS {
                            return Err(invalid(
                                "external-link entry collection exceeds 65,535 items",
                            ));
                        }
                        workbook_entries.push(entry);
                    },
                    EXTERNAL_REFERENCE_DDE => {
                        let mut item = DdeItem::new(name)?
                            .with_advise(bits[0] & DATA_ITEM_WANT_ADVISE != 0)
                            .with_picture(bits[0] & DATA_ITEM_WANT_PICTURE != 0)
                            .with_ole_support(bits[0] & DDE_ITEM_SUPPORTS_OLE != 0);
                        if let Some(cache) = current_cache.take() {
                            item = item.with_cached_values(cache);
                        }
                        if dde_entries.len() >= MAX_COLLECTION_ITEMS {
                            return Err(invalid(
                                "external-link entry collection exceeds 65,535 items",
                            ));
                        }
                        dde_entries.push(item);
                    },
                    EXTERNAL_REFERENCE_OLE => {
                        let mut item = OleItem::new(name)?
                            .with_advise(bits[0] & DATA_ITEM_WANT_ADVISE != 0)
                            .with_picture(bits[0] & DATA_ITEM_WANT_PICTURE != 0)
                            .with_icon(bits[0] & OLE_ITEM_DISPLAY_AS_ICON != 0);
                        if let Some(cache) = current_cache.take() {
                            item = item.with_cached_values(cache);
                        }
                        if ole_entries.len() >= MAX_COLLECTION_ITEMS {
                            return Err(invalid(
                                "external-link entry collection exceeds 65,535 items",
                            ));
                        }
                        ole_entries.push(item);
                    },
                    _ => unreachable!("external link kind was validated above"),
                }
                sup_name_state = 0;
            },
            crate::raw::kind::END_SUP_BOOK => {
                if !record.payload().is_empty() {
                    return Err(Error::InvalidLength {
                        expected: 0,
                        found: record.payload().len(),
                    });
                }
                if sup_name_state != 0 {
                    return Err(invalid(
                        "BrtEndSupBook occurs inside an external-name block",
                    ));
                }
                saw_end = true;
            },
            _ => {
                if sup_name_state == 4
                    || (link_type == Some(EXTERNAL_REFERENCE_WORKBOOK) && sup_name_state != 0)
                {
                    return Err(invalid(
                        "unexpected record inside an external name or cache",
                    ));
                }
            },
        }
    }

    let kind = link_type.ok_or_else(|| invalid("external link has no BrtBeginSupBook"))?;
    if !saw_end {
        return Err(invalid("external link has no BrtEndSupBook"));
    }
    if kind == EXTERNAL_REFERENCE_WORKBOOK && !saw_sup_tabs {
        return Err(invalid("external workbook link has no BrtSupTabs"));
    }
    let link_kind = match kind {
        EXTERNAL_REFERENCE_WORKBOOK => Kind::Workbook,
        EXTERNAL_REFERENCE_DDE => Kind::Dde,
        EXTERNAL_REFERENCE_OLE => Kind::Ole,
        _ => unreachable!("external link kind was validated above"),
    };
    let relationship_id = match link_kind {
        Kind::Dde => None,
        Kind::Workbook | Kind::Ole => Some(target_key.clone()),
    };
    let entries = match kind {
        EXTERNAL_REFERENCE_WORKBOOK => Entries::Workbook(workbook_entries),
        EXTERNAL_REFERENCE_DDE => Entries::Dde(dde_entries),
        EXTERNAL_REFERENCE_OLE => Entries::Ole(ole_entries),
        _ => unreachable!("external link kind was validated above"),
    };
    let link = Link {
        kind: link_kind,
        source: target_key,
        detail: match link_kind {
            Kind::Dde | Kind::Ole => Some(target_detail),
            Kind::Workbook => None,
        },
        sheet_names,
        entries,
    };
    link.validate()?;
    Ok(Parsed {
        link,
        relationship_id,
    })
}

/// Explicitly named alias for callers that need the unresolved relationship
/// metadata as well as the typed link.
pub fn parse_external_link_with_relationship(data: &[u8]) -> Result<Parsed> {
    parse_external_link(data)
}

/// Parse a stream when the caller only needs its inert semantic model.
pub fn parse_external_link_model(data: &[u8]) -> Result<Link> {
    parse_external_link(data).map(Parsed::into_link)
}

fn parse_external_cached_value(
    record_type: crate::raw::Kind,
    data: &[u8],
    limits: crate::raw::Limits,
) -> Result<CachedValue> {
    match record_type {
        crate::raw::kind::SUP_NAME_NIL if data.is_empty() => Ok(CachedValue::Empty),
        crate::raw::kind::SUP_NAME_NUM if data.len() == 8 => {
            let number = f64::from_le_bytes(data.try_into().expect("length was checked"));
            validate_number(number)?;
            Ok(CachedValue::Number(number))
        },
        crate::raw::kind::SUP_NAME_BOOL if data.len() == 1 && data[0] <= 1 => {
            Ok(CachedValue::Boolean(data[0] != 0))
        },
        crate::raw::kind::SUP_NAME_ERROR if data.len() == 1 => {
            Ok(CachedValue::Error(ErrorValue::from_code(data[0])?))
        },
        crate::raw::kind::SUP_NAME_STRING => {
            if data.len() < 4 {
                return Err(Error::InvalidLength {
                    expected: 4,
                    found: data.len(),
                });
            }
            let mut cursor = crate::raw::Cursor::with_limits(data, "BrtSupNameSt", limits);
            let value = cursor.read_wide_string()?;
            cursor.finish()?;
            Ok(CachedValue::String(value))
        },
        _ => Err(invalid(format!(
            "invalid cached external value record {record_type}"
        ))),
    }
}

fn parse_external_sheet_names(data: &[u8], limits: crate::raw::Limits) -> Result<Vec<String>> {
    if data.len() < 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: data.len(),
        });
    }
    let mut cursor = crate::raw::Cursor::with_limits(data, "BrtSupTabs", limits);
    let count = usize::try_from(cursor.read_u32()?)
        .map_err(|_| invalid("external sheet-name count overflow"))?;
    if count >= MAX_COLLECTION_ITEMS {
        return Err(invalid(format!(
            "external sheet-name count {count} exceeds 65,534"
        )));
    }
    let mut names = Vec::new();
    names
        .try_reserve(count)
        .map_err(|source| Error::Allocation {
            resource: "external sheet names",
            source,
        })?;
    for _ in 0..count {
        let name = cursor.read_wide_string()?;
        let name_len = name.encode_utf16().count();
        if name_len == 0
            || name_len > 31
            || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
            || name.starts_with('\'')
            || name.ends_with('\'')
        {
            return Err(invalid(format!(
                "external sheet name {name:?} does not follow sheet-name grammar"
            )));
        }
        if names
            .iter()
            .any(|existing: &String| excel_name_eq(existing, &name))
        {
            return Err(invalid(format!("duplicate external sheet name {name:?}")));
        }
        names.push(name);
    }
    cursor.finish()?;
    Ok(names)
}

fn validate_external_name_bits(kind: u16, bits: &[u8; 7]) -> Result<()> {
    let reserved_word = &bits[2..6];
    let valid = match kind {
        EXTERNAL_REFERENCE_WORKBOOK => {
            bits[0] & EXTERNAL_NAME_RESERVED_MASK == 0
                && bits[6] & DATA_ITEM_REQUIRED_TRAILING_FLAG == 0
        },
        EXTERNAL_REFERENCE_DDE => {
            bits[0] & DDE_ITEM_RESERVED_MASK == 0
                && reserved_word == [0, 0, 0, 0]
                && bits[6] & DATA_ITEM_REQUIRED_TRAILING_FLAG != 0
        },
        EXTERNAL_REFERENCE_OLE => {
            bits[0] & OLE_ITEM_RESERVED_MASK == 0
                && bits[0] & OLE_ITEM_REQUIRED_CLASS_FLAG != 0
                && reserved_word == [0, 0, 0, 0]
                && bits[6] & DATA_ITEM_REQUIRED_TRAILING_FLAG != 0
        },
        _ => false,
    };
    if !valid {
        return Err(invalid(format!(
            "invalid BrtSupNameBits properties for external-link kind {kind}"
        )));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormula(message.into())
}

fn excel_name_eq(left: &str, right: &str) -> bool {
    left.chars()
        .flat_map(char::to_lowercase)
        .eq(right.chars().flat_map(char::to_lowercase))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::external_link::model::{
        EXT_PTG_AREA, EXT_PTG_AREA_ERROR, EXT_PTG_ERROR, EXT_PTG_REFERENCE,
        EXT_PTG_REFERENCE_ERROR, REFERENCE_ERROR_CODE,
    };
    use crate::raw::Writer;

    #[test]
    fn accepts_exactly_the_five_external_name_token_structures() {
        for (tokens, kind) in [
            (
                vec![EXT_PTG_REFERENCE, 0, 0, 0, 0, 3, 0, 2, 0],
                NameFormulaKind::CellReference,
            ),
            (
                vec![EXT_PTG_AREA, 0, 0, 0, 0, 1, 0, 3, 0, 2, 0, 4, 0],
                NameFormulaKind::AreaReference,
            ),
            (
                vec![EXT_PTG_REFERENCE_ERROR, 0, 0, 0, 0, 0, 0, 0, 0],
                NameFormulaKind::CellReferenceError,
            ),
            (
                vec![EXT_PTG_AREA_ERROR, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                NameFormulaKind::AreaReferenceError,
            ),
            (
                vec![EXT_PTG_ERROR, REFERENCE_ERROR_CODE],
                NameFormulaKind::ReferenceError,
            ),
        ] {
            assert_eq!(NameFormula::from_tokens(tokens).unwrap().kind(), kind);
        }
    }

    #[test]
    fn rejects_external_name_tokens_outside_the_restricted_grammar() {
        assert!(NameFormula::from_tokens(vec![0x1E, 1, 0]).is_err());
        assert!(NameFormula::from_tokens(vec![EXT_PTG_ERROR, 0x2A]).is_err());
        assert!(
            NameFormula::from_tokens(vec![EXT_PTG_AREA, 0, 0, 0, 0, 3, 0, 1, 0, 0, 0, 0, 0,])
                .is_err()
        );
    }

    #[test]
    fn typed_external_name_formula_constructors_emit_canonical_tokens() {
        let sheets = SheetRange::sheets(0, 1).unwrap();
        let first = CellLocation::new(2, 3)
            .with_column_relative(true)
            .with_row_relative(true);
        let last = CellLocation::new(4, 5);
        let cell = NameFormula::cell_reference(CellReference::new(sheets, first));
        assert_eq!(
            cell.tokens(),
            [EXT_PTG_REFERENCE, 0, 0, 1, 0, 2, 0, 3, 0xC0]
        );
        assert_eq!(cell.cell().unwrap().location(), first);
        let area = NameFormula::area_reference(AreaReference::new(sheets, first, last).unwrap());
        assert_eq!(
            area.tokens(),
            [EXT_PTG_AREA, 0, 0, 1, 0, 2, 0, 4, 0, 3, 0xC0, 5, 0]
        );
        assert_eq!(area.area().unwrap().first(), first);
        assert_eq!(area.sheets(), Some(sheets));
        assert_eq!(
            NameFormula::cell_reference_error(SheetRange::Missing).kind(),
            NameFormulaKind::CellReferenceError
        );
        assert_eq!(
            NameFormula::area_reference_error(SheetRange::Missing).kind(),
            NameFormulaKind::AreaReferenceError
        );
        assert_eq!(
            NameFormula::reference_error().tokens(),
            [EXT_PTG_ERROR, REFERENCE_ERROR_CODE]
        );
    }

    #[test]
    fn owner_stream_writer_and_parser_preserve_relationship_ids_and_metadata() {
        let sheets = SheetRange::sheets(0, 1).unwrap();
        let formula =
            NameFormula::cell_reference(CellReference::new(sheets, CellLocation::new(3, 2)));
        let link = Link::workbook_with_defined_names(
            "Book.xlsx",
            vec!["Data".to_string(), "Rates".to_string()],
            vec![
                DefinedName::new("ExchangeRate")
                    .unwrap()
                    .with_formula(formula)
                    .with_built_in(true)
                    .with_sheet_scope(1),
            ],
        )
        .unwrap();

        let bytes = write_external_link_stream(&link, Some("rIdPath")).unwrap();
        let parsed = parse_external_link(&bytes).unwrap();
        assert_eq!(parsed.relationship_id(), Some("rIdPath"));
        assert_eq!(parsed.link().source(), "rIdPath");
        assert_eq!(parsed.link().sheet_names(), link.sheet_names());
        assert_eq!(parsed.link().defined_names(), link.defined_names());
        assert_eq!(parsed.resolve_source("Book.xlsx").unwrap(), link);
    }

    #[test]
    fn owner_stream_codec_preserves_inert_dde_and_ole_caches() {
        let cache = ValueMatrix::new(
            1,
            3,
            vec![
                CachedValue::Number(7.0),
                CachedValue::Boolean(true),
                CachedValue::String("Ready".to_string()),
            ],
        )
        .unwrap();
        let dde = Link::dde_with_items(
            "Excel",
            "System",
            vec![
                DdeItem::new("StatusItem")
                    .unwrap()
                    .with_advise(true)
                    .with_picture(true)
                    .with_cached_values(cache.clone()),
            ],
        )
        .unwrap();
        let dde_bytes = write_external_link_stream(&dde, None).unwrap();
        let parsed_dde = parse_external_link(&dde_bytes).unwrap().into_link();
        assert_eq!(parsed_dde, dde);

        let ole = Link::ole_with_items(
            "Model.xlsx",
            "Acme.Server",
            vec![
                OleItem::new("ReportItem")
                    .unwrap()
                    .with_advise(true)
                    .with_picture(true)
                    .with_icon(true)
                    .with_cached_values(cache),
            ],
        )
        .unwrap();
        let ole_bytes = write_external_link_stream(&ole, Some("rIdOle")).unwrap();
        let parsed_ole = parse_external_link(&ole_bytes).unwrap();
        assert_eq!(parsed_ole.relationship_id(), Some("rIdOle"));
        assert_eq!(parsed_ole.resolve_source("Model.xlsx").unwrap(), ole);
    }

    #[test]
    fn owner_parser_rejects_trailing_records_after_end() {
        let link = Link::dde("Excel", "System", vec!["StatusItem".to_string()]).unwrap();
        let mut bytes = write_external_link_stream(&link, None).unwrap();
        Writer::new(&mut bytes)
            .write_record(crate::raw::kind::SUP_NAME_END, &[])
            .unwrap();
        assert!(parse_external_link(&bytes).is_err());
    }
}
