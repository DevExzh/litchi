//! BIFF8 payload codecs for external-link records.

use super::model::{
    CacheRow, CachedValue, ClipboardFormat, ErrorValue, Name, NameBody, Sheet, SheetReference,
    SupportingBook, ValueMatrix, Workbook,
};
use super::{
    CRN_RECORD_TYPE, EXTERN_NAME_RECORD_TYPE, EXTERN_SHEET_RECORD_TYPE, MAX_DDE_OLE_VALUES,
    MAX_EXTERNAL_SHEETS, SUP_BOOK_RECORD_TYPE,
};
use crate::{Error, Result};

impl ClipboardFormat {
    pub(super) fn parse(value: i16) -> Result<Self> {
        match value {
            -1 => Ok(Self::None),
            0 => Ok(Self::Text),
            2 => Ok(Self::EnhancedMetafile),
            5 => Ok(Self::Csv),
            6 => Ok(Self::Sylk),
            7 => Ok(Self::RichText),
            8 => Ok(Self::Biff8),
            9 => Ok(Self::Bitmap),
            16 => Ok(Self::ApplicationTable),
            20 => Ok(Self::Biff3),
            30 => Ok(Self::Biff4),
            36 => Ok(Self::MetafilePicture),
            44 => Ok(Self::UnicodeText),
            63 => Ok(Self::Biff12),
            _ => invalid(
                EXTERN_NAME_RECORD_TYPE,
                "ExternName clipboard format is invalid",
            ),
        }
    }
}

impl ErrorValue {
    pub(super) fn parse(code: u8) -> Result<Self> {
        match code {
            0x00 => Ok(Self::Null),
            0x07 => Ok(Self::DivisionByZero),
            0x0f => Ok(Self::Value),
            0x17 => Ok(Self::Reference),
            0x1d => Ok(Self::Name),
            0x24 => Ok(Self::Number),
            0x2a => Ok(Self::NotAvailable),
            0x2b => Ok(Self::GettingData),
            _ => invalid(
                CRN_RECORD_TYPE,
                format!("invalid cached error code 0x{code:02X}"),
            ),
        }
    }
}

pub(super) fn parse_external_name(
    data: &[u8],
    book_index: usize,
    book: &SupportingBook,
) -> Result<Name> {
    if data.len() < 8 || data.len() > 8224 {
        return invalid(
            EXTERN_NAME_RECORD_TYPE,
            "ExternName payload length must be 8..=8224 bytes",
        );
    }
    let flags = read_u16(data, 0);
    let built_in = flags & 0x0001 != 0;
    let automatic = flags & 0x0002 != 0;
    let picture = flags & 0x0004 != 0;
    let standard_document_name = flags & 0x0008 != 0;
    let ole_link = flags & 0x0010 != 0;
    if standard_document_name && ole_link {
        return invalid(
            EXTERN_NAME_RECORD_TYPE,
            "ExternName fOle and fOleLink are mutually exclusive",
        );
    }
    let clipboard_bits = (flags >> 5) & 0x03ff;
    let clipboard_format = if clipboard_bits & 0x0200 != 0 {
        crate::utils::wrap_u16_to_i16(clipboard_bits) - 1024
    } else {
        crate::utils::wrap_u16_to_i16(clipboard_bits)
    };
    let clipboard_format = ClipboardFormat::parse(clipboard_format)?;
    let displayed_as_icon = flags & 0x8000 != 0;
    let (name, offset) = parse_short_unicode_string(data, 6, EXTERN_NAME_RECORD_TYPE)?;

    let body = match book {
        SupportingBook::ExternalWorkbook(book) => {
            if flags & !0x0001 != 0 {
                return invalid(
                    EXTERN_NAME_RECORD_TYPE,
                    "external defined name has non-name flags",
                );
            }
            if read_u16(data, 4) != 0 {
                return invalid(
                    EXTERN_NAME_RECORD_TYPE,
                    "ExternDocName reserved field must be zero",
                );
            }
            let ixals = read_u16(data, 2);
            if usize::from(ixals) > book.sheets.len() {
                return invalid(
                    EXTERN_NAME_RECORD_TYPE,
                    "ExternDocName sheet scope exceeds SupBook sheet count",
                );
            }
            let formula_len =
                usize::from(*data.get(offset).ok_or_else(|| Error::InvalidRecord {
                    record_type: EXTERN_NAME_RECORD_TYPE,
                    message: "ExtNameParsedFormula length is missing".to_string(),
                })?) | (usize::from(*data.get(offset + 1).ok_or_else(|| {
                    Error::InvalidRecord {
                        record_type: EXTERN_NAME_RECORD_TYPE,
                        message: "ExtNameParsedFormula length is truncated".to_string(),
                    }
                })?) << 8);
            let formula_start = offset + 2;
            if formula_start.checked_add(formula_len) != Some(data.len()) {
                return invalid(
                    EXTERN_NAME_RECORD_TYPE,
                    "ExtNameParsedFormula length does not consume ExternName",
                );
            }
            let formula_bytes = data[formula_start..].to_vec();
            if formula_bytes
                .first()
                .is_some_and(|token| !matches!(token, 0x1c | 0x3a | 0x3b | 0x3c | 0x3d))
            {
                return invalid(
                    EXTERN_NAME_RECORD_TYPE,
                    "ExtNameParsedFormula token kind is invalid",
                );
            }
            NameBody::ExternalDefinedName {
                sheet_index: ixals.checked_sub(1),
                name,
                formula_bytes,
            }
        },
        SupportingBook::AddIn => {
            if flags != 0 || read_u16(data, 2) != 0 || read_u16(data, 4) != 0 {
                return invalid(
                    EXTERN_NAME_RECORD_TYPE,
                    "AddinUdf flags and reserved fields must be zero",
                );
            }
            let unused_len = usize::from(read_u16_checked(data, offset, EXTERN_NAME_RECORD_TYPE)?);
            let unused_start = offset + 2;
            if unused_start.checked_add(unused_len) != Some(data.len()) {
                return invalid(
                    EXTERN_NAME_RECORD_TYPE,
                    "AddinUdf unused byte count does not consume ExternName",
                );
            }
            NameBody::AddInFunction {
                name,
                unused_data: data[unused_start..].to_vec(),
            }
        },
        SupportingBook::DdeOrOle { .. } => {
            let storage_id = u32::from_le_bytes(data[2..6].try_into().unwrap());
            if built_in {
                return invalid(
                    EXTERN_NAME_RECORD_TYPE,
                    "DDE/OLE ExternName has incompatible flags",
                );
            }
            if standard_document_name {
                if displayed_as_icon
                    || storage_id != 0
                    || name != "StdDocumentName"
                    || offset != data.len()
                {
                    return invalid(
                        EXTERN_NAME_RECORD_TYPE,
                        "ExternDdeLinkNoOper body is invalid",
                    );
                }
                NameBody::DdeStandardDocumentName { name }
            } else {
                if !ole_link && storage_id != 0 {
                    return invalid(
                        EXTERN_NAME_RECORD_TYPE,
                        "DDE item has a nonzero OLE link-storage identifier",
                    );
                }
                if displayed_as_icon && !ole_link {
                    return invalid(
                        EXTERN_NAME_RECORD_TYPE,
                        "only an OLE item can be displayed as an icon",
                    );
                }
                let remaining = &data[offset..];
                let matrix = if remaining.is_empty() {
                    None
                } else {
                    if remaining.len() < 3 {
                        return invalid(
                            EXTERN_NAME_RECORD_TYPE,
                            "DDE/OLE MOper header is truncated",
                        );
                    }
                    let last_column = remaining[0];
                    let last_row = read_u16(remaining, 1);
                    let expected = (usize::from(last_column) + 1)
                        .checked_mul(usize::from(last_row) + 1)
                        .ok_or_else(|| Error::InvalidRecord {
                            record_type: EXTERN_NAME_RECORD_TYPE,
                            message: "DDE/OLE matrix size overflows".to_string(),
                        })?;
                    if expected > MAX_DDE_OLE_VALUES {
                        return invalid(
                            EXTERN_NAME_RECORD_TYPE,
                            "DDE/OLE matrix exceeds the resource bound",
                        );
                    }
                    let values =
                        parse_ser_ar_values(&remaining[3..], expected, EXTERN_NAME_RECORD_TYPE)?;
                    Some(ValueMatrix {
                        last_column,
                        last_row,
                        values,
                    })
                };
                NameBody::DdeOrOle {
                    storage_id,
                    name,
                    matrix,
                }
            }
        },
        _ => {
            return invalid(
                EXTERN_NAME_RECORD_TYPE,
                "SupBook variant cannot own ExternName records",
            );
        },
    };
    Ok(Name {
        supporting_book_index: u16::try_from(book_index).map_err(|_error| {
            Error::InvalidRecord {
                record_type: EXTERN_NAME_RECORD_TYPE,
                message: "supporting-book index exceeds u16".to_string(),
            }
        })?,
        built_in,
        automatic,
        picture,
        standard_document_name,
        ole_link,
        clipboard_format,
        displayed_as_icon,
        body,
    })
}

pub(super) fn parse_ser_ar_values(
    data: &[u8],
    maximum: usize,
    record_type: u16,
) -> Result<Vec<CachedValue>> {
    let mut offset = 0usize;
    let mut values = Vec::new();
    while offset < data.len() {
        if values.len() == maximum {
            return invalid(record_type, "SerAr value count exceeds matrix dimensions");
        }
        let tag = data[offset];
        match tag {
            0x00 => {
                require_available_for(data, offset, 9, record_type)?;
                values.push(CachedValue::Blank);
                offset += 9;
            },
            0x01 => {
                require_available_for(data, offset, 9, record_type)?;
                values.push(CachedValue::Number(f64::from_le_bytes(
                    data[offset + 1..offset + 9].try_into().unwrap(),
                )));
                offset += 9;
            },
            0x02 => {
                let (value, next) = parse_unicode_string(data, offset + 1, 255, record_type)?;
                values.push(CachedValue::Text(value));
                offset = next;
            },
            0x04 => {
                require_available_for(data, offset, 9, record_type)?;
                if data[offset + 2] != 0 {
                    return invalid(record_type, "SerBool reserved byte must be zero");
                }
                values.push(CachedValue::Boolean(match data[offset + 1] {
                    0 => false,
                    1 => true,
                    _ => return invalid(record_type, "SerBool value must be zero or one"),
                }));
                offset += 9;
            },
            0x10 => {
                require_available_for(data, offset, 9, record_type)?;
                if data[offset + 2] != 0 {
                    return invalid(record_type, "SerErr reserved byte must be zero");
                }
                values.push(CachedValue::Error(ErrorValue::parse(data[offset + 1])?));
                offset += 9;
            },
            _ => return invalid(record_type, format!("unknown SerAr tag 0x{tag:02X}")),
        }
    }
    Ok(values)
}

fn require_available_for(
    data: &[u8],
    offset: usize,
    length: usize,
    record_type: u16,
) -> Result<()> {
    if offset
        .checked_add(length)
        .is_some_and(|end| end <= data.len())
    {
        Ok(())
    } else {
        invalid(
            record_type,
            "SerAr value is truncated or split across records",
        )
    }
}

fn parse_short_unicode_string(
    data: &[u8],
    offset: usize,
    record_type: u16,
) -> Result<(String, usize)> {
    let count = usize::from(*data.get(offset).ok_or_else(|| Error::InvalidRecord {
        record_type,
        message: "short Unicode string length is missing".to_string(),
    })?);
    parse_unicode_no_cch(data, offset + 1, count, record_type)
}

fn read_u16_checked(data: &[u8], offset: usize, record_type: u16) -> Result<u16> {
    let bytes = data
        .get(offset..offset + 2)
        .ok_or_else(|| Error::InvalidRecord {
            record_type,
            message: "two-byte field is truncated".to_string(),
        })?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

pub(super) fn parse_sup_book(data: &[u8]) -> Result<SupportingBook> {
    if data.len() < 4 || data.len() > 8224 {
        return invalid(
            SUP_BOOK_RECORD_TYPE,
            "SupBook payload length is outside BIFF8 bounds",
        );
    }
    let sheet_count = usize::from(read_u16(data, 0));
    let cch = read_u16(data, 2);
    if data.len() == 4 {
        return match cch {
            0x0401 => Ok(SupportingBook::SelfReference),
            0x3a01 if sheet_count == 1 => Ok(SupportingBook::AddIn),
            0x3a01 => invalid(
                SUP_BOOK_RECORD_TYPE,
                "add-in SupBook sheet count must be one",
            ),
            _ => invalid(SUP_BOOK_RECORD_TYPE, "invalid four-byte SupBook marker"),
        };
    }
    if !(1..=255).contains(&cch) {
        return invalid(
            SUP_BOOK_RECORD_TYPE,
            "SupBook virtual path length must be 1..=255",
        );
    }
    let (path, mut offset) = parse_unicode_no_cch(data, 4, usize::from(cch), SUP_BOOK_RECORD_TYPE)?;
    if path == "\0" {
        if sheet_count != 0 || offset != data.len() {
            return invalid(
                SUP_BOOK_RECORD_TYPE,
                "same-sheet SupBook must have zero sheets and no rgst",
            );
        }
        return Ok(SupportingBook::SameSheet);
    }
    if sheet_count == 0 {
        if offset != data.len() {
            return invalid(
                SUP_BOOK_RECORD_TYPE,
                "DDE/OLE SupBook has trailing sheet names",
            );
        }
        return Ok(SupportingBook::DdeOrOle {
            encoded_virtual_path: path,
        });
    }
    if sheet_count > MAX_EXTERNAL_SHEETS {
        return invalid(
            SUP_BOOK_RECORD_TYPE,
            "external SupBook sheet count exceeds resource bound",
        );
    }
    let mut names = Vec::with_capacity(sheet_count);
    for _ in 0..sheet_count {
        let (name, next) = parse_unicode_string(data, offset, 31, SUP_BOOK_RECORD_TYPE)?;
        names.push(name);
        offset = next;
    }
    if offset != data.len() {
        return invalid(
            SUP_BOOK_RECORD_TYPE,
            "SupBook sheet count does not consume payload exactly",
        );
    }
    if path == " " {
        if names.iter().any(|name| name != " ") {
            return invalid(
                SUP_BOOK_RECORD_TYPE,
                "unused SupBook sheet placeholders must be spaces",
            );
        }
        return Ok(SupportingBook::Unused { sheet_names: names });
    }
    Ok(SupportingBook::ExternalWorkbook(Workbook {
        encoded_virtual_path: path,
        sheets: names
            .into_iter()
            .map(|name| Sheet {
                name,
                cache_valid: true,
                cache_rows: Vec::new(),
                cache_declared: false,
            })
            .collect(),
    }))
}

pub(super) fn parse_extern_sheet(data: &[u8]) -> Result<Vec<SheetReference>> {
    if data.len() < 2 {
        return Err(Error::InvalidLength {
            expected: 2,
            found: data.len(),
        });
    }
    let count = usize::from(read_u16(data, 0));
    let expected = 2usize
        .checked_add(count.checked_mul(6).ok_or_else(|| Error::InvalidRecord {
            record_type: EXTERN_SHEET_RECORD_TYPE,
            message: "ExternSheet count overflows".to_string(),
        })?)
        .ok_or_else(|| Error::InvalidRecord {
            record_type: EXTERN_SHEET_RECORD_TYPE,
            message: "ExternSheet size overflows".to_string(),
        })?;
    if data.len() != expected {
        return Err(Error::InvalidLength {
            expected,
            found: data.len(),
        });
    }
    Ok(data[2..]
        .chunks_exact(6)
        .map(|entry| SheetReference {
            supporting_book_index: read_u16(entry, 0),
            first_sheet_index: i16::from_le_bytes([entry[2], entry[3]]),
            last_sheet_index: i16::from_le_bytes([entry[4], entry[5]]),
        })
        .collect())
}

pub(super) fn validate_reference(
    reference: &SheetReference,
    books: &[SupportingBook],
    internal_sheet_count: usize,
) -> Result<()> {
    let book = books
        .get(usize::from(reference.supporting_book_index))
        .ok_or_else(|| Error::InvalidRecord {
            record_type: EXTERN_SHEET_RECORD_TYPE,
            message: "XTI supporting-book index is out of range".to_string(),
        })?;
    let (first, last) = (reference.first_sheet_index, reference.last_sheet_index);
    match book {
        SupportingBook::AddIn | SupportingBook::SameSheet | SupportingBook::DdeOrOle { .. } => {
            if first != -2 || last != -2 {
                return invalid(
                    EXTERN_SHEET_RECORD_TYPE,
                    "unscoped supporting link requires -2/-2 XTI scope",
                );
            }
        },
        SupportingBook::SelfReference => validate_sheet_scope(first, last, internal_sheet_count)?,
        SupportingBook::ExternalWorkbook(book) => {
            validate_sheet_scope(first, last, book.sheets.len())?;
        },
        SupportingBook::Unused { sheet_names } => {
            validate_sheet_scope(first, last, sheet_names.len())?;
        },
    }
    Ok(())
}

fn validate_sheet_scope(first: i16, last: i16, sheet_count: usize) -> Result<()> {
    if first == -2 {
        if last != -2 {
            return invalid(EXTERN_SHEET_RECORD_TYPE, "workbook XTI scope must be -2/-2");
        }
        return Ok(());
    }
    if first == -1 || last == -1 {
        if first < -1 || last < -1 {
            return invalid(EXTERN_SHEET_RECORD_TYPE, "invalid missing-sheet XTI scope");
        }
        return Ok(());
    }
    if first < 0 || last < first || usize::try_from(last).unwrap_or(usize::MAX) >= sheet_count {
        return invalid(
            EXTERN_SHEET_RECORD_TYPE,
            "XTI sheet scope is outside supporting book",
        );
    }
    Ok(())
}

pub(super) fn parse_cache_row(data: &[u8]) -> Result<CacheRow> {
    if data.len() < 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: data.len(),
        });
    }
    let last_column = data[0];
    let first_column = data[1];
    if last_column < first_column {
        return invalid(CRN_RECORD_TYPE, "CRN column range is reversed");
    }
    let value_count = usize::from(last_column - first_column) + 1;
    let mut offset = 4;
    let mut values = Vec::with_capacity(value_count);
    for _ in 0..value_count {
        let tag = *data.get(offset).ok_or_else(|| Error::InvalidRecord {
            record_type: CRN_RECORD_TYPE,
            message: "CRN cached value is truncated".to_string(),
        })?;
        match tag {
            0x00 => {
                require_available(data, offset, 9)?;
                values.push(CachedValue::Blank);
                offset += 9;
            },
            0x01 => {
                require_available(data, offset, 9)?;
                let value = f64::from_le_bytes(data[offset + 1..offset + 9].try_into().unwrap());
                values.push(CachedValue::Number(value));
                offset += 9;
            },
            0x02 => {
                let (value, next) = parse_unicode_string(data, offset + 1, 255, CRN_RECORD_TYPE)?;
                values.push(CachedValue::Text(value));
                offset = next;
            },
            0x04 => {
                require_available(data, offset, 9)?;
                if data[offset + 2] != 0 {
                    return invalid(CRN_RECORD_TYPE, "SerBool reserved byte must be zero");
                }
                let value = match data[offset + 1] {
                    0 => false,
                    1 => true,
                    _ => return invalid(CRN_RECORD_TYPE, "SerBool value must be zero or one"),
                };
                values.push(CachedValue::Boolean(value));
                offset += 9;
            },
            0x10 => {
                require_available(data, offset, 9)?;
                if data[offset + 2] != 0 {
                    return invalid(CRN_RECORD_TYPE, "SerErr reserved byte must be zero");
                }
                values.push(CachedValue::Error(ErrorValue::parse(data[offset + 1])?));
                offset += 9;
            },
            _ => return invalid(CRN_RECORD_TYPE, format!("unknown SerAr tag 0x{tag:02X}")),
        }
    }
    if offset != data.len() {
        return invalid(
            CRN_RECORD_TYPE,
            "CRN cached values do not consume payload exactly",
        );
    }
    Ok(CacheRow {
        row: read_u16(data, 2),
        first_column,
        values,
    })
}

fn parse_unicode_string(
    data: &[u8],
    offset: usize,
    max_characters: usize,
    record_type: u16,
) -> Result<(String, usize)> {
    require_available(data, offset, 3)?;
    let count = usize::from(read_u16(data, offset));
    parse_unicode_no_cch(data, offset + 2, count, record_type).and_then(|(value, next)| {
        if count > max_characters {
            invalid(
                record_type,
                format!("string exceeds {max_characters} characters"),
            )
        } else {
            Ok((value, next))
        }
    })
}

fn parse_unicode_no_cch(
    data: &[u8],
    offset: usize,
    count: usize,
    record_type: u16,
) -> Result<(String, usize)> {
    let flags = *data.get(offset).ok_or_else(|| Error::InvalidRecord {
        record_type,
        message: "Unicode string options are missing".to_string(),
    })?;
    if flags & !1 != 0 {
        return invalid(record_type, "Unicode string reserved flags must be zero");
    }
    let wide = flags == 1;
    let bytes =
        count
            .checked_mul(if wide { 2 } else { 1 })
            .ok_or_else(|| Error::InvalidRecord {
                record_type,
                message: "Unicode string length overflows".to_string(),
            })?;
    let start = offset + 1;
    let end = start
        .checked_add(bytes)
        .ok_or_else(|| Error::InvalidRecord {
            record_type,
            message: "Unicode string end overflows".to_string(),
        })?;
    let encoded = data.get(start..end).ok_or_else(|| Error::InvalidRecord {
        record_type,
        message: "Unicode string is truncated".to_string(),
    })?;
    let value = if wide {
        let units = encoded
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units).map_err(|error| Error::InvalidRecord {
            record_type,
            message: format!("invalid UTF-16 string: {error}"),
        })?
    } else {
        encoded.iter().map(|byte| char::from(*byte)).collect()
    };
    Ok((value, end))
}

fn require_available(data: &[u8], offset: usize, length: usize) -> Result<()> {
    let end = offset
        .checked_add(length)
        .ok_or_else(|| Error::InvalidRecord {
            record_type: CRN_RECORD_TYPE,
            message: "cached value length overflows".to_string(),
        })?;
    if end > data.len() {
        return Err(Error::InvalidLength {
            expected: end,
            found: data.len(),
        });
    }
    Ok(())
}

pub(super) fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

/// Re-encodes only the specified supporting-book metadata.  Sheet cache
/// records and all records outside this `SupBook` remain owned by the source
/// package.
pub(super) fn encode_supporting_book(
    book: &SupportingBook,
    target: Option<&str>,
    sheet_edit: Option<(usize, &str)>,
) -> Result<Vec<u8>> {
    let target = match book {
        SupportingBook::ExternalWorkbook(book) => {
            let target = target.unwrap_or(book.encoded_virtual_path());
            validate_target(target, false)?;
            target
        },
        SupportingBook::DdeOrOle {
            encoded_virtual_path,
        } => {
            if sheet_edit.is_some() {
                return invalid(
                    SUP_BOOK_RECORD_TYPE,
                    "DDE/OLE supporting books do not own external sheet names",
                );
            }
            let target = target.unwrap_or(encoded_virtual_path);
            validate_target(target, true)?;
            target
        },
        SupportingBook::Unused { .. } => {
            return invalid(
                SUP_BOOK_RECORD_TYPE,
                "unused supporting-book placeholders do not own editable target metadata",
            );
        },
        SupportingBook::SelfReference | SupportingBook::AddIn | SupportingBook::SameSheet => {
            return invalid(
                SUP_BOOK_RECORD_TYPE,
                "this supporting-book kind does not own editable target metadata",
            );
        },
    };

    let mut payload = Vec::new();
    match book {
        SupportingBook::ExternalWorkbook(book) => {
            let sheet_count = u16::try_from(book.sheets.len()).map_err(|_error| {
                Error::Allocation("encoding external supporting-book sheet count")
            })?;
            if book.sheets.len() > MAX_EXTERNAL_SHEETS {
                return invalid(
                    SUP_BOOK_RECORD_TYPE,
                    "external SupBook sheet count exceeds resource bound",
                );
            }
            payload.extend_from_slice(&sheet_count.to_le_bytes());
            append_counted_unicode(&mut payload, target, 255, SUP_BOOK_RECORD_TYPE)?;
            for (index, sheet) in book.sheets.iter().enumerate() {
                let name = sheet_edit
                    .filter(|(edited, _)| *edited == index)
                    .map_or(sheet.name(), |(_, name)| name);
                append_counted_unicode(&mut payload, name, 31, SUP_BOOK_RECORD_TYPE)?;
            }
        },
        SupportingBook::DdeOrOle { .. } => {
            payload.extend_from_slice(&0u16.to_le_bytes());
            append_counted_unicode(&mut payload, target, 255, SUP_BOOK_RECORD_TYPE)?;
        },
        _ => unreachable!(),
    }
    if payload.len() > 8_224 {
        return invalid(
            SUP_BOOK_RECORD_TYPE,
            "encoded supporting-book payload exceeds BIFF8 bounds",
        );
    }
    Ok(payload)
}

/// Replaces the `stName` field in one existing `ExternName` while retaining
/// all flags, formula bytes, matrix bytes, and the record's following Continue
/// payloads exactly.
pub(super) fn encode_external_name(data: &[u8], name: &str) -> Result<Vec<u8>> {
    if data.len() < 8 {
        return invalid(
            EXTERN_NAME_RECORD_TYPE,
            "ExternName payload is too short for a name edit",
        );
    }
    let (_, name_end) = parse_short_unicode_string(data, 6, EXTERN_NAME_RECORD_TYPE)?;
    let encoded = encode_short_unicode(name, EXTERN_NAME_RECORD_TYPE)?;
    let mut payload = Vec::with_capacity(data.len() + encoded.len().saturating_sub(name_end - 6));
    payload.extend_from_slice(&data[..6]);
    payload.extend_from_slice(&encoded);
    payload.extend_from_slice(&data[name_end..]);
    if payload.len() > 8_224 {
        return invalid(
            EXTERN_NAME_RECORD_TYPE,
            "encoded ExternName payload exceeds BIFF8 bounds",
        );
    }
    Ok(payload)
}

fn validate_target(target: &str, allow_space: bool) -> Result<()> {
    let count = target.encode_utf16().count();
    if !(1..=255).contains(&count) {
        return invalid(
            SUP_BOOK_RECORD_TYPE,
            "supporting-book target must contain 1..=255 UTF-16 code units",
        );
    }
    if target.chars().any(|character| character == '\0') {
        return invalid(
            SUP_BOOK_RECORD_TYPE,
            "supporting-book target cannot contain NUL characters",
        );
    }
    if !allow_space && target == " " {
        return invalid(
            SUP_BOOK_RECORD_TYPE,
            "external-workbook target cannot become the unused-book marker",
        );
    }
    Ok(())
}

fn append_counted_unicode(
    output: &mut Vec<u8>,
    value: &str,
    maximum: usize,
    record_type: u16,
) -> Result<()> {
    if value.chars().any(|character| character == '\0') {
        return invalid(record_type, "Unicode string cannot contain NUL characters");
    }
    let count = value.encode_utf16().count();
    if count > maximum {
        return invalid(
            record_type,
            format!("Unicode string exceeds {maximum} UTF-16 code units"),
        );
    }
    let count =
        u16::try_from(count).map_err(|_error| Error::Allocation("encoding BIFF8 string"))?;
    output.extend_from_slice(&count.to_le_bytes());
    append_unicode_no_cch(output, value);
    Ok(())
}

fn append_unicode_no_cch(output: &mut Vec<u8>, value: &str) {
    let compressed = value.chars().all(|character| u32::from(character) <= 0xFF);
    output.push(u8::from(!compressed));
    if compressed {
        output.extend(value.chars().map(|character| character as u8));
    } else {
        for unit in value.encode_utf16() {
            output.extend_from_slice(&unit.to_le_bytes());
        }
    }
}

fn encode_short_unicode(value: &str, record_type: u16) -> Result<Vec<u8>> {
    let count = value.encode_utf16().count();
    let count = u8::try_from(count).map_err(|_error| Error::InvalidRecord {
        record_type,
        message: "ExternName short Unicode name exceeds 255 UTF-16 code units".into(),
    })?;
    if value.chars().any(|character| character == '\0') {
        return invalid(
            record_type,
            "ExternName short Unicode name cannot contain NUL characters",
        );
    }
    let mut output = Vec::new();
    output.push(count);
    append_unicode_no_cch(&mut output, value);
    Ok(output)
}

pub(super) fn invalid<T>(record_type: u16, message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidRecord {
        record_type,
        message: message.into(),
    })
}
