//! BIFF8 workbook function-category metadata.

use std::collections::HashSet;

use super::{Error, Result};

pub(crate) const FN_GROUP_NAME_RECORD_TYPE: u16 = 0x009a;
pub(crate) const BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE: u16 = 0x009c;
pub(crate) const FN_GRP12_RECORD_TYPE: u16 = 0x0898;

const FN_GRP12_HEADER_LEN: usize = 12;
const UNICODE_STRING_HEADER_LEN: usize = 3;
const MAX_FUNCTION_CATEGORY_NAME_LENGTH: usize = 32;
const MAX_FUNCTION_CATEGORIES: usize = 256;

/// The built-in function category set recorded by BIFF8.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BuiltInFunctionCategories {
    Fourteen,
    Sixteen,
    /// Interoperability value emitted by established BIFF8 producers even
    /// though current MS-XLS only enumerates 14 and 16.
    SeventeenCompatibility,
}

impl BuiltInFunctionCategories {
    pub const fn count(self) -> u16 {
        match self {
            Self::Fourteen => 14,
            Self::Sixteen => 16,
            Self::SeventeenCompatibility => 17,
        }
    }
}

/// Built-in and user-defined workbook function categories.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FunctionGroups {
    built_in: BuiltInFunctionCategories,
    custom_categories: Vec<String>,
    classic_category_count: usize,
}

impl FunctionGroups {
    pub fn built_in(&self) -> BuiltInFunctionCategories {
        self.built_in
    }
    pub fn custom_categories(&self) -> &[String] {
        &self.custom_categories
    }
    pub fn classic_categories(&self) -> &[String] {
        &self.custom_categories[..self.classic_category_count]
    }
    pub fn extended_categories(&self) -> &[String] {
        &self.custom_categories[self.classic_category_count..]
    }
}

#[derive(Debug, Default)]
pub(crate) struct FunctionGroupCollector {
    built_in: Option<BuiltInFunctionCategories>,
    classic: Vec<String>,
    extended: Vec<String>,
    closed: bool,
}

impl FunctionGroupCollector {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> Result<()> {
        let target = matches!(
            record_type,
            BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE | FN_GROUP_NAME_RECORD_TYPE | FN_GRP12_RECORD_TYPE
        );
        if self.built_in.is_some() && !target {
            self.closed = true;
            return Ok(());
        }
        if !target {
            return Ok(());
        }
        if self.closed {
            return invalid(
                record_type,
                "function-group record is outside its contiguous BIFF8 collection",
            );
        }

        match record_type {
            BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE => {
                if self.built_in.is_some() {
                    return invalid(record_type, "duplicate BuiltInFnGroupCount record");
                }
                if data.len() != 2 {
                    return Err(Error::InvalidLength {
                        expected: 2,
                        found: data.len(),
                    });
                }
                self.built_in = Some(match read_u16(data, 0)? {
                    14 => BuiltInFunctionCategories::Fourteen,
                    16 => BuiltInFunctionCategories::Sixteen,
                    17 => BuiltInFunctionCategories::SeventeenCompatibility,
                    value => {
                        return invalid(
                            record_type,
                            format!(
                                "built-in function category count must be 14, 16, or compatibility value 17, got {value}"
                            ),
                        );
                    },
                });
            },
            FN_GROUP_NAME_RECORD_TYPE => {
                if self.built_in.is_none() {
                    return invalid(
                        record_type,
                        "FnGroupName appears without BuiltInFnGroupCount",
                    );
                }
                if !self.extended.is_empty() {
                    return invalid(record_type, "FnGroupName must precede all FnGrp12 records");
                }
                self.classic.push(parse_unicode_string(record_type, data)?);
            },
            FN_GRP12_RECORD_TYPE => {
                if self.built_in.is_none() {
                    return invalid(record_type, "FnGrp12 appears without BuiltInFnGroupCount");
                }
                if data.len() < FN_GRP12_HEADER_LEN + UNICODE_STRING_HEADER_LEN {
                    return Err(Error::InvalidLength {
                        expected: FN_GRP12_HEADER_LEN + UNICODE_STRING_HEADER_LEN,
                        found: data.len(),
                    });
                }
                let reserved = data
                    .get(4..FN_GRP12_HEADER_LEN)
                    .ok_or(Error::InvalidLength {
                        expected: FN_GRP12_HEADER_LEN,
                        found: data.len(),
                    })?;
                if read_u16(data, 0)? != FN_GRP12_RECORD_TYPE
                    || read_u16(data, 2)? != 0
                    || reserved.iter().any(|byte| *byte != 0)
                {
                    return invalid(record_type, "FnGrp12 future-record header is invalid");
                }
                let name = data
                    .get(FN_GRP12_HEADER_LEN..)
                    .ok_or(Error::InvalidLength {
                        expected: FN_GRP12_HEADER_LEN,
                        found: data.len(),
                    })?;
                self.extended.push(parse_unicode_string(record_type, name)?);
            },
            _ => {
                return invalid(record_type, "unsupported function-group record");
            },
        }
        self.validate_resource_bounds(record_type)
    }

    fn validate_resource_bounds(&self, record_type: u16) -> Result<()> {
        let built_in = self.built_in.map_or(0, BuiltInFunctionCategories::count) as usize;
        let total = built_in
            .checked_add(self.classic.len())
            .and_then(|total| total.checked_add(self.extended.len()))
            .ok_or_else(|| Error::InvalidRecord {
                record_type,
                message: "function category count overflows usize".to_string(),
            })?;
        if total > MAX_FUNCTION_CATEGORIES {
            return invalid(record_type, "function category count exceeds 256");
        }
        Ok(())
    }

    pub(crate) fn finish(self) -> Result<Option<FunctionGroups>> {
        let Some(built_in) = self.built_in else {
            if self.classic.is_empty() && self.extended.is_empty() {
                return Ok(None);
            }
            return invalid(
                BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE,
                "function categories lack BuiltInFnGroupCount",
            );
        };
        let classic_limit = 32usize
            .checked_sub(usize::from(built_in.count()))
            .ok_or_else(|| Error::InvalidRecord {
                record_type: BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE,
                message: "built-in function category count exceeds the classic range".to_string(),
            })?;
        if self.classic.len() > classic_limit {
            return invalid(
                FN_GROUP_NAME_RECORD_TYPE,
                "FnGroupName category index exceeds 31",
            );
        }
        if !self.extended.is_empty() && self.classic.len() != classic_limit {
            return invalid(
                FN_GRP12_RECORD_TYPE,
                "FnGrp12 categories must begin at category index 32",
            );
        }
        let unique_capacity = self
            .classic
            .len()
            .checked_add(self.extended.len())
            .ok_or_else(|| Error::InvalidRecord {
                record_type: FN_GROUP_NAME_RECORD_TYPE,
                message: "function category count overflows usize".to_string(),
            })?;
        let mut unique = HashSet::with_capacity(unique_capacity);
        for name in self.classic.iter().chain(&self.extended) {
            if !unique.insert(name) {
                return invalid(
                    FN_GROUP_NAME_RECORD_TYPE,
                    "function category names must be unique",
                );
            }
        }
        let classic_category_count = self.classic.len();
        let mut custom_categories = self.classic;
        custom_categories.extend(self.extended);
        Ok(Some(FunctionGroups {
            built_in,
            custom_categories,
            classic_category_count,
        }))
    }
}

fn parse_unicode_string(record_type: u16, data: &[u8]) -> Result<String> {
    if data.len() < UNICODE_STRING_HEADER_LEN {
        return Err(Error::InvalidLength {
            expected: UNICODE_STRING_HEADER_LEN,
            found: data.len(),
        });
    }
    let character_count = usize::from(read_u16(data, 0)?);
    if character_count > MAX_FUNCTION_CATEGORY_NAME_LENGTH {
        return invalid(record_type, "function category name exceeds 32 characters");
    }
    let flags = data.get(2).copied().ok_or(Error::InvalidLength {
        expected: UNICODE_STRING_HEADER_LEN,
        found: data.len(),
    })?;
    if flags & !1 != 0 {
        return invalid(record_type, "XLUnicodeString reserved flags must be zero");
    }
    let wide = flags == 1;
    let character_bytes = character_count
        .checked_mul(if wide { 2 } else { 1 })
        .ok_or_else(|| Error::InvalidRecord {
            record_type,
            message: "function category name byte length overflows usize".to_string(),
        })?;
    let expected = UNICODE_STRING_HEADER_LEN
        .checked_add(character_bytes)
        .ok_or_else(|| Error::InvalidRecord {
            record_type,
            message: "function category record length overflows usize".to_string(),
        })?;
    if data.len() != expected {
        return Err(Error::InvalidLength {
            expected,
            found: data.len(),
        });
    }
    let encoded = data
        .get(UNICODE_STRING_HEADER_LEN..)
        .ok_or(Error::InvalidLength {
            expected: UNICODE_STRING_HEADER_LEN,
            found: data.len(),
        })?;
    if wide {
        let units = encoded
            .chunks_exact(2)
            .map(|chunk| read_u16(chunk, 0))
            .collect::<Result<Vec<_>>>()?;
        String::from_utf16(&units).map_err(|error| Error::InvalidRecord {
            record_type,
            message: format!("invalid UTF-16 function category name: {error}"),
        })
    } else {
        Ok(encoded.iter().map(|byte| char::from(*byte)).collect())
    }
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16> {
    let end = offset.checked_add(2).ok_or(Error::InvalidLength {
        expected: 2,
        found: data.len(),
    })?;
    let bytes = data.get(offset..end).ok_or(Error::InvalidLength {
        expected: end,
        found: data.len(),
    })?;
    let bytes = bytes.try_into().map_err(|_| Error::InvalidLength {
        expected: 2,
        found: bytes.len(),
    })?;
    Ok(u16::from_le_bytes(bytes))
}

fn invalid<T>(record_type: u16, message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidRecord {
        record_type,
        message: message.into(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_invalid_count_strings_and_future_header() {
        let mut collector = FunctionGroupCollector::new();
        collector
            .feed_record(BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE, &[17, 0])
            .unwrap();
        assert_eq!(
            collector.finish().unwrap().unwrap().built_in(),
            BuiltInFunctionCategories::SeventeenCompatibility,
        );

        let mut collector = FunctionGroupCollector::new();
        assert!(
            collector
                .feed_record(BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE, &[15, 0])
                .is_err()
        );

        let mut collector = FunctionGroupCollector::new();
        collector
            .feed_record(BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE, &[14, 0])
            .unwrap();
        assert!(
            collector
                .feed_record(FN_GROUP_NAME_RECORD_TYPE, &[1, 0, 0x80, b'A'])
                .is_err()
        );

        let mut collector = FunctionGroupCollector::new();
        collector
            .feed_record(BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE, &[16, 0])
            .unwrap();
        let mut invalid_header = vec![0; 15];
        invalid_header[0..2].copy_from_slice(&FN_GRP12_RECORD_TYPE.to_le_bytes());
        invalid_header[2] = 1;
        assert!(
            collector
                .feed_record(FN_GRP12_RECORD_TYPE, &invalid_header)
                .is_err()
        );
    }

    #[test]
    fn rejects_order_cardinality_and_duplicates() {
        let mut collector = FunctionGroupCollector::new();
        collector
            .feed_record(BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE, &[14, 0])
            .unwrap();
        collector.feed_record(0x0018, &[]).unwrap();
        assert!(
            collector
                .feed_record(FN_GROUP_NAME_RECORD_TYPE, &[1, 0, 0, b'A'])
                .is_err()
        );

        let mut collector = FunctionGroupCollector::new();
        collector
            .feed_record(BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE, &[14, 0])
            .unwrap();
        collector
            .feed_record(FN_GROUP_NAME_RECORD_TYPE, &[1, 0, 0, b'A'])
            .unwrap();
        collector
            .feed_record(FN_GROUP_NAME_RECORD_TYPE, &[1, 0, 0, b'A'])
            .unwrap();
        assert!(collector.finish().is_err());
    }

    #[test]
    fn rejects_truncated_and_overflowing_fixed_width_reads() {
        assert!(read_u16(&[], 0).is_err());
        assert!(read_u16(&[], usize::MAX).is_err());

        let mut collector = FunctionGroupCollector::new();
        assert!(
            collector
                .feed_record(BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE, &[14])
                .is_err()
        );

        collector
            .feed_record(BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE, &[14, 0])
            .unwrap();
        assert!(
            collector
                .feed_record(FN_GRP12_RECORD_TYPE, &[0; 14])
                .is_err()
        );
    }
}
