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
    #[must_use]
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
    #[must_use]
    pub fn built_in(&self) -> BuiltInFunctionCategories {
        self.built_in
    }
    #[must_use]
    pub fn custom_categories(&self) -> &[String] {
        &self.custom_categories
    }
    #[must_use]
    pub fn classic_categories(&self) -> &[String] {
        &self.custom_categories[..self.classic_category_count]
    }
    #[must_use]
    pub fn extended_categories(&self) -> &[String] {
        &self.custom_categories[self.classic_category_count..]
    }
}

/// Where the workbook-global record stream stands relative to its one
/// function-group block (BuiltInFnGroupCount followed directly by the
/// FnGroupName and FnGrp12 records).
///
/// ```text
/// Before --BuiltInFnGroupCount--> Open --any other record--> Closed
/// Open   --FnGroupName / FnGrp12--> Open
/// ```
///
/// A name record is refused in `Before`, a second BuiltInFnGroupCount in
/// `Open`, and every function-group record in `Closed`.
#[derive(Debug, Default, Clone, Copy)]
enum FunctionGroupBlockState {
    #[default]
    Before,
    Open(BuiltInFunctionCategories),
    Closed(BuiltInFunctionCategories),
}

#[derive(Debug, Default)]
pub(crate) struct FunctionGroupCollector {
    state: FunctionGroupBlockState,
    classic: Vec<String>,
    extended: Vec<String>,
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
        if !target {
            if let FunctionGroupBlockState::Open(built_in) = self.state {
                self.state = FunctionGroupBlockState::Closed(built_in);
            }
            return Ok(());
        }
        let built_in = match self.state {
            FunctionGroupBlockState::Before => None,
            FunctionGroupBlockState::Open(built_in) => Some(built_in),
            FunctionGroupBlockState::Closed(_) => {
                return invalid(
                    record_type,
                    "function-group record is outside its contiguous BIFF8 collection",
                );
            },
        };

        match record_type {
            BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE => {
                if built_in.is_some() {
                    return invalid(record_type, "duplicate BuiltInFnGroupCount record");
                }
                if data.len() != 2 {
                    return Err(Error::InvalidLength {
                        expected: 2,
                        found: data.len(),
                    });
                }
                self.state = FunctionGroupBlockState::Open(match read_u16(data, 0)? {
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
                if built_in.is_none() {
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
                if built_in.is_none() {
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
        let built_in = match self.state {
            FunctionGroupBlockState::Before => 0,
            FunctionGroupBlockState::Open(built_in) | FunctionGroupBlockState::Closed(built_in) => {
                built_in.count()
            },
        } as usize;
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
        let built_in = match self.state {
            // No name is collected before the block opens, so there is nothing
            // to report.
            FunctionGroupBlockState::Before => return Ok(None),
            FunctionGroupBlockState::Open(built_in) | FunctionGroupBlockState::Closed(built_in) => {
                built_in
            },
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
            .as_chunks::<2>()
            .0
            .iter()
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
    let bytes = bytes.try_into().map_err(|_error| Error::InvalidLength {
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

    /// One row per record sequence: the outcome of every record and the classic
    /// category count (`None` when no collection was found).
    #[test]
    fn block_contiguity_accepts_and_rejects_the_same_sequences() {
        type Record = (u16, &'static [u8]);
        type Case = (&'static [Record], &'static [&'static str], Option<usize>);
        const COUNT: Record = (BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE, &[14, 0]);
        const BAD_COUNT: Record = (BUILT_IN_FN_GROUP_COUNT_RECORD_TYPE, &[15, 0]);
        const NAME_A: Record = (FN_GROUP_NAME_RECORD_TYPE, &[1, 0, 0, b'A']);
        const NAME_B: Record = (FN_GROUP_NAME_RECORD_TYPE, &[1, 0, 0, b'B']);
        const GRP12: Record = (FN_GRP12_RECORD_TYPE, &[0; 14]);
        const OTHER: Record = (0x0018, &[]);
        const OUTSIDE_COUNT: &str = "Invalid record 0x009C: function-group record is outside its contiguous BIFF8 collection";
        const OUTSIDE_NAME: &str = "Invalid record 0x009A: function-group record is outside its contiguous BIFF8 collection";
        const OUTSIDE_GRP12: &str = "Invalid record 0x0898: function-group record is outside its contiguous BIFF8 collection";
        const DUPLICATE: &str = "Invalid record 0x009C: duplicate BuiltInFnGroupCount record";
        const NAME_ALONE: &str =
            "Invalid record 0x009A: FnGroupName appears without BuiltInFnGroupCount";
        const GRP12_ALONE: &str =
            "Invalid record 0x0898: FnGrp12 appears without BuiltInFnGroupCount";
        const BAD_COUNT_ERROR: &str = "Invalid record 0x009C: built-in function category count must be 14, 16, or compatibility value 17, got 15";
        let cases: &[Case] = &[
            (&[], &[], None),
            (&[OTHER, OTHER], &["ok", "ok"], None),
            (&[COUNT], &["ok"], Some(0)),
            (&[OTHER, COUNT, NAME_A, NAME_B], &["ok"; 4], Some(2)),
            (&[COUNT, NAME_A, OTHER, OTHER], &["ok"; 4], Some(1)),
            (
                &[COUNT, NAME_A, OTHER, NAME_B],
                &["ok", "ok", "ok", OUTSIDE_NAME],
                Some(1),
            ),
            (
                &[COUNT, OTHER, COUNT],
                &["ok", "ok", OUTSIDE_COUNT],
                Some(0),
            ),
            (
                &[COUNT, OTHER, GRP12],
                &["ok", "ok", OUTSIDE_GRP12],
                Some(0),
            ),
            (&[COUNT, COUNT, NAME_A], &["ok", DUPLICATE, "ok"], Some(1)),
            (&[NAME_A], &[NAME_ALONE], None),
            (&[GRP12], &[GRP12_ALONE], None),
            (
                &[OTHER, NAME_A, OTHER, COUNT, NAME_A],
                &["ok", NAME_ALONE, "ok", "ok", "ok"],
                Some(1),
            ),
            (
                &[BAD_COUNT, OTHER, NAME_A, COUNT, NAME_A],
                &[BAD_COUNT_ERROR, "ok", NAME_ALONE, "ok", "ok"],
                Some(1),
            ),
        ];
        for (records, expected, classic) in cases {
            let mut collector = FunctionGroupCollector::new();
            let outcomes: Vec<String> = records
                .iter()
                .map(
                    |(record_type, data)| match collector.feed_record(*record_type, data) {
                        Ok(()) => "ok".to_string(),
                        Err(error) => error.to_string(),
                    },
                )
                .collect();
            assert_eq!(outcomes, *expected, "{records:?}");
            let groups = collector.finish().unwrap();
            assert_eq!(
                groups.map(|groups| groups.classic_categories().len()),
                *classic,
                "{records:?}"
            );
        }
    }
}
