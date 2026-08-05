//! BIFF8 `BookExt` record (MS-XLS 2.4.23): workbook extension flags.
//!
//! The record carries AutoRecover/privacy/smart-tag/recovery flags and,
//! depending on the declared record size, the `BookExt_Conditional11`
//! (MS-XLS 2.5.12) and `BookExt_Conditional12` (MS-XLS 2.5.13) extensions.

use super::{Error, Result};

/// Record type of the `BookExt` record.
pub(crate) const BOOK_EXT_RECORD_TYPE: u16 = 0x0863;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// Fixed payload size: `FrtHeader` + `cb` + the 4-byte flag bitfield.
const BASE_LEN: usize = 20;
/// Maximum payload size with both conditional extensions.
const MAX_LEN: usize = 22;

// Flag bitfield (4 bytes).
const DONT_AUTO_RECOVER: u32 = 0x0001;
const HIDE_PIVOT_LIST: u32 = 0x0002;
const FILTER_PRIVACY: u32 = 0x0004;
const EMBED_FACTOIDS: u32 = 0x0008;
const FACTOID_DISPLAY_SHIFT: u32 = 4;
const FACTOID_DISPLAY_MASK: u32 = 0x3;
const SAVED_DURING_RECOVERY: u32 = 0x0040;
const CREATED_VIA_MINIMAL_SAVE: u32 = 0x0080;
const OPENED_VIA_DATA_RECOVERY: u32 = 0x0100;
const OPENED_VIA_SAFE_LOAD: u32 = 0x0200;

// BookExt_Conditional11 (grbit1).
const BUGGED_USER_ABOUT_SOLUTION: u8 = 0x01;
const SHOW_INK_ANNOTATION: u8 = 0x02;

// BookExt_Conditional12 (grbit2).
const PUBLISHED_BOOK_ITEMS: u8 = 0x02;

/// How smart tags are displayed in the workbook (`mdFactoidDisplay`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum FactoidDisplay {
    /// Show the smart-tag actions button and the smart-tag indicator.
    #[default]
    IndicatorAndButton,
    /// Show the smart-tag actions button only.
    ButtonOnly,
    /// Show neither the actions button nor the indicator.
    Hidden,
}

impl FactoidDisplay {
    fn from_code(value: u32) -> Result<Self> {
        Ok(match value {
            0 => Self::IndicatorAndButton,
            1 => Self::ButtonOnly,
            2 => Self::Hidden,
            _ => {
                return Err(Error::InvalidRecord {
                    record_type: BOOK_EXT_RECORD_TYPE,
                    message: "mdFactoidDisplay value 3 is reserved".to_string(),
                });
            },
        })
    }

    fn code(self) -> u32 {
        match self {
            Self::IndicatorAndButton => 0,
            Self::ButtonOnly => 1,
            Self::Hidden => 2,
        }
    }
}

/// `BookExt_Conditional11` extension flags (MS-XLS 2.5.12).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct BookExtConditional11 {
    /// Whether a warning is requested before loading a smart-document
    /// manifest (`fBuggedUserAboutSolution`).
    pub bugged_user_about_solution: bool,
    /// Whether ink comments are visible in the workbook (`fShowInkAnnotation`).
    pub show_ink_annotation: bool,
}

/// `BookExt_Conditional12` extension flags (MS-XLS 2.5.13).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct BookExtConditional12 {
    /// Whether only selected items are shown when the workbook is published
    /// to a server (`fPublishedBookItems`).
    pub published_book_items: bool,
}

/// Typed `BookExt` record content (MS-XLS 2.4.23).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct BookExt {
    /// Whether AutoRecover is disabled for the workbook.
    pub dont_auto_recover: bool,
    /// Whether the PivotTable field list is hidden.
    pub hide_pivot_list: bool,
    /// Whether personal information is removed on save.
    pub filter_privacy: bool,
    /// Whether smart tags are embedded on save.
    pub embed_factoids: bool,
    /// How smart tags are displayed.
    pub factoid_display: FactoidDisplay,
    /// Whether the workbook was saved during AutoRecover.
    pub saved_during_recovery: bool,
    /// Whether the workbook was created by a minimal save during recovery.
    pub created_via_minimal_save: bool,
    /// Whether the workbook was opened through data recovery.
    pub opened_via_data_recovery: bool,
    /// Whether the workbook was opened in safe-load mode.
    pub opened_via_safe_load: bool,
    /// Conditional11 extension, present iff the record size demands it.
    pub conditional11: Option<BookExtConditional11>,
    /// Conditional12 extension, present iff the record size demands it.
    pub conditional12: Option<BookExtConditional12>,
}

impl BookExt {
    /// Parse a `BookExt` record payload.
    pub(crate) fn parse(data: &[u8]) -> Result<Self> {
        let invalid = |message: &str| Error::InvalidRecord {
            record_type: BOOK_EXT_RECORD_TYPE,
            message: message.to_string(),
        };
        if data.len() < BASE_LEN {
            return Err(Error::InvalidLength {
                expected: BASE_LEN,
                found: data.len(),
            });
        }
        if u16::from_le_bytes([data[0], data[1]]) != BOOK_EXT_RECORD_TYPE {
            return Err(invalid("BookExt FrtHeader.rt mismatch"));
        }
        let cb = u32::from_le_bytes(data[12..16].try_into().expect("length checked")) as usize;
        if !(BASE_LEN..=MAX_LEN).contains(&cb) {
            return Err(invalid("BookExt declares an unsupported record size"));
        }
        if data.len() != cb {
            return Err(Error::InvalidLength {
                expected: cb,
                found: data.len(),
            });
        }
        let flags = u32::from_le_bytes(data[16..20].try_into().expect("length checked"));
        let conditional11 = if cb > BASE_LEN {
            let grbit = data[BASE_LEN];
            Some(BookExtConditional11 {
                bugged_user_about_solution: grbit & BUGGED_USER_ABOUT_SOLUTION != 0,
                show_ink_annotation: grbit & SHOW_INK_ANNOTATION != 0,
            })
        } else {
            None
        };
        let conditional12 = if cb > BASE_LEN + 1 {
            let grbit = data[BASE_LEN + 1];
            Some(BookExtConditional12 {
                published_book_items: grbit & PUBLISHED_BOOK_ITEMS != 0,
            })
        } else {
            None
        };
        Ok(Self {
            dont_auto_recover: flags & DONT_AUTO_RECOVER != 0,
            hide_pivot_list: flags & HIDE_PIVOT_LIST != 0,
            filter_privacy: flags & FILTER_PRIVACY != 0,
            embed_factoids: flags & EMBED_FACTOIDS != 0,
            factoid_display: FactoidDisplay::from_code(
                (flags >> FACTOID_DISPLAY_SHIFT) & FACTOID_DISPLAY_MASK,
            )?,
            saved_during_recovery: flags & SAVED_DURING_RECOVERY != 0,
            created_via_minimal_save: flags & CREATED_VIA_MINIMAL_SAVE != 0,
            opened_via_data_recovery: flags & OPENED_VIA_DATA_RECOVERY != 0,
            opened_via_safe_load: flags & OPENED_VIA_SAFE_LOAD != 0,
            conditional11,
            conditional12,
        })
    }

    /// Serialize back to a complete `BookExt` record payload.
    pub(crate) fn to_payload(&self) -> Vec<u8> {
        let len = BASE_LEN
            + usize::from(self.conditional11.is_some())
            + usize::from(self.conditional12.is_some());
        let mut flags = self.factoid_display.code() << FACTOID_DISPLAY_SHIFT;
        if self.dont_auto_recover {
            flags |= DONT_AUTO_RECOVER;
        }
        if self.hide_pivot_list {
            flags |= HIDE_PIVOT_LIST;
        }
        if self.filter_privacy {
            flags |= FILTER_PRIVACY;
        }
        if self.embed_factoids {
            flags |= EMBED_FACTOIDS;
        }
        if self.saved_during_recovery {
            flags |= SAVED_DURING_RECOVERY;
        }
        if self.created_via_minimal_save {
            flags |= CREATED_VIA_MINIMAL_SAVE;
        }
        if self.opened_via_data_recovery {
            flags |= OPENED_VIA_DATA_RECOVERY;
        }
        if self.opened_via_safe_load {
            flags |= OPENED_VIA_SAFE_LOAD;
        }
        let mut payload = Vec::with_capacity(len);
        payload.extend_from_slice(&BOOK_EXT_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        payload.extend_from_slice(&(len as u32).to_le_bytes());
        payload.extend_from_slice(&flags.to_le_bytes());
        if let Some(conditional11) = &self.conditional11 {
            let mut grbit = 0u8;
            if conditional11.bugged_user_about_solution {
                grbit |= BUGGED_USER_ABOUT_SOLUTION;
            }
            if conditional11.show_ink_annotation {
                grbit |= SHOW_INK_ANNOTATION;
            }
            payload.push(grbit);
        }
        if let Some(conditional12) = &self.conditional12 {
            payload.push(if conditional12.published_book_items {
                PUBLISHED_BOOK_ITEMS
            } else {
                0
            });
        }
        payload
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(flags: u32, conditional: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&BOOK_EXT_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; 10]);
        data.extend_from_slice(&((BASE_LEN + conditional.len()) as u32).to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(conditional);
        data
    }

    #[test]
    fn parses_base_record() {
        let flags = DONT_AUTO_RECOVER
            | FILTER_PRIVACY
            | (2 << FACTOID_DISPLAY_SHIFT)
            | SAVED_DURING_RECOVERY
            | OPENED_VIA_SAFE_LOAD;
        let parsed = BookExt::parse(&record(flags, &[])).unwrap();
        assert!(parsed.dont_auto_recover);
        assert!(!parsed.hide_pivot_list);
        assert!(parsed.filter_privacy);
        assert!(!parsed.embed_factoids);
        assert_eq!(parsed.factoid_display, FactoidDisplay::Hidden);
        assert!(parsed.saved_during_recovery);
        assert!(!parsed.created_via_minimal_save);
        assert!(!parsed.opened_via_data_recovery);
        assert!(parsed.opened_via_safe_load);
        assert!(parsed.conditional11.is_none());
        assert!(parsed.conditional12.is_none());
        assert_eq!(parsed.to_payload(), record(flags, &[]));
    }

    #[test]
    fn parses_conditional_extensions() {
        let parsed = BookExt::parse(&record(0, &[0x03, 0x02])).unwrap();
        let conditional11 = parsed.conditional11.unwrap();
        assert!(conditional11.bugged_user_about_solution);
        assert!(conditional11.show_ink_annotation);
        assert!(parsed.conditional12.unwrap().published_book_items);
        assert_eq!(parsed.to_payload(), record(0, &[0x03, 0x02]));

        // Conditional11 without Conditional12.
        let parsed = BookExt::parse(&record(0, &[0x02])).unwrap();
        assert!(parsed.conditional11.unwrap().show_ink_annotation);
        assert!(parsed.conditional12.is_none());
        assert_eq!(parsed.to_payload(), record(0, &[0x02]));
    }

    #[test]
    fn rejects_malformed_records() {
        // Truncated.
        assert!(BookExt::parse(&record(0, &[])[..10]).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = record(0, &[]);
        wrong_rt[0..2].copy_from_slice(&0x0862u16.to_le_bytes());
        assert!(BookExt::parse(&wrong_rt).is_err());
        // Unsupported record size.
        let mut wrong_cb = record(0, &[0, 0]);
        wrong_cb[12..16].copy_from_slice(&23u32.to_le_bytes());
        wrong_cb.push(0);
        assert!(BookExt::parse(&wrong_cb).is_err());
        // Reserved mdFactoidDisplay value.
        assert!(BookExt::parse(&record(3 << FACTOID_DISPLAY_SHIFT, &[])).is_err());
        // Declared size disagreeing with the payload length.
        let mut padded = record(0, &[]);
        padded.push(0);
        assert!(BookExt::parse(&padded).is_err());
    }
}
