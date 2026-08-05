//! `TextSpecialInfoAtom` per-run special information and `MasterTextPropAtom`
//! indent levels (MS-PPT 2.9.54/2.9.55 and 2.9.79/2.9.80).
//!
//! Both atoms annotate the *corresponding text* named by the most closely
//! preceding `TextHeaderAtom`: the first carries language, spelling, and
//! smart-tag information per character run, the second carries the master
//! indent level per character run. They are inert: languages and smart-tag
//! indices are never resolved or applied.

use super::records::record::Record;
use super::text_si_exception::SpellingFlags;
use crate::consts::RecordType;
use crate::package::{Error, Result};

/// `RT_TextSpecialInfoAtom` record type (MS-PPT 2.9.54).
const TEXT_SPECIAL_INFO_TYPE: u16 = 0x0FAA;
/// `RT_MasterTextPropAtom` record type (MS-PPT 2.9.79).
const MASTER_TEXT_PROP_TYPE: u16 = 0x0FA2;

// TextSIException mask bits (MS-PPT 2.9.32).
const MASK_SPELL: u32 = 0x0001;
const MASK_LANG: u32 = 0x0002;
const MASK_ALT_LANG: u32 = 0x0004;
const MASK_PP10_EXT: u32 = 0x0020;
const MASK_BIDI: u32 = 0x0040;
const MASK_SMART_TAG: u32 = 0x0200;
/// Bits that must be zero in a `TextSIException`: `reserved1` and
/// `reserved2`. The `unused1`, `unused2`, and `unused3` bits are undefined
/// and ignored.
const MASK_FORBIDDEN: u32 = 0xFFFF_FD00;

/// `SpellingFlags.grammar`: must be zero in a `TextSIException`.
const SPELL_GRAMMAR: u16 = 0x0004;
/// `SpellingFlags` reserved bits: must be zero.
const SPELL_RESERVED: u16 = 0xFFF8;
/// `SpellingFlags.error`.
const SPELL_ERROR: u16 = 0x0001;
/// `SpellingFlags.clean`.
const SPELL_CLEAN: u16 = 0x0002;

// PP10 extension word layout (MS-PPT 2.9.32): `pp10runid` occupies bits 0-3,
// `reserved3` occupies bits 4-30 and must be zero, `grammarError` is bit 31.
const PP10_RUN_ID_MASK: u32 = 0x0000_000F;
const PP10_RESERVED_MASK: u32 = 0x7FFF_FFF0;
const PP10_GRAMMAR_ERROR: u32 = 0x8000_0000;

/// Size in bytes of one `SmartTagIndex` entry (MS-PPT 2.2.26).
const SMART_TAG_INDEX_LEN: usize = 4;
/// Size in bytes of one `MasterTextPropRun` (MS-PPT 2.9.80).
const MASTER_TEXT_PROP_RUN_LEN: usize = 6;
/// Largest valid `IndentLevel` value (MS-PPT 2.2.13).
const MAX_INDENT_LEVEL: u16 = 0x0004;

fn corrupted(message: impl Into<String>) -> Error {
    Error::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(data[offset..offset + 2].try_into().expect("length checked"))
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().expect("length checked"))
}

fn require_bytes(data: &[u8], offset: usize, needed: usize, field: &str) -> Result<()> {
    if data.len() < offset.saturating_add(needed) {
        return Err(corrupted(format!("{field} is truncated")));
    }
    Ok(())
}

/// A full `TextSIException` structure (MS-PPT 2.9.32) as stored in a
/// `TextSIRun`. Unlike the defaults atom (MS-PPT 2.9.31), run-level special
/// information may carry bidirectional flags, PowerPoint 10 extension data,
/// and smart-tag indices.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextSIException {
    /// Spelling status, when present.
    spelling: Option<SpellingFlags>,
    /// Primary language identifier (`TxLCID`), when present.
    language: Option<u16>,
    /// Alternate language identifier (`TxLCID`), when present.
    alternate_language: Option<u16>,
    /// Whether the text contains bidirectional characters, when present.
    bidi: Option<bool>,
    /// Four-bit run identifier used by `StyleTextProp11Atom`, when present.
    pp10_run_id: Option<u8>,
    /// Whether a grammar error is flagged, when the PP10 extension exists.
    grammar_error: Option<bool>,
    /// Zero-based indices into the document-wide smart-tag store
    /// (`SmartTagStore11Container`, MS-PPT 2.11.28).
    smart_tag_indices: Vec<u32>,
}

impl TextSIException {
    /// Spelling status, when present.
    pub const fn spelling(&self) -> Option<SpellingFlags> {
        self.spelling
    }
    /// Primary language identifier, when present.
    pub const fn language(&self) -> Option<u16> {
        self.language
    }
    /// Alternate language identifier, when present.
    pub const fn alternate_language(&self) -> Option<u16> {
        self.alternate_language
    }
    /// Whether the text contains bidirectional characters, when present.
    pub const fn bidi(&self) -> Option<bool> {
        self.bidi
    }
    /// Four-bit `StyleTextProp11Atom` run identifier, when present.
    pub const fn pp10_run_id(&self) -> Option<u8> {
        self.pp10_run_id
    }
    /// Whether a grammar error is flagged, when the PP10 extension exists.
    pub const fn grammar_error(&self) -> Option<bool> {
        self.grammar_error
    }
    /// Zero-based indices into the document-wide smart-tag store.
    pub fn smart_tag_indices(&self) -> &[u32] {
        &self.smart_tag_indices
    }

    /// Parse one `TextSIException` from the start of `data`.
    ///
    /// Returns the decoded structure and the number of bytes consumed.
    pub(crate) fn parse_prefix(data: &[u8]) -> Result<(Self, usize)> {
        require_bytes(data, 0, 4, "TextSIException mask")?;
        let mask = read_u32(data, 0);
        if mask & MASK_FORBIDDEN != 0 {
            return Err(corrupted("TextSIException has reserved mask bits set"));
        }
        let mut offset = 4usize;

        let spelling = if mask & MASK_SPELL != 0 {
            require_bytes(data, offset, 2, "TextSIException spellInfo")?;
            let raw = read_u16(data, offset);
            offset += 2;
            if raw & (SPELL_GRAMMAR | SPELL_RESERVED) != 0 {
                return Err(corrupted(
                    "TextSIException spellInfo has grammar or reserved bits set",
                ));
            }
            Some(SpellingFlags::from_bits(
                raw & SPELL_ERROR != 0,
                raw & SPELL_CLEAN != 0,
            ))
        } else {
            None
        };
        let language = if mask & MASK_LANG != 0 {
            require_bytes(data, offset, 2, "TextSIException lid")?;
            let value = read_u16(data, offset);
            offset += 2;
            Some(value)
        } else {
            None
        };
        let alternate_language = if mask & MASK_ALT_LANG != 0 {
            require_bytes(data, offset, 2, "TextSIException altLid")?;
            let value = read_u16(data, offset);
            offset += 2;
            Some(value)
        } else {
            None
        };
        let bidi = if mask & MASK_BIDI != 0 {
            require_bytes(data, offset, 2, "TextSIException bidi")?;
            let value = read_u16(data, offset);
            offset += 2;
            match value {
                0 => Some(false),
                1 => Some(true),
                _ => return Err(corrupted("TextSIException has an invalid bidi flag")),
            }
        } else {
            None
        };
        let (pp10_run_id, grammar_error) = if mask & MASK_PP10_EXT != 0 {
            require_bytes(data, offset, 4, "TextSIException PP10 extension")?;
            let value = read_u32(data, offset);
            offset += 4;
            if value & PP10_RESERVED_MASK != 0 {
                return Err(corrupted(
                    "TextSIException PP10 extension has reserved bits set",
                ));
            }
            (
                Some((value & PP10_RUN_ID_MASK) as u8),
                Some(value & PP10_GRAMMAR_ERROR != 0),
            )
        } else {
            (None, None)
        };
        let smart_tag_indices = if mask & MASK_SMART_TAG != 0 {
            require_bytes(data, offset, 4, "TextSIException smart-tag count")?;
            let count = read_u32(data, offset) as usize;
            offset += 4;
            let remaining = (data.len() - offset) / SMART_TAG_INDEX_LEN;
            if count > remaining {
                return Err(corrupted("TextSIException smart-tag indices are truncated"));
            }
            let mut indices = Vec::with_capacity(count);
            for _ in 0..count {
                indices.push(read_u32(data, offset));
                offset += SMART_TAG_INDEX_LEN;
            }
            indices
        } else {
            Vec::new()
        };

        Ok((
            Self {
                spelling,
                language,
                alternate_language,
                bidi,
                pp10_run_id,
                grammar_error,
                smart_tag_indices,
            },
            offset,
        ))
    }

    /// Serialize the structure without any record header.
    pub(crate) fn to_bytes(&self) -> Vec<u8> {
        let mut mask = 0u32;
        if self.spelling.is_some() {
            mask |= MASK_SPELL;
        }
        if self.language.is_some() {
            mask |= MASK_LANG;
        }
        if self.alternate_language.is_some() {
            mask |= MASK_ALT_LANG;
        }
        if self.bidi.is_some() {
            mask |= MASK_BIDI;
        }
        if self.pp10_run_id.is_some() {
            mask |= MASK_PP10_EXT;
        }
        if !self.smart_tag_indices.is_empty() {
            mask |= MASK_SMART_TAG;
        }
        let mut data = Vec::with_capacity(16);
        data.extend_from_slice(&mask.to_le_bytes());
        if let Some(spelling) = self.spelling {
            let raw = u16::from(spelling.error()) | (u16::from(spelling.clean()) << 1);
            data.extend_from_slice(&raw.to_le_bytes());
        }
        if let Some(language) = self.language {
            data.extend_from_slice(&language.to_le_bytes());
        }
        if let Some(alternate_language) = self.alternate_language {
            data.extend_from_slice(&alternate_language.to_le_bytes());
        }
        if let Some(bidi) = self.bidi {
            data.extend_from_slice(&u16::from(bidi).to_le_bytes());
        }
        if let Some(pp10_run_id) = self.pp10_run_id {
            let mut value = u32::from(pp10_run_id);
            if self.grammar_error == Some(true) {
                value |= PP10_GRAMMAR_ERROR;
            }
            data.extend_from_slice(&value.to_le_bytes());
        }
        if !self.smart_tag_indices.is_empty() {
            data.extend_from_slice(&(self.smart_tag_indices.len() as u32).to_le_bytes());
            for index in &self.smart_tag_indices {
                data.extend_from_slice(&index.to_le_bytes());
            }
        }
        data
    }
}

/// One `TextSIRun` of language and spelling information (MS-PPT 2.9.55).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextSIRun {
    /// Number of characters the special information applies to.
    count: u32,
    /// Language and spelling information for the run.
    special_info: TextSIException,
}

impl TextSIRun {
    /// Number of characters the special information applies to.
    pub const fn count(&self) -> u32 {
        self.count
    }
    /// Language and spelling information for the run.
    pub const fn special_info(&self) -> &TextSIException {
        &self.special_info
    }
}

/// Parsed payload of a `TextSpecialInfoAtom` (MS-PPT 2.9.54).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TextSpecialInfoRuns {
    runs: Vec<TextSIRun>,
}

impl TextSpecialInfoRuns {
    /// The parsed runs, in source order.
    pub fn runs(&self) -> &[TextSIRun] {
        &self.runs
    }

    /// Total number of characters covered by all runs.
    ///
    /// Per MS-PPT 2.9.54 the sum of the `count` fields must equal the number
    /// of characters in the corresponding text; callers holding that length
    /// can compare it against this total.
    pub fn total_count(&self) -> u64 {
        self.runs.iter().map(|run| u64::from(run.count)).sum()
    }

    /// Parse a complete `TextSpecialInfoAtom` record.
    pub fn parse_record(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::TextSpecInfoAtom
            || record.record_type_raw != TEXT_SPECIAL_INFO_TYPE
            || record.version != 0
            || record.instance != 0
        {
            return Err(corrupted("TextSpecialInfoAtom has an invalid header"));
        }
        Self::parse(&record.data)
    }

    /// Parse the `rgSIRun` payload of a `TextSpecialInfoAtom`.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut offset = 0usize;
        let mut runs = Vec::new();
        while offset < data.len() {
            require_bytes(data, offset, 4, "TextSIRun count")?;
            let count = read_u32(data, offset);
            if count == 0 {
                return Err(corrupted("TextSIRun count must be at least one"));
            }
            offset += 4;
            let (special_info, consumed) = TextSIException::parse_prefix(&data[offset..])?;
            offset += consumed;
            runs.push(TextSIRun {
                count,
                special_info,
            });
        }
        Ok(Self { runs })
    }

    /// Serialize the complete record, including its header.
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut payload = Vec::new();
        for run in &self.runs {
            payload.extend_from_slice(&run.count.to_le_bytes());
            payload.extend_from_slice(&run.special_info.to_bytes());
        }
        let mut data = Vec::with_capacity(8 + payload.len());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&TEXT_SPECIAL_INFO_TYPE.to_le_bytes());
        data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        data.extend_from_slice(&payload);
        data
    }
}

/// One `MasterTextPropRun` of master indent-level information (MS-PPT 2.9.80).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MasterTextPropRun {
    /// Number of characters the indent level applies to.
    count: u32,
    /// `IndentLevel` of the characters; at most 0x0004 (MS-PPT 2.2.13).
    indent_level: u16,
}

impl MasterTextPropRun {
    /// Number of characters the indent level applies to.
    pub const fn count(&self) -> u32 {
        self.count
    }
    /// The master indent level of the characters.
    pub const fn indent_level(&self) -> u16 {
        self.indent_level
    }
}

/// Parsed payload of a `MasterTextPropAtom` (MS-PPT 2.9.79).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MasterTextPropLevels {
    runs: Vec<MasterTextPropRun>,
}

impl MasterTextPropLevels {
    /// The parsed runs, in source order.
    pub fn runs(&self) -> &[MasterTextPropRun] {
        &self.runs
    }

    /// Total number of characters covered by all runs.
    pub fn total_count(&self) -> u64 {
        self.runs.iter().map(|run| u64::from(run.count)).sum()
    }

    /// Parse a complete `MasterTextPropAtom` record.
    pub fn parse_record(record: &Record) -> Result<Self> {
        if record.record_type != RecordType::MasterTextPropAtom
            || record.record_type_raw != MASTER_TEXT_PROP_TYPE
            || record.version != 0
            || record.instance != 0
        {
            return Err(corrupted("MasterTextPropAtom has an invalid header"));
        }
        Self::parse(&record.data)
    }

    /// Parse the `rgMasterTextPropRun` payload of a `MasterTextPropAtom`.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if !data.len().is_multiple_of(MASTER_TEXT_PROP_RUN_LEN) {
            return Err(corrupted(
                "MasterTextPropAtom payload is not a whole number of runs",
            ));
        }
        let mut runs = Vec::with_capacity(data.len() / MASTER_TEXT_PROP_RUN_LEN);
        let mut offset = 0usize;
        while offset < data.len() {
            let count = read_u32(data, offset);
            let indent_level = read_u16(data, offset + 4);
            if indent_level > MAX_INDENT_LEVEL {
                return Err(corrupted("MasterTextPropRun indent level exceeds 0x0004"));
            }
            runs.push(MasterTextPropRun {
                count,
                indent_level,
            });
            offset += MASTER_TEXT_PROP_RUN_LEN;
        }
        Ok(Self { runs })
    }

    /// Serialize the complete record, including its header.
    pub fn to_bytes(&self) -> Vec<u8> {
        let payload_len = self.runs.len() * MASTER_TEXT_PROP_RUN_LEN;
        let mut data = Vec::with_capacity(8 + payload_len);
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&MASTER_TEXT_PROP_TYPE.to_le_bytes());
        data.extend_from_slice(&(payload_len as u32).to_le_bytes());
        for run in &self.runs {
            data.extend_from_slice(&run.count.to_le_bytes());
            data.extend_from_slice(&run.indent_level.to_le_bytes());
        }
        data
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn special_info_atom(data: &[u8]) -> Record {
        Record {
            record_type: RecordType::TextSpecInfoAtom,
            record_type_raw: TEXT_SPECIAL_INFO_TYPE,
            version: 0,
            instance: 0,
            data_length: data.len() as u32,
            data: data.to_vec(),
            children: Vec::new(),
        }
    }

    fn master_prop_atom(data: &[u8]) -> Record {
        Record {
            record_type: RecordType::MasterTextPropAtom,
            record_type_raw: MASTER_TEXT_PROP_TYPE,
            version: 0,
            instance: 0,
            data_length: data.len() as u32,
            data: data.to_vec(),
            children: Vec::new(),
        }
    }

    /// Two runs: the first with spelling and languages, the second with
    /// bidi, a PP10 extension, and smart-tag indices.
    fn sample_special_info_payload() -> Vec<u8> {
        let mut data = Vec::new();
        // Run 1: count = 5, mask = spell | lang | altLang.
        data.extend_from_slice(&5u32.to_le_bytes());
        data.extend_from_slice(&0x0007u32.to_le_bytes());
        data.extend_from_slice(&0x0002u16.to_le_bytes()); // clean
        data.extend_from_slice(&0x0409u16.to_le_bytes()); // en-US
        data.extend_from_slice(&0x0809u16.to_le_bytes()); // en-GB
        // Run 2: count = 3, mask = fPp10ext | fBidi | smartTag.
        data.extend_from_slice(&3u32.to_le_bytes());
        data.extend_from_slice(&0x0260u32.to_le_bytes());
        data.extend_from_slice(&1u16.to_le_bytes()); // bidi
        data.extend_from_slice(&0x8000_0003u32.to_le_bytes()); // run id 3, grammar error
        data.extend_from_slice(&2u32.to_le_bytes()); // two smart tags
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&7u32.to_le_bytes());
        data
    }

    #[test]
    fn parses_special_info_runs_and_round_trips() {
        let payload = sample_special_info_payload();
        let parsed = TextSpecialInfoRuns::parse_record(&special_info_atom(&payload)).unwrap();
        assert_eq!(parsed.runs().len(), 2);
        assert_eq!(parsed.total_count(), 8);

        let first = &parsed.runs()[0];
        assert_eq!(first.count(), 5);
        let si = first.special_info();
        assert!(si.spelling().unwrap().clean());
        assert!(!si.spelling().unwrap().error());
        assert_eq!(si.language(), Some(0x0409));
        assert_eq!(si.alternate_language(), Some(0x0809));
        assert_eq!(si.bidi(), None);
        assert_eq!(si.pp10_run_id(), None);
        assert!(si.smart_tag_indices().is_empty());

        let second = &parsed.runs()[1];
        assert_eq!(second.count(), 3);
        let si = second.special_info();
        assert_eq!(si.spelling(), None);
        assert_eq!(si.bidi(), Some(true));
        assert_eq!(si.pp10_run_id(), Some(3));
        assert_eq!(si.grammar_error(), Some(true));
        assert_eq!(si.smart_tag_indices(), &[0, 7]);

        assert_eq!(parsed.to_bytes()[8..], payload[..]);
    }

    #[test]
    fn parses_empty_special_info_payload() {
        let parsed = TextSpecialInfoRuns::parse_record(&special_info_atom(&[])).unwrap();
        assert!(parsed.runs().is_empty());
        assert_eq!(parsed.total_count(), 0);
        assert_eq!(parsed.to_bytes().len(), 8);
    }

    #[test]
    fn rejects_malformed_special_info() {
        // Wrong record type.
        assert!(TextSpecialInfoRuns::parse_record(&master_prop_atom(&[])).is_err());
        // Truncated run count.
        assert!(TextSpecialInfoRuns::parse_record(&special_info_atom(&[1, 0])).is_err());
        // Zero run count.
        let mut data = Vec::new();
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        assert!(TextSpecialInfoRuns::parse_record(&special_info_atom(&data)).is_err());
        // Reserved mask bits.
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0x0100u32.to_le_bytes());
        assert!(TextSpecialInfoRuns::parse_record(&special_info_atom(&data)).is_err());
        // Grammar bit in spellInfo.
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0x0001u32.to_le_bytes());
        data.extend_from_slice(&0x0004u16.to_le_bytes());
        assert!(TextSpecialInfoRuns::parse_record(&special_info_atom(&data)).is_err());
        // Invalid bidi value.
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0x0040u32.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        assert!(TextSpecialInfoRuns::parse_record(&special_info_atom(&data)).is_err());
        // PP10 reserved bits.
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0x0020u32.to_le_bytes());
        data.extend_from_slice(&0x0010u32.to_le_bytes());
        assert!(TextSpecialInfoRuns::parse_record(&special_info_atom(&data)).is_err());
        // Truncated smart-tag indices.
        let mut data = Vec::new();
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0x0200u32.to_le_bytes());
        data.extend_from_slice(&3u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        assert!(TextSpecialInfoRuns::parse_record(&special_info_atom(&data)).is_err());
    }

    #[test]
    fn parses_master_text_prop_runs_and_round_trips() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&12u32.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(&4u32.to_le_bytes());
        payload.extend_from_slice(&4u16.to_le_bytes());
        let parsed = MasterTextPropLevels::parse_record(&master_prop_atom(&payload)).unwrap();
        assert_eq!(parsed.runs().len(), 2);
        assert_eq!(parsed.total_count(), 16);
        assert_eq!(parsed.runs()[0].count(), 12);
        assert_eq!(parsed.runs()[0].indent_level(), 0);
        assert_eq!(parsed.runs()[1].count(), 4);
        assert_eq!(parsed.runs()[1].indent_level(), 4);
        assert_eq!(parsed.to_bytes()[8..], payload[..]);
    }

    #[test]
    fn rejects_malformed_master_text_prop() {
        // Wrong record type.
        assert!(MasterTextPropLevels::parse_record(&special_info_atom(&[])).is_err());
        // Payload not a whole number of runs.
        assert!(MasterTextPropLevels::parse_record(&master_prop_atom(&[0; 5])).is_err());
        // Indent level above the maximum.
        let mut payload = Vec::new();
        payload.extend_from_slice(&1u32.to_le_bytes());
        payload.extend_from_slice(&5u16.to_le_bytes());
        assert!(MasterTextPropLevels::parse_record(&master_prop_atom(&payload)).is_err());
        // Nonzero instance.
        let mut record = master_prop_atom(&[]);
        record.instance = 1;
        assert!(MasterTextPropLevels::parse_record(&record).is_err());
    }
}
