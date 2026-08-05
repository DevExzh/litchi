//! Word spelling, grammar, and language auto-detection proofing-state PLCFs.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;

const SPELLING_FIB_INDEX: usize = 55;
const GRAMMAR_FIB_INDEX: usize = 90;
/// Table-pointer index of `fcPlcfLad`/`lcbPlcfLad` (MS-DOC 2.5.7 FibRgFcLcb2000).
const LANGUAGE_DETECTION_FIB_INDEX: usize = 98;
const MAX_PROOFING_ENTRIES: usize = 1_000_000;
const MAX_PROOFING_TABLE_BYTES: usize = 4 + MAX_PROOFING_ENTRIES * 6;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// The checker whose state is described by a proofing PLCF.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ProofingFeature {
    Spelling,
    Grammar,
    /// Language auto-detection (`Plcflad`, MS-DOC 2.8.24).
    LanguageAutoDetect,
}

/// Allowed `SPLS.splf` proofing states (MS-DOC 2.9.256).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum ProofingState {
    MaybeDirty = 0x2,
    Dirty = 0x3,
    Edit = 0x4,
    Foreign = 0x5,
    Clean = 0x7,
    /// `splfNoLAD`: language auto-detection is disabled for the range.
    NoLad = 0x8,
    Error = 0xA,
    RepeatWord = 0xB,
    UnknownWord = 0xC,
}

impl ProofingState {
    fn from_raw(value: u8) -> Result<Self> {
        match value {
            0x2 => Ok(Self::MaybeDirty),
            0x3 => Ok(Self::Dirty),
            0x4 => Ok(Self::Edit),
            0x5 => Ok(Self::Foreign),
            0x7 => Ok(Self::Clean),
            0x8 => Ok(Self::NoLad),
            0xA => Ok(Self::Error),
            0xB => Ok(Self::RepeatWord),
            0xC => Ok(Self::UnknownWord),
            _ => Err(corrupted(format!("invalid SPLS state 0x{value:X}"))),
        }
    }

    fn requires_error(self) -> bool {
        matches!(self, Self::Error | Self::RepeatWord | Self::UnknownWord)
    }

    fn permits_error(self) -> bool {
        matches!(
            self,
            Self::Dirty | Self::Edit | Self::Error | Self::RepeatWord | Self::UnknownWord
        )
    }
}

/// Decoded two-byte `SPLS` state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ProofingStatus {
    state: ProofingState,
    error: bool,
    extend: bool,
    typo: bool,
}

impl ProofingStatus {
    pub fn try_new(
        feature: ProofingFeature,
        state: ProofingState,
        error: bool,
        extend: bool,
        typo: bool,
    ) -> Result<Self> {
        let status = Self {
            state,
            error,
            extend,
            typo,
        };
        status.validate(feature)?;
        Ok(status)
    }

    pub fn state(&self) -> ProofingState {
        self.state
    }
    pub fn is_error(&self) -> bool {
        self.error
    }
    pub fn extend_on_recheck(&self) -> bool {
        self.extend
    }
    pub fn is_typo(&self) -> bool {
        self.typo
    }

    pub fn from_raw(feature: ProofingFeature, raw: u16) -> Result<Self> {
        if raw & 0xFF80 != 0 {
            return Err(corrupted("SPLS unused bits must be zero"));
        }
        let status = Self {
            state: ProofingState::from_raw((raw & 0xF) as u8)?,
            error: raw & 0x10 != 0,
            extend: raw & 0x20 != 0,
            typo: raw & 0x40 != 0,
        };
        status.validate(feature)?;
        Ok(status)
    }

    pub fn to_raw(self, feature: ProofingFeature) -> Result<u16> {
        self.validate(feature)?;
        Ok(self.state as u16
            | u16::from(self.error) << 4
            | u16::from(self.extend) << 5
            | u16::from(self.typo) << 6)
    }

    fn validate(self, feature: ProofingFeature) -> Result<()> {
        if self.state.requires_error() && !self.error {
            return Err(corrupted("SPLS error state requires fError"));
        }
        if self.error && !self.state.permits_error() {
            return Err(corrupted("SPLS fError is invalid for this state"));
        }
        match feature {
            ProofingFeature::Spelling => {
                if self.state == ProofingState::Error {
                    return Err(corrupted("SpellingSpls does not permit splfErrorMin"));
                }
                if self.state == ProofingState::NoLad {
                    return Err(corrupted("SpellingSpls does not permit splfNoLAD"));
                }
                if self.extend || self.typo {
                    return Err(corrupted("SpellingSpls fExtend and fTypo must be zero"));
                }
            },
            ProofingFeature::Grammar => {
                if self.state == ProofingState::NoLad {
                    return Err(corrupted("GrammarSpls does not permit splfNoLAD"));
                }
                if self.extend && !self.error {
                    return Err(corrupted("GrammarSpls fExtend requires fError"));
                }
            },
            ProofingFeature::LanguageAutoDetect => {
                if matches!(
                    self.state,
                    ProofingState::Error | ProofingState::RepeatWord | ProofingState::UnknownWord
                ) {
                    return Err(corrupted("LadSpls does not permit error states"));
                }
                if self.extend || self.typo {
                    return Err(corrupted("LadSpls fExtend and fTypo must be zero"));
                }
            },
        }
        Ok(())
    }
}

/// One proofing state beginning at `start_cp`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ProofingEntry {
    start_cp: u32,
    status: ProofingStatus,
}

impl ProofingEntry {
    pub const fn new(start_cp: u32, status: ProofingStatus) -> Self {
        Self { start_cp, status }
    }

    pub fn start_cp(&self) -> u32 {
        self.start_cp
    }
    pub fn status(&self) -> ProofingStatus {
        self.status
    }
}

/// A resolved proofing range. Zero-length ranges represent insertion/deletion points.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ProofingRange {
    start_cp: u32,
    end_cp: u32,
    status: ProofingStatus,
}

impl ProofingRange {
    pub fn start_cp(&self) -> u32 {
        self.start_cp
    }
    pub fn end_cp(&self) -> u32 {
        self.end_cp
    }
    pub fn is_point(&self) -> bool {
        self.start_cp == self.end_cp
    }
    pub fn status(&self) -> ProofingStatus {
        self.status
    }
}

/// A typed `Plcfspl` or `Plcfgram` table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProofingStateTable {
    feature: ProofingFeature,
    entries: Vec<ProofingEntry>,
    terminal_cp: u32,
}

impl ProofingStateTable {
    pub fn try_new(
        feature: ProofingFeature,
        entries: Vec<ProofingEntry>,
        terminal_cp: u32,
    ) -> Result<Self> {
        validate_entries(feature, &entries, terminal_cp, None)?;
        Ok(Self {
            feature,
            entries,
            terminal_cp,
        })
    }

    pub fn parse_bytes(feature: ProofingFeature, data: &[u8]) -> Result<Self> {
        Self::parse_bytes_with_max_cp(feature, data, None)
    }

    fn parse_bytes_with_max_cp(
        feature: ProofingFeature,
        data: &[u8],
        maximum_cp: Option<u32>,
    ) -> Result<Self> {
        if data.len() < 4 || !(data.len() - 4).is_multiple_of(6) {
            return Err(corrupted("proofing PLCF length must have form 6n + 4"));
        }
        if data.len() > MAX_PROOFING_TABLE_BYTES {
            return Err(corrupted("proofing PLCF exceeds one-million-entry cap"));
        }
        let count = (data.len() - 4) / 6;
        let cp_bytes = count
            .checked_add(1)
            .and_then(|value| value.checked_mul(4))
            .ok_or_else(|| corrupted("proofing PLCF CP array size overflows"))?;
        let mut positions = Vec::with_capacity(count + 1);
        for index in 0..=count {
            positions.push(read_u32(data, index * 4, "proofing CP")?);
        }
        let terminal_cp = positions[count];
        let mut entries = Vec::with_capacity(count);
        for (index, &start_cp) in positions[..count].iter().enumerate() {
            let raw = read_u16(data, cp_bytes + index * 2, "proofing SPLS")?;
            entries.push(ProofingEntry::new(
                start_cp,
                ProofingStatus::from_raw(feature, raw)?,
            ));
        }
        validate_entries(feature, &entries, terminal_cp, maximum_cp)?;
        Ok(Self {
            feature,
            entries,
            terminal_cp,
        })
    }

    pub fn feature(&self) -> ProofingFeature {
        self.feature
    }
    pub fn entries(&self) -> &[ProofingEntry] {
        &self.entries
    }
    pub fn terminal_cp(&self) -> u32 {
        self.terminal_cp
    }
    pub fn len(&self) -> usize {
        self.entries.len()
    }
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub fn range(&self, index: usize) -> Option<ProofingRange> {
        let entry = self.entries.get(index)?;
        let end_cp = self
            .entries
            .get(index + 1)
            .map(ProofingEntry::start_cp)
            .unwrap_or(self.terminal_cp);
        Some(ProofingRange {
            start_cp: entry.start_cp,
            end_cp,
            status: entry.status,
        })
    }

    pub fn ranges(&self) -> impl ExactSizeIterator<Item = ProofingRange> + '_ {
        (0..self.entries.len()).map(|index| self.range(index).unwrap())
    }

    /// Serialize the complete PLC deterministically.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_bytes_with_max_cp(None)
    }

    pub(crate) fn to_bytes_for_document(&self, maximum_cp: u32) -> Result<Vec<u8>> {
        self.to_bytes_with_max_cp(Some(maximum_cp))
    }

    fn to_bytes_with_max_cp(&self, maximum_cp: Option<u32>) -> Result<Vec<u8>> {
        validate_entries(self.feature, &self.entries, self.terminal_cp, maximum_cp)?;
        let size = self
            .entries
            .len()
            .checked_mul(6)
            .and_then(|value| value.checked_add(4))
            .ok_or_else(|| corrupted("proofing PLCF serialized size overflows"))?;
        let mut data = Vec::with_capacity(size);
        for entry in &self.entries {
            data.extend_from_slice(&entry.start_cp.to_le_bytes());
        }
        data.extend_from_slice(&self.terminal_cp.to_le_bytes());
        for entry in &self.entries {
            data.extend_from_slice(&entry.status.to_raw(self.feature)?.to_le_bytes());
        }
        Ok(data)
    }
}

fn validate_entries(
    feature: ProofingFeature,
    entries: &[ProofingEntry],
    terminal_cp: u32,
    maximum_cp: Option<u32>,
) -> Result<()> {
    if entries.len() > MAX_PROOFING_ENTRIES {
        return Err(corrupted("proofing PLCF exceeds one-million-entry cap"));
    }
    let mut previous = None;
    for (index, entry) in entries.iter().enumerate() {
        entry.status.validate(feature)?;
        if entry.start_cp > i32::MAX as u32 {
            return Err(corrupted(format!(
                "proofing CP {index} exceeds signed CP range"
            )));
        }
        if previous.is_some_and(|value| entry.start_cp < value) {
            return Err(corrupted("proofing PLCF CPs are not nondecreasing"));
        }
        previous = Some(entry.start_cp);
    }
    if terminal_cp > i32::MAX as u32 {
        return Err(corrupted("proofing terminal CP exceeds signed CP range"));
    }
    if previous.is_some_and(|value| terminal_cp < value) {
        return Err(corrupted("proofing terminal CP precedes the final entry"));
    }
    if maximum_cp.is_some_and(|maximum| terminal_cp > maximum) {
        return Err(corrupted("proofing terminal CP exceeds the document parts"));
    }
    Ok(())
}

/// Optional spelling, grammar, and language auto-detection proofing tables for a document.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ProofingTables {
    spelling: Option<ProofingStateTable>,
    grammar: Option<ProofingStateTable>,
    language_detection: Option<ProofingStateTable>,
}

impl ProofingTables {
    /// Create a pair of optional proofing tables.
    ///
    /// Each supplied table must describe the matching proofing feature. A
    /// language auto-detection table can be added with [`ProofingTables::set`].
    pub fn try_new(
        spelling: Option<ProofingStateTable>,
        grammar: Option<ProofingStateTable>,
    ) -> Result<Self> {
        if spelling
            .as_ref()
            .is_some_and(|table| table.feature() != ProofingFeature::Spelling)
        {
            return Err(corrupted("spelling slot requires a Plcfspl table"));
        }
        if grammar
            .as_ref()
            .is_some_and(|table| table.feature() != ProofingFeature::Grammar)
        {
            return Err(corrupted("grammar slot requires a Plcfgram table"));
        }
        Ok(Self {
            spelling,
            grammar,
            language_detection: None,
        })
    }

    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let maximum_cp = fib
            .get_document_parts_end()
            .and_then(|value| value.checked_add(2))
            .ok_or_else(|| corrupted("document-parts proofing CP ceiling overflows"))?;
        Ok(Self {
            spelling: parse_fib_table(
                fib,
                table_stream,
                SPELLING_FIB_INDEX,
                ProofingFeature::Spelling,
                maximum_cp,
            )?,
            grammar: parse_fib_table(
                fib,
                table_stream,
                GRAMMAR_FIB_INDEX,
                ProofingFeature::Grammar,
                maximum_cp,
            )?,
            language_detection: parse_fib_table(
                fib,
                table_stream,
                LANGUAGE_DETECTION_FIB_INDEX,
                ProofingFeature::LanguageAutoDetect,
                maximum_cp,
            )?,
        })
    }

    pub fn spelling(&self) -> Option<&ProofingStateTable> {
        self.spelling.as_ref()
    }
    pub fn grammar(&self) -> Option<&ProofingStateTable> {
        self.grammar.as_ref()
    }

    /// Language auto-detection state ranges (`Plcflad`, MS-DOC 2.8.24).
    pub fn language_detection(&self) -> Option<&ProofingStateTable> {
        self.language_detection.as_ref()
    }

    pub fn get(&self, feature: ProofingFeature) -> Option<&ProofingStateTable> {
        match feature {
            ProofingFeature::Spelling => self.spelling(),
            ProofingFeature::Grammar => self.grammar(),
            ProofingFeature::LanguageAutoDetect => self.language_detection(),
        }
    }

    /// Insert or replace the table for its typed proofing feature.
    pub fn set(&mut self, table: ProofingStateTable) -> Option<ProofingStateTable> {
        match table.feature() {
            ProofingFeature::Spelling => self.spelling.replace(table),
            ProofingFeature::Grammar => self.grammar.replace(table),
            ProofingFeature::LanguageAutoDetect => self.language_detection.replace(table),
        }
    }

    /// Remove and return one proofing table.
    pub fn remove(&mut self, feature: ProofingFeature) -> Option<ProofingStateTable> {
        match feature {
            ProofingFeature::Spelling => self.spelling.take(),
            ProofingFeature::Grammar => self.grammar.take(),
            ProofingFeature::LanguageAutoDetect => self.language_detection.take(),
        }
    }
}

fn parse_fib_table(
    fib: &FileInformationBlock,
    table_stream: &[u8],
    index: usize,
    feature: ProofingFeature,
    maximum_cp: u32,
) -> Result<Option<ProofingStateTable>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted("proofing PLCF offset is too large"))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted("proofing PLCF length is too large"))?;
    if length > MAX_PROOFING_TABLE_BYTES {
        return Err(corrupted("proofing PLCF exceeds one-million-entry cap"));
    }
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted("proofing PLCF range overflows"))?;
    let data = table_stream
        .get(start..end)
        .ok_or_else(|| corrupted("proofing PLCF extends beyond the table stream"))?;
    ProofingStateTable::parse_bytes_with_max_cp(feature, data, Some(maximum_cp)).map(Some)
}

#[cfg(test)]
mod tests {
    use super::*;

    const POI_SPELLING: [u8; 22] = [
        0, 0, 0, 0, 33, 0, 0, 0, 39, 0, 0, 0, 162, 0, 0, 0, 7, 0, 4, 0, 7, 0,
    ];
    const POI_GRAMMAR: [u8; 34] = [
        0, 0, 0, 0, 18, 0, 0, 0, 40, 0, 0, 0, 68, 0, 0, 0, 98, 0, 0, 0, 162, 0, 0, 0, 7, 0, 4, 0,
        7, 0, 51, 0, 7, 0,
    ];

    #[test]
    fn parses_and_round_trips_poi_reference_tables() {
        let spelling =
            ProofingStateTable::parse_bytes(ProofingFeature::Spelling, &POI_SPELLING).unwrap();
        assert_eq!(spelling.terminal_cp(), 162);
        assert_eq!(spelling.len(), 3);
        assert_eq!(spelling.range(1).unwrap().start_cp(), 33);
        assert_eq!(spelling.range(1).unwrap().end_cp(), 39);
        assert_eq!(
            spelling.range(1).unwrap().status().state(),
            ProofingState::Edit
        );
        assert_eq!(spelling.to_bytes().unwrap(), POI_SPELLING);

        let grammar =
            ProofingStateTable::parse_bytes(ProofingFeature::Grammar, &POI_GRAMMAR).unwrap();
        let error = grammar.range(3).unwrap();
        assert_eq!((error.start_cp(), error.end_cp()), (68, 98));
        assert_eq!(error.status().state(), ProofingState::Dirty);
        assert!(error.status().is_error());
        assert!(error.status().extend_on_recheck());
        assert_eq!(grammar.to_bytes().unwrap(), POI_GRAMMAR);
    }

    #[test]
    fn preserves_duplicate_cp_point_ranges() {
        let clean = ProofingStatus::try_new(
            ProofingFeature::Spelling,
            ProofingState::Clean,
            false,
            false,
            false,
        )
        .unwrap();
        let table = ProofingStateTable::try_new(
            ProofingFeature::Spelling,
            vec![ProofingEntry::new(10, clean), ProofingEntry::new(10, clean)],
            20,
        )
        .unwrap();
        assert!(table.range(0).unwrap().is_point());
        assert_eq!(
            ProofingStateTable::parse_bytes(ProofingFeature::Spelling, &table.to_bytes().unwrap())
                .unwrap(),
            table
        );
    }

    #[test]
    fn rejects_malformed_plc_shapes_and_positions() {
        assert!(ProofingStateTable::parse_bytes(ProofingFeature::Spelling, &[]).is_err());
        assert!(
            ProofingStateTable::parse_bytes(ProofingFeature::Spelling, &POI_SPELLING[..21],)
                .is_err()
        );
        let mut bytes = POI_SPELLING;
        bytes[4..8].copy_from_slice(&40u32.to_le_bytes());
        assert!(ProofingStateTable::parse_bytes(ProofingFeature::Spelling, &bytes).is_err());
        let mut bytes = POI_SPELLING;
        bytes[12..16].copy_from_slice(&38u32.to_le_bytes());
        assert!(ProofingStateTable::parse_bytes(ProofingFeature::Spelling, &bytes).is_err());
        let mut bytes = POI_SPELLING;
        bytes[0..4].copy_from_slice(&0x8000_0000u32.to_le_bytes());
        assert!(ProofingStateTable::parse_bytes(ProofingFeature::Spelling, &bytes).is_err());
    }

    #[test]
    fn rejects_invalid_spls_states_and_flags() {
        assert!(ProofingStatus::from_raw(ProofingFeature::Spelling, 0).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::Spelling, 0x8007).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::Spelling, 0x0A).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::Spelling, 0x1A).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::Spelling, 0x27).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::Grammar, 0x25).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::Grammar, 0x1A).is_ok());
        assert!(ProofingStatus::from_raw(ProofingFeature::Grammar, 0x33).is_ok());
    }

    /// `Plcflad` sample: three LadSpls ranges ending at CP 40.
    const POI_LANGUAGE_DETECTION: [u8; 22] = [
        0, 0, 0, 0, 12, 0, 0, 0, 30, 0, 0, 0, 40, 0, 0, 0, 7, 0, 8, 0, 4, 0,
    ];

    #[test]
    fn parses_and_round_trips_language_detection_table() {
        let table = ProofingStateTable::parse_bytes(
            ProofingFeature::LanguageAutoDetect,
            &POI_LANGUAGE_DETECTION,
        )
        .unwrap();
        assert_eq!(table.feature(), ProofingFeature::LanguageAutoDetect);
        assert_eq!(table.terminal_cp(), 40);
        assert_eq!(table.len(), 3);
        assert_eq!(
            table.range(0).unwrap().status().state(),
            ProofingState::Clean
        );
        assert_eq!(
            table.range(1).unwrap().status().state(),
            ProofingState::NoLad
        );
        assert_eq!(
            table.range(2).unwrap().status().state(),
            ProofingState::Edit
        );
        assert_eq!(table.to_bytes().unwrap(), POI_LANGUAGE_DETECTION);
    }

    #[test]
    fn enforces_language_detection_state_restrictions() {
        // splfNoLAD is exclusive to language auto-detection.
        assert!(ProofingStatus::from_raw(ProofingFeature::Spelling, 0x8).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::Grammar, 0x8).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::LanguageAutoDetect, 0x8).is_ok());
        // LadSpls forbids the error states and the fExtend/fTypo flags.
        assert!(ProofingStatus::from_raw(ProofingFeature::LanguageAutoDetect, 0x1A).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::LanguageAutoDetect, 0x1B).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::LanguageAutoDetect, 0x27).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::LanguageAutoDetect, 0x47).is_err());
        assert!(ProofingStatus::from_raw(ProofingFeature::LanguageAutoDetect, 0x13).is_ok());
    }

    fn fib_with_lad_pointer(offset: u32, length: u32) -> FileInformationBlock {
        let mut data = vec![0u8; 154 + 99 * 8];
        data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        data[2..4].copy_from_slice(&0x00D9u16.to_le_bytes());
        data[152..154].copy_from_slice(&99u16.to_le_bytes());
        // FibRgLw97.ccpText at offset 0x4C bounds proofing CPs.
        data[0x4C..0x50].copy_from_slice(&100u32.to_le_bytes());
        let pointer = 154 + LANGUAGE_DETECTION_FIB_INDEX * 8;
        data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
        data[pointer + 4..pointer + 8].copy_from_slice(&length.to_le_bytes());
        FileInformationBlock::parse(&data).unwrap()
    }

    #[test]
    fn parses_language_detection_table_through_fib() {
        let fib = fib_with_lad_pointer(4, POI_LANGUAGE_DETECTION.len() as u32);
        let mut table_stream = vec![0u8; 4];
        table_stream.extend_from_slice(&POI_LANGUAGE_DETECTION);
        let tables = ProofingTables::parse(&fib, &table_stream).unwrap();
        assert!(tables.spelling().is_none());
        assert!(tables.grammar().is_none());
        let lad = tables.language_detection().unwrap();
        assert_eq!(lad.len(), 3);
        assert_eq!(
            tables.get(ProofingFeature::LanguageAutoDetect).unwrap(),
            lad
        );
    }

    #[test]
    fn rejects_language_detection_cp_beyond_document_parts() {
        let fib = fib_with_lad_pointer(0, POI_LANGUAGE_DETECTION.len() as u32);
        let mut bytes = POI_LANGUAGE_DETECTION.to_vec();
        bytes[12..16].copy_from_slice(&500u32.to_le_bytes());
        assert!(ProofingTables::parse(&fib, &bytes).is_err());
    }

    #[test]
    fn language_detection_slot_round_trips_through_set_and_remove() {
        let lad = ProofingStateTable::parse_bytes(
            ProofingFeature::LanguageAutoDetect,
            &POI_LANGUAGE_DETECTION,
        )
        .unwrap();
        let mut tables = ProofingTables::default();
        assert!(tables.set(lad.clone()).is_none());
        assert_eq!(tables.language_detection(), Some(&lad));
        assert_eq!(
            tables.remove(ProofingFeature::LanguageAutoDetect),
            Some(lad)
        );
        assert!(tables.language_detection().is_none());
    }
}
