//! BIFF8 worksheet scenario manager and scenario records.

use super::{XlsError, XlsResult};

pub(crate) const SCEN_MAN_RECORD_TYPE: u16 = 0x00AE;
pub(crate) const SCENARIO_RECORD_TYPE: u16 = 0x00AF;
const CONTINUE_RECORD_TYPE: u16 = 0x003C;
const DIMENSIONS_RECORD_TYPE: u16 = 0x0200;
const MAX_SCENARIO_CELLS: usize = 32;
const MAX_RESULT_RANGES: usize = 32;
const MAX_SCENARIO_BYTES: usize = 4_200_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsScenarioRange {
    pub(crate) first_row: u16,
    pub(crate) last_row: u16,
    pub(crate) first_column: u8,
    pub(crate) last_column: u8,
}

impl XlsScenarioRange {
    pub fn new(
        first_row: u16,
        last_row: u16,
        first_column: u8,
        last_column: u8,
    ) -> XlsResult<Self> {
        if first_row > last_row || first_column > last_column {
            return Err(XlsError::InvalidData(
                "scenario range is reversed".to_string(),
            ));
        }
        Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
        })
    }
    pub fn first_row(&self) -> u16 {
        self.first_row
    }
    pub fn last_row(&self) -> u16 {
        self.last_row
    }
    pub fn first_column(&self) -> u8 {
        self.first_column
    }
    pub fn last_column(&self) -> u8 {
        self.last_column
    }
}

/// One inert changed-cell value from a scenario.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsScenarioCell {
    pub(crate) row: u16,
    pub(crate) column: u8,
    pub(crate) deleted: bool,
    pub(crate) value: String,
}

impl XlsScenarioCell {
    pub fn new(row: u16, column: u8, value: impl Into<String>) -> Self {
        Self {
            row,
            column,
            deleted: false,
            value: value.into(),
        }
    }
    pub fn deleted(row: u16, column: u8, value: impl Into<String>) -> Self {
        Self {
            row,
            column,
            deleted: true,
            value: value.into(),
        }
    }
    pub fn row(&self) -> u16 {
        self.row
    }
    pub fn column(&self) -> u8 {
        self.column
    }
    pub fn is_deleted(&self) -> bool {
        self.deleted
    }
    /// Raw scenario value. It is never parsed or evaluated as a formula or macro.
    pub fn value(&self) -> &str {
        &self.value
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsScenario {
    pub(crate) name: String,
    pub(crate) creator: Option<String>,
    pub(crate) comment: Option<String>,
    pub(crate) locked: bool,
    pub(crate) hidden: bool,
    pub(crate) cells: Vec<XlsScenarioCell>,
}

impl XlsScenario {
    pub fn new(name: impl Into<String>, cells: Vec<XlsScenarioCell>) -> Self {
        Self {
            name: name.into(),
            creator: None,
            comment: None,
            locked: false,
            hidden: false,
            cells,
        }
    }
    pub fn set_creator(&mut self, creator: Option<String>) {
        self.creator = creator;
    }
    pub fn set_comment(&mut self, comment: Option<String>) {
        self.comment = comment;
    }
    pub fn set_locked(&mut self, locked: bool) {
        self.locked = locked;
    }
    pub fn set_hidden(&mut self, hidden: bool) {
        self.hidden = hidden;
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn creator(&self) -> Option<&str> {
        self.creator.as_deref()
    }
    pub fn comment(&self) -> Option<&str> {
        self.comment.as_deref()
    }
    pub fn is_locked(&self) -> bool {
        self.locked
    }
    pub fn is_hidden(&self) -> bool {
        self.hidden
    }
    pub fn cells(&self) -> &[XlsScenarioCell] {
        &self.cells
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsScenarioManager {
    pub(crate) current_scenario: Option<usize>,
    pub(crate) shown_scenario: Option<usize>,
    pub(crate) result_ranges: Vec<XlsScenarioRange>,
    pub(crate) scenarios: Vec<XlsScenario>,
}

impl XlsScenarioManager {
    pub fn new(scenarios: Vec<XlsScenario>) -> Self {
        Self {
            current_scenario: None,
            shown_scenario: None,
            result_ranges: Vec::new(),
            scenarios,
        }
    }
    pub fn set_current_scenario(&mut self, index: Option<usize>) {
        self.current_scenario = index;
    }
    pub fn set_shown_scenario(&mut self, index: Option<usize>) {
        self.shown_scenario = index;
    }
    pub fn set_result_ranges(&mut self, ranges: Vec<XlsScenarioRange>) {
        self.result_ranges = ranges;
    }
    pub fn current_scenario(&self) -> Option<usize> {
        self.current_scenario
    }
    pub fn shown_scenario(&self) -> Option<usize> {
        self.shown_scenario
    }
    pub fn result_ranges(&self) -> &[XlsScenarioRange] {
        &self.result_ranges
    }
    pub fn scenarios(&self) -> &[XlsScenario] {
        &self.scenarios
    }

    pub(crate) fn validate_for_write(&self) -> XlsResult<()> {
        if self.scenarios.len() > i16::MAX as usize {
            return invalid_data("scenario count exceeds 32767");
        }
        if self.result_ranges.len() > MAX_RESULT_RANGES {
            return invalid_data("scenario result range count exceeds 32");
        }
        for index in [self.current_scenario, self.shown_scenario]
            .into_iter()
            .flatten()
        {
            if index >= self.scenarios.len() {
                return invalid_data("selected scenario index is out of range");
            }
        }
        for scenario in &self.scenarios {
            let name_units = scenario.name.encode_utf16().count();
            if name_units > u8::MAX as usize {
                return invalid_data("scenario name exceeds 255 UTF-16 code units");
            }
            if !(1..=MAX_SCENARIO_CELLS).contains(&scenario.cells.len()) {
                return invalid_data("scenario changed-cell count must be 1..=32");
            }
            if scenario
                .creator
                .as_ref()
                .is_some_and(|value| value.encode_utf16().count() > 52)
            {
                return invalid_data("scenario creator exceeds 52 UTF-16 code units");
            }
            if scenario
                .comment
                .as_ref()
                .is_some_and(|value| value.encode_utf16().count() > 255)
            {
                return invalid_data("scenario comment exceeds 255 UTF-16 code units");
            }
            for cell in &scenario.cells {
                if cell.value.encode_utf16().count() > u16::MAX as usize {
                    return invalid_data("scenario value exceeds 65535 UTF-16 code units");
                }
            }
        }
        Ok(())
    }
}

struct ScenarioManagerHeader {
    declared_count: usize,
    current_scenario: Option<usize>,
    shown_scenario: Option<usize>,
    result_ranges: Vec<XlsScenarioRange>,
}

pub(crate) struct ScenarioCollector {
    header: Option<ScenarioManagerHeader>,
    scenarios: Vec<XlsScenario>,
    pending_scenario: Option<Vec<u8>>,
    scenario_slot_closed: bool,
}

impl ScenarioCollector {
    pub(crate) fn new() -> Self {
        Self {
            header: None,
            scenarios: Vec::new(),
            pending_scenario: None,
            scenario_slot_closed: false,
        }
    }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        if record_type == CONTINUE_RECORD_TYPE && self.pending_scenario.is_some() {
            let pending = self.pending_scenario.as_mut().unwrap();
            let new_len = pending.len().checked_add(data.len()).ok_or_else(|| {
                XlsError::InvalidData("scenario continuation size overflow".to_string())
            })?;
            if new_len > MAX_SCENARIO_BYTES {
                return invalid(
                    SCENARIO_RECORD_TYPE,
                    "scenario continuation exceeds resource bound",
                );
            }
            pending.extend_from_slice(data);
            return Ok(());
        }

        self.finish_pending()?;
        match record_type {
            SCEN_MAN_RECORD_TYPE => {
                if self.header.is_some() || self.scenario_slot_closed {
                    return invalid(record_type, "duplicate or out-of-order ScenMan record");
                }
                self.header = Some(parse_scen_man(data)?);
            },
            SCENARIO_RECORD_TYPE => {
                let header = self
                    .header
                    .as_ref()
                    .ok_or_else(|| XlsError::InvalidRecord {
                        record_type,
                        message: "Scenario record appears without ScenMan".to_string(),
                    })?;
                if self.scenarios.len() >= header.declared_count {
                    return invalid(
                        record_type,
                        "more Scenario records than declared by ScenMan",
                    );
                }
                if data.len() > MAX_SCENARIO_BYTES {
                    return invalid(record_type, "scenario exceeds resource bound");
                }
                self.pending_scenario = Some(data.to_vec());
            },
            DIMENSIONS_RECORD_TYPE => {
                self.ensure_declared_count()?;
                self.scenario_slot_closed = true;
            },
            _ => {
                if self
                    .header
                    .as_ref()
                    .is_some_and(|header| self.scenarios.len() < header.declared_count)
                {
                    return invalid(
                        record_type,
                        "ScenMan must be followed immediately by its Scenario records",
                    );
                }
            },
        }
        Ok(())
    }

    fn finish_pending(&mut self) -> XlsResult<()> {
        if let Some(data) = self.pending_scenario.take() {
            self.scenarios.push(parse_scenario(&data)?);
        }
        Ok(())
    }

    fn ensure_declared_count(&self) -> XlsResult<()> {
        if let Some(header) = &self.header
            && self.scenarios.len() != header.declared_count
        {
            return invalid(
                SCEN_MAN_RECORD_TYPE,
                "ScenMan scenario count does not match Scenario records",
            );
        }
        Ok(())
    }

    pub(crate) fn finish(mut self) -> XlsResult<Option<XlsScenarioManager>> {
        self.finish_pending()?;
        self.ensure_declared_count()?;
        let Some(header) = self.header else {
            return Ok(None);
        };
        Ok(Some(XlsScenarioManager {
            current_scenario: header.current_scenario,
            shown_scenario: header.shown_scenario,
            result_ranges: header.result_ranges,
            scenarios: self.scenarios,
        }))
    }
}

fn parse_scen_man(data: &[u8]) -> XlsResult<ScenarioManagerHeader> {
    if data.len() < 8 || (data.len() - 8) % 8 != 0 {
        return invalid(SCEN_MAN_RECORD_TYPE, "ScenMan payload length is invalid");
    }
    let declared = read_i16(data, 0);
    let current = read_i16(data, 2);
    let shown = read_i16(data, 4);
    let result_count = read_i16(data, 6);
    if declared < 0 {
        return invalid(SCEN_MAN_RECORD_TYPE, "negative scenario count");
    }
    if !(0..=32).contains(&result_count) {
        return invalid(
            SCEN_MAN_RECORD_TYPE,
            "scenario result range count must be 0..=32",
        );
    }
    let declared_count = declared as usize;
    let expected = 8 + result_count as usize * 8;
    if data.len() != expected {
        return invalid(
            SCEN_MAN_RECORD_TYPE,
            "ScenMan result range count does not match length",
        );
    }
    let parse_index = |value: i16| -> XlsResult<Option<usize>> {
        if value == -1 {
            return Ok(None);
        }
        if value < 0 || value as usize >= declared_count {
            return invalid(
                SCEN_MAN_RECORD_TYPE,
                "selected scenario index is out of range",
            );
        }
        Ok(Some(value as usize))
    };
    let current_scenario = parse_index(current)?;
    let shown_scenario = parse_index(shown)?;
    let mut result_ranges = Vec::with_capacity(result_count as usize);
    for chunk in data[8..].chunks_exact(8) {
        let first_row = read_u16(chunk, 0);
        let last_row = read_u16(chunk, 2);
        let first_column = read_u16(chunk, 4);
        let last_column = read_u16(chunk, 6);
        if first_row > last_row || first_column > last_column || last_column > 255 {
            return invalid(SCEN_MAN_RECORD_TYPE, "invalid ScenMan result range");
        }
        result_ranges.push(XlsScenarioRange {
            first_row,
            last_row,
            first_column: first_column as u8,
            last_column: last_column as u8,
        });
    }
    Ok(ScenarioManagerHeader {
        declared_count,
        current_scenario,
        shown_scenario,
        result_ranges,
    })
}

fn parse_scenario(data: &[u8]) -> XlsResult<XlsScenario> {
    let mut cursor = Cursor::new(data);
    let count = cursor.u16()? as usize;
    if !(1..=MAX_SCENARIO_CELLS).contains(&count) {
        return invalid(
            SCENARIO_RECORD_TYPE,
            "Scenario changed-cell count must be 1..=32",
        );
    }
    let locked = cursor.boolean8()?;
    let hidden = cursor.boolean8()?;
    let name_count = cursor.u8()? as usize;
    let comment_count = cursor.u8()? as usize;
    let creator_count = cursor.u8()? as usize;
    if creator_count > 52 {
        return invalid(
            SCENARIO_RECORD_TYPE,
            "Scenario creator exceeds 52 characters",
        );
    }
    let name = cursor.unicode_no_cch(name_count)?;
    let creator = if creator_count == 0 {
        None
    } else {
        Some(cursor.unicode_with_expected_count(creator_count)?)
    };
    let comment = if comment_count == 0 {
        None
    } else {
        Some(cursor.unicode_with_expected_count(comment_count)?)
    };
    let mut refs = Vec::with_capacity(count);
    for _ in 0..count {
        let row = cursor.u16()?;
        let column_flags = cursor.u16()?;
        if column_flags & 0x8000 != 0 || column_flags & 0x3FFF > 255 {
            return invalid(
                SCENARIO_RECORD_TYPE,
                "Scenario cell column flags are invalid",
            );
        }
        refs.push((
            row,
            (column_flags & 0x00FF) as u8,
            column_flags & 0x4000 != 0,
        ));
    }
    let mut values = Vec::with_capacity(count);
    for _ in 0..count {
        values.push(cursor.unicode()?);
    }
    let unused = cursor.take(count * 2)?;
    let _ = unused;
    if cursor.remaining() != 0 {
        return invalid(SCENARIO_RECORD_TYPE, "Scenario payload has trailing bytes");
    }
    let cells = refs
        .into_iter()
        .zip(values)
        .map(|((row, column, deleted), value)| XlsScenarioCell {
            row,
            column,
            deleted,
            value,
        })
        .collect();
    Ok(XlsScenario {
        name,
        creator,
        comment,
        locked,
        hidden,
        cells,
    })
}

struct Cursor<'a> {
    data: &'a [u8],
    position: usize,
}
impl<'a> Cursor<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, position: 0 }
    }
    fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.position)
    }
    fn take(&mut self, count: usize) -> XlsResult<&'a [u8]> {
        let end = self
            .position
            .checked_add(count)
            .ok_or_else(|| XlsError::InvalidData("Scenario size overflow".to_string()))?;
        let value = self
            .data
            .get(self.position..end)
            .ok_or_else(|| XlsError::InvalidRecord {
                record_type: SCENARIO_RECORD_TYPE,
                message: "truncated Scenario record".to_string(),
            })?;
        self.position = end;
        Ok(value)
    }
    fn u8(&mut self) -> XlsResult<u8> {
        Ok(self.take(1)?[0])
    }
    fn u16(&mut self) -> XlsResult<u16> {
        Ok(u16::from_le_bytes(self.take(2)?.try_into().unwrap()))
    }
    fn boolean8(&mut self) -> XlsResult<bool> {
        match self.u8()? {
            0 => Ok(false),
            1 => Ok(true),
            _ => invalid(SCENARIO_RECORD_TYPE, "Scenario Boolean must be 0 or 1"),
        }
    }
    fn unicode(&mut self) -> XlsResult<String> {
        let count = self.u16()? as usize;
        self.unicode_no_cch(count)
    }
    fn unicode_with_expected_count(&mut self, expected: usize) -> XlsResult<String> {
        let count = self.u16()? as usize;
        if count != expected {
            return invalid(
                SCENARIO_RECORD_TYPE,
                "Scenario string count does not match header",
            );
        }
        self.unicode_no_cch(count)
    }
    fn unicode_no_cch(&mut self, count: usize) -> XlsResult<String> {
        let options = self.u8()?;
        if options & 0xFE != 0 {
            return invalid(
                SCENARIO_RECORD_TYPE,
                "Scenario string contains reserved option bits",
            );
        }
        if options & 1 == 0 {
            Ok(self
                .take(count)?
                .iter()
                .map(|byte| char::from(*byte))
                .collect())
        } else {
            let bytes = self.take(count.checked_mul(2).ok_or_else(|| {
                XlsError::InvalidData("Scenario string size overflow".to_string())
            })?)?;
            let units: Vec<u16> = bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect();
            String::from_utf16(&units).map_err(|_| XlsError::InvalidRecord {
                record_type: SCENARIO_RECORD_TYPE,
                message: "Scenario string contains invalid UTF-16".to_string(),
            })
        }
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(data[offset..offset + 2].try_into().unwrap())
}
fn read_i16(data: &[u8], offset: usize) -> i16 {
    i16::from_le_bytes(data[offset..offset + 2].try_into().unwrap())
}
fn invalid<T>(record_type: u16, message: impl Into<String>) -> XlsResult<T> {
    Err(XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    })
}
fn invalid_data<T>(message: impl Into<String>) -> XlsResult<T> {
    Err(XlsError::InvalidData(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_empty_manager_fixture_shape() {
        let manager = parse_scen_man(&[0, 0, 0xFF, 0xFF, 0xFF, 0xFF, 0, 0]).unwrap();
        assert_eq!(manager.declared_count, 0);
        assert_eq!(manager.current_scenario, None);
    }

    #[test]
    fn rejects_bad_lengths_counts_flags_and_order() {
        assert!(parse_scen_man(&[0; 7]).is_err());
        assert!(parse_scen_man(&[1, 0, 1, 0, 0xFF, 0xFF, 0, 0]).is_err());
        let mut collector = ScenarioCollector::new();
        assert!(
            collector
                .feed_record(SCENARIO_RECORD_TYPE, &[0; 8])
                .is_err()
        );

        let mut payload = vec![1, 0, 2, 0, 0, 0, 0];
        payload.extend_from_slice(&[0, 0, 0, 0]);
        payload.extend_from_slice(&[0, 0, 0]);
        payload.extend_from_slice(&[0, 0]);
        assert!(parse_scenario(&payload).is_err());
    }

    #[test]
    fn rejects_partial_declared_collection() {
        let mut collector = ScenarioCollector::new();
        collector
            .feed_record(SCEN_MAN_RECORD_TYPE, &[1, 0, 0xFF, 0xFF, 0xFF, 0xFF, 0, 0])
            .unwrap();
        assert!(
            collector
                .feed_record(DIMENSIONS_RECORD_TYPE, &[0; 14])
                .is_err()
        );
    }
}
