//! Package-level aggregation for the `SXVIEWEX` record sequence.

use crate::error::{Error, Result};

use super::model::{PivotFieldOlapExt, PivotHierarchy, PivotPageItemOlapExt, PivotViewOlapHeader};
use super::{SX_VIEW_EX_RECORD_TYPE, SXPI_EX_RECORD_TYPE, SXTH_RECORD_TYPE, SXVDT_EX_RECORD_TYPE};

/// A complete inert `SXViewEx`/`SXTH`/`SXPIEx`/`SXVDTEx` sequence.
///
/// The sequence owns its typed records and does not evaluate MDX or contact
/// an OLAP provider. `parse` and `to_records` use the record order and counts
/// mandated by MS-XLS 2.1. A record payload is represented without the outer
/// BIFF `(type, length)` header, matching the individual record APIs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OlapSequence {
    /// The leading `SXViewEx` count and future-byte record.
    pub header: PivotViewOlapHeader,
    /// Hierarchy records following the header (`SXTH`).
    pub hierarchies: Vec<PivotHierarchy>,
    /// Page-axis OLAP extensions (`SXPIEx`).
    pub page_extensions: Vec<PivotPageItemOlapExt>,
    /// Pivot-field OLAP extensions (`SXVDTEx`).
    pub field_extensions: Vec<PivotFieldOlapExt>,
}

impl OlapSequence {
    /// Build a complete sequence and derive the count fields in its header.
    pub fn from_parts(
        hierarchies: Vec<PivotHierarchy>,
        page_extensions: Vec<PivotPageItemOlapExt>,
        field_extensions: Vec<PivotFieldOlapExt>,
        future_bytes: Vec<u8>,
    ) -> Result<Self> {
        let header = PivotViewOlapHeader {
            hierarchy_count: u32::try_from(hierarchies.len())
                .map_err(|_| Error::InvalidData("SXViewEx hierarchy count exceeds u32".into()))?,
            page_extension_count: u32::try_from(page_extensions.len()).map_err(|_| {
                Error::InvalidData("SXViewEx page extension count exceeds u32".into())
            })?,
            field_extension_count: u32::try_from(field_extensions.len()).map_err(|_| {
                Error::InvalidData("SXViewEx field extension count exceeds u32".into())
            })?,
            future_bytes,
        };
        let sequence = Self {
            header,
            hierarchies,
            page_extensions,
            field_extensions,
        };
        // This also enforces the MS-XLS requirement that at least one SXTH is
        // present and bounds the future blob before any bytes are emitted.
        sequence.header.to_payload()?;
        Ok(sequence)
    }

    /// Parse a complete sequence of `(record type, payload)` pairs.
    pub fn parse(records: &[(u16, &[u8])]) -> Result<Self> {
        let (header_type, header_payload) = records.first().ok_or_else(|| {
            Error::InvalidData("SXVIEWEX sequence is missing its SXViewEx header".to_string())
        })?;
        if *header_type != SX_VIEW_EX_RECORD_TYPE {
            return Err(Error::UnexpectedRecordType {
                expected: SX_VIEW_EX_RECORD_TYPE,
                found: *header_type,
            });
        }
        let header = PivotViewOlapHeader::parse(header_payload)?;
        let hierarchy_count = usize::try_from(header.hierarchy_count)
            .map_err(|_| Error::InvalidData("SXViewEx hierarchy count overflow".to_string()))?;
        let page_count = usize::try_from(header.page_extension_count).map_err(|_| {
            Error::InvalidData("SXViewEx page extension count overflow".to_string())
        })?;
        let field_count = usize::try_from(header.field_extension_count).map_err(|_| {
            Error::InvalidData("SXViewEx field extension count overflow".to_string())
        })?;
        let expected_len = 1usize
            .checked_add(hierarchy_count)
            .and_then(|value| value.checked_add(page_count))
            .and_then(|value| value.checked_add(field_count))
            .ok_or_else(|| Error::InvalidData("SXVIEWEX sequence length overflow".to_string()))?;
        if records.len() != expected_len {
            return Err(Error::InvalidData(format!(
                "SXVIEWEX count fields require {expected_len} records, found {}",
                records.len()
            )));
        }

        let mut offset = 1usize;
        let mut hierarchies = Vec::with_capacity(hierarchy_count);
        for _ in 0..hierarchy_count {
            let (record_type, payload) = records[offset];
            if record_type != SXTH_RECORD_TYPE {
                return Err(Error::UnexpectedRecordType {
                    expected: SXTH_RECORD_TYPE,
                    found: record_type,
                });
            }
            hierarchies.push(PivotHierarchy::parse(payload)?);
            offset += 1;
        }

        let mut page_extensions = Vec::with_capacity(page_count);
        for _ in 0..page_count {
            let (record_type, payload) = records[offset];
            if record_type != SXPI_EX_RECORD_TYPE {
                return Err(Error::UnexpectedRecordType {
                    expected: SXPI_EX_RECORD_TYPE,
                    found: record_type,
                });
            }
            page_extensions.push(PivotPageItemOlapExt::parse(payload)?);
            offset += 1;
        }

        let mut field_extensions = Vec::with_capacity(field_count);
        for _ in 0..field_count {
            let (record_type, payload) = records[offset];
            if record_type != SXVDT_EX_RECORD_TYPE {
                return Err(Error::UnexpectedRecordType {
                    expected: SXVDT_EX_RECORD_TYPE,
                    found: record_type,
                });
            }
            field_extensions.push(PivotFieldOlapExt::parse(payload)?);
            offset += 1;
        }

        Ok(Self {
            header,
            hierarchies,
            page_extensions,
            field_extensions,
        })
    }

    /// Serialize the complete sequence as `(record type, payload)` pairs.
    pub fn to_records(&self) -> Result<Vec<(u16, Vec<u8>)>> {
        let hierarchy_count = usize::try_from(self.header.hierarchy_count)
            .map_err(|_| Error::InvalidData("SXViewEx hierarchy count overflow".to_string()))?;
        let page_count = usize::try_from(self.header.page_extension_count).map_err(|_| {
            Error::InvalidData("SXViewEx page extension count overflow".to_string())
        })?;
        let field_count = usize::try_from(self.header.field_extension_count).map_err(|_| {
            Error::InvalidData("SXViewEx field extension count overflow".to_string())
        })?;
        if hierarchy_count != self.hierarchies.len()
            || page_count != self.page_extensions.len()
            || field_count != self.field_extensions.len()
        {
            return Err(Error::InvalidData(
                "SXViewEx count fields disagree with the typed sequence".to_string(),
            ));
        }

        let mut records = Vec::with_capacity(1 + hierarchy_count + page_count + field_count);
        records.push((SX_VIEW_EX_RECORD_TYPE, self.header.to_payload()?));
        for hierarchy in &self.hierarchies {
            records.push((SXTH_RECORD_TYPE, hierarchy.to_payload()?));
        }
        for page_extension in &self.page_extensions {
            records.push((SXPI_EX_RECORD_TYPE, page_extension.to_payload()?));
        }
        for field_extension in &self.field_extensions {
            records.push((SXVDT_EX_RECORD_TYPE, field_extension.to_payload()?));
        }
        Ok(records)
    }
}
