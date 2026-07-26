//! Legacy Word spelling and grammar proofing PLCF authoring.

use super::core::DocWriteError;
use super::fib::FibBuilder;
use crate::doc::{ProofingStateTable, ProofingTables};

#[derive(Debug, Default)]
struct SerializedProofingTables {
    spelling: Option<Vec<u8>>,
    grammar: Option<Vec<u8>>,
}

impl SerializedProofingTables {
    fn build(tables: &ProofingTables, maximum_cp: u32) -> Result<Self, DocWriteError> {
        fn serialize(
            table: Option<&ProofingStateTable>,
            maximum_cp: u32,
        ) -> Result<Option<Vec<u8>>, DocWriteError> {
            table
                .map(|table| {
                    table
                        .to_bytes_for_document(maximum_cp)
                        .map_err(|error| DocWriteError::InvalidData(error.to_string()))
                })
                .transpose()
        }

        Ok(Self {
            spelling: serialize(tables.spelling(), maximum_cp)?,
            grammar: serialize(tables.grammar(), maximum_cp)?,
        })
    }
}

fn checked_length(data: &[u8]) -> Result<u32, DocWriteError> {
    u32::try_from(data.len())
        .map_err(|_| DocWriteError::InvalidData("proofing PLCF exceeds 4 GiB".into()))
}

fn checked_end(offset: u32, length: u32) -> Result<u32, DocWriteError> {
    offset
        .checked_add(length)
        .ok_or_else(|| DocWriteError::InvalidData("DOC table offset overflows".into()))
}

fn validate_start_offset(table_stream: &[u8], offset: u32) -> Result<(), DocWriteError> {
    let actual_offset = u32::try_from(table_stream.len())
        .map_err(|_| DocWriteError::InvalidData("DOC table stream exceeds 4 GiB".into()))?;
    if actual_offset != offset {
        return Err(DocWriteError::InvalidData(
            "DOC table-stream offset plan is inconsistent".into(),
        ));
    }
    Ok(())
}

/// Validate, serialize, and append both optional tables.
///
/// Serialization finishes before either table is appended, so validation failure
/// cannot leave a partially authored proofing pair in the table stream.
pub(super) fn append_proofing_tables(
    fib: &mut FibBuilder,
    table_stream: &mut Vec<u8>,
    tables: &ProofingTables,
    table_offset: u32,
    maximum_cp: u32,
) -> Result<u32, DocWriteError> {
    let serialized = SerializedProofingTables::build(tables, maximum_cp)?;
    validate_start_offset(table_stream, table_offset)?;

    let spelling = serialized
        .spelling
        .map(|data| checked_length(&data).map(|length| (data, length)))
        .transpose()?;
    let grammar = serialized
        .grammar
        .map(|data| checked_length(&data).map(|length| (data, length)))
        .transpose()?;
    let spelling_length = spelling.as_ref().map_or(0, |(_, length)| *length);
    let grammar_length = grammar.as_ref().map_or(0, |(_, length)| *length);
    let grammar_offset = checked_end(table_offset, spelling_length)?;
    let final_offset = checked_end(grammar_offset, grammar_length)?;

    if let Some((data, length)) = spelling {
        fib.set_plcfspl(table_offset, length);
        table_stream.extend_from_slice(&data);
    }
    if let Some((data, length)) = grammar {
        fib.set_plcfgram(grammar_offset, length);
        table_stream.extend_from_slice(&data);
    }
    Ok(final_offset)
}
