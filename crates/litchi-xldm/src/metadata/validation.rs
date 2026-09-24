//! Section 2.5 validation facade.
//!
//! Cross-file validation is kept behind this module so callers can discover
//! the semantic verification operation without reaching into the XML codec.

use super::super::generated::SystemGeneratedFile;
use super::super::native::NativeFile;
use super::codec::{validate_columns, validate_hierarchies, validate_relationships};
use super::model::{MetadataModel, MetadataResult};

pub fn validate_files(
    metadata: &MetadataModel<'_>,
    native: &[NativeFile<'_>],
    generated: &[SystemGeneratedFile<'_>],
) -> MetadataResult<()> {
    validate_columns(metadata, native)?;
    validate_hierarchies(metadata, native, generated)?;
    validate_relationships(metadata, generated)?;
    Ok(())
}
