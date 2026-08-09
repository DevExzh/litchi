//! Revision-save/session provenance metadata.

use crate::{RtfError, RtfResult};
use std::collections::HashSet;

pub(crate) const MAX_REVISION_SAVE_IDS: usize = 65_536;

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RevisionSaveMetadata {
    ids: Vec<u32>,
    root: Option<u32>,
}

impl RevisionSaveMetadata {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn new(ids: Vec<u32>, root: Option<u32>) -> RtfResult<Self> {
        let metadata = Self { ids, root };
        metadata.validate()?;
        Ok(metadata)
    }

    #[must_use]
    pub fn ids(&self) -> &[u32] {
        &self.ids
    }

    #[must_use]
    pub fn root(&self) -> Option<u32> {
        self.root
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn push_id(&mut self, id: u32) -> RtfResult<()> {
        validate_id(id)?;
        if self.ids.len() >= MAX_REVISION_SAVE_IDS {
            return Err(RtfError::MalformedDocument(
                "RTF revision-save ID count exceeds the safety limit".to_string(),
            ));
        }
        if self.ids.contains(&id) {
            return Err(RtfError::MalformedDocument(
                "RTF revision-save IDs must be unique".to_string(),
            ));
        }
        self.ids.push(id);
        Ok(())
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn set_root(&mut self, root: Option<u32>) -> RtfResult<()> {
        if let Some(root_id) = root {
            validate_id(root_id)?;
            if !self.ids.is_empty() && !self.ids.contains(&root_id) {
                return Err(RtfError::MalformedDocument(
                    "RTF revision root must occur in the revision-save table".to_string(),
                ));
            }
        }
        self.root = root;
        Ok(())
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if self.ids.len() > MAX_REVISION_SAVE_IDS {
            return Err(RtfError::MalformedDocument(
                "RTF revision-save ID count exceeds the safety limit".to_string(),
            ));
        }
        let mut seen = HashSet::with_capacity(self.ids.len());
        for &id in &self.ids {
            validate_id(id)?;
            if !seen.insert(id) {
                return Err(RtfError::MalformedDocument(
                    "RTF revision-save IDs must be unique".to_string(),
                ));
            }
        }
        if let Some(root) = self.root {
            validate_id(root)?;
            if !self.ids.is_empty() && !seen.contains(&root) {
                return Err(RtfError::MalformedDocument(
                    "RTF revision root must occur in the revision-save table".to_string(),
                ));
            }
        }
        Ok(())
    }
}

fn validate_id(id: u32) -> RtfResult<()> {
    if id == 0 || id > i32::MAX as u32 {
        return Err(RtfError::MalformedDocument(
            "RTF revision-save IDs must be in 1..=2147483647".to_string(),
        ));
    }
    Ok(())
}
