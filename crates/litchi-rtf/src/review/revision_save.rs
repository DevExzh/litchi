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
    pub fn new(ids: Vec<u32>, root: Option<u32>) -> RtfResult<Self> {
        let metadata = Self { ids, root };
        metadata.validate()?;
        Ok(metadata)
    }

    pub fn ids(&self) -> &[u32] {
        &self.ids
    }

    pub fn root(&self) -> Option<u32> {
        self.root
    }

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

    pub fn set_root(&mut self, root: Option<u32>) -> RtfResult<()> {
        if let Some(root) = root {
            validate_id(root)?;
            if !self.ids.is_empty() && !self.ids.contains(&root) {
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
