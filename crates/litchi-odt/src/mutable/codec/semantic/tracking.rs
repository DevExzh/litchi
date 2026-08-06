//! Tracked-change declarations, policies, and marker edits.

use super::super::super::model::MutableDocument;
use litchi_core::Result;

impl MutableDocument {
    /// Return declarations, policy, and marker-correlated content from current XML.
    pub fn tracked_changes(&self) -> Result<crate::TrackedChanges> {
        self.with_content_xml(crate::parser::Parser::parse_tracked_changes)
    }

    /// Atomically replace the declaration table and policy metadata.
    pub fn set_tracked_changes(&mut self, tracked: crate::TrackedChanges) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::set_tracked_changes_xml(xml, Some(&tracked)))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Atomically update tracking policy and inert protection metadata.
    pub fn set_tracked_change_policy(
        &mut self,
        track_changes: Option<bool>,
        protection_key: Option<String>,
        digest_algorithm: Option<String>,
    ) -> Result<()> {
        let mut tracked = self.tracked_changes()?;
        tracked.track_changes = track_changes;
        tracked.protection_key = protection_key;
        tracked.protection_key_digest_algorithm = digest_algorithm;
        self.set_tracked_changes(tracked)
    }

    /// Atomically append a declaration in insertion order.
    pub fn add_tracked_change(&mut self, change: crate::TrackChange) -> Result<()> {
        let mut tracked = self.tracked_changes()?;
        tracked.changes.push(change);
        self.set_tracked_changes(tracked)
    }

    /// Atomically replace a declaration without changing marker identity.
    pub fn update_tracked_change(
        &mut self,
        id: &str,
        replacement: crate::TrackChange,
    ) -> Result<crate::TrackChange> {
        if replacement.id != id {
            return Err(litchi_core::Error::InvalidFormat(
                "tracked-change update cannot change its stable ID".to_string(),
            ));
        }
        let mut tracked = self.tracked_changes()?;
        let index = tracked
            .changes
            .iter()
            .position(|change| change.id == id)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "tracked-change declaration '{id}' was not found"
                ))
            })?;
        let old = std::mem::replace(&mut tracked.changes[index], replacement);
        self.set_tracked_changes(tracked)?;
        Ok(old)
    }

    /// Remove a declaration and all of its correlated markers atomically.
    pub fn remove_tracked_change(&mut self, id: &str) -> Result<crate::TrackChange> {
        let mut tracked = self.tracked_changes()?;
        let index = tracked
            .changes
            .iter()
            .position(|change| change.id == id)
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "tracked-change declaration '{id}' was not found"
                ))
            })?;
        let removed = tracked.changes.remove(index);
        let updated = self.with_content_xml(|xml| {
            let unmarked = crate::unmark_tracked_change_xml(xml, id)?;
            crate::set_tracked_changes_xml(&unmarked, Some(&tracked))
        })?;
        self.content_xml = Some(updated);
        Ok(removed)
    }

    /// Remove all declarations, policy, and correlated markers.
    pub fn clear_tracked_changes(&mut self) -> Result<()> {
        let tracked = self.tracked_changes()?;
        let updated = self.with_content_xml(|xml| {
            let mut candidate = xml.to_string();
            for change in &tracked.changes {
                candidate = crate::unmark_tracked_change_xml(&candidate, &change.id)?;
            }
            crate::set_tracked_changes_xml(&candidate, None)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Mark a live insertion or format-change range using Unicode character offsets.
    pub fn mark_tracked_change_range(
        &mut self,
        change_id: &str,
        start: crate::Position,
        end: crate::Position,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::mark_tracked_change_range_xml(xml, change_id, &start, &end)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Place a point deletion marker using a Unicode character offset.
    pub fn mark_tracked_deletion(
        &mut self,
        change_id: &str,
        position: crate::Position,
    ) -> Result<()> {
        let updated = self
            .with_content_xml(|xml| crate::mark_tracked_deletion_xml(xml, change_id, &position))?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Remove every marker for one declaration while retaining its live text.
    pub fn unmark_tracked_change(&mut self, change_id: &str) -> Result<()> {
        let updated =
            self.with_content_xml(|xml| crate::unmark_tracked_change_xml(xml, change_id))?;
        self.content_xml = Some(updated);
        Ok(())
    }
}
