//! Section-scoped body text reading and mutation.

use std::ops::Range;
use std::sync::Arc;

use super::PagesEditor;
use crate::text::{
    TextBookmark, TextBookmarkId, TextBookmarkSettings, TextDateTimeField, TextDateTimeFieldId,
    TextPosition, TextRange,
};
use crate::{Error, Result};
use litchi_iwa_text::date_time::{DisplayText, Settings};
use litchi_pages::{
    Package, PackageError, SectionSelector, SectionTextError, section::SectionType,
};

impl PagesEditor {
    /// Read every native ranged bookmark in the main body.
    pub fn body_bookmarks(&self) -> Result<Vec<TextBookmark>> {
        self.text.text_bookmarks(self.body_storage_id)
    }

    /// Create a native body bookmark over a nonempty UTF-16 range.
    pub fn add_body_bookmark(
        &mut self,
        range: TextRange,
        settings: TextBookmarkSettings,
    ) -> Result<TextBookmark> {
        self.text
            .add_text_bookmark(self.body_storage_id, range, settings)
    }

    /// Atomically update a body bookmark's range and settings.
    pub fn update_body_bookmark(
        &mut self,
        id: TextBookmarkId,
        range: TextRange,
        settings: TextBookmarkSettings,
    ) -> Result<TextBookmark> {
        self.text
            .update_text_bookmark(self.body_storage_id, id, range, settings)
    }

    /// Delete one native body bookmark and reclaim its owned field object.
    pub fn remove_body_bookmark(&mut self, id: TextBookmarkId) -> Result<TextBookmark> {
        self.text.remove_text_bookmark(self.body_storage_id, id)
    }

    /// Read every native Date & Time field in the main body.
    pub fn body_date_time_fields(&self) -> Result<Vec<TextDateTimeField>> {
        self.text.text_date_time_fields(self.body_storage_id)
    }

    /// Attach a Date & Time field to existing body text.
    pub fn add_body_date_time_field(
        &mut self,
        range: TextRange,
        settings: Settings,
    ) -> Result<TextDateTimeField> {
        self.text
            .add_text_date_time_field(self.body_storage_id, range, settings)
    }

    /// Atomically insert exact display text and its Date & Time field.
    pub fn insert_body_date_time_field(
        &mut self,
        position: TextPosition,
        display_text: DisplayText,
        settings: Settings,
    ) -> Result<TextDateTimeField> {
        self.text.insert_text_date_time_field(
            self.body_storage_id,
            position,
            display_text,
            settings,
        )
    }

    /// Atomically update a body Date & Time field's range and formatter payload.
    pub fn update_body_date_time_field(
        &mut self,
        id: TextDateTimeFieldId,
        range: TextRange,
        settings: Settings,
    ) -> Result<TextDateTimeField> {
        self.text
            .update_text_date_time_field(self.body_storage_id, id, range, settings)
    }

    /// Delete one body Date & Time field while retaining its visible text.
    pub fn remove_body_date_time_field(
        &mut self,
        id: TextDateTimeFieldId,
    ) -> Result<TextDateTimeField> {
        self.text
            .remove_text_date_time_field(self.body_storage_id, id)
    }

    /// Replace a UTF-16 range in the body without creating or deleting section boundaries.
    ///
    /// Ranges may edit content on either side of a boundary, but cannot consume the native
    /// U+0004 section-break marker. Use [`Self::insert_section`] or [`Self::remove_section`] to
    /// change the section graph.
    pub fn replace_body_text(&mut self, range: Range<usize>, replacement: &str) -> Result<()> {
        self.validate_body_edit(&range, replacement)?;
        let footnotes =
            super::footnotes::body_footnote_graphs(self.package(), self.body_storage_id.get())?;
        let mut staged = self.text.clone();
        staged.replace_text(self.body_storage_id, range, replacement)?;
        let mut package = staged.into_package();
        super::footnotes::cleanup_removed_body_footnotes(
            &mut package,
            self.body_storage_id.get(),
            &footnotes,
        )?;
        *self = Self::from_package(package)?;
        Ok(())
    }

    /// Replace the complete body of a single-section document.
    ///
    /// Multi-section documents must be edited with [`Self::set_section_text`] so their native
    /// section breaks cannot be discarded accidentally.
    pub fn set_body_text(&mut self, replacement: &str) -> Result<()> {
        if self.sections.len() > 1 {
            return Err(Error::ParseError(
                "Cannot replace a multi-section Pages body; use set_section_text".to_owned(),
            ));
        }
        let body_length = self.body_text()?.encode_utf16().count();
        self.replace_body_text(0..body_length, replacement)
    }

    /// Clear the complete body of a single-section document.
    pub fn clear_body(&mut self) -> Result<()> {
        self.set_body_text("")
    }

    /// Read the text owned by one reachable section, excluding native section-break markers.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Pages section-text API; use litchi_pages::Package::section_text with SectionSelector; retained for migration-host compatibility"
    )]
    pub fn section_text(&self, section_id: u64) -> Result<String> {
        let position = self.section_position(section_id)?;
        if let Some(focused) = self.focused_section_package()? {
            match focused
                .package
                .section_text(SectionSelector::index(position))
            {
                Ok(text) => return Ok(text.to_owned()),
                Err(error) if semantic_read_fallback(&error) => {},
                Err(error) => return Err(map_section_text_error(error)),
            }
        }
        self.legacy_section_text(section_id)
    }

    /// Read one section through the legacy editor's native graph.
    ///
    /// This remains the compatibility path for nested `Index.zip` sources and
    /// focused semantic shapes that cannot yet be represented without losing
    /// supported host behavior.
    fn legacy_section_text(&self, section_id: u64) -> Result<String> {
        let body = self.body_text()?;
        let units = body.encode_utf16().collect::<Vec<_>>();
        let range = self.section_content_range(section_id, &units)?;
        String::from_utf16(&units[range]).map_err(|_| {
            Error::InvalidFormat(format!(
                "Pages section {section_id} boundary splits a UTF-16 surrogate pair"
            ))
        })
    }

    /// Replace a section-relative UTF-16 text range.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Pages section-text API; use litchi_pages::Package::edit_section_text with SectionSelector; retained for migration-host compatibility"
    )]
    pub fn replace_section_text(
        &mut self,
        section_id: u64,
        range: Range<usize>,
        replacement: &str,
    ) -> Result<()> {
        if range.start > range.end {
            return Err(Error::ParseError(
                "Section text replacement range starts after it ends".to_owned(),
            ));
        }
        let body = self.body_text()?;
        let units = body.encode_utf16().collect::<Vec<_>>();
        let content = self.section_content_range(section_id, &units)?;
        let content_length = content.end - content.start;
        if range.end > content_length {
            return Err(Error::ParseError(format!(
                "Pages section {section_id} text range {}..{} exceeds its UTF-16 length {content_length}",
                range.start, range.end
            )));
        }
        let absolute_start = content
            .start
            .checked_add(range.start)
            .ok_or_else(|| Error::ParseError("Pages section text range overflow".to_owned()))?;
        let absolute_end = content
            .start
            .checked_add(range.end)
            .ok_or_else(|| Error::ParseError("Pages section text range overflow".to_owned()))?;
        self.replace_body_text(absolute_start..absolute_end, replacement)
    }

    /// Replace all text owned by one section while preserving its layout and neighboring sections.
    #[deprecated(
        since = "0.0.1",
        note = "legacy raw-ID Pages section-text API; use litchi_pages::Package::set_section_text with SectionSelector; retained for migration-host compatibility"
    )]
    pub fn set_section_text(&mut self, section_id: u64, replacement: &str) -> Result<()> {
        let position = self.section_position(section_id)?;
        if let Some(mut focused) = self.focused_section_package()? {
            let mut edit = match focused
                .package
                .edit_section_text(SectionSelector::index(position))
            {
                Ok(edit) => edit,
                Err(error) if semantic_edit_fallback(&error) => {
                    return self.legacy_set_section_text(section_id, replacement);
                },
                Err(error) => return Err(map_section_text_error(error)),
            };
            match edit.set(replacement) {
                Ok(_) => {},
                Err(error) if semantic_edit_fallback(&error) => {
                    return self.legacy_set_section_text(section_id, replacement);
                },
                Err(error) => return Err(map_section_text_error(error)),
            }
            match edit.commit() {
                Ok(commit) => {
                    if commit.patch().is_noop() {
                        return Ok(());
                    }
                    self.publish_focused_section_commit(
                        section_id,
                        replacement,
                        commit,
                        &mut focused.budget,
                    )
                },
                Err(error) if semantic_edit_fallback(&error) => {
                    self.legacy_set_section_text(section_id, replacement)
                },
                Err(error) => Err(map_section_text_error(error)),
            }
        } else {
            self.legacy_set_section_text(section_id, replacement)
        }
    }

    /// Clear all text owned by one section while preserving the section itself.
    #[allow(deprecated)]
    pub fn clear_section_text(&mut self, section_id: u64) -> Result<()> {
        self.set_section_text(section_id, "")
    }

    #[allow(deprecated)]
    fn legacy_set_section_text(&mut self, section_id: u64, replacement: &str) -> Result<()> {
        let length = self.legacy_section_text(section_id)?.encode_utf16().count();
        self.replace_section_text(section_id, 0..length, replacement)
    }

    fn section_position(&self, section_id: u64) -> Result<usize> {
        self.sections
            .iter()
            .position(|section| section.object_id == section_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Section {section_id} is not reachable from the Pages body"
                ))
            })
    }

    /// Prepare the focused immutable Pages package for a selector-first text
    /// transaction. Nested legacy bundles deliberately stay on the host
    /// writer because the focused package refuses changed non-exact sources.
    fn focused_section_package(&self) -> Result<Option<FocusedSectionPackage>> {
        if !self.package().source_is_exact() {
            return Ok(None);
        }

        let source = self.package().exact_source_bytes().ok_or_else(|| {
            Error::InvalidFormat(
                "Pages focused section-text source provenance disappeared".to_owned(),
            )
        })?;
        let mut budget = FocusedBridgeBudget::new(self.package().limits(), source.len())?;
        // The focused package parser walks the retained ZIP/IWA source. Charge
        // a conservative source-sized reservation before entering it so an
        // exhausted bridge refuses before any uncharged ingress scan.
        budget.charge_scaled(source.len(), FOCUSED_INGRESS_FACTOR, "ingress")?;
        let limits = focused_pages_limits(self.package().limits())?;
        match Package::from_bytes_with_limits(source, limits) {
            Ok(package) if self.focused_sections_match(&package, &mut budget)? => {
                Ok(Some(FocusedSectionPackage { package, budget }))
            },
            Ok(_) => Ok(None),
            Err(error) if focused_source_fallback(&error) => Ok(None),
            Err(error) => Err(Error::InvalidFormat(format!(
                "Pages focused section-text ingress failed: {error}"
            ))),
        }
    }

    fn focused_sections_match(
        &self,
        package: &Package,
        budget: &mut FocusedBridgeBudget,
    ) -> Result<bool> {
        let source = self.package().exact_source_bytes().ok_or_else(|| {
            Error::InvalidFormat(
                "Pages focused section-text source provenance disappeared".to_owned(),
            )
        })?;
        // Matching traverses every projected section and reconstructs the
        // host-owned ranges. Reserve the source-sized bound before doing that
        // work; a mismatch must not become an uncharged fallback scan.
        budget.charge(source.len(), "projection matching")?;

        if package.sections().len() != self.sections.len() {
            return Ok(false);
        }
        let body = self.body_text()?;
        let units = body.encode_utf16().collect::<Vec<_>>();
        let mut expected_start = 0usize;
        for (index, (focused, host)) in package.sections().iter().zip(&self.sections).enumerate() {
            if focused.index() != index
                || focused.section_type() != SectionType::Body
                || focused.name() != host.name.as_deref()
                || usize::try_from(host.character_index).ok() != Some(expected_start)
            {
                return Ok(false);
            }

            // The matcher already has the host section's position. Reuse it so
            // each section range does not scan the section list from the start.
            let content = self.section_content_range_at(index, &units)?;
            let content_length = content.end - content.start;
            let host_text = String::from_utf16(&units[content]).map_err(|_| {
                Error::InvalidFormat(format!(
                    "Pages section {} boundary splits a UTF-16 surrogate pair",
                    host.object_id
                ))
            })?;
            let Some(focused_text) = focused.body_text() else {
                return Ok(false);
            };
            if focused_text != host_text {
                return Ok(false);
            }

            expected_start = expected_start.checked_add(content_length).ok_or_else(|| {
                Error::InvalidFormat(
                    "Pages focused section-text topology exceeds the platform index range"
                        .to_owned(),
                )
            })?;
            if index + 1 < self.sections.len() {
                expected_start = expected_start.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat(
                        "Pages focused section-text topology exceeds the platform index range"
                            .to_owned(),
                    )
                })?;
            }
        }

        // The section-range checks above validate native U+0004 separators;
        // this final length check also proves that no unowned body suffix or
        // prefix was silently omitted by either projection.
        if expected_start != units.len() {
            return Ok(false);
        }
        Ok(true)
    }

    fn publish_focused_section_commit(
        &mut self,
        section_id: u64,
        replacement: &str,
        commit: litchi_pages::SectionTextCommit,
        budget: &mut FocusedBridgeBudget,
    ) -> Result<()> {
        let source_limits = self.package().limits();
        let target_source = commit.package().source_bytes();
        budget.charge_scaled(
            target_source.len(),
            FOCUSED_READBACK_FACTOR,
            "candidate readback",
        )?;
        let mut target_bytes = Vec::new();
        target_bytes
            .try_reserve_exact(target_source.len())
            .map_err(|_| {
                Error::InvalidFormat(
                    "Pages focused section-text candidate allocation was refused".to_owned(),
                )
            })?;
        target_bytes.extend_from_slice(target_source);
        let target: Arc<[u8]> = target_bytes.into();
        let candidate =
            crate::package::IWorkPackage::from_shared_bytes_with_limits(target, source_limits)?;
        let verified = Self::from_package(candidate)?;
        let old_length = self.legacy_section_text(section_id)?.encode_utf16().count();
        if !host_section_topology_matches(
            &self.sections,
            &verified.sections,
            section_id,
            old_length,
            replacement,
        )? || verified.legacy_section_text(section_id)? != replacement
        {
            return Err(Error::InvalidFormat(
                "Pages focused section-text candidate readback failed".into(),
            ));
        }
        *self = verified;
        Ok(())
    }

    fn validate_body_edit(&self, range: &Range<usize>, replacement: &str) -> Result<()> {
        if range.start > range.end {
            return Err(Error::ParseError(
                "Text replacement range starts after it ends".to_owned(),
            ));
        }
        if replacement.contains('\u{4}') {
            return Err(Error::ParseError(
                "Pages section breaks must be changed through section CRUD APIs".to_owned(),
            ));
        }
        if replacement.contains('\u{e}') {
            return Err(Error::ParseError(
                "Pages footnote anchors must be changed through footnote CRUD APIs".to_owned(),
            ));
        }
        if replacement.contains('\u{fffc}') {
            return Err(Error::ParseError(
                "Pages inline-object markers must be changed through object CRUD APIs".to_owned(),
            ));
        }
        let body = self.body_text()?;
        let units = body.encode_utf16().collect::<Vec<_>>();
        if range.end > units.len() {
            return Err(Error::ParseError(format!(
                "Text replacement range {}..{} exceeds body UTF-16 length {}",
                range.start,
                range.end,
                units.len()
            )));
        }
        for section in self.sections.iter().skip(1) {
            let boundary = usize::try_from(section.character_index).map_err(|_| {
                Error::InvalidFormat(format!(
                    "Pages section {} boundary exceeds the platform index range",
                    section.object_id
                ))
            })?;
            let marker = boundary.checked_sub(1).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages section {} has an invalid zero boundary",
                    section.object_id
                ))
            })?;
            if units.get(marker) != Some(&0x0004) {
                return Err(Error::InvalidFormat(format!(
                    "Pages section {} is not preceded by a native section-break marker",
                    section.object_id
                )));
            }
            if range.start <= marker && marker < range.end {
                return Err(Error::ParseError(format!(
                    "Text replacement range {}..{} crosses the section break before section {}",
                    range.start, range.end, section.object_id
                )));
            }
        }
        Ok(())
    }

    fn section_content_range(&self, section_id: u64, body: &[u16]) -> Result<Range<usize>> {
        let index = self
            .sections
            .iter()
            .position(|section| section.object_id == section_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Section {section_id} is not reachable from the Pages body"
                ))
            })?;
        self.section_content_range_at(index, body)
    }

    fn section_content_range_at(&self, index: usize, body: &[u16]) -> Result<Range<usize>> {
        let section = self.sections.get(index).ok_or_else(|| {
            Error::ParseError(format!(
                "Section index {index} is not reachable from the Pages body"
            ))
        })?;
        let section_id = section.object_id;
        let start = usize::try_from(section.character_index).map_err(|_| {
            Error::InvalidFormat(format!(
                "Pages section {section_id} boundary exceeds the platform index range"
            ))
        })?;
        let end = if let Some(next) = self.sections.get(index + 1) {
            let boundary = usize::try_from(next.character_index).map_err(|_| {
                Error::InvalidFormat(format!(
                    "Pages section {} boundary exceeds the platform index range",
                    next.object_id
                ))
            })?;
            let marker = boundary.checked_sub(1).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Pages section {} has an invalid zero boundary",
                    next.object_id
                ))
            })?;
            if body.get(marker) != Some(&0x0004) {
                return Err(Error::InvalidFormat(format!(
                    "Pages section {} is not preceded by a native section-break marker",
                    next.object_id
                )));
            }
            marker
        } else {
            body.len()
        };
        if start > end || end > body.len() {
            return Err(Error::InvalidFormat(format!(
                "Pages section {section_id} has invalid body range {start}..{end}"
            )));
        }
        Ok(start..end)
    }
}

struct FocusedSectionPackage {
    package: Package,
    budget: FocusedBridgeBudget,
}

#[derive(Debug, Clone, Copy)]
struct FocusedBridgeBudget {
    maximum: usize,
    consumed: usize,
}

const FOCUSED_INGRESS_FACTOR: usize = 2;
const FOCUSED_READBACK_FACTOR: usize = 2;

impl FocusedBridgeBudget {
    fn new(limits: crate::package::PackageLimits, _source_len: usize) -> Result<Self> {
        let maximum = usize::try_from(limits.max_total_bytes())
            .unwrap_or(usize::MAX)
            .min(litchi_iwa_common::WireLimits::MAX_REWRITE_WORK);
        Ok(Self {
            maximum,
            consumed: 0,
        })
    }

    fn charge(&mut self, amount: usize, stage: &str) -> Result<()> {
        let observed = self.consumed.checked_add(amount).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages focused section-text {stage} exceeds its bounded work budget"
            ))
        })?;
        if observed > self.maximum {
            return Err(Error::InvalidFormat(format!(
                "Pages focused section-text {stage} exceeds its bounded work budget"
            )));
        }
        self.consumed = observed;
        Ok(())
    }

    fn charge_scaled(&mut self, amount: usize, factor: usize, stage: &str) -> Result<()> {
        let amount = amount.checked_mul(factor).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages focused section-text {stage} exceeds its bounded work budget"
            ))
        })?;
        self.charge(amount, stage)
    }
}

fn host_section_topology_matches(
    before: &[super::PagesSectionInfo],
    after: &[super::PagesSectionInfo],
    section_id: u64,
    old_length: usize,
    replacement: &str,
) -> Result<bool> {
    if before.len() != after.len() {
        return Ok(false);
    }
    let selected = before
        .iter()
        .position(|section| section.object_id == section_id)
        .ok_or_else(|| Error::InvalidFormat("Pages focused section is not reachable".to_owned()))?;
    let replacement_units = u32::try_from(replacement.encode_utf16().count()).map_err(|_| {
        Error::InvalidFormat("Pages focused section replacement is too large".to_owned())
    })?;
    let old_length = u32::try_from(old_length)
        .map_err(|_| Error::InvalidFormat("Pages focused section text is too large".to_owned()))?;
    let delta = i64::from(replacement_units) - i64::from(old_length);

    for (index, (old, new)) in before.iter().zip(after).enumerate() {
        if old.object_id != new.object_id
            || old.name != new.name
            || old.first_template_id != new.first_template_id
            || old.even_template_id != new.even_template_id
            || old.odd_template_id != new.odd_template_id
        {
            return Ok(false);
        }
        let expected = if index <= selected {
            i64::from(old.character_index)
        } else {
            i64::from(old.character_index) + delta
        };
        if expected < 0 || u32::try_from(expected).ok() != Some(new.character_index) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn focused_pages_limits(source: crate::package::PackageLimits) -> Result<litchi_pages::Limits> {
    let limits = litchi_pages::Limits::new(
        source.max_input_bytes(),
        source.max_entries(),
        source.max_entry_bytes(),
        source.max_total_bytes(),
        source.max_iwa_stream_bytes(),
    )
    .map_err(|error| Error::InvalidFormat(format!("Pages focused limits are invalid: {error}")))?;
    limits
        .with_archive_limits(source.archive_limits())
        .map_err(|error| Error::InvalidFormat(format!("Pages focused limits are invalid: {error}")))
}

fn focused_source_fallback(error: &PackageError) -> bool {
    matches!(error, PackageError::NotPages | PackageError::Detection(_))
        || matches!(
            error,
            PackageError::InvalidFormat(reason)
                if opaque_optional_footnote_error(reason)
                    || malformed_optional_footnote_payload_error(reason)
        )
}

fn opaque_optional_footnote_error(reason: &str) -> bool {
    reason.starts_with("Pages body footnote")
        || reason.starts_with("Pages footnote reference")
        || reason.starts_with("Pages footnote storage")
        || reason.starts_with("Pages footnote marker")
}

fn malformed_optional_footnote_payload_error(reason: &str) -> bool {
    let Some(reason) = reason.strip_prefix("Pages body text payload failed bounded validation: ")
    else {
        return false;
    };
    reason.starts_with("storage table field 16:")
        || reason.starts_with("TSWP storage table field 16 ")
        || reason.starts_with("singular TSWP storage field 16 ")
}

fn semantic_read_fallback(error: &SectionTextError) -> bool {
    matches!(
        error,
        SectionTextError::UnsupportedSource | SectionTextError::InvalidSource
    )
}

fn semantic_edit_fallback(error: &SectionTextError) -> bool {
    matches!(
        error,
        SectionTextError::UnsupportedSource
            | SectionTextError::DependentContent
            | SectionTextError::FootnoteAnchorReplacement
    )
}

fn map_section_text_error(error: SectionTextError) -> Error {
    Error::InvalidFormat(format!("Pages section text transaction failed: {error}"))
}

#[cfg(test)]
mod tests {
    use super::FocusedBridgeBudget;
    use crate::package::PackageLimits;

    #[test]
    fn focused_bridge_budget_refuses_exhaustion_before_scan() {
        let limits = PackageLimits::new_with_limits(1, 1, 1, 1, 1).unwrap();
        let mut budget = FocusedBridgeBudget::new(limits, 1).unwrap();
        assert!(budget.charge_scaled(1, 2, "test").is_err());
    }

    #[test]
    fn focused_bridge_budget_respects_shared_rewrite_work_ceiling() {
        let mut budget = FocusedBridgeBudget::new(PackageLimits::default(), 0).unwrap();
        assert_eq!(
            budget.maximum,
            litchi_iwa_common::WireLimits::MAX_REWRITE_WORK
        );
        assert!(
            budget
                .charge(litchi_iwa_common::WireLimits::MAX_REWRITE_WORK, "test")
                .is_ok()
        );
        assert!(budget.charge(1, "test").is_err());
    }
}
