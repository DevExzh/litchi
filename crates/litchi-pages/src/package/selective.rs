//! Opt-in, section-scoped Pages text ingress.
//!
//! The ordinary [`super::Package`] reader publishes a complete semantic
//! document. This module deliberately owns a smaller lifecycle for callers
//! that need one text object only: physical ZIP/IWA framing is still checked
//! for the complete source, while application payload interpretation stops at
//! the Pages root, one body storage, and its section-boundary table.

#![allow(
    clippy::module_name_repetitions,
    reason = "Selective lifecycle types retain their explicit API-family names."
)]

use std::cmp::{max, min};
use std::str;

use litchi_iwa_archive::{ComponentCatalog, SourceCatalog, SourceProvenance};
use litchi_iwa_core::ArchiveObject;
use litchi_iwa_index::{ByteSpan, FragmentId, IndexBuilder, ObjectId, ObjectIndex, ObjectRecord};

use super::{
    Limits, MAX_SECTIONS, PackageError, PackageResult, effective_text_limit,
    native_section_references, parse_wire, preflight_body_wire, root_references_with_limits,
    unique_message_payload, unique_text_payload,
};

/// A positional selector for the opt-in selective text lifecycle.
///
/// The selector intentionally has no native object identifier or name lookup
/// variant. Resolving a name would require decoding every selected section's
/// settings payload, which is outside this bounded candidate's ownership.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SelectiveTextSelector {
    index: usize,
}

impl SelectiveTextSelector {
    /// Construct a selector for one zero-based semantic section position.
    #[must_use]
    pub const fn section(index: usize) -> Self {
        Self { index }
    }

    /// Return the selected zero-based semantic section position.
    #[must_use]
    pub const fn index(self) -> usize {
        self.index
    }
}

impl From<usize> for SelectiveTextSelector {
    fn from(index: usize) -> Self {
        Self::section(index)
    }
}

/// Finite policy for one selective section-text lifecycle.
///
/// The physical [`Limits`] profile remains authoritative for ZIP, Snappy, and
/// IWA ingress. `max_text_bytes` is an additional output ceiling for the one
/// selected section and may only tighten the hard Pages text envelope.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SelectiveTextOptions {
    limits: Limits,
    max_text_bytes: usize,
    collect_source_metrics: bool,
}

impl Default for SelectiveTextOptions {
    fn default() -> Self {
        Self {
            limits: Limits::default(),
            max_text_bytes: super::effective_text_limit(Limits::default()),
            collect_source_metrics: false,
        }
    }
}

impl SelectiveTextOptions {
    /// Construct a selective profile from a checked physical limit candidate.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::Archive`] when the physical profile is invalid.
    pub fn new(limits: Limits) -> PackageResult<Self> {
        let limits = limits.validate()?;
        Ok(Self {
            max_text_bytes: effective_text_limit(limits),
            limits,
            collect_source_metrics: false,
        })
    }

    /// Replace the physical profile and reset the selected-text ceiling to its
    /// maximum value for that profile.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::Archive`] when the physical profile is invalid.
    pub fn with_limits(self, limits: Limits) -> PackageResult<Self> {
        Self::new(limits).map(|replacement| Self {
            collect_source_metrics: self.collect_source_metrics,
            ..replacement
        })
    }

    /// Tighten the selected section's retained UTF-8 ceiling.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::PayloadLimit`] when `maximum` is zero or would
    /// exceed the hard text envelope authorized by the physical profile.
    pub fn with_max_text_bytes(mut self, maximum: usize) -> PackageResult<Self> {
        let hard_maximum = effective_text_limit(self.limits);
        if maximum == 0 || maximum > hard_maximum {
            return Err(PackageError::PayloadLimit {
                observed: maximum,
                limit: hard_maximum,
            });
        }
        self.max_text_bytes = maximum;
        Ok(self)
    }

    /// Enable or disable content-free source metrics for the returned handle.
    #[must_use]
    pub const fn with_source_metrics(mut self, enabled: bool) -> Self {
        self.collect_source_metrics = enabled;
        self
    }

    /// Return the checked physical profile retained by this lifecycle.
    #[must_use]
    pub const fn limits(self) -> Limits {
        self.limits
    }

    /// Return the selected output UTF-8 ceiling.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }

    /// Return whether source metrics are requested.
    #[must_use]
    pub const fn collects_source_metrics(self) -> bool {
        self.collect_source_metrics
    }
}

/// Content-free source accounting for one selective lifecycle.
///
/// These counters describe physical traversal and selected payload shape;
/// they never retain document text, native identifiers, or package paths.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceMetrics {
    source_bytes: usize,
    component_count: usize,
    object_count: usize,
    message_count: usize,
    selected_message_count: usize,
    selected_payload_bytes: usize,
    opaque_member_count: usize,
}

impl SourceMetrics {
    /// Return the exact retained package source length.
    #[must_use]
    pub const fn source_bytes(self) -> usize {
        self.source_bytes
    }

    /// Return the number of physically parsed IWA components.
    #[must_use]
    pub const fn component_count(self) -> usize {
        self.component_count
    }

    /// Return the number of physically parsed IWA objects.
    #[must_use]
    pub const fn object_count(self) -> usize {
        self.object_count
    }

    /// Return the number of raw IWA application messages retained by framing.
    #[must_use]
    pub const fn message_count(self) -> usize {
        self.message_count
    }

    /// Return the count of application payloads interpreted by this lifecycle.
    #[must_use]
    pub const fn selected_message_count(self) -> usize {
        self.selected_message_count
    }

    /// Return bytes in the root and selected body payloads interpreted by the
    /// lifecycle.
    #[must_use]
    pub const fn selected_payload_bytes(self) -> usize {
        self.selected_payload_bytes
    }

    /// Return the number of unsupported-compression package members retained
    /// as opaque source data.
    #[must_use]
    pub const fn opaque_member_count(self) -> usize {
        self.opaque_member_count
    }
}

/// Exact source-backed result of one selective section-text lifecycle.
///
/// The text is the only owned semantic payload. The original source catalog
/// remains authoritative for exact no-op/preservation workflows, including
/// unsupported ZIP members and unknown IWA message bytes.
#[derive(Debug)]
pub struct SelectedSectionText {
    source: SourceCatalog,
    selector: SelectiveTextSelector,
    text: Box<str>,
    source_metrics: Option<SourceMetrics>,
}

impl SelectedSectionText {
    /// Borrow the selected section's UTF-8 text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Return the lifecycle selector used to produce this result.
    #[must_use]
    pub const fn selector(&self) -> SelectiveTextSelector {
        self.selector
    }

    /// Borrow the exact retained source bytes.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        self.source.source_bytes()
    }

    /// Return whether the logical source remains an exact ZIP representation.
    #[must_use]
    pub const fn source_is_exact(&self) -> bool {
        self.source.source_is_exact()
    }

    /// Return the source provenance retained by the physical catalog.
    #[must_use]
    pub const fn source_provenance(&self) -> SourceProvenance {
        self.source.source_provenance()
    }

    /// Borrow optional source metrics requested by the lifecycle profile.
    #[must_use]
    pub const fn source_metrics(&self) -> Option<SourceMetrics> {
        self.source_metrics
    }
}

pub(super) fn select_section_text(
    bytes: &[u8],
    selector: SelectiveTextSelector,
    options: SelectiveTextOptions,
) -> PackageResult<SelectedSectionText> {
    let limits = options.limits.validate()?;
    let source = SourceCatalog::from_bytes_with_limits(bytes, limits)?;
    let components = source.components();
    let object_count = super::validate_components(components)?;
    let root = root_references_with_limits(components, limits)?;
    let object_index = build_object_index(components, object_count, root)?;
    let body_identifier = root
        .body
        .ok_or_else(|| PackageError::InvalidFormat("Pages root has no body storage".to_owned()))?;
    let body_id = ObjectId::try_from(body_identifier.get()).map_err(|error| {
        PackageError::InvalidFormat(format!(
            "Pages body storage object {body_identifier} has an invalid identifier: {error}"
        ))
    })?;
    let body = object_index.object(components, body_id).ok_or_else(|| {
        PackageError::InvalidFormat(format!(
            "Pages body storage object {body_identifier} is missing"
        ))
    })?;
    let body_payload = unique_text_payload(&body.messages, body_identifier)?;
    let preflight = preflight_body_wire(
        body_payload,
        body_identifier,
        MAX_SECTIONS,
        effective_text_limit(limits),
        limits,
    )?;
    let references = native_section_references(
        preflight.section_references,
        root.initial_section,
        MAX_SECTIONS,
    )?;
    if selector.index() >= references.len() {
        return Err(PackageError::SelectiveSectionNotFound {
            index: selector.index(),
        });
    }
    let text = selected_text_from_body(
        body_payload,
        &references,
        selector.index(),
        body_identifier,
        options.max_text_bytes,
    )?;

    let source_metrics = options.collects_source_metrics().then(|| {
        let root_payload_bytes = components
            .get("Index/Document.iwa")
            .and_then(|component| component.archive().object(1))
            .and_then(|object| {
                unique_message_payload(&object.messages, 10_000, "Pages root object 1").ok()
            })
            .map_or(0, <[u8]>::len);
        collect_source_metrics(
            &source,
            object_count,
            root_payload_bytes,
            body_payload.len(),
        )
    });

    Ok(SelectedSectionText {
        source,
        selector,
        text: text.into_boxed_str(),
        source_metrics,
    })
}

/// Neutral IWA object locations plus the adapter-owned map back to raw
/// payload storage. The index never owns or decodes payload bytes; the
/// sidecar is only a private locator for the selected body object.
#[derive(Debug)]
struct PageObjectIndex {
    index: ObjectIndex,
    locations: Box<[IndexedObjectLocation]>,
}

#[derive(Debug, Clone, Copy)]
struct PendingObjectLocation {
    id: ObjectId,
    component_index: usize,
    object_index: usize,
}

#[derive(Debug, Clone, Copy)]
struct IndexedObjectLocation {
    component_index: usize,
    object_index: usize,
}

impl PageObjectIndex {
    fn object<'a>(
        &self,
        components: &'a ComponentCatalog,
        id: ObjectId,
    ) -> Option<&'a ArchiveObject> {
        let (position, _record) = self.index.object_with_position(id)?;
        let location = self.locations.get(position)?;
        let component = components.get_index(location.component_index)?;
        let object = component.archive().objects.get(location.object_index)?;
        (object.archive_info.identifier == Some(id.get())).then_some(object)
    }
}

fn build_object_index(
    components: &ComponentCatalog,
    object_count: usize,
    root: super::RootReferences,
) -> PackageResult<PageObjectIndex> {
    let mut builder = IndexBuilder::new();
    let mut pending_locations = Vec::new();
    pending_locations
        .try_reserve_exact(object_count)
        .map_err(|_allocation| PackageError::Allocation {
            amount: object_count,
        })?;

    for (component_index, component) in components.iter().enumerate() {
        let ordinal = component_index
            .checked_add(1)
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| {
                PackageError::InvalidFormat(
                    "Pages IWA component ordinal does not fit the neutral index".to_owned(),
                )
            })?;
        let fragment = FragmentId::try_from(ordinal).map_err(|error| {
            PackageError::InvalidFormat(format!(
                "Pages IWA component {} has an invalid neutral fragment: {error}",
                component.name()
            ))
        })?;
        builder.add_fragment(fragment).map_err(map_index_error)?;

        for (object_index, object) in component.archive().objects.iter().enumerate() {
            let identifier = object.archive_info.identifier.ok_or_else(|| {
                PackageError::InvalidFormat(format!(
                    "Pages component {} contains an object without an identifier",
                    component.name()
                ))
            })?;
            let id = ObjectId::try_from(identifier).map_err(|error| {
                PackageError::InvalidFormat(format!(
                    "Pages component {} contains an invalid object identifier {identifier}: {error}",
                    component.name()
                ))
            })?;
            let span = ByteSpan::new(object.data_offset, object.data_length).map_err(|error| {
                PackageError::InvalidFormat(format!(
                    "Pages object {identifier} source span is invalid: {error}"
                ))
            })?;
            builder
                .add_object(ObjectRecord::new(id, fragment, span))
                .map_err(map_index_error)?;
            pending_locations.push(PendingObjectLocation {
                id,
                component_index,
                object_index,
            });

            for message_info in &object.archive_info.message_infos {
                add_index_references(&mut builder, id, &message_info.object_references)?;
            }
        }
    }

    let root_id = ObjectId::try_from(1_u64).map_err(|error| {
        PackageError::InvalidFormat(format!("Pages root object identity is invalid: {error}"))
    })?;
    for target in [root.body, root.initial_section].into_iter().flatten() {
        let target = ObjectId::try_from(target.get()).map_err(|error| {
            PackageError::InvalidFormat(format!(
                "Pages root reference {target} is invalid: {error}"
            ))
        })?;
        builder
            .add_reference_if_absent(root_id, target)
            .map_err(map_index_error)?;
    }

    pending_locations.sort_unstable_by_key(|location| location.id);
    let mut locations = Vec::new();
    locations
        .try_reserve_exact(pending_locations.len())
        .map_err(|_allocation| PackageError::Allocation {
            amount: pending_locations.len(),
        })?;
    locations.extend(
        pending_locations
            .into_iter()
            .map(|location| IndexedObjectLocation {
                component_index: location.component_index,
                object_index: location.object_index,
            }),
    );
    let index = builder
        .build_allow_missing_targets()
        .map_err(map_index_error)?;
    Ok(PageObjectIndex {
        index,
        locations: locations.into_boxed_slice(),
    })
}

fn add_index_references(
    builder: &mut IndexBuilder,
    source: ObjectId,
    targets: &[u64],
) -> PackageResult<()> {
    for &target in targets {
        // A native zero reference is retained in the source catalog but has
        // no neutral graph edge; ObjectId intentionally rejects that sentinel.
        let Some(target) = ObjectId::new(target) else {
            continue;
        };
        builder
            .add_reference_if_absent(source, target)
            .map_err(map_index_error)?;
    }
    Ok(())
}

fn map_index_error(error: litchi_iwa_index::IndexError) -> PackageError {
    match error {
        litchi_iwa_index::IndexError::Allocation { requested, .. } => {
            PackageError::Allocation { amount: requested }
        },
        other => PackageError::InvalidFormat(format!(
            "Pages IWA object index rejected validated archive: {other}"
        )),
    }
}

fn collect_source_metrics(
    source: &SourceCatalog,
    object_count: usize,
    root_payload_bytes: usize,
    body_payload_bytes: usize,
) -> SourceMetrics {
    let message_count = source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .fold(0usize, |count, object| {
            count.saturating_add(object.messages.len())
        });
    let opaque_member_count = source
        .package()
        .iter()
        .filter(|entry| entry.is_opaque())
        .count();
    SourceMetrics {
        source_bytes: source.source_bytes().len(),
        component_count: source.components().len(),
        object_count,
        message_count,
        selected_message_count: 2,
        selected_payload_bytes: root_payload_bytes.saturating_add(body_payload_bytes),
        opaque_member_count,
    }
}

fn selected_text_from_body(
    payload: &[u8],
    references: &[super::NativeSectionReference],
    selected_index: usize,
    body_identifier: std::num::NonZeroU64,
    max_text_bytes: usize,
) -> PackageResult<String> {
    let context = format!("Pages body object {body_identifier}");
    let view = parse_wire(payload, &context)?;
    let mut points = Vec::new();
    points
        .try_reserve_exact(references.len())
        .map_err(|_allocation| PackageError::Allocation {
            amount: references.len(),
        })?;
    let mut reference_index = 0usize;
    let mut utf16_offset = 0usize;
    let mut byte_offset = 0usize;
    let mut preceding_character = None;

    for field in view.fields() {
        if field.number() != 3 {
            continue;
        }
        super::validate_wire_field(field, 2, &context)?;
        let fragment = str::from_utf8(field.payload()).map_err(|error| {
            PackageError::InvalidFormat(format!(
                "{context} text fragment is not valid UTF-8: {error}"
            ))
        })?;
        for character in fragment.chars() {
            super::capture_section_boundary(
                references,
                &mut reference_index,
                utf16_offset,
                byte_offset,
                preceding_character,
                &mut points,
            )?;
            let next_utf16 = utf16_offset
                .checked_add(character.len_utf16())
                .ok_or_else(|| {
                    PackageError::InvalidFormat(format!("{context} UTF-16 length overflows usize"))
                })?;
            if references.get(reference_index).is_some_and(|reference| {
                let target = reference.character_index as usize;
                target > utf16_offset && target < next_utf16
            }) {
                return Err(PackageError::InvalidFormat(format!(
                    "{context} section boundary splits a UTF-16 surrogate pair"
                )));
            }
            utf16_offset = next_utf16;
            byte_offset = byte_offset
                .checked_add(character.len_utf8())
                .ok_or_else(|| {
                    PackageError::InvalidFormat(format!("{context} UTF-8 length overflows usize"))
                })?;
            preceding_character = Some(character);
        }
    }
    super::capture_section_boundary(
        references,
        &mut reference_index,
        utf16_offset,
        byte_offset,
        preceding_character,
        &mut points,
    )?;
    if reference_index != references.len() {
        let reference = references[reference_index];
        return Err(PackageError::InvalidFormat(format!(
            "{context} section {} boundary {} exceeds body UTF-16 length {utf16_offset}",
            reference.identifier, reference.character_index
        )));
    }

    for (index, point) in points.iter().enumerate().skip(1) {
        if point.preceding_character != Some('\u{0004}') {
            return Err(PackageError::InvalidFormat(format!(
                "{context} section {} boundary {} is not preceded by a native section-break marker",
                references[index].identifier, references[index].character_index
            )));
        }
    }
    let start = points
        .get(selected_index)
        .map(|point| point.byte_offset)
        .ok_or(PackageError::SelectiveSectionNotFound {
            index: selected_index,
        })?;
    let end = points
        .get(selected_index + 1)
        .map(|point| point.byte_offset.saturating_sub(1))
        .unwrap_or(byte_offset);
    if end < start {
        return Err(PackageError::InvalidFormat(format!(
            "{context} selected section has an invalid text range"
        )));
    }
    let selected_bytes = end - start;
    if selected_bytes > max_text_bytes {
        return Err(PackageError::Semantic(super::SemanticError::TextTooLarge {
            observed: selected_bytes,
            limit: max_text_bytes,
        }));
    }

    let mut text = String::new();
    text.try_reserve_exact(selected_bytes)
        .map_err(|_allocation| PackageError::Allocation {
            amount: selected_bytes,
        })?;
    let mut cursor = 0usize;
    for field in view.fields() {
        if field.number() != 3 {
            continue;
        }
        let fragment = str::from_utf8(field.payload()).map_err(|error| {
            PackageError::InvalidFormat(format!(
                "{context} text fragment is not valid UTF-8: {error}"
            ))
        })?;
        let fragment_end = cursor.checked_add(fragment.len()).ok_or_else(|| {
            PackageError::InvalidFormat(format!("{context} UTF-8 length overflows usize"))
        })?;
        let overlap_start = max(start, cursor);
        let overlap_end = min(end, fragment_end);
        if overlap_start < overlap_end {
            let local_start = overlap_start - cursor;
            let local_end = overlap_end - cursor;
            let selected = fragment.get(local_start..local_end).ok_or_else(|| {
                PackageError::InvalidFormat(format!(
                    "{context} selected section boundary is not on a UTF-8 character boundary"
                ))
            })?;
            text.push_str(selected);
        }
        cursor = fragment_end;
    }
    if text.len() != selected_bytes {
        return Err(PackageError::InvalidFormat(format!(
            "{context} selected text length disagrees with its validated range"
        )));
    }
    Ok(text)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn body_payload(text: &[&str]) -> Vec<u8> {
        let mut output = Vec::new();
        for fragment in text {
            output.push(0x1a);
            output.push(u8::try_from(fragment.len()).expect("test fragment is short"));
            output.extend_from_slice(fragment.as_bytes());
        }
        output
    }

    fn references() -> Vec<super::super::NativeSectionReference> {
        vec![
            super::super::NativeSectionReference {
                character_index: 0,
                identifier: std::num::NonZeroU64::new(11).expect("non-zero test identifier"),
            },
            super::super::NativeSectionReference {
                character_index: 6,
                identifier: std::num::NonZeroU64::new(12).expect("non-zero test identifier"),
            },
        ]
    }

    #[test]
    fn selected_text_excludes_native_break_and_unselected_text() {
        let payload = body_payload(&["alpha\u{0004}", "beta"]);
        let selected = selected_text_from_body(
            &payload,
            &references(),
            1,
            std::num::NonZeroU64::new(9).expect("non-zero body identifier"),
            64,
        )
        .expect("bounded selected text");
        assert_eq!(selected, "beta");
    }

    #[test]
    fn selected_text_limit_is_checked_before_owned_publication() {
        let payload = body_payload(&["alpha\u{0004}", "beta"]);
        let error = selected_text_from_body(
            &payload,
            &references(),
            1,
            std::num::NonZeroU64::new(9).expect("non-zero body identifier"),
            3,
        )
        .expect_err("selected text exceeds its finite output profile");
        assert!(matches!(
            error,
            PackageError::Semantic(super::super::SemanticError::TextTooLarge {
                observed: 4,
                limit: 3,
            })
        ));
    }

    #[test]
    fn options_make_source_metrics_explicit_and_bounded() {
        let options = SelectiveTextOptions::new(Limits::default())
            .expect("default physical limits")
            .with_source_metrics(true);
        assert!(options.collects_source_metrics());
        assert!(options.max_text_bytes() <= effective_text_limit(Limits::default()));
        assert!(!SelectiveTextOptions::default().collects_source_metrics());
    }
}
