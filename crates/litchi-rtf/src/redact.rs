//! Bounded, source-proven redaction of passive top-level RTF references.
//!
//! This module intentionally covers only the two document-format destinations
//! already modeled by [`crate::DocumentExternalReferences`]: `nextfile` and
//! `template`.  It does not claim to sanitize fields, objects, pictures,
//! shapes, opaque syntax, signatures, or protection metadata.  Those surfaces
//! remain explicit diagnostics and make strict planning fail closed.

use crate::edit::Error;
use crate::{Document, RtfError};
use litchi_core::patch::{BlobBundle, BlobId, ForwardOnly, Patch, PatchLimits, PatchOperation};
use serde_json::Value;
use std::collections::BTreeMap;
use std::ops::Range;

const REDACTION_OPERATION: &str = "external-reference.remove";
const REDACTION_FORMAT: &str = "litchi-rtf";
const SOURCE_SPAN_PRECONDITION: &str = "source_span";
const MAX_REFERENCES: usize = 2;
const MAX_DIAGNOSTICS: usize = 64;

/// One modeled passive top-level RTF reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[non_exhaustive]
pub enum ReferenceKind {
    /// The `nextfile` document destination.
    NextFile,
    /// The `template` document destination.
    Template,
}

impl ReferenceKind {
    /// Stable diagnostic/patch spelling.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::NextFile => "next-file",
            Self::Template => "template",
        }
    }

    fn target(self) -> &'static str {
        match self {
            Self::NextFile => "document:external-reference:next-file",
            Self::Template => "document:external-reference:template",
        }
    }
}

/// One bounded inventory entry.  The value is retained because the existing
/// typed metadata API already exposes these passive names; no reference is
/// opened, resolved, or otherwise acted upon.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferenceOccurrence {
    kind: ReferenceKind,
    value: String,
    source_range: Option<Range<usize>>,
}

impl ReferenceOccurrence {
    /// Reference destination kind.
    #[must_use]
    pub const fn kind(&self) -> ReferenceKind {
        self.kind
    }

    /// Borrow the inert stored name.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }

    /// Exact whole-group byte range when the source is uncompressed ASCII.
    #[must_use]
    pub fn source_range(&self) -> Option<Range<usize>> {
        self.source_range.clone()
    }

    /// Whether this entry can be removed without canonical rewriting.
    #[must_use]
    pub const fn has_exact_source_range(&self) -> bool {
        self.source_range.is_some()
    }
}

/// Surfaces deliberately outside this narrow redaction closure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[non_exhaustive]
pub enum UnsupportedReference {
    /// A modeled reference has no exact source span (compressed/non-ASCII
    /// transport or unavailable provenance).
    MissingSourceSpan(ReferenceKind),
    /// An external or active field is retained.
    Field,
    /// A field safety classification was unknown or out of sync.
    UnknownField,
    /// A linked or otherwise unrecognized object is retained.
    Object,
    /// A picture carries hyperlink metadata.
    Picture,
    /// A shape carries hyperlink metadata.
    Shape,
    /// Opaque syntax was retained.
    OpaqueSyntax,
    /// Parser retained an unknown syntax marker.
    UnknownSyntax,
    /// An RTF file table is an external-reference surface outside this API.
    FileTable,
    /// A hyperlink base is an unmodeled external-reference owner.
    HyperlinkBase,
    /// Mail merge data may resolve external content.
    MailMerge,
    /// XSL metadata or its requested use is outside this closure.
    XslTransform,
    /// Linked-template policy is not the modeled `template` destination.
    LinkedTemplatePolicy,
    /// Any explicit protection metadata is a publication policy boundary.
    Protection,
    /// Write reservations are authentication-like metadata and are not
    /// rewritten by this closure.
    WriteReservation,
    /// Revision-save metadata is retained but not covered by this closure.
    RevisionSave,
}

/// Strict planning refuses when this mode leaves any unsupported surface.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Mode {
    /// Remove all modeled references and require a complete supported closure.
    Strict,
    /// Remove only proven modeled references and return incomplete diagnostics.
    BestEffort,
}

/// Effects and explicit incompleteness evidence returned by a redaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Diagnostics {
    removed_references: usize,
    unsupported: Box<[UnsupportedReference]>,
}

impl Diagnostics {
    /// Number of passive top-level reference groups removed.
    #[must_use]
    pub const fn removed_references(&self) -> usize {
        self.removed_references
    }

    /// Unsupported surfaces which remain in the ordinary returned document.
    #[must_use]
    pub fn unsupported(&self) -> &[UnsupportedReference] {
        &self.unsupported
    }

    /// Whether the requested closure was complete.
    #[must_use]
    pub const fn is_complete(&self) -> bool {
        self.unsupported.is_empty()
    }

    /// Whether this result is explicitly incomplete.
    #[must_use]
    pub const fn is_incomplete(&self) -> bool {
        !self.is_complete()
    }
}

/// Exact immutable source inventory for the narrow external-reference closure.
#[derive(Debug, Clone)]
pub struct Snapshot {
    source: Document,
    references: Vec<ReferenceOccurrence>,
    unsupported: Box<[UnsupportedReference]>,
}

impl Snapshot {
    /// Capture a bounded inventory from one immutable document snapshot.
    pub(crate) fn from_document(source: &Document) -> Result<Self, Error> {
        let model = source.model();
        let references = model.external_references();
        let spans = model.external_reference_source_spans();
        let source_bytes = model.preserved_source();

        let count = usize::from(references.next_file.is_some())
            .saturating_add(usize::from(references.template.is_some()));
        if count > MAX_REFERENCES {
            return Err(Error::Rtf(RtfError::LimitExceeded {
                resource: "RTF external-reference inventory entries",
                observed: count,
                limit: MAX_REFERENCES,
            }));
        }
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(count)
            .map_err(|_error| allocation("external-reference inventory", count))?;

        push_occurrence(
            &mut entries,
            ReferenceKind::NextFile,
            references.next_file.as_deref(),
            spans.next_file.clone(),
            source_bytes,
        )?;
        push_occurrence(
            &mut entries,
            ReferenceKind::Template,
            references.template.as_deref(),
            spans.template.clone(),
            source_bytes,
        )?;

        let mut unsupported = Vec::new();
        unsupported
            .try_reserve(8)
            .map_err(|_error| allocation("external-reference diagnostics", 8))?;
        for entry in &entries {
            if !entry.has_exact_source_range() {
                push_unique(
                    &mut unsupported,
                    UnsupportedReference::MissingSourceSpan(entry.kind),
                )?;
            }
        }
        collect_unsupported(model, &mut unsupported)?;

        Ok(Self {
            source: source.clone(),
            references: entries,
            unsupported: unsupported.into_boxed_slice(),
        })
    }

    /// Borrow all bounded modeled references in stable destination order.
    #[must_use]
    pub fn references(&self) -> &[ReferenceOccurrence] {
        &self.references
    }

    /// Borrow unsupported surfaces found during the non-mutating inventory.
    #[must_use]
    pub fn unsupported(&self) -> &[UnsupportedReference] {
        &self.unsupported
    }

    /// Build a non-mutating redaction plan under an explicit policy and patch
    /// bound.  The plan never changes the source snapshot.
    pub fn plan(&self, mode: Mode, limits: PatchLimits) -> Result<Plan, Error> {
        if matches!(mode, Mode::Strict) && !self.unsupported.is_empty() {
            return Err(Error::UnsupportedSource(
                "external-reference redaction has unsupported remaining surfaces",
            ));
        }
        Ok(Plan {
            source: self.clone(),
            mode,
            limits,
        })
    }
}

/// Non-mutating source-bound plan for removing all proven modeled references.
#[derive(Debug, Clone)]
pub struct Plan {
    source: Snapshot,
    mode: Mode,
    limits: PatchLimits,
}

impl Plan {
    /// Apply the plan atomically to an immutable candidate and seal its
    /// forward-only durable patch.  No inverse data is retained.
    pub fn apply(self) -> Result<Commit, Error> {
        let selected = selected_entries(&self.source.references)?;
        let (candidate, diagnostics) = rewrite_candidate(
            &self.source.source,
            &selected,
            self.mode,
            &self.source.unsupported,
        )?;
        let patch = build_patch(&self.source.source, &selected, self.mode, self.limits)?;
        Ok(Commit {
            document: candidate,
            patch,
            diagnostics,
        })
    }

    /// Source inventory used by this plan.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.source
    }
}

/// Ordinary document plus an intentionally irreversible forward-only patch.
#[derive(Debug, Clone)]
pub struct Commit {
    document: Document,
    patch: Patch<ForwardOnly>,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the ordinary candidate document.  This is not a certification
    /// type; callers must inspect [`Self::diagnostics`] before treating the
    /// result as complete.
    #[must_use]
    pub const fn document(&self) -> &Document {
        &self.document
    }

    /// Consume the commit and return its ordinary candidate document.
    #[must_use]
    pub fn into_document(self) -> Document {
        self.document
    }

    /// Borrow the sealed forward-only durable patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch<ForwardOnly> {
        &self.patch
    }

    /// Borrow explicit completeness/effect diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

fn push_occurrence(
    entries: &mut Vec<ReferenceOccurrence>,
    kind: ReferenceKind,
    value: Option<&str>,
    source_range: Option<Range<usize>>,
    source: Option<&[u8]>,
) -> Result<(), Error> {
    let Some(value) = value else { return Ok(()) };
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_error| allocation("external-reference value", value.len()))?;
    owned.push_str(value);
    let source_range = source_range.filter(|range| {
        source
            .and_then(|bytes| bytes.get(range.clone()))
            .is_some_and(|group| group.starts_with(b"{\\*\\") && group.ends_with(b"}"))
    });
    entries.push(ReferenceOccurrence {
        kind,
        value: owned,
        source_range,
    });
    Ok(())
}

fn collect_unsupported(
    model: &crate::raw::Document<'_>,
    unsupported: &mut Vec<UnsupportedReference>,
) -> Result<(), Error> {
    if !model.opaque_nodes().is_empty() {
        push_unique(unsupported, UnsupportedReference::OpaqueSyntax)?;
    }
    if model.unknown_syntax_markers() != 0 {
        push_unique(unsupported, UnsupportedReference::UnknownSyntax)?;
    }
    if model.file_table().is_some() {
        push_unique(unsupported, UnsupportedReference::FileTable)?;
    }
    if model.info().hyperlink_base.is_some() {
        push_unique(unsupported, UnsupportedReference::HyperlinkBase)?;
    }
    if model.mail_merge().is_some() {
        push_unique(unsupported, UnsupportedReference::MailMerge)?;
    }
    if model.xsl_transform().is_some() || model.xsl_transform_usage().is_requested() {
        push_unique(unsupported, UnsupportedReference::XslTransform)?;
    }
    if model.style_policies().update_styles_from_template {
        push_unique(unsupported, UnsupportedReference::LinkedTemplatePolicy)?;
    }
    let protection = model.protection();
    if protection.forms.is_some()
        || protection.annotations.is_some()
        || protection.revisions.is_some()
        || protection.read_only.is_some()
        || protection.all.is_some()
        || protection.enforced.is_some()
        || protection.level.is_some()
        || protection.password_hash.is_some()
        || !model.protection_ranges().is_empty()
        || !model.editable_regions().is_empty()
        || model.protection_user_table().is_some()
    {
        push_unique(unsupported, UnsupportedReference::Protection)?;
    }
    if !model.write_reservations().is_empty() {
        push_unique(unsupported, UnsupportedReference::WriteReservation)?;
    }
    if model.revision_save_metadata().is_some() {
        push_unique(unsupported, UnsupportedReference::RevisionSave)?;
    }

    let safety = model.field_safety();
    if safety.len() != model.fields().len() {
        push_unique(unsupported, UnsupportedReference::UnknownField)?;
    } else {
        for item in safety {
            match item {
                crate::validation::FieldSafety::Neutral => {},
                crate::validation::FieldSafety::External
                | crate::validation::FieldSafety::Active
                | crate::validation::FieldSafety::ExternalAndActive => {
                    push_unique(unsupported, UnsupportedReference::Field)?;
                },
                crate::validation::FieldSafety::ExternalUnknown
                | crate::validation::FieldSafety::ActiveUnknown
                | crate::validation::FieldSafety::ExternalAndActiveUnknown => {
                    push_unique(unsupported, UnsupportedReference::UnknownField)?;
                },
            }
        }
    }
    // Legacy FORMTEXT/FORMCHECKBOX/FORMDROPDOWN instructions live in the
    // generic field store; their separate `form_fields()` destination is not
    // sufficient to detect them.  They can carry macros, recalculation, and
    // host interaction flags, so a neutral generic field classification is
    // not a proof that the form surface is passive.
    if !model.form_fields().is_empty()
        || model.fields().iter().any(|field| {
            field.legacy_form_field().is_some()
                || matches!(
                    field.field_type,
                    crate::FieldType::FormText
                        | crate::FieldType::FormCheckbox
                        | crate::FieldType::FormDropdown
                )
        })
    {
        push_unique(unsupported, UnsupportedReference::Field)?;
    }

    // Every object destination is an active/unmodeled policy surface for this
    // narrow redaction closure.  Even `linkself` objects and nominally
    // embedded/OLE-control kinds can carry update behavior or payloads whose
    // ownership is not proven by the passive metadata model.
    if !model.objects().is_empty() {
        push_unique(unsupported, UnsupportedReference::Object)?;
    }
    for picture in model.pictures() {
        if picture.shape_properties.as_ref().is_some_and(|properties| {
            properties
                .properties
                .iter()
                .any(|property| property.hyperlink.is_some())
        }) {
            push_unique(unsupported, UnsupportedReference::Picture)?;
        }
    }
    if model_has_hyperlink_shapes(model) {
        push_unique(unsupported, UnsupportedReference::Shape)?;
    }
    Ok(())
}

/// Inspect every retained story owner.  A shape hyperlink is still an active
/// external-reference surface when it lives outside the body root: headers,
/// footers, notes, annotations, separators, table cells, field results, and
/// legacy text-box stories all retain independent drawing stores.
fn model_has_hyperlink_shapes(model: &crate::raw::Document<'_>) -> bool {
    story_shapes_have_hyperlink(model.shapes(), model.shape_groups())
        || model.tables().iter().any(table_has_hyperlink)
        || model
            .fields()
            .iter()
            .any(|field| story_shapes_have_hyperlink(&field.shapes, &field.shape_groups))
        || model.sections().iter().any(|section| {
            section
                .headers_footers
                .iter()
                .any(|story| story_shapes_have_hyperlink(&story.shapes, &story.shape_groups))
        })
        || model
            .notes()
            .iter()
            .any(|note| story_shapes_have_hyperlink(&note.shapes, &note.shape_groups))
        || model.annotations().iter().any(|annotation| {
            story_shapes_have_hyperlink(&annotation.shapes, &annotation.shape_groups)
        })
        || model.note_separators().entries().iter().any(|separator| {
            story_shapes_have_hyperlink(&separator.shapes, &separator.shape_groups)
        })
        || model
            .legacy_text_boxes()
            .iter()
            .any(|text_box| story_shapes_have_hyperlink(&text_box.shapes, &text_box.shape_groups))
}

fn story_shapes_have_hyperlink(
    shapes: &[crate::Shape<'_>],
    shape_groups: &[crate::ShapeGroup<'_>],
) -> bool {
    shapes.iter().any(shape_has_hyperlink) || shape_groups.iter().any(group_has_hyperlink)
}

fn table_has_hyperlink(table: &crate::Table<'_>) -> bool {
    table.rows().iter().any(|row| {
        row.cells().iter().any(|cell| {
            story_shapes_have_hyperlink(cell.shapes(), cell.shape_groups())
                || cell
                    .nested_tables()
                    .iter()
                    .any(|nested| table_has_hyperlink(&nested.table))
        })
    })
}

fn shape_has_hyperlink(shape: &crate::Shape<'_>) -> bool {
    shape
        .properties
        .iter()
        .any(|property| property.hyperlink.is_some())
        || shape.text_shapes.iter().any(shape_has_hyperlink)
        || shape.text_shape_groups.iter().any(group_has_hyperlink)
}

fn group_has_hyperlink(group: &crate::ShapeGroup<'_>) -> bool {
    group
        .properties
        .iter()
        .any(|property| property.hyperlink.is_some())
        || group.shapes.iter().any(shape_has_hyperlink)
        || group.groups.iter().any(group_has_hyperlink)
}

fn push_unique(
    unsupported: &mut Vec<UnsupportedReference>,
    value: UnsupportedReference,
) -> Result<(), Error> {
    if unsupported.contains(&value) {
        return Ok(());
    }
    if unsupported.len() >= MAX_DIAGNOSTICS {
        return Err(Error::Rtf(RtfError::LimitExceeded {
            resource: "external-reference diagnostics",
            observed: unsupported.len().saturating_add(1),
            limit: MAX_DIAGNOSTICS,
        }));
    }
    unsupported
        .try_reserve(1)
        .map_err(|_error| allocation("external-reference diagnostics", 1))?;
    unsupported.push(value);
    unsupported.sort_unstable();
    Ok(())
}

fn selected_entries(entries: &[ReferenceOccurrence]) -> Result<Vec<&ReferenceOccurrence>, Error> {
    let count = entries
        .iter()
        .filter(|entry| entry.source_range.is_some())
        .count();
    let mut selected = Vec::new();
    selected
        .try_reserve_exact(count)
        .map_err(|_error| allocation("external-reference selection", count))?;
    selected.extend(entries.iter().filter(|entry| entry.source_range.is_some()));
    selected.sort_unstable_by_key(|entry| entry.source_range.as_ref().map(|range| range.start));
    Ok(selected)
}

fn rewrite_candidate(
    source: &Document,
    selected: &[&ReferenceOccurrence],
    mode: Mode,
    unsupported: &[UnsupportedReference],
) -> Result<(Document, Diagnostics), Error> {
    if matches!(mode, Mode::Strict) && !unsupported.is_empty() {
        return Err(Error::UnsupportedSource(
            "external-reference redaction has unsupported remaining surfaces",
        ));
    }
    let source_bytes = source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("RTF source bytes are unavailable"))?;
    let diagnostics = Diagnostics {
        removed_references: selected.len(),
        unsupported: clone_unsupported(unsupported)?,
    };
    if selected.is_empty() {
        return Ok((source.clone(), diagnostics));
    }
    let output = splice_groups(source_bytes, selected)?;
    if output == source_bytes {
        return Ok((source.clone(), diagnostics));
    }
    let candidate =
        Document::from_bytes_with_limits(&output, source.limits()).map_err(Error::Rtf)?;
    if candidate.text() != source.text() {
        return Err(Error::UnsupportedSource(
            "external-reference redaction changed visible text",
        ));
    }
    if candidate
        .source_bytes()
        .is_none_or(|candidate_bytes| candidate_bytes != output.as_slice())
    {
        return Err(Error::UnsupportedSource(
            "external-reference redaction candidate did not retain exact bytes",
        ));
    }
    for entry in selected {
        let present = match entry.kind {
            ReferenceKind::NextFile => candidate.external_references().next_file.is_some(),
            ReferenceKind::Template => candidate.external_references().template.is_some(),
        };
        if present {
            return Err(Error::UnsupportedSource(
                "external-reference redaction candidate retained a selected group",
            ));
        }
    }
    if matches!(mode, Mode::Strict) && !candidate.external_references().is_empty() {
        return Err(Error::UnsupportedSource(
            "external-reference redaction candidate retained a modeled reference",
        ));
    }
    Ok((candidate, diagnostics))
}

fn clone_unsupported(
    unsupported: &[UnsupportedReference],
) -> Result<Box<[UnsupportedReference]>, Error> {
    let mut cloned = Vec::new();
    cloned
        .try_reserve_exact(unsupported.len())
        .map_err(|_error| allocation("external-reference diagnostics", unsupported.len()))?;
    cloned.extend_from_slice(unsupported);
    Ok(cloned.into_boxed_slice())
}

fn splice_groups(source: &[u8], selected: &[&ReferenceOccurrence]) -> Result<Vec<u8>, Error> {
    let removed = selected.iter().try_fold(0usize, |total, entry| {
        let range = entry.source_range.as_ref().ok_or(Error::UnsupportedSource(
            "external-reference source span is unavailable",
        ))?;
        let length = range
            .end
            .checked_sub(range.start)
            .ok_or(Error::UnsupportedSource(
                "external-reference source span is reversed",
            ))?;
        total.checked_add(length).ok_or(Error::DurablePatch(
            "external-reference source span size overflow".to_string(),
        ))
    })?;
    let retained = source
        .len()
        .checked_sub(removed)
        .ok_or(Error::UnsupportedSource(
            "external-reference source span exceeds source",
        ))?;
    // Several adjacent destination groups can be removed together.  Inspect
    // the retained suffix after the whole contiguous run; checking each group
    // independently misses `\\ansi{...}{...}Body` -> `\\ansiBody`.
    let delimiter = needs_control_delimiter_before_suffix(source, selected);
    let delimiters = usize::from(delimiter);
    let output_len = retained.checked_add(delimiters).ok_or(Error::DurablePatch(
        "external-reference candidate size overflow".to_string(),
    ))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_error| allocation("external-reference candidate bytes", output_len))?;
    let mut cursor = 0usize;
    for entry in selected {
        let range = entry.source_range.as_ref().ok_or(Error::UnsupportedSource(
            "external-reference source span is unavailable",
        ))?;
        if range.start < cursor || range.end > source.len() || range.start >= range.end {
            return Err(Error::UnsupportedSource(
                "external-reference source spans overlap or fall outside the source",
            ));
        }
        output.extend_from_slice(source.get(cursor..range.start).ok_or(
            Error::UnsupportedSource("external-reference source prefix is unavailable"),
        )?);
        cursor = range.end;
    }
    if delimiter {
        // Removing a destination immediately after a control word can
        // otherwise merge the following body text into that control word
        // (`\\ansi{...}Body` -> `\\ansiBody`).  The inserted delimiter is
        // syntax-only; every retained source byte and visible scalar is
        // still preserved.
        output.push(b' ');
    }
    output.extend_from_slice(source.get(cursor..).ok_or(Error::UnsupportedSource(
        "external-reference source suffix is unavailable",
    ))?);
    Ok(output)
}

fn needs_control_delimiter(source: &[u8], start: usize, end: usize) -> bool {
    let Some(next) = source.get(end).copied() else {
        return false;
    };
    if !next.is_ascii_alphanumeric() {
        return false;
    }
    let mut cursor = start;
    while cursor > 0
        && source
            .get(cursor - 1)
            .is_some_and(u8::is_ascii_alphanumeric)
    {
        cursor -= 1;
    }
    cursor > 0 && source.get(cursor - 1) == Some(&b'\\')
}

fn needs_control_delimiter_before_suffix(source: &[u8], selected: &[&ReferenceOccurrence]) -> bool {
    let Some(last) = selected
        .last()
        .and_then(|entry| entry.source_range.as_ref())
    else {
        return false;
    };
    let mut run_start = last.start;
    while let Some(previous) = selected.iter().find_map(|entry| {
        entry
            .source_range
            .as_ref()
            .filter(|range| range.end == run_start)
    }) {
        run_start = previous.start;
    }
    needs_control_delimiter(source, run_start, last.end)
}

fn build_patch(
    source: &Document,
    selected: &[&ReferenceOccurrence],
    mode: Mode,
    limits: PatchLimits,
) -> Result<Patch<ForwardOnly>, Error> {
    let source_bytes = source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("RTF source bytes are unavailable"))?;
    let artifact = BlobId::of(source_bytes).as_hex();
    let mut operations = Vec::new();
    operations
        .try_reserve_exact(selected.len())
        .map_err(|_error| allocation("external-reference patch operations", selected.len()))?;
    for entry in selected {
        let source_range = entry.source_range.as_ref().ok_or(Error::UnsupportedSource(
            "external-reference source span is unavailable",
        ))?;
        let mut preconditions = BTreeMap::new();
        // `BTreeMap` has no fallible reserve API on the supported Rust
        // versions.  This map is deliberately bounded to four fixed keys per
        // operation (and at most two operations); all variable-sized strings
        // below use fallible copies before entering the core patch envelope.
        preconditions.insert(
            owned_string("artifact_sha256", "external-reference patch key")?,
            Value::String(owned_string(
                &artifact,
                "external-reference artifact precondition",
            )?),
        );
        preconditions.insert(
            owned_string("value", "external-reference patch key")?,
            Value::String(owned_string(
                &entry.value,
                "external-reference value precondition",
            )?),
        );
        preconditions.insert(
            owned_string("mode", "external-reference patch key")?,
            Value::String(owned_string(
                match mode {
                    Mode::Strict => "strict",
                    Mode::BestEffort => "best-effort",
                },
                "external-reference redaction mode",
            )?),
        );
        preconditions.insert(
            owned_string(SOURCE_SPAN_PRECONDITION, "external-reference patch key")?,
            Value::String(source_span_text(source_range)?),
        );
        let operation = PatchOperation::new(
            limits,
            owned_string(REDACTION_OPERATION, "external-reference operation name")?,
            owned_string(entry.kind.target(), "external-reference operation target")?,
            preconditions,
            Value::Null,
        )
        .map_err(|error| Error::DurablePatch(error.to_string()))?;
        operations.push(operation);
    }
    let patch = Patch::<ForwardOnly>::new(
        limits,
        REDACTION_FORMAT,
        operations,
        BlobBundle::new(limits.blobs()),
    )
    .map_err(|error| Error::DurablePatch(error.to_string()))?;
    patch
        .to_deterministic_json()
        .map_err(|error| Error::DurablePatch(error.to_string()))?;
    Ok(patch)
}

/// Apply a previously sealed forward-only external-reference patch.
pub(crate) fn apply_forward(
    source: &Document,
    patch: &Patch<ForwardOnly>,
) -> Result<Document, Error> {
    if patch.format() != REDACTION_FORMAT {
        return Err(Error::DurablePatch(
            "external-reference patch format is not litchi-rtf".to_string(),
        ));
    }
    if !patch.blobs().is_empty() {
        return Err(Error::DurablePatch(
            "external-reference redaction patches cannot carry blobs".to_string(),
        ));
    }
    if patch.operations().len() > MAX_REFERENCES {
        return Err(Error::Rtf(RtfError::LimitExceeded {
            resource: "external-reference patch operations",
            observed: patch.operations().len(),
            limit: MAX_REFERENCES,
        }));
    }
    let snapshot = Snapshot::from_document(source)?;
    if patch.operations().is_empty() {
        return Ok(source.clone());
    }
    let source_bytes = source
        .source_bytes()
        .ok_or(Error::UnsupportedSource("RTF source bytes are unavailable"))?;
    let artifact = BlobId::of(source_bytes).as_hex();
    let mut selected = Vec::new();
    selected
        .try_reserve_exact(patch.operations().len())
        .map_err(|_error| {
            allocation(
                "external-reference patch selection",
                patch.operations().len(),
            )
        })?;
    let mut selected_next_file = false;
    let mut selected_template = false;
    let mut mode = None;
    for operation in patch.operations() {
        if operation.op != REDACTION_OPERATION || operation.value != Value::Null {
            return Err(Error::DurablePatch(
                "unsupported external-reference redaction operation".to_string(),
            ));
        }
        let kind = match operation.target.as_str() {
            "document:external-reference:next-file" => ReferenceKind::NextFile,
            "document:external-reference:template" => ReferenceKind::Template,
            _ => {
                return Err(Error::DurablePatch(
                    "unknown external-reference redaction target".to_string(),
                ));
            },
        };
        let selected_kind = match kind {
            ReferenceKind::NextFile => &mut selected_next_file,
            ReferenceKind::Template => &mut selected_template,
        };
        if *selected_kind {
            return Err(Error::DurablePatch(
                "duplicate external-reference redaction target".to_string(),
            ));
        }
        *selected_kind = true;
        let expected_artifact = operation
            .preconditions
            .get("artifact_sha256")
            .and_then(Value::as_str)
            .ok_or_else(|| {
                Error::DurablePatch("missing external-reference artifact precondition".to_string())
            })?;
        if expected_artifact != artifact {
            return Err(Error::PatchConflict);
        }
        let expected_mode = operation
            .preconditions
            .get("mode")
            .and_then(Value::as_str)
            .ok_or_else(|| {
                Error::DurablePatch("missing external-reference redaction mode".to_string())
            })?;
        let operation_mode = match expected_mode {
            "strict" => Mode::Strict,
            "best-effort" => Mode::BestEffort,
            _ => {
                return Err(Error::DurablePatch(
                    "invalid external-reference redaction mode".to_string(),
                ));
            },
        };
        if let Some(existing) = mode
            && existing != operation_mode
        {
            return Err(Error::DurablePatch(
                "external-reference redaction operations disagree on mode".to_string(),
            ));
        }
        mode = Some(operation_mode);
        let expected_value = operation
            .preconditions
            .get("value")
            .and_then(Value::as_str)
            .ok_or_else(|| {
                Error::DurablePatch("missing external-reference value precondition".to_string())
            })?;
        let expected_span = operation
            .preconditions
            .get(SOURCE_SPAN_PRECONDITION)
            .and_then(Value::as_str)
            .ok_or_else(|| {
                Error::DurablePatch(
                    "missing external-reference source span precondition".to_string(),
                )
            })?;
        let expected_range = parse_source_span(expected_span)?;
        let entry = snapshot
            .references
            .iter()
            .find(|entry| entry.kind == kind)
            .ok_or(Error::PatchConflict)?;
        if entry.value != expected_value
            || entry
                .source_range
                .as_ref()
                .is_none_or(|range| range != &expected_range)
        {
            return Err(Error::PatchConflict);
        }
        selected.push(entry);
    }
    selected.sort_unstable_by_key(|entry| entry.source_range.as_ref().map(|range| range.start));
    let mode = mode.unwrap_or(Mode::BestEffort);
    if matches!(mode, Mode::Strict) && !snapshot.unsupported.is_empty() {
        return Err(Error::UnsupportedSource(
            "external-reference redaction has unsupported remaining surfaces",
        ));
    }
    let (candidate, _diagnostics) =
        rewrite_candidate(source, &selected, mode, &snapshot.unsupported)?;
    Ok(candidate)
}

fn allocation(resource: &'static str, requested: usize) -> Error {
    Error::Rtf(RtfError::AllocationFailed {
        resource,
        requested,
    })
}

fn owned_string(value: &str, resource: &'static str) -> Result<String, Error> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_error| allocation(resource, value.len()))?;
    owned.push_str(value);
    Ok(owned)
}

fn source_span_text(range: &Range<usize>) -> Result<String, Error> {
    use std::fmt::Write as _;

    let capacity = (usize::BITS as usize * 2).saturating_add(1);
    let mut text = String::new();
    text.try_reserve_exact(capacity)
        .map_err(|_error| allocation("external-reference source span", capacity))?;
    write!(&mut text, "{}:{}", range.start, range.end)
        .map_err(|_error| Error::DurablePatch("failed to encode source span".to_string()))?;
    Ok(text)
}

fn parse_source_span(value: &str) -> Result<Range<usize>, Error> {
    let (start, end) = value
        .split_once(':')
        .ok_or_else(|| Error::DurablePatch("invalid external-reference source span".to_string()))?;
    let start = start.parse::<usize>().map_err(|_error| {
        Error::DurablePatch("invalid external-reference source span".to_string())
    })?;
    let end = end.parse::<usize>().map_err(|_error| {
        Error::DurablePatch("invalid external-reference source span".to_string())
    })?;
    if start >= end {
        return Err(Error::DurablePatch(
            "invalid external-reference source span".to_string(),
        ));
    }
    Ok(start..end)
}
