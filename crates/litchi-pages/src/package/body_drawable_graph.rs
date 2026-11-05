//! Private proof of the drawables anchored in the rooted Pages body, and the
//! publication of in-place message rewrites.
//!
//! The families that edit body images, movies, audio and shapes share this
//! module: it resolves the body storage from the document root, walks its
//! attachment table in text-flow order, proves every drawable is owned by the
//! body, and publishes a candidate in which only the rewritten messages
//! changed. Native identifiers never leave this module except as private
//! targets the families keep inside their `Edit` and `Patch` values.

use std::collections::HashMap;
use std::mem::size_of;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_iwa_archive::package::{Entry, EntryEdit};
use litchi_iwa_common::wire::{WireFieldView, WireView};
use litchi_iwa_common::{WireLimits, decode_varint_from_bytes, varint::encoded_len};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{pages_body_codec, pages_media_codec};

use super::Package;
use crate::selector::DrawableSelector;
use crate::{Error as SemanticError, PackageError};

pub(crate) const DOCUMENT_COMPONENT: &str = "Index/Document.iwa";
pub(crate) const ROOT_OBJECT_IDENTIFIER: u64 = 1;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const BODY_TEXT_FIELD: u32 = 3;
const BODY_ATTACHMENTS_FIELD: u32 = 9;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_CHARACTER_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
pub(crate) const DRAWABLE_ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const ATTACHMENT_DRAWABLE_FIELD: u32 = 1;

/// A private failure of the shared body-drawable proof; each family maps it
/// to its own typed error.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum GraphError {
    /// The rooted body graph is malformed or ambiguous.
    InvalidSource,
    /// A resource ceiling was exceeded.
    Limit {
        kind: GraphLimit,
        observed: u64,
        maximum: u64,
    },
    /// A bounded allocation failed.
    Allocation { amount: usize },
}

/// The resource a [`GraphError::Limit`] refers to.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum GraphLimit {
    InputBytes,
    OutputBytes,
    WireFields,
    WireNesting,
    WireWork,
    /// Semantic section or body-storage cardinality.
    SemanticItems,
    /// Semantic UTF-8 text retained by the reopened document.
    SemanticTextBytes,
}

pub(crate) type GraphResult<T> = Result<T, GraphError>;

/// Per-operation ledger over the package's physical limits.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Budget {
    limits: WireLimits,
    max_input: usize,
    max_work: usize,
    max_fields: usize,
    input: usize,
    work: usize,
    fields: usize,
}

impl Budget {
    pub(crate) fn new(package: &Package) -> GraphResult<Self> {
        let physical = package.state.source.limits();
        let archive = physical
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let message = archive.max_message_bytes().max(1);
        let limits = WireLimits::default()
            .with_input_bytes(message.min(WireLimits::MAX_INPUT_BYTES))
            .and_then(|value| value.with_output_bytes(message.min(WireLimits::MAX_OUTPUT_BYTES)))
            .and_then(|value| {
                value.with_rewrite_work(
                    message
                        .saturating_mul(16)
                        .clamp(1, WireLimits::MAX_REWRITE_WORK),
                )
            })
            .map_err(map_wire_error)?;
        Ok(Self {
            limits,
            max_input: physical.max_iwa_stream_bytes().saturating_mul(4).max(1),
            max_work: usize::try_from(physical.max_total_bytes())
                .unwrap_or(usize::MAX)
                .saturating_mul(32)
                .max(limits.max_rewrite_work()),
            max_fields: limits.max_fields(),
            input: 0,
            work: 0,
            fields: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: GraphLimit,
    ) -> GraphResult<()> {
        let observed = current
            .checked_add(amount)
            .ok_or(GraphError::InvalidSource)?;
        if observed > maximum {
            return Err(GraphError::Limit {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    /// Charge aggregate traversal work.
    pub(crate) fn work(&mut self, amount: usize) -> GraphResult<()> {
        Self::add(&mut self.work, amount, self.max_work, GraphLimit::WireWork)
    }

    /// Parse one payload and charge its bytes, fields and depth.
    pub(crate) fn parse<'a>(
        &mut self,
        payload: &'a [u8],
        depth: usize,
    ) -> GraphResult<WireView<'a>> {
        if depth > self.limits.max_nesting() {
            return Err(GraphError::Limit {
                kind: GraphLimit::WireNesting,
                observed: depth as u64,
                maximum: self.limits.max_nesting() as u64,
            });
        }
        let view = WireView::parse(payload, self.limits).map_err(map_wire_error)?;
        Self::add(
            &mut self.input,
            payload.len(),
            self.max_input,
            GraphLimit::InputBytes,
        )?;
        Self::add(
            &mut self.fields,
            view.len(),
            self.max_fields,
            GraphLimit::WireFields,
        )?;
        self.work(payload.len().saturating_add(view.len()))?;
        Ok(view)
    }

    /// The wire limits a codec call may use for one payload.
    pub(crate) const fn wire_limits(&self) -> WireLimits {
        self.limits
    }
}

/// One object found in the package's components.
#[derive(Clone, Copy)]
pub(crate) struct Located<'a> {
    pub(crate) component_index: usize,
    pub(crate) object: &'a ArchiveObject,
}

/// Identifier → object index over every component, built once per operation.
pub(crate) struct ObjectIndex<'a> {
    package: &'a Package,
    objects: HashMap<u64, (usize, usize)>,
}

impl<'a> ObjectIndex<'a> {
    pub(crate) fn new(package: &'a Package, budget: &mut Budget) -> GraphResult<Self> {
        budget.work(package.state.object_count)?;
        let mut objects = HashMap::new();
        objects
            .try_reserve(package.state.object_count)
            .map_err(|_| GraphError::Allocation {
                amount: package.state.object_count,
            })?;
        for (component_index, component) in package.state.source.components().iter().enumerate() {
            for (object_index, object) in component.archive().objects.iter().enumerate() {
                let Some(identifier) = object.archive_info.identifier else {
                    continue;
                };
                if objects
                    .insert(identifier, (component_index, object_index))
                    .is_some()
                {
                    return Err(GraphError::InvalidSource);
                }
            }
        }
        Ok(Self { package, objects })
    }

    pub(crate) fn get(&self, identifier: NonZeroU64) -> GraphResult<Located<'a>> {
        let &(component_index, object_index) = self
            .objects
            .get(&identifier.get())
            .ok_or(GraphError::InvalidSource)?;
        let component = self
            .package
            .state
            .source
            .components()
            .get_index(component_index)
            .ok_or(GraphError::InvalidSource)?;
        let object = component
            .archive()
            .objects
            .get(object_index)
            .ok_or(GraphError::InvalidSource)?;
        Ok(Located {
            component_index,
            object,
        })
    }

    pub(crate) fn component_name(&self, component_index: usize) -> GraphResult<Arc<str>> {
        self.package
            .state
            .source
            .components()
            .get_index(component_index)
            .map(|component| Arc::from(component.name()))
            .ok_or(GraphError::InvalidSource)
    }
}

/// Facts read from the document root.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Root {
    pub(crate) body: NonZeroU64,
    pub(crate) left_margin: Option<f32>,
}

/// One drawable attachment of the body, in text-flow order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Anchor {
    /// The UTF-16 index of the anchoring U+FFFC in the body text.
    pub(crate) character_index: u32,
    pub(crate) attachment: NonZeroU64,
    pub(crate) attachment_message: usize,
    pub(crate) drawable: NonZeroU64,
}

/// The body's drawable anchors after the ownership proof.
pub(crate) struct Body {
    pub(crate) root: Root,
    pub(crate) anchors: Vec<Anchor>,
}

/// Resolve the rooted body and its drawable attachments.
///
/// Entries whose attachment is not a `DrawableAttachmentArchive` (inline
/// text, footnotes, number attachments) are skipped; every listed entry must
/// still sit on an object-replacement character.
pub(crate) fn resolve_body(index: &ObjectIndex<'_>, budget: &mut Budget) -> GraphResult<Body> {
    let package = index.package;
    let document = package
        .state
        .source
        .components()
        .get(DOCUMENT_COMPONENT)
        .ok_or(GraphError::InvalidSource)?;
    let root = document
        .archive()
        .object(ROOT_OBJECT_IDENTIFIER)
        .ok_or(GraphError::InvalidSource)?;
    let (_, root_message) = unique_message(root, ROOT_MESSAGE_TYPE)?;
    let root_payload = root_message.data.as_slice();
    let options = body_options(budget, root_payload)?;
    budget.work(root_payload.len())?;
    let body_facts = pages_body_codec::decode_document_body(root_payload, options)
        .map_err(|_| GraphError::InvalidSource)?;
    let root_facts = pages_body_codec::decode_document_root(root_payload, options)
        .map_err(|_| GraphError::InvalidSource)?;
    let body = body_facts
        .body_storage()
        .map(|reference| reference.identifier())
        .ok_or(GraphError::InvalidSource)?;
    let root = Root {
        body,
        left_margin: root_facts.left_margin(),
    };
    let storage = index.get(body)?;
    let payload = unique_text_payload(storage.object)?;
    let view = budget.parse(payload, 1)?;
    let Some(table) = unique_field(&view, BODY_ATTACHMENTS_FIELD, 2)? else {
        return Ok(Body {
            root,
            anchors: Vec::new(),
        });
    };
    let table_view = budget.parse(table.payload(), 2)?;
    let mut entries = Vec::new();
    entries
        .try_reserve(table_view.len())
        .map_err(|_| GraphError::Allocation {
            amount: table_view.len(),
        })?;
    for field in table_view
        .fields()
        .filter(|field| field.number() == TABLE_ENTRIES_FIELD)
    {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        entries.push(parse_entry(field.payload(), budget)?);
    }
    entries.sort_unstable_by_key(|entry: &(u32, NonZeroU64)| entry.0);
    if entries.windows(2).any(|pair| pair[0].0 == pair[1].0) {
        return Err(GraphError::InvalidSource);
    }
    validate_anchors(payload, &entries, budget)?;
    let mut anchors = Vec::new();
    anchors
        .try_reserve(entries.len())
        .map_err(|_| GraphError::Allocation {
            amount: entries.len(),
        })?;
    let mut drawables = std::collections::HashSet::new();
    for (character_index, attachment) in entries {
        let located = index.get(attachment)?;
        let Some((attachment_message, message)) =
            unique_optional_message(located.object, DRAWABLE_ATTACHMENT_MESSAGE_TYPE)?
        else {
            continue;
        };
        validate_message_metadata(located.object, attachment_message)?;
        let attachment_view = budget.parse(message.data.as_slice(), 2)?;
        let Some(field) = unique_field(&attachment_view, ATTACHMENT_DRAWABLE_FIELD, 2)? else {
            continue;
        };
        let drawable = parse_reference(field.payload(), budget)?;
        if drawable.get() == ROOT_OBJECT_IDENTIFIER
            || drawable == body
            || drawable == attachment
            || !drawables.insert(drawable)
        {
            return Err(GraphError::InvalidSource);
        }
        anchors.push(Anchor {
            character_index,
            attachment,
            attachment_message,
            drawable,
        });
    }
    Ok(Body { root, anchors })
}

/// The document root's drawable z-order object.
pub(crate) fn root_zorder(
    package: &Package,
    budget: &mut Budget,
) -> GraphResult<Option<NonZeroU64>> {
    let document = package
        .state
        .source
        .components()
        .get(DOCUMENT_COMPONENT)
        .ok_or(GraphError::InvalidSource)?;
    let root = document
        .archive()
        .object(ROOT_OBJECT_IDENTIFIER)
        .ok_or(GraphError::InvalidSource)?;
    let (_, root_message) = unique_message(root, ROOT_MESSAGE_TYPE)?;
    let payload = root_message.data.as_slice();
    let options = body_options(budget, payload)?;
    budget.work(payload.len())?;
    let facts = pages_body_codec::decode_document_root(payload, options)
        .map_err(|_| GraphError::InvalidSource)?;
    Ok(facts
        .drawables_zorder()
        .map(|reference| reference.identifier()))
}

/// The document root's theme and document language.
pub(crate) fn root_theme(
    package: &Package,
    budget: &mut Budget,
) -> GraphResult<(Option<NonZeroU64>, Option<String>)> {
    let document = package
        .state
        .source
        .components()
        .get(DOCUMENT_COMPONENT)
        .ok_or(GraphError::InvalidSource)?;
    let root = document
        .archive()
        .object(ROOT_OBJECT_IDENTIFIER)
        .ok_or(GraphError::InvalidSource)?;
    let (_, root_message) = unique_message(root, ROOT_MESSAGE_TYPE)?;
    let payload = root_message.data.as_slice();
    let options = body_options(budget, payload)?;
    budget.work(payload.len())?;
    let facts = pages_body_codec::decode_document_root(payload, options)
        .map_err(|_| GraphError::InvalidSource)?;
    Ok((
        facts.theme().map(|reference| reference.identifier()),
        facts.document_language().map(str::to_owned),
    ))
}

fn body_options(budget: &Budget, source: &[u8]) -> GraphResult<pages_body_codec::DecodeOptions> {
    let limits = budget.wire_limits();
    let recursion = u32::try_from(limits.max_nesting()).map_err(|_| GraphError::InvalidSource)?;
    Ok(pages_body_codec::DecodeOptions::new(
        limits.max_input_bytes().min(source.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    ))
}

fn parse_entry(source: &[u8], budget: &mut Budget) -> GraphResult<(u32, NonZeroU64)> {
    let view = budget.parse(source, 3)?;
    let character =
        unique_field(&view, ENTRY_CHARACTER_INDEX_FIELD, 0)?.ok_or(GraphError::InvalidSource)?;
    let character = canonical_varint(character.payload())?;
    let character = u32::try_from(character).map_err(|_| GraphError::InvalidSource)?;
    let object = unique_field(&view, ENTRY_OBJECT_FIELD, 2)?.ok_or(GraphError::InvalidSource)?;
    Ok((character, parse_reference(object.payload(), budget)?))
}

fn validate_anchors(
    payload: &[u8],
    entries: &[(u32, NonZeroU64)],
    budget: &mut Budget,
) -> GraphResult<()> {
    let view = budget.parse(payload, 1)?;
    let mut next = 0usize;
    let mut utf16 = 0usize;
    for field in view
        .fields()
        .filter(|field| field.number() == BODY_TEXT_FIELD)
    {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let text = std::str::from_utf8(field.payload()).map_err(|_| GraphError::InvalidSource)?;
        for character in text.chars() {
            if entries
                .get(next)
                .is_some_and(|entry| entry.0 as usize == utf16)
            {
                if character != '\u{fffc}' {
                    return Err(GraphError::InvalidSource);
                }
                next += 1;
            }
            utf16 = utf16
                .checked_add(character.len_utf16())
                .ok_or(GraphError::InvalidSource)?;
        }
    }
    if next != entries.len() {
        return Err(GraphError::InvalidSource);
    }
    budget.work(utf16)
}

/// Decode one canonical varint payload.
pub(crate) fn canonical_varint(payload: &[u8]) -> GraphResult<u64> {
    let (value, width) =
        decode_varint_from_bytes(payload).map_err(|_| GraphError::InvalidSource)?;
    if width != encoded_len(value) || width != payload.len() {
        return Err(GraphError::InvalidSource);
    }
    Ok(value)
}

/// Parse a `TSP.Reference` payload into its non-zero identifier.
pub(crate) fn parse_reference(source: &[u8], budget: &mut Budget) -> GraphResult<NonZeroU64> {
    let view = budget.parse(source, 4)?;
    let field = unique_field(&view, 1, 0)?.ok_or(GraphError::InvalidSource)?;
    NonZeroU64::new(canonical_varint(field.payload())?).ok_or(GraphError::InvalidSource)
}

/// The single field `number` of `view`, which must use `wire_type`.
pub(crate) fn unique_field<'a>(
    view: &WireView<'a>,
    number: u32,
    wire_type: u8,
) -> GraphResult<Option<WireFieldView<'a>>> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != wire_type {
            return Err(GraphError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if found.replace(field).is_some() {
            return Err(GraphError::InvalidSource);
        }
    }
    Ok(found)
}

/// The single message of `message_type` in `object`, if any, after checking
/// the object's message table.
pub(crate) fn unique_optional_message(
    object: &ArchiveObject,
    message_type: u32,
) -> GraphResult<Option<(usize, &RawMessage)>> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(GraphError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        let info = &object.archive_info.message_infos[index];
        if info.type_ != message.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(GraphError::InvalidSource);
        }
        if message.type_ == message_type && selected.replace((index, message)).is_some() {
            return Err(GraphError::InvalidSource);
        }
    }
    Ok(selected)
}

/// The single message of `message_type` in `object`.
pub(crate) fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> GraphResult<(usize, &RawMessage)> {
    unique_optional_message(object, message_type)?.ok_or(GraphError::InvalidSource)
}

fn unique_text_payload(object: &ArchiveObject) -> GraphResult<&[u8]> {
    let mut selected = None;
    for message in &object.messages {
        if matches!(message.type_, 2_001 | 2_022)
            && selected.replace(message.data.as_slice()).is_some()
        {
            return Err(GraphError::InvalidSource);
        }
    }
    selected.ok_or(GraphError::InvalidSource)
}

/// Refuse merge/diff message metadata the in-place rewrite cannot preserve.
pub(crate) fn validate_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> GraphResult<()> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(GraphError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(GraphError::InvalidSource);
    }
    Ok(())
}

/// One message replacement inside one component.
pub(crate) struct MessageRewrite {
    pub(crate) component: Arc<str>,
    pub(crate) object: NonZeroU64,
    pub(crate) message_index: usize,
    pub(crate) message_type: u32,
    pub(crate) data: Vec<u8>,
}

/// Publish a candidate in which only the given messages changed, and reopen
/// it under the source's limits.
pub(crate) fn publish(source: &Package, rewrites: Vec<MessageRewrite>) -> GraphResult<Package> {
    publish_with_replacements(source, rewrites, &[])
}

/// [`publish`], also replacing the given non-IWA members (name, bytes).
pub(crate) fn publish_with_replacements(
    source: &Package,
    rewrites: Vec<MessageRewrite>,
    replacements: &[(&str, &[u8])],
) -> GraphResult<Package> {
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();
    for rewrite in rewrites {
        let position = match archives
            .iter()
            .position(|(name, _)| name.as_ref() == rewrite.component.as_ref())
        {
            Some(position) => position,
            None => {
                archives.push((
                    Arc::clone(&rewrite.component),
                    editable_archive(source, &rewrite.component)?,
                ));
                archives.len() - 1
            },
        };
        archives[position]
            .1
            .object_mut(rewrite.object.get())
            .ok_or(GraphError::InvalidSource)?
            .replace_message_preserving_header(
                rewrite.message_index,
                RawMessage {
                    type_: rewrite.message_type,
                    data: rewrite.data,
                },
                archive_limits,
            )
            .map_err(map_core_error)?;
    }
    publish_changes(source, archives, &[], replacements, &[])
}

/// A clone of one component's archive for staging.
pub(crate) fn editable_archive(source: &Package, component: &str) -> GraphResult<Archive> {
    source
        .state
        .source
        .components()
        .get(component)
        .map(|component| component.archive().clone())
        .ok_or(GraphError::InvalidSource)
}

/// Publish a candidate in which only the given component archives changed,
/// and reopen it under the source's limits.
pub(crate) fn publish_archives(
    source: &Package,
    archives: Vec<(Arc<str>, Archive)>,
) -> GraphResult<Package> {
    publish_changes(source, archives, &[], &[], &[])
}

/// Publish a candidate in which the given component archives changed, the
/// given members were inserted (name, bytes) and the named members deleted,
/// and reopen it under the source's limits.
pub(crate) fn publish_changes(
    source: &Package,
    archives: Vec<(Arc<str>, Archive)>,
    insertions: &[(&str, &[u8])],
    replacements: &[(&str, &[u8])],
    deletions: &[&str],
) -> GraphResult<Package> {
    let limits = source.state.source.limits();
    let archive_limits = limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = limits.snappy_limits().map_err(map_archive_error)?;
    let mut payloads = Vec::with_capacity(archives.len());
    for (name, archive) in archives {
        let bytes = archive.to_bytes(archive_limits).map_err(map_core_error)?;
        let compressed = SnappyStream::compress(&bytes, snappy_limits).map_err(map_core_error)?;
        if compressed.len() as u64 > limits.max_entry_bytes() {
            return Err(GraphError::Limit {
                kind: GraphLimit::OutputBytes,
                observed: compressed.len() as u64,
                maximum: limits.max_entry_bytes(),
            });
        }
        payloads.push((name, compressed));
    }
    let edits = payloads
        .iter()
        .map(|(name, data)| EntryEdit::new(name.as_ref(), data))
        .chain(
            replacements
                .iter()
                .map(|(name, data)| EntryEdit::new(name, data)),
        )
        .collect::<Vec<_>>();
    let output = source
        .state
        .source
        .package()
        .reassemble_with_changes_to_bytes(
            &insertions
                .iter()
                .map(|(name, data)| litchi_iwa_archive::package::EntryInsertion::new(name, data))
                .collect::<Vec<_>>(),
            &edits,
            deletions,
            limits,
        )
        .map_err(map_archive_error)?;
    if output.len() as u64 > limits.max_input_bytes() {
        return Err(GraphError::Limit {
            kind: GraphLimit::OutputBytes,
            observed: output.len() as u64,
            maximum: limits.max_input_bytes(),
        });
    }
    let catalog = source
        .state
        .source
        .reopen(output.into())
        .map_err(map_archive_error)?;
    Package::from_source_catalog_with_semantic_limits(catalog, source.state.semantic_limits)
        .map_err(map_package_error)
}

/// Reopen the target artifact of an exact patch under the source's limits.
pub(crate) fn reopen(source: &Package, target: Arc<[u8]>) -> GraphResult<Package> {
    let catalog = source
        .state
        .source
        .reopen(target)
        .map_err(map_archive_error)?;
    Package::from_source_catalog_with_semantic_limits(catalog, source.state.semantic_limits)
        .map_err(map_package_error)
}

/// Prove that an exact-source candidate changed only one archive message.
///
/// The candidate's semantic readback is necessary but not sufficient: this
/// witness also compares every ZIP member, every untouched archive object,
/// and every untouched message in the selected object against the source.
/// The selected ZIP record is allowed to change only in CRC/size fields.
pub(crate) fn verify_message_locality(
    source: &Package,
    candidate: &Package,
    component_name: &str,
    object_index: usize,
    object_identifier: NonZeroU64,
    message_index: usize,
    message_type: u32,
    expected_payload: &[u8],
    budget: &mut Budget,
) -> GraphResult<()> {
    let source_entries = source.state.source.package();
    let candidate_entries = candidate.state.source.package();
    if source_entries.len() != candidate_entries.len() {
        return Err(GraphError::InvalidSource);
    }
    for (before, after) in source_entries.iter().zip(candidate_entries.iter()) {
        budget.work(zip_entry_work(before)?.saturating_add(zip_entry_work(after)?))?;
        if before.name() != after.name()
            || before.raw_name() != after.raw_name()
            || before.is_opaque() != after.is_opaque()
            || before.metadata().local() != after.metadata().local()
            || before.metadata().central() != after.metadata().central()
        {
            return Err(GraphError::InvalidSource);
        }
        if before.name() == component_name {
            if !zip_local_header_preserved(
                before.raw_record().local_record(),
                after.raw_record().local_record(),
            ) || !zip_record_equal_outside(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
                &[(16, 28), (42, 46)],
            ) {
                return Err(GraphError::InvalidSource);
            }
        } else if before.data() != after.data()
            || before.raw_record().local_record() != after.raw_record().local_record()
            || !zip_record_equal_outside(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
                &[(42, 46)],
            )
            || before.raw_record().compressed_data() != after.raw_record().compressed_data()
        {
            return Err(GraphError::InvalidSource);
        }
    }

    let source_components = source.state.source.components();
    let candidate_components = candidate.state.source.components();
    if source_components.len() != candidate_components.len() {
        return Err(GraphError::InvalidSource);
    }
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut selected_seen = false;
    for (before_component, after_component) in
        source_components.iter().zip(candidate_components.iter())
    {
        if before_component.name() != after_component.name()
            || before_component.archive().objects.len() != after_component.archive().objects.len()
        {
            return Err(GraphError::InvalidSource);
        }
        for (index, (before, after)) in before_component
            .archive()
            .objects
            .iter()
            .zip(&after_component.archive().objects)
            .enumerate()
        {
            budget
                .work(archive_object_work(before)?.saturating_add(archive_object_work(after)?))?;
            let selected = before_component.name() == component_name && index == object_index;
            if !selected {
                if !before.same_content_ignoring_offsets(after) {
                    return Err(GraphError::InvalidSource);
                }
                continue;
            }
            selected_seen = true;
            if before.archive_info.identifier != Some(object_identifier.get())
                || after.archive_info.identifier != Some(object_identifier.get())
            {
                return Err(GraphError::InvalidSource);
            }
            let message = after
                .messages
                .get(message_index)
                .filter(|message| message.type_ == message_type)
                .ok_or(GraphError::InvalidSource)?;
            if message.data.as_slice() != expected_payload {
                return Err(GraphError::InvalidSource);
            }
            let mut expected = before.clone();
            expected
                .replace_message_preserving_header(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data: message.data.clone(),
                    },
                    archive_limits,
                )
                .map_err(map_core_error)?;
            expected.header_length = after.header_length;
            expected.data_length = after.data_length;
            if !expected.same_content_ignoring_offsets(after) {
                return Err(GraphError::InvalidSource);
            }
        }
    }
    if selected_seen {
        Ok(())
    } else {
        Err(GraphError::InvalidSource)
    }
}

fn zip_entry_work(entry: &Entry) -> GraphResult<usize> {
    entry
        .name()
        .len()
        .checked_add(entry.raw_name().len())
        .and_then(|amount| amount.checked_add(entry.data().len()))
        .and_then(|amount| amount.checked_add(entry.raw_record().local_record().len()))
        .and_then(|amount| amount.checked_add(entry.raw_record().central_directory_record().len()))
        .and_then(|amount| amount.checked_add(entry.raw_record().compressed_data().len()))
        .ok_or(GraphError::InvalidSource)
}

fn archive_object_work(object: &ArchiveObject) -> GraphResult<usize> {
    let mut work = size_of::<ArchiveObject>()
        .checked_add(object.messages.len())
        .and_then(|amount| amount.checked_add(object.archive_info.message_infos.len()))
        .ok_or(GraphError::InvalidSource)?;
    for message in &object.messages {
        work = work
            .checked_add(size_of::<RawMessage>())
            .and_then(|amount| amount.checked_add(message.data.len()))
            .ok_or(GraphError::InvalidSource)?;
    }
    for info in &object.archive_info.message_infos {
        work = work
            .checked_add(size_of::<litchi_iwa_core::MessageInfo>())
            .and_then(|amount| amount.checked_add(info.versions.len()))
            .and_then(|amount| amount.checked_add(info.field_infos.len()))
            .and_then(|amount| amount.checked_add(info.object_references.len()))
            .and_then(|amount| amount.checked_add(info.data_references.len()))
            .and_then(|amount| amount.checked_add(info.diff_merge_version.len()))
            .and_then(|amount| amount.checked_add(info.fields_to_remove.len()))
            .and_then(|amount| amount.checked_add(info.diff_read_version.len()))
            .ok_or(GraphError::InvalidSource)?;
        if let Some(path) = info.diff_field_path.as_ref() {
            work = work
                .checked_add(path.path.len())
                .ok_or(GraphError::InvalidSource)?;
        }
        for path in &info.fields_to_remove {
            work = work
                .checked_add(path.path.len())
                .ok_or(GraphError::InvalidSource)?;
        }
        for field in &info.field_infos {
            work = work
                .checked_add(size_of::<litchi_iwa_core::FieldInfo>())
                .and_then(|amount| amount.checked_add(field.path.path.len()))
                .and_then(|amount| amount.checked_add(field.object_references.len()))
                .and_then(|amount| amount.checked_add(field.data_references.len()))
                .and_then(|amount| amount.checked_add(field.known_field_version.len()))
                .ok_or(GraphError::InvalidSource)?;
            if let Some(identifier) = field.known_field_feature_identifier.as_ref() {
                work = work
                    .checked_add(identifier.len())
                    .ok_or(GraphError::InvalidSource)?;
            }
        }
    }
    Ok(work)
}

fn zip_record_equal_outside(before: &[u8], after: &[u8], excluded: &[(usize, usize)]) -> bool {
    if before.len() != after.len() {
        return false;
    }
    let mut cursor = 0;
    for &(start, end) in excluded {
        if start < cursor
            || end < start
            || end > before.len()
            || before[cursor..start] != after[cursor..start]
        {
            return false;
        }
        cursor = end;
    }
    before[cursor..] == after[cursor..]
}

fn zip_local_header_preserved(before: &[u8], after: &[u8]) -> bool {
    let Some(before_header) = zip_local_header_len(before) else {
        return false;
    };
    let Some(after_header) = zip_local_header_len(after) else {
        return false;
    };
    before_header == after_header
        && zip_record_equal_outside(
            &before[..before_header],
            &after[..after_header],
            &[(14, 26)],
        )
}

fn zip_local_header_len(record: &[u8]) -> Option<usize> {
    if record.len() < 30 || record.get(..4) != Some(b"PK\x03\x04") {
        return None;
    }
    let name = usize::from(u16::from_le_bytes([record[26], record[27]]));
    let extra = usize::from(u16::from_le_bytes([record[28], record[29]]));
    30usize
        .checked_add(name)?
        .checked_add(extra)
        .filter(|length| *length <= record.len())
}

pub(crate) fn map_wire_error(error: litchi_iwa_common::Error) -> GraphError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => GraphError::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => GraphLimit::InputBytes,
                litchi_iwa_common::LimitKind::Fields => GraphLimit::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => GraphLimit::OutputBytes,
                litchi_iwa_common::LimitKind::Nesting => GraphLimit::WireNesting,
                _ => GraphLimit::WireWork,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => GraphError::Allocation { amount },
        _ => GraphError::InvalidSource,
    }
}

pub(crate) fn map_archive_error(error: litchi_iwa_archive::Error) -> GraphError {
    match error {
        litchi_iwa_archive::Error::Limit {
            observed,
            maximum,
            kind,
        } => GraphError::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::OutputBytes => GraphLimit::OutputBytes,
                _ => GraphLimit::InputBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => GraphError::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => GraphError::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> GraphError {
    match error {
        PackageError::Archive(error) => map_archive_error(error),
        PackageError::Limit {
            observed, maximum, ..
        } => GraphError::Limit {
            kind: GraphLimit::InputBytes,
            observed,
            maximum,
        },
        PackageError::PayloadLimit {
            observed,
            limit: maximum,
        }
        | PackageError::SectionNamesTooLarge {
            observed,
            limit: maximum,
        }
        | PackageError::ObjectLimit {
            observed,
            limit: maximum,
        } => GraphError::Limit {
            kind: GraphLimit::InputBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        PackageError::Semantic(SemanticError::TooManySections {
            actual: observed,
            limit: maximum,
        })
        | PackageError::Semantic(SemanticError::TooManyBodyStorages {
            actual: observed,
            limit: maximum,
        }) => GraphError::Limit {
            kind: GraphLimit::SemanticItems,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        PackageError::Semantic(SemanticError::TextTooLarge {
            observed,
            limit: maximum,
        }) => GraphError::Limit {
            kind: GraphLimit::SemanticTextBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        },
        PackageError::Allocation { amount } => GraphError::Allocation { amount },
        _ => GraphError::InvalidSource,
    }
}

pub(crate) fn map_core_error(error: litchi_iwa_core::Error) -> GraphError {
    match error {
        litchi_iwa_core::Error::Limit {
            observed,
            maximum,
            kind,
        } => GraphError::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes => GraphLimit::InputBytes,
                _ => GraphLimit::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            GraphError::Allocation { amount: requested }
        },
        _ => GraphError::InvalidSource,
    }
}

const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const WRAPPED_DRAWABLE_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const DRAWABLE_PENCIL_ANNOTATIONS_FIELD: u32 = 9;
const SHAPE_INFO_DEPRECATED_STORAGE_FIELD: u32 = 2;
const SHAPE_INFO_OWNED_STORAGE_FIELD: u32 = 4;
const SHAPE_INFO_IS_TEXT_BOX_FIELD: u32 = 6;

/// The native message type that carries a drawable of the selected kind.
pub(crate) const fn message_type(selector: DrawableSelector) -> u32 {
    match selector {
        DrawableSelector::Image(_) => IMAGE_MESSAGE_TYPE,
        DrawableSelector::Movie(_) | DrawableSelector::Audio(_) => MOVIE_MESSAGE_TYPE,
        DrawableSelector::Shape(_) | DrawableSelector::TextBox(_) => SHAPE_INFO_MESSAGE_TYPE,
    }
}

/// The selector's kind as a static word for refusals.
pub(crate) const fn kind_name(selector: DrawableSelector) -> &'static str {
    match selector {
        DrawableSelector::Image(_) => "image",
        DrawableSelector::Movie(_) => "movie",
        DrawableSelector::Audio(_) => "audio",
        DrawableSelector::Shape(_) => "shape",
        DrawableSelector::TextBox(_) => "text box",
    }
}

/// One body drawable selected by kind and text-flow position, after the
/// ownership proof (its parent is the body storage).
pub(crate) struct Selected<'a> {
    pub(crate) anchor: Anchor,
    pub(crate) drawable: Located<'a>,
    pub(crate) attachment: Located<'a>,
    pub(crate) message_index: usize,
    pub(crate) payload: &'a [u8],
    /// The `TSD.DrawableArchive` inside `payload`.
    pub(crate) drawable_payload: &'a [u8],
    /// A shape's owned text storage.
    pub(crate) storage: Option<NonZeroU64>,
}

/// Select the drawable of `selector`'s kind at its position among the body's
/// drawables of that kind, in text-flow order; `None` when there is none.
pub(crate) fn select_drawable<'a>(
    index: &ObjectIndex<'a>,
    body: &Body,
    selector: DrawableSelector,
    budget: &mut Budget,
) -> GraphResult<Option<Selected<'a>>> {
    let requested = selector.position().get();
    let message_type = message_type(selector);
    let mut seen = 0usize;
    for anchor in &body.anchors {
        let drawable = index.get(anchor.drawable)?;
        let Some((message_index, message)) =
            unique_optional_message(drawable.object, message_type)?
        else {
            continue;
        };
        let payload = message.data.as_slice();
        let mut storage = None;
        match selector {
            DrawableSelector::Movie(_) | DrawableSelector::Audio(_) => {
                let flags = pages_media_codec::decode_movie_media_flags(
                    payload,
                    pages_media_codec::DecodeOptions::for_source(payload),
                )
                .map_err(|_| GraphError::InvalidSource)?;
                if flags.is_live_video() == Some(true)
                    || (flags.audio_only() == Some(true))
                        != matches!(selector, DrawableSelector::Audio(_))
                {
                    continue;
                }
            },
            DrawableSelector::Shape(_) => {
                let Some(owned) = shape_storage(payload, false, budget)? else {
                    continue;
                };
                storage = Some(owned);
            },
            DrawableSelector::TextBox(_) => {
                let Some(owned) = shape_storage(payload, true, budget)? else {
                    continue;
                };
                storage = Some(owned);
            },
            DrawableSelector::Image(_) => {},
        }
        if seen != requested {
            seen += 1;
            continue;
        }
        validate_message_metadata(drawable.object, message_index)?;
        let drawable_payload = drawable_payload(selector, payload, budget)?;
        let view = budget.parse(drawable_payload, 3)?;
        let parent = unique_field(&view, DRAWABLE_PARENT_FIELD, 2)?
            .map(|field| parse_reference(field.payload(), budget))
            .transpose()?;
        if parent != Some(body.root.body) {
            return Err(GraphError::InvalidSource);
        }
        let attachment = index.get(anchor.attachment)?;
        return Ok(Some(Selected {
            anchor: *anchor,
            drawable,
            attachment,
            message_index,
            payload,
            drawable_payload,
            storage,
        }));
    }
    Ok(None)
}

/// The `TSD.DrawableArchive` of a drawable message: field 1 of an
/// `ImageArchive`/`MovieArchive`, field 1 of field 1 of a `ShapeInfoArchive`.
fn drawable_payload<'a>(
    selector: DrawableSelector,
    payload: &'a [u8],
    budget: &mut Budget,
) -> GraphResult<&'a [u8]> {
    let view = budget.parse(payload, 1)?;
    let outer = unique_field(&view, WRAPPED_DRAWABLE_FIELD, 2)?
        .ok_or(GraphError::InvalidSource)?
        .payload();
    if !matches!(
        selector,
        DrawableSelector::Shape(_) | DrawableSelector::TextBox(_)
    ) {
        return Ok(outer);
    }
    let shape = budget.parse(outer, 2)?;
    let drawable = unique_field(&shape, WRAPPED_DRAWABLE_FIELD, 2)?
        .ok_or(GraphError::InvalidSource)?
        .payload();
    let view = budget.parse(drawable, 3)?;
    if view
        .fields()
        .any(|field| field.number() == DRAWABLE_PENCIL_ANNOTATIONS_FIELD)
    {
        return Err(GraphError::InvalidSource);
    }
    Ok(drawable)
}

/// The owned storage of an ordinary body shape (`text_box == false`: an
/// explicit non-text-box flag and one storage recorded in both storage
/// slots) or of a body text box (`text_box == true`: the text-box flag, an
/// owned storage and a deprecated slot that is absent or names the same
/// storage, the host editor's text-box ownership proof). `None` for a shape
/// of the other role or without the explicit flag.
fn shape_storage(
    payload: &[u8],
    text_box: bool,
    budget: &mut Budget,
) -> GraphResult<Option<NonZeroU64>> {
    let view = budget.parse(payload, 1)?;
    let flag = unique_field(&view, SHAPE_INFO_IS_TEXT_BOX_FIELD, 0)?
        .map(|field| canonical_varint(field.payload()))
        .transpose()?;
    if flag != Some(u64::from(text_box)) {
        return Ok(None);
    }
    let owned =
        unique_field(&view, SHAPE_INFO_OWNED_STORAGE_FIELD, 2)?.ok_or(GraphError::InvalidSource)?;
    let owned = parse_reference(owned.payload(), budget)?;
    match unique_field(&view, SHAPE_INFO_DEPRECATED_STORAGE_FIELD, 2)? {
        Some(deprecated) => {
            if owned != parse_reference(deprecated.payload(), budget)? {
                return Err(GraphError::InvalidSource);
            }
        },
        None if text_box => {},
        None => return Err(GraphError::InvalidSource),
    }
    Ok(Some(owned))
}

#[cfg(test)]
mod tests {
    use super::{GraphError, GraphLimit, map_package_error, zip_record_equal_outside};
    use crate::{Error as SemanticError, PackageError};

    #[test]
    fn semantic_package_limits_keep_resource_kinds() {
        let cases = [
            (
                PackageError::Semantic(SemanticError::TooManySections {
                    actual: 9,
                    limit: 8,
                }),
                GraphLimit::SemanticItems,
                9,
                8,
            ),
            (
                PackageError::Semantic(SemanticError::TooManyBodyStorages {
                    actual: 11,
                    limit: 10,
                }),
                GraphLimit::SemanticItems,
                11,
                10,
            ),
            (
                PackageError::Semantic(SemanticError::TextTooLarge {
                    observed: 257,
                    limit: 256,
                }),
                GraphLimit::SemanticTextBytes,
                257,
                256,
            ),
        ];

        for (error, expected_kind, observed, maximum) in cases {
            assert!(matches!(
                map_package_error(error),
                GraphError::Limit {
                    kind,
                    observed: actual,
                    maximum: limit,
                } if kind == expected_kind && actual == observed && limit == maximum
            ));
        }
    }

    #[test]
    fn locality_negative_control_rejects_unrelated_zip_mutation() {
        let source = b"untouched-member";
        let mut candidate = source.to_vec();
        candidate[3] ^= 1;
        assert!(!zip_record_equal_outside(source, &candidate, &[]));
        assert!(zip_record_equal_outside(source, &candidate, &[(3, 4)]));
    }
}
