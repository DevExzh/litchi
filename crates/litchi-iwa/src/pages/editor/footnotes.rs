//! Native body-footnote CRUD for Pages documents.

use std::collections::HashSet;
use std::hash::Hash;
use std::str;

use litchi_iwa_common::{
    LimitKind, WireLimits,
    varint::{decode_varint_from_bytes, encoded_len as varint_len},
    wire::{WireDescent, WireFieldView, WirePreflight, WireView, preflight_wire_tree_with_limits},
};
use litchi_iwa_protos::pages_body_codec;
use litchi_iwa_protos::pages_footnote_codec;
use litchi_iwa_protos::pages_footnote_marker_codec;
use prost::Message;

use super::text_box_create::body_text_storage;
use super::{
    DOCUMENT_OBJECT_ID, PagesEditor, STORAGE_MESSAGE_TYPES, find_object_archive,
    package_references_object,
};
use crate::archive::{ArchiveObject, RawMessage};
use crate::package_metadata::{
    add_component_object_uuids, component_identifier_for_object_uuid, next_object_identifier,
    release_package_identifier_suffix, remove_component_external_references_to_object,
    remove_component_object_uuids, set_package_last_object_identifier,
};
use crate::protobuf::{tsp, tswp};
use crate::text::IWorkTextEditor;
use crate::text::editor::storage_object_references;
#[cfg(test)]
use crate::wire::repeated_length_delimited_payloads;
use crate::wire::{patch_length_delimited_field, rewrite_repeated_length_delimited_fields};
use crate::{Error, IWorkPackage, Result};
use litchi_pages::footnote::body::{Footnote, Position, Selector};

const FOOTNOTE_REFERENCE_MESSAGE_TYPE: u32 = 2_008;
const TEXTUAL_ATTACHMENT_MESSAGE_TYPE: u32 = 2_004;
#[cfg(test)]
const FOOTNOTE_SUPER_FIELD: u32 = 1;
const FOOTNOTE_TABLE_FIELD: u32 = 16;
const TABLE_ENTRIES_FIELD: u32 = 1;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const FOOTNOTE_ANCHOR: char = '\u{000e}';
const FOOTNOTE_ANCHOR_TEXT: &str = "\u{000e}";
const FOOTNOTE_ANCHOR_UNIT: u16 = 0x000e;
const FOOTNOTE_MARK: char = '\u{fffc}';
const FOOTNOTE_CONTENT_PREFIX: &str = "\u{fffc} ";
const FOOTNOTE_REFERENCE_CODEC_RECURSION_LIMIT: u32 = 64;
const MAX_BODY_FOOTNOTES: usize = 4096;
const MAX_FOOTNOTE_TABLE_BYTES: usize = WireLimits::MAX_INPUT_BYTES;

/// Native Pages footnote data plus the private objects it owns.
#[derive(Debug, Clone)]
pub(super) struct BodyFootnoteGraph {
    pub(super) footnote: Footnote,
    reference_id: u64,
    storage_id: u64,
    marker_id: u64,
}

#[derive(Debug)]
struct FootnoteTableEntry {
    index: u32,
    reference_id: u64,
    raw: Vec<u8>,
}

#[derive(Debug, Clone, Copy)]
struct FootnoteObjectIds {
    reference: u64,
    storage: u64,
    marker: u64,
}

impl FootnoteObjectIds {
    fn allocate(first: u64) -> Result<Self> {
        let identifier = |offset| {
            first.checked_add(offset).ok_or_else(|| {
                Error::ParseError("Pages footnote object identifier overflow".to_owned())
            })
        };
        Ok(Self {
            reference: identifier(0)?,
            storage: identifier(1)?,
            marker: identifier(2)?,
        })
    }

    const fn last(self) -> u64 {
        self.marker
    }
}

impl PagesEditor {
    /// Read every native footnote attached to the main Pages body.
    pub fn body_footnotes(&self) -> Result<Vec<Footnote>> {
        let graphs = body_footnote_graphs(self.package(), self.body_storage_id.get())?;
        let mut footnotes = Vec::new();
        footnotes.try_reserve_exact(graphs.len()).map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Pages body footnotes",
                amount: graphs.len(),
            })
        })?;
        footnotes.extend(graphs.into_iter().map(|graph| graph.footnote));
        Ok(footnotes)
    }

    /// Insert a native Pages footnote at a UTF-16 body position.
    ///
    /// The inserted body character is Pages' private U+000E footnote anchor;
    /// use [`Self::body_footnotes`] instead of treating that character as text.
    pub fn insert_body_footnote(
        &mut self,
        position: Position,
        text: impl AsRef<str>,
    ) -> Result<Footnote> {
        let text = text.as_ref();
        validate_footnote_text(text)?;
        let position_u32 = position.utf16_index();
        let position_index = usize::try_from(position_u32).map_err(|_| {
            Error::ParseError("Pages footnote position exceeds the platform index range".to_owned())
        })?;
        body_footnote_graphs(self.package(), self.body_storage_id.get())?;

        let mut text_editor = IWorkTextEditor::from_package(self.package().clone());
        text_editor.replace_text(
            self.body_storage_id,
            position_index..position_index,
            FOOTNOTE_ANCHOR_TEXT,
        )?;
        let mut staged = text_editor.into_package();
        let ids = FootnoteObjectIds::allocate(next_object_identifier(&staged)?)?;
        let body = storage_at(&staged, self.body_storage_id.get(), "Pages body")?.1;
        let archive_name = find_object_archive(&staged, self.body_storage_id.get())?;
        let objects = new_footnote_objects(ids, text, &body)?;

        insert_footnote_reference(
            &mut staged,
            &archive_name,
            self.body_storage_id.get(),
            position_u32,
            ids.reference,
        )?;
        staged.update_archive(&archive_name, |archive| {
            for object in objects {
                archive.insert_object(object)?;
            }
            Ok(())
        })?;
        add_component_object_uuids(&mut staged, DOCUMENT_OBJECT_ID, &[ids.storage])?;
        set_package_last_object_identifier(&mut staged, ids.last())?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let created = body_footnote_by_selector(&verified, Selector::At(position))?.footnote;
        if created.position.utf16_index() != position_u32
            || created.text.as_ref() != text
            || created.custom_mark.is_some()
        {
            return Err(Error::InvalidFormat(
                "Pages footnote insertion failed validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(created)
    }

    /// Delete one native body footnote, its body anchor, and its owned objects.
    pub fn remove_body_footnote(&mut self, selector: Selector) -> Result<Footnote> {
        let removed = body_footnote_by_selector(self, selector)?;
        let start = usize::try_from(removed.footnote.position.utf16_index()).map_err(|_| {
            Error::ParseError("Pages footnote position exceeds the platform index range".to_owned())
        })?;
        let end = start
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Pages footnote anchor range overflow".to_owned()))?;
        // Keep the legacy graph edit and its cleanup off the live editor until
        // the removed native reference is absent. Position-only validation is
        // incorrect when a following footnote shifts into the deleted anchor.
        let mut staged = self.clone();
        staged.replace_body_text(start..end, "")?;
        if body_footnote_graphs(staged.package(), staged.body_storage_id.get())?
            .iter()
            .any(|graph| graph.reference_id == removed.reference_id)
        {
            return Err(Error::InvalidFormat(
                "Pages footnote deletion failed validation".to_owned(),
            ));
        }
        let result = removed.footnote;
        *self = staged;
        Ok(result)
    }
}

pub(super) fn body_footnote_graphs(
    package: &IWorkPackage,
    body_storage_id: u64,
) -> Result<Vec<BodyFootnoteGraph>> {
    let mut wire_budget = FootnoteGraphBudget::default();
    let entries = with_storage_projection(
        package,
        body_storage_id,
        "Pages body",
        &mut wire_budget,
        |body, wire_budget| {
            footnote_table_entries_from_projection(body_storage_id, body, wire_budget)
        },
    )?;
    let mut seen = HashSet::new();
    let limits = wire_budget.limits;
    reserve_footnote_set(
        &mut seen,
        entries.len(),
        limits,
        "Pages body footnote references",
    )?;
    wire_budget.charge_allocation(entries.len(), "Pages body footnote references")?;
    let mut footnotes = Vec::new();
    reserve_footnote_collection(
        &mut footnotes,
        entries.len(),
        limits,
        "Pages body footnote graphs",
    )?;
    wire_budget.charge_allocation(entries.len(), "Pages body footnote graphs")?;
    for entry in entries {
        if !seen.insert(entry.reference_id) {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {body_storage_id} references footnote object {} more than once",
                entry.reference_id
            )));
        }
        footnotes.push(decode_footnote_graph(
            package,
            entry.index,
            entry.reference_id,
            &mut wire_budget,
        )?);
    }
    Ok(footnotes)
}

/// Reclaim footnote graphs whose anchors were removed by an ordinary body edit.
pub(super) fn cleanup_removed_body_footnotes(
    package: &mut IWorkPackage,
    body_storage_id: u64,
    before: &[BodyFootnoteGraph],
) -> Result<()> {
    if before.is_empty() {
        return Ok(());
    }
    let remaining_graphs = body_footnote_graphs(package, body_storage_id)?;
    let mut remaining = HashSet::new();
    remaining.try_reserve(remaining_graphs.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "Pages body remaining footnote references",
            amount: remaining_graphs.len(),
        })
    })?;
    remaining.extend(remaining_graphs.into_iter().map(|graph| graph.reference_id));

    let removed_count = before.iter().try_fold(0usize, |count, graph| {
        if remaining.contains(&graph.reference_id) {
            Ok(count)
        } else {
            count.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat(
                    "Pages removed footnote graph count overflows usize".to_owned(),
                )
            })
        }
    })?;
    if removed_count == 0 {
        return Ok(());
    }
    let identifier_count = footnote_cleanup_identifier_count(removed_count)?;

    let mut staged = package.clone();
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(identifier_count)
        .map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Pages removed footnote object identifiers",
                amount: identifier_count,
            })
        })?;
    for graph in before
        .iter()
        .filter(|graph| !remaining.contains(&graph.reference_id))
    {
        identifiers.extend(remove_unreferenced_footnote_graph(&mut staged, graph)?);
    }
    release_package_identifier_suffix(&mut staged, &identifiers)?;
    IWorkPackage::from_bytes(&staged.to_bytes()?)?;
    *package = staged;
    Ok(())
}

fn footnote_cleanup_identifier_count(removed_count: usize) -> Result<usize> {
    removed_count.checked_mul(3).ok_or_else(|| {
        Error::InvalidFormat("Pages removed footnote identifier count overflows usize".to_owned())
    })
}

fn body_footnote_by_selector(
    editor: &PagesEditor,
    selector: Selector,
) -> Result<BodyFootnoteGraph> {
    let footnotes = body_footnote_graphs(editor.package(), editor.body_storage_id.get())?;
    match selector {
        Selector::Index(index) => footnotes.into_iter().nth(index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages body has no footnote at source index {index}"
            ))
        }),
        Selector::At(position) => {
            let mut matches = footnotes
                .into_iter()
                .filter(|graph| graph.footnote.position == position);
            let Some(graph) = matches.next() else {
                return Err(Error::InvalidFormat(format!(
                    "Pages body has no footnote at UTF-16 position {}",
                    position.utf16_index()
                )));
            };
            if matches.next().is_some() {
                return Err(Error::InvalidFormat(format!(
                    "Pages body has more than one footnote at UTF-16 position {}",
                    position.utf16_index()
                )));
            }
            Ok(graph)
        },
    }
}

/// Aggregate wire accounting for one rooted body-footnote graph read.
///
/// The strict codecs receive a per-message `DecodeOptions`, so keeping their
/// options source-sized is not sufficient when one body names many reference,
/// storage, and marker payloads.  This coordinator carries the checked
/// counters across the complete graph and charges the borrowed `WireView`
/// scans before any generated projection is allowed to run.
#[derive(Debug, Clone, Copy)]
struct FootnoteGraphBudget {
    limits: WireLimits,
    input_bytes: usize,
    fields: usize,
    work: usize,
    max_depth: usize,
    allocations: usize,
}

impl Default for FootnoteGraphBudget {
    fn default() -> Self {
        Self {
            limits: WireLimits::default(),
            input_bytes: 0,
            fields: 0,
            work: 0,
            max_depth: 0,
            allocations: 0,
        }
    }
}

impl FootnoteGraphBudget {
    fn remaining_limits(&self) -> Result<WireLimits> {
        let input_bytes = self.remaining(
            self.input_bytes,
            self.limits.max_input_bytes(),
            LimitKind::InputBytes,
        )?;
        let fields = self.remaining(self.fields, self.limits.max_fields(), LimitKind::Fields)?;
        let work = self.remaining(
            self.work,
            self.limits.max_rewrite_work(),
            LimitKind::RewriteWork,
        )?;
        self.limits
            .with_input_bytes(input_bytes)
            .and_then(|limits| limits.with_fields(fields))
            .and_then(|limits| limits.with_rewrite_work(work))
            .map_err(Into::into)
    }

    fn remaining(&self, current: usize, maximum: usize, kind: LimitKind) -> Result<usize> {
        let remaining = maximum.checked_sub(current).ok_or_else(|| {
            Error::InvalidFormat("Pages footnote aggregate budget overflows usize".to_owned())
        })?;
        if remaining == 0 {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed: current,
                limit: maximum,
            }));
        }
        Ok(remaining)
    }

    fn charge(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: LimitKind,
        message: &'static str,
    ) -> Result<()> {
        let observed = current
            .checked_add(amount)
            .ok_or_else(|| Error::InvalidFormat(message.to_owned()))?;
        if observed > maximum {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit: maximum,
            }));
        }
        *current = observed;
        Ok(())
    }

    fn charge_input(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.input_bytes,
            amount,
            self.limits.max_input_bytes(),
            LimitKind::InputBytes,
            "Pages footnote aggregate input-byte count overflows usize",
        )
    }

    fn charge_fields(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.fields,
            amount,
            self.limits.max_fields(),
            LimitKind::Fields,
            "Pages footnote aggregate field count overflows usize",
        )
    }

    fn charge_work(&mut self, amount: usize) -> Result<()> {
        Self::charge(
            &mut self.work,
            amount,
            self.limits.max_rewrite_work(),
            LimitKind::RewriteWork,
            "Pages footnote aggregate work count overflows usize",
        )
    }

    fn charge_depth(&mut self, depth: usize) -> Result<()> {
        if depth > self.limits.max_nesting() {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Nesting,
                observed: depth,
                limit: self.limits.max_nesting(),
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }

    fn charge_allocation(&mut self, amount: usize, resource: &'static str) -> Result<()> {
        Self::charge(
            &mut self.allocations,
            amount,
            MAX_BODY_FOOTNOTES,
            LimitKind::Fields,
            "Pages footnote allocation count overflows usize",
        )?;
        self.charge_work(amount).map_err(|error| match error {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded { .. }) => error,
            _ => Error::IwaCommon(litchi_iwa_common::Error::Allocation { resource, amount }),
        })
    }

    fn scan<'a, F>(
        &mut self,
        source: &'a [u8],
        mut descent: F,
    ) -> Result<(WireView<'a>, WirePreflight)>
    where
        F: FnMut(litchi_iwa_common::wire::WireVisit<'_, '_>) -> Result<WireDescent>,
    {
        let limits = self.remaining_limits()?;
        let report = preflight_wire_tree_with_limits(source, limits, |visit| {
            descent(visit).map_err(|error| match error {
                Error::IwaCommon(error) => error,
                Error::InvalidFormat(message) => litchi_iwa_common::Error::InvalidFormat(message),
                other => litchi_iwa_common::Error::InvalidFormat(other.to_string()),
            })
        })?;
        let view = WireView::parse_with_limits(source, limits)?;
        self.charge_input(report.scanned_bytes())?;
        self.charge_fields(view.len())?;
        let scan_work = source.len().checked_add(view.len()).ok_or_else(|| {
            Error::InvalidFormat("Pages footnote wire scan work overflows usize".to_owned())
        })?;
        self.charge_work(scan_work)?;
        Ok((view, report))
    }

    fn reference_decode_options(
        source: &[u8],
        report: WirePreflight,
    ) -> Result<pages_footnote_codec::DecodeOptions> {
        let work = report.scanned_bytes().checked_mul(2).ok_or_else(|| {
            Error::InvalidFormat("Pages footnote reference work overflows usize".to_owned())
        })?;
        Ok(pages_footnote_codec::DecodeOptions::new(
            source.len().max(1),
            report.fields().max(1),
            work.max(1),
            FOOTNOTE_REFERENCE_CODEC_RECURSION_LIMIT,
        ))
    }

    fn body_decode_options(
        source: &[u8],
        report: WirePreflight,
    ) -> Result<pages_body_codec::DecodeOptions> {
        let work = report.scanned_bytes().checked_mul(2).ok_or_else(|| {
            Error::InvalidFormat("Pages footnote body work overflows usize".to_owned())
        })?;
        Ok(pages_body_codec::DecodeOptions::new(
            source.len().max(1),
            report.fields().max(1),
            work.max(1),
            FOOTNOTE_REFERENCE_CODEC_RECURSION_LIMIT,
        ))
    }

    fn marker_decode_options(
        source: &[u8],
        report: WirePreflight,
    ) -> Result<pages_footnote_marker_codec::DecodeOptions> {
        let work = report.scanned_bytes().checked_mul(2).ok_or_else(|| {
            Error::InvalidFormat("Pages footnote marker work overflows usize".to_owned())
        })?;
        Ok(pages_footnote_marker_codec::DecodeOptions::new(
            source.len().max(1),
            report.fields().max(1),
            work.max(1),
            FOOTNOTE_REFERENCE_CODEC_RECURSION_LIMIT,
        ))
    }

    fn charge_codec(&mut self, report: WirePreflight) -> Result<()> {
        self.charge_fields(report.fields())?;
        let work = report.scanned_bytes().checked_mul(2).ok_or_else(|| {
            Error::InvalidFormat("Pages footnote codec work overflows usize".to_owned())
        })?;
        self.charge_work(work)
    }
}

/// Borrowed read/validation facts from one legacy `TSWP.StorageArchive`.
///
/// The original payload is retained as the only preservation representation;
/// this view owns no generated archive and never reconstructs a protobuf value.
/// Its selected text fragments and table payloads borrow the package-owned
/// source, so unknown fields remain in the original bytes rather than being
/// reconstructed. Prost remains intentionally scoped to the mutation/template
/// helpers below.
#[derive(Debug)]
struct FootnoteStorageProjection<'source> {
    kind: Option<i32>,
    text: Vec<&'source str>,
    table_attachment: Option<&'source [u8]>,
    table_footnote: Option<&'source [u8]>,
}

impl<'source> FootnoteStorageProjection<'source> {
    fn decode(
        source: &'source [u8],
        storage_id: u64,
        label: &str,
        wire_budget: &mut FootnoteGraphBudget,
    ) -> Result<Self> {
        let validation = litchi_iwa_text_wire::validate_storage_with_limits(
            source,
            litchi_iwa_text_wire::RewriteLimits::default(),
        )
        .map_err(|error| map_storage_wire_error(storage_id, label, error))?;
        wire_budget.charge_input(source.len())?;
        wire_budget.charge_fields(validation.fields())?;
        wire_budget.charge_work(validation.validation_work())?;
        // The full text-storage preflight admits the four schema levels
        // (storage, table, entry, and reference/range). Keep that depth in the
        // aggregate graph budget before retaining any selected projection.
        wire_budget.charge_depth(4)?;

        let (view, _) = wire_budget.scan(source, no_footnote_descent)?;
        let mut kind = None;
        let mut text = Vec::new();
        let mut table_attachment = None;
        let mut table_footnote = None;
        for field in view.fields() {
            field.validate_canonical_framing().map_err(|error| {
                Error::InvalidFormat(format!(
                    "Pages {label} storage {storage_id} has noncanonical wire framing: {error}"
                ))
            })?;
            match field.number() {
                1 => {
                    if kind.is_some() {
                        return Err(Error::InvalidFormat(format!(
                            "Pages {label} storage {storage_id} repeats its kind field"
                        )));
                    }
                    kind = Some(storage_kind(field)?);
                },
                3 => {
                    let fragment = field.canonical_payload().map_err(|error| {
                        Error::InvalidFormat(format!(
                            "Pages {label} storage {storage_id} has invalid text framing: {error}"
                        ))
                    })?;
                    let fragment = str::from_utf8(fragment).map_err(|error| {
                        Error::InvalidFormat(format!(
                            "Pages {label} storage {storage_id} has invalid UTF-8 text: {error}"
                        ))
                    })?;
                    let requested = text.len().checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat(
                            "Pages footnote storage text fragment count overflows usize".to_owned(),
                        )
                    })?;
                    text.try_reserve(1).map_err(|_| {
                        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                            resource: "Pages footnote storage text fragments",
                            amount: requested,
                        })
                    })?;
                    wire_budget.charge_work(fragment.len())?;
                    text.push(fragment);
                },
                9 => {
                    // Keep attachment-table payloads opaque here. The outer
                    // key/length framing is canonical, but unknown nested
                    // fields are intentionally not interpreted or claimed
                    // canonical by this read projection.
                    if table_attachment.is_some() {
                        return Err(Error::InvalidFormat(format!(
                            "Pages {label} storage {storage_id} repeats its attachment table"
                        )));
                    }
                    table_attachment = Some(field.canonical_payload().map_err(|error| {
                        Error::InvalidFormat(format!(
                            "Pages {label} storage {storage_id} has invalid attachment-table framing: {error}"
                        ))
                    })?);
                },
                FOOTNOTE_TABLE_FIELD => {
                    // As with the attachment table above, only the outer
                    // framing is canonical here; unknown nested table fields
                    // remain opaque to this projection.
                    if table_footnote.is_some() {
                        return Err(Error::InvalidFormat(format!(
                            "Pages {label} storage {storage_id} repeats its footnote table"
                        )));
                    }
                    table_footnote = Some(field.canonical_payload().map_err(|error| {
                        Error::InvalidFormat(format!(
                            "Pages {label} storage {storage_id} has invalid footnote-table framing: {error}"
                        ))
                    })?);
                },
                _ => {},
            }
        }
        Ok(Self {
            kind,
            text,
            table_attachment,
            table_footnote,
        })
    }

    fn kind(&self) -> Option<i32> {
        self.kind
    }

    fn text(&self) -> &[&'source str] {
        &self.text
    }

    fn table_attachment(&self) -> Option<&'source [u8]> {
        self.table_attachment
    }

    fn table_footnote(&self) -> Option<&'source [u8]> {
        self.table_footnote
    }
}

fn storage_kind(field: WireFieldView<'_>) -> Result<i32> {
    if field.wire_type() != 0 {
        return Err(Error::InvalidFormat(
            "Pages storage kind is not a canonical varint".to_owned(),
        ));
    }
    let value = canonical_varint(field, "Pages storage kind")?;
    if value > 0x7fff_ffff && value < 0xffff_ffff_8000_0000 {
        return Err(Error::InvalidFormat(
            "Pages storage kind is not a canonical int32".to_owned(),
        ));
    }
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_possible_wrap,
        reason = "The preceding range check proves a canonical sign-extended int32."
    )]
    Ok(value as i32)
}

fn canonical_varint(field: WireFieldView<'_>, label: &str) -> Result<u64> {
    let (value, length) = decode_varint_from_bytes(field.payload()).map_err(|error| {
        Error::InvalidFormat(format!("{label} contains an invalid varint: {error}"))
    })?;
    if length != field.payload().len() || length != varint_len(value) {
        return Err(Error::InvalidFormat(format!(
            "{label} contains a noncanonical varint"
        )));
    }
    Ok(value)
}

fn footnote_reference_descent(
    visit: litchi_iwa_common::wire::WireVisit<'_, '_>,
) -> Result<WireDescent> {
    let field = visit.field();
    if visit.path().is_empty() && matches!(field.number(), 1 | 2) {
        if field.wire_type() != 2 {
            return Err(Error::InvalidFormat(
                "Pages footnote reference nested field is not length-delimited".to_owned(),
            ));
        }
        field.validate_canonical_framing()?;
        return Ok(WireDescent::Descend);
    }
    Ok(WireDescent::Skip)
}

fn footnote_body_descent(visit: litchi_iwa_common::wire::WireVisit<'_, '_>) -> Result<WireDescent> {
    let field = visit.field();
    if visit.path().is_empty() && field.number() == 2 {
        if field.wire_type() != 2 {
            return Err(Error::InvalidFormat(
                "Pages footnote body nested field is not length-delimited".to_owned(),
            ));
        }
        field.validate_canonical_framing()?;
        return Ok(WireDescent::Descend);
    }
    Ok(WireDescent::Skip)
}

fn no_footnote_descent(_visit: litchi_iwa_common::wire::WireVisit<'_, '_>) -> Result<WireDescent> {
    Ok(WireDescent::Skip)
}

fn decode_footnote_graph(
    package: &IWorkPackage,
    position: u32,
    reference_id: u64,
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<BodyFootnoteGraph> {
    let reference_archive = find_object_archive(package, reference_id)?;
    let reference_archive_data = package.archive(&reference_archive)?;
    let reference_object = reference_archive_data.object(reference_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages footnote object {reference_id} is missing"))
    })?;
    let reference_data = object_message_data(
        reference_object,
        FOOTNOTE_REFERENCE_MESSAGE_TYPE,
        "Pages footnote reference",
    )?;
    let (_, reference_report) = wire_budget.scan(reference_data, footnote_reference_descent)?;
    let reference = pages_footnote_codec::decode_footnote_reference(
        reference_data,
        FootnoteGraphBudget::reference_decode_options(reference_data, reference_report)?,
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "Pages footnote reference object {reference_id} failed strict validation: {error}"
        ))
    })?;
    wire_budget.charge_codec(reference_report)?;
    let custom_mark = owned_custom_mark(reference.custom_mark_string(), wire_budget)?;
    if reference.super_kind().is_some_and(|kind| {
        kind != tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32
    }) {
        return Err(Error::InvalidFormat(format!(
            "Pages footnote object {reference_id} has the wrong attachment kind"
        )));
    }
    let storage_id = reference
        .contained_storage()
        .map(|value| value.identifier().get())
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages footnote object {reference_id} has no contained storage"
            ))
        })?;
    let (content, marker_id) = with_storage_projection(
        package,
        storage_id,
        "Pages footnote",
        wire_budget,
        |storage, wire_budget| {
            if storage.kind() != Some(tswp::storage_archive::KindType::Footnote as i32) {
                return Err(Error::InvalidFormat(format!(
                    "Pages footnote storage {storage_id} is not a native footnote storage"
                )));
            }
            let content = storage_text(storage, wire_budget)?;
            let marker_id = footnote_marker_id(storage_id, storage, wire_budget)?;
            Ok((content, marker_id))
        },
    )?;
    let text = content
        .strip_prefix(FOOTNOTE_CONTENT_PREFIX)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages footnote storage {storage_id} lacks its native marker prefix"
            ))
        })?;
    validate_footnote_marker(package, marker_id, wire_budget)?;
    let text = owned_boxed_str(text, "Pages footnote text", wire_budget)?;

    let position = Position::from_utf16_index(usize::try_from(position).map_err(|_| {
        Error::ParseError("Pages footnote position exceeds the platform index range".to_owned())
    })?)
    .map_err(|error| Error::ParseError(format!("invalid Pages footnote position: {error}")))?;
    let footnote = Footnote::with_custom_mark(position, text, custom_mark)
        .map_err(|error| Error::ParseError(format!("invalid Pages footnote value: {error}")))?;

    Ok(BodyFootnoteGraph {
        footnote,
        reference_id,
        storage_id,
        marker_id,
    })
}

fn owned_custom_mark(
    custom_mark: Option<&str>,
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<Option<Box<str>>> {
    let Some(custom_mark) = custom_mark else {
        return Ok(None);
    };
    if custom_mark.len() > litchi_pages::footnote::body::MAX_CUSTOM_MARK_BYTES {
        return Err(Error::ParseError(
            "Pages footnote custom marker exceeds its semantic byte budget".to_owned(),
        ));
    }
    Ok(Some(owned_boxed_str(
        custom_mark,
        "Pages footnote custom marker",
        wire_budget,
    )?))
}

fn owned_boxed_str(
    value: &str,
    resource: &'static str,
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<Box<str>> {
    wire_budget.charge_work(value.len())?;
    let mut owned = String::new();
    owned.try_reserve_exact(value.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource,
            amount: value.len(),
        })
    })?;
    owned.push_str(value);
    Ok(owned.into_boxed_str())
}

fn reserve_footnote_collection<T>(
    values: &mut Vec<T>,
    additional: usize,
    limits: WireLimits,
    resource: &'static str,
) -> Result<()> {
    let requested = values
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflows usize")))?;
    let limit = limits.max_fields().min(MAX_BODY_FOOTNOTES);
    if requested > limit {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: requested,
            limit,
        }));
    }
    values.try_reserve_exact(additional).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource,
            amount: requested,
        })
    })
}

fn reserve_footnote_set<T: Eq + Hash>(
    values: &mut HashSet<T>,
    additional: usize,
    limits: WireLimits,
    resource: &'static str,
) -> Result<()> {
    let requested = values
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflows usize")))?;
    let limit = limits.max_fields().min(MAX_BODY_FOOTNOTES);
    if requested > limit {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: requested,
            limit,
        }));
    }
    values.try_reserve(additional).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource,
            amount: requested,
        })
    })
}

fn reserve_footnote_entries<T>(
    values: &mut Vec<T>,
    additional: usize,
    resource: &'static str,
) -> Result<()> {
    let requested = values
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::InvalidFormat(format!("{resource} size overflows usize")))?;
    if requested > MAX_BODY_FOOTNOTES {
        return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: LimitKind::Fields,
            observed: requested,
            limit: MAX_BODY_FOOTNOTES,
        }));
    }
    values.try_reserve_exact(additional).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource,
            amount: requested,
        })
    })
}

fn validate_footnote_marker(
    package: &IWorkPackage,
    marker_id: u64,
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<()> {
    let archive_name = find_object_archive(package, marker_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(marker_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages footnote marker object {marker_id} is missing"
        ))
    })?;
    let marker_data = object_message_data(
        object,
        TEXTUAL_ATTACHMENT_MESSAGE_TYPE,
        "Pages footnote marker",
    )?;
    let (_, report) = wire_budget.scan(marker_data, no_footnote_descent)?;
    let marker = pages_footnote_marker_codec::decode_textual_attachment(
        marker_data,
        FootnoteGraphBudget::marker_decode_options(marker_data, report)?,
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "Pages footnote marker object {marker_id} failed strict validation: {error}"
        ))
    })?;
    wire_budget.charge_codec(report)?;
    if marker.kind() != Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32) {
        return Err(Error::InvalidFormat(format!(
            "Pages footnote marker object {marker_id} has the wrong attachment kind"
        )));
    }
    Ok(())
}

fn footnote_marker_id(
    storage_id: u64,
    storage: &FootnoteStorageProjection<'_>,
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<u64> {
    let table = storage.table_attachment().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Pages footnote storage {storage_id} has no marker attachment table"
        ))
    })?;
    let (view, _) = wire_budget.scan(table, no_footnote_descent)?;
    let mut marker = None;
    for field in view
        .fields()
        .filter(|field| field.number() == TABLE_ENTRIES_FIELD)
    {
        field.validate_canonical_framing()?;
        let entry = field.canonical_payload()?;
        let (index, object) = marker_table_entry(entry, wire_budget)?;
        if index != 0 {
            continue;
        }
        if marker.replace(object).is_some() {
            return Err(Error::InvalidFormat(format!(
                "Pages footnote storage {storage_id} must have exactly one marker attachment at index zero"
            )));
        }
    }
    let Some(marker_id) = marker else {
        return Err(Error::InvalidFormat(format!(
            "Pages footnote storage {storage_id} must have exactly one marker attachment at index zero"
        )));
    };
    Ok(marker_id)
}

fn marker_table_entry(source: &[u8], wire_budget: &mut FootnoteGraphBudget) -> Result<(u32, u64)> {
    let (view, _) = wire_budget.scan(source, no_footnote_descent)?;
    let mut character_index = None;
    let mut object = None;
    for field in view.fields() {
        field.validate_canonical_framing()?;
        match field.number() {
            1 => {
                if character_index.is_some() {
                    return Err(Error::InvalidFormat(
                        "Pages marker table entry repeats character index".to_owned(),
                    ));
                }
                let value = canonical_varint(field, "Pages marker character index")?;
                character_index = Some(u32::try_from(value).map_err(|_| {
                    Error::InvalidFormat("Pages marker character index exceeds u32".to_owned())
                })?);
            },
            2 => {
                if object.is_some() {
                    return Err(Error::InvalidFormat(
                        "Pages marker table entry repeats object reference".to_owned(),
                    ));
                }
                let payload = field.canonical_payload()?;
                object = Some(marker_reference(payload, wire_budget)?);
            },
            _ => {},
        }
    }
    Ok((
        character_index.ok_or_else(|| {
            Error::InvalidFormat("Pages marker table entry has no character index".to_owned())
        })?,
        object.ok_or_else(|| {
            Error::InvalidFormat("Pages marker table entry has no object reference".to_owned())
        })?,
    ))
}

fn marker_reference(source: &[u8], wire_budget: &mut FootnoteGraphBudget) -> Result<u64> {
    let (view, _) = wire_budget.scan(source, no_footnote_descent)?;
    let mut identifier = None;
    for field in view.fields() {
        field.validate_canonical_framing()?;
        if field.number() != 1 {
            continue;
        }
        if identifier.is_some() {
            return Err(Error::InvalidFormat(
                "Pages marker reference repeats identifier".to_owned(),
            ));
        }
        let value = canonical_varint(field, "Pages marker reference identifier")?;
        if value == 0 {
            return Err(Error::InvalidFormat(
                "Pages marker reference identifier is zero".to_owned(),
            ));
        }
        identifier = Some(value);
    }
    identifier.ok_or_else(|| {
        Error::InvalidFormat("Pages marker table entry has no object reference".to_owned())
    })
}

fn storage_text(
    storage: &FootnoteStorageProjection<'_>,
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<String> {
    let length = storage.text().iter().try_fold(0usize, |total, fragment| {
        total.checked_add(fragment.len()).ok_or_else(|| {
            Error::InvalidFormat("Pages footnote text length overflows usize".to_owned())
        })
    })?;
    wire_budget.charge_work(length)?;
    let mut text = String::new();
    text.try_reserve_exact(length).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "Pages footnote storage text",
            amount: length,
        })
    })?;
    for fragment in storage.text() {
        text.push_str(fragment);
    }
    Ok(text)
}

fn new_footnote_objects(
    ids: FootnoteObjectIds,
    text: &str,
    body: &tswp::StorageArchive,
) -> Result<[ArchiveObject; 3]> {
    let mut content = String::with_capacity(FOOTNOTE_CONTENT_PREFIX.len() + text.len());
    content.push_str(FOOTNOTE_CONTENT_PREFIX);
    content.push_str(text);
    let mut storage = body_text_storage(&content, body);
    storage.kind = Some(tswp::storage_archive::KindType::Footnote as i32);
    storage.table_attachment = Some(tswp::ObjectAttributeTable {
        entries: vec![tswp::object_attribute_table::ObjectAttribute {
            character_index: 0,
            object: Some(reference(ids.marker)),
        }],
    });
    let marker = tswp::TextualAttachmentArchive {
        string_equivalent: None,
        kind: Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32),
    };
    let attachment = tswp::FootnoteReferenceAttachmentArchive {
        super_: None,
        contained_storage: Some(reference(ids.storage)),
        custom_mark_string: None,
    };
    let storage_references = storage_object_references(&storage);
    Ok([
        pages_object(
            ids.reference,
            FOOTNOTE_REFERENCE_MESSAGE_TYPE,
            attachment,
            &[ids.storage],
        )?,
        pages_object(
            ids.storage,
            STORAGE_MESSAGE_TYPES[0],
            storage,
            &storage_references,
        )?,
        pages_object(ids.marker, TEXTUAL_ATTACHMENT_MESSAGE_TYPE, marker, &[])?,
    ])
}

fn pages_object(
    identifier: u64,
    message_type: u32,
    message: impl Message,
    references: &[u64],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: message.encode_to_vec(),
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec();
    info.object_references = references.to_vec();
    Ok(object)
}

fn insert_footnote_reference(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    position: u32,
    reference_id: u64,
) -> Result<()> {
    let mut wire_budget = FootnoteGraphBudget::default();
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages body storage {storage_id} is missing"))
        })?;
        let message_index = unique_storage_message_index(object, storage_id)?;
        let original = &object.messages[message_index];
        let storage = decode_storage_for_mutation(
            original.data.as_slice(),
            storage_id,
            "Pages body storage",
            &mut wire_budget,
        )?;
        let table = footnote_table_payload(
            storage_id,
            original.data.as_slice(),
            &mut wire_budget,
        )?;
        let mut entries = footnote_table_entries_from_table(
            storage_id,
            table,
            &storage,
            &mut wire_budget,
        )?;
        if entries.iter().any(|entry| entry.index == position) {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} already has a footnote at UTF-16 index {position}"
            )));
        }
        require_text_boundary(storage_id, position, &storage.text)?;
        if utf16_unit_at(&storage.text, position) != Some(FOOTNOTE_ANCHOR_UNIT) {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} has no U+000E footnote anchor at UTF-16 index {position}"
            )));
        }
        let new_entry = tswp::object_attribute_table::ObjectAttribute {
            character_index: position,
            object: Some(reference(reference_id)),
        };
        reserve_footnote_entries(&mut entries, 1, "Pages body footnote table entries")?;
        let existing_raw_bytes = entries.iter().try_fold(0usize, |total, entry| {
            total.checked_add(entry.raw.len()).ok_or_else(|| {
                Error::InvalidFormat("Pages body footnote table raw size overflows usize".to_owned())
            })
        })?;
        let new_raw = new_entry.encode_to_vec();
        let raw_bytes = existing_raw_bytes
            .checked_add(new_raw.len())
            .ok_or_else(|| {
                Error::InvalidFormat("Pages body footnote table raw size overflows usize".to_owned())
            })?;
        if raw_bytes > MAX_FOOTNOTE_TABLE_BYTES {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed: raw_bytes,
                limit: MAX_FOOTNOTE_TABLE_BYTES,
            }));
        }
        entries.push(FootnoteTableEntry {
            index: position,
            reference_id,
            raw: new_raw,
        });
        entries.sort_by_key(|entry| entry.index);
        let mut encoded_entries = Vec::new();
        reserve_footnote_entries(
            &mut encoded_entries,
            entries.len(),
            "Pages body footnote table encoded entries",
        )?;
        encoded_entries.extend(entries.into_iter().map(|entry| entry.raw));
        let has_table = table.is_some();
        let table = match table {
            Some(table) => rewrite_repeated_length_delimited_fields(
                table,
                TABLE_ENTRIES_FIELD,
                &encoded_entries,
            )?,
            None => rewrite_repeated_length_delimited_fields(
                &[],
                TABLE_ENTRIES_FIELD,
                &encoded_entries,
            )?,
        };
        let data = patch_length_delimited_field(
            original.data.as_slice(),
            FOOTNOTE_TABLE_FIELD,
            has_table,
            Some(&table),
        )?;
        let verified = decode_storage_for_mutation(
            data.as_slice(),
            storage_id,
            "Pages body storage patch",
            &mut wire_budget,
        )?;
        if footnote_table_entries(storage_id, &data, &verified, &mut wire_budget)?
            .iter()
            .all(|entry| entry.reference_id != reference_id)
        {
            return Err(Error::InvalidFormat(
                "Pages body footnote table patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: original.type_,
                data,
            },
        )?;
        let references = &mut object.archive_info.message_infos[message_index].object_references;
        if references.contains(&reference_id) {
            return Err(Error::InvalidFormat(format!(
                "Pages body metadata already references footnote object {reference_id}"
            )));
        }
        references.push(reference_id);
        Ok(())
    })
}

/// Mutation-only compatibility route. The caller already owns a temporary
/// generated storage template; read/validation callers use the borrowed
/// [`FootnoteStorageProjection`] route below.
fn footnote_table_entries(
    storage_id: u64,
    data: &[u8],
    storage: &tswp::StorageArchive,
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<Vec<FootnoteTableEntry>> {
    let table = footnote_table_payload(storage_id, data, wire_budget)?;
    footnote_table_entries_from_table(storage_id, table, storage, wire_budget)
}

fn footnote_table_entries_from_projection(
    storage_id: u64,
    storage: &FootnoteStorageProjection<'_>,
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<Vec<FootnoteTableEntry>> {
    let Some(table) = storage.table_footnote() else {
        return Ok(Vec::new());
    };
    parse_footnote_table_entries(storage_id, table, storage.text(), wire_budget)
}

fn footnote_table_payload<'a>(
    storage_id: u64,
    data: &'a [u8],
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<Option<&'a [u8]>> {
    let (view, _) = wire_budget.scan(data, no_footnote_descent)?;
    for field in view.fields() {
        field.validate_canonical_framing()?;
    }
    let mut count = 0usize;
    let mut table = None;
    for field in view
        .fields()
        .filter(|field| field.number() == FOOTNOTE_TABLE_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} footnote table is not length-delimited"
            )));
        }
        count = count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("Pages footnote table count overflows usize".to_owned())
        })?;
        table = Some(field.canonical_payload()?);
    }
    if count > 1 {
        return Err(Error::InvalidFormat(format!(
            "Pages body storage {storage_id} contains {count} footnote tables"
        )));
    }
    Ok(table)
}

fn footnote_table_entries_from_table(
    storage_id: u64,
    table: Option<&[u8]>,
    storage: &tswp::StorageArchive,
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<Vec<FootnoteTableEntry>> {
    let Some(table) = table else {
        return if storage.table_footnote.is_none() {
            Ok(Vec::new())
        } else {
            Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} footnote table wire state is inconsistent"
            )))
        };
    };
    if storage.table_footnote.is_none() {
        return Err(Error::InvalidFormat(format!(
            "Pages body storage {storage_id} footnote table wire state is inconsistent"
        )));
    }
    parse_footnote_table_entries(storage_id, table, &storage.text, wire_budget)
}

fn parse_footnote_table_entries<F: AsRef<str>>(
    storage_id: u64,
    table: &[u8],
    text: &[F],
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<Vec<FootnoteTableEntry>> {
    let (view, _) = wire_budget.scan(table, no_footnote_descent)?;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(|error| {
            Error::InvalidFormat(format!(
                "Pages body storage {storage_id} footnote table has invalid framing: {error}"
            ))
        })?;
    }
    let mut entries = Vec::new();
    reserve_footnote_entries(
        &mut entries,
        table.len().min(MAX_BODY_FOOTNOTES),
        "Pages body footnote table entries",
    )?;
    for field in view
        .fields()
        .filter(|field| field.number() == TABLE_ENTRIES_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} footnote attachment is not length-delimited"
            )));
        }
        let next_count = entries.len().checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("Pages footnote entry count overflows usize".to_owned())
        })?;
        if next_count > MAX_BODY_FOOTNOTES {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed: next_count,
                limit: MAX_BODY_FOOTNOTES,
            }));
        }
        let raw = field.payload();
        wire_budget.charge_input(raw.len())?;
        let (.., report) = wire_budget.scan(raw, footnote_body_descent)?;
        reserve_footnote_entries(&mut entries, 1, "Pages body footnote table entries")?;
        let entry = pages_body_codec::decode_section_boundary(
            raw,
            FootnoteGraphBudget::body_decode_options(raw, report)?,
        )
        .map_err(|error| {
            Error::InvalidFormat(format!(
                "Pages body storage {storage_id} footnote attachment failed strict validation: {error}"
            ))
        })?;
        wire_budget.charge_codec(report)?;
        wire_budget.charge_work(raw.len())?;
        let mut raw_copy = Vec::new();
        raw_copy.try_reserve_exact(raw.len()).map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Pages body footnote table raw payload",
                amount: raw.len(),
            })
        })?;
        raw_copy.extend_from_slice(raw);
        entries.push(FootnoteTableEntry {
            index: entry.character_index(),
            reference_id: entry.section().map_or(0, |value| value.identifier().get()),
            raw: raw_copy,
        });
    }
    validate_footnote_table_entries(storage_id, &entries, text)?;
    Ok(entries)
}

fn validate_footnote_table_entries<F: AsRef<str>>(
    storage_id: u64,
    entries: &[FootnoteTableEntry],
    text: &[F],
) -> Result<()> {
    let text_length = text_utf16_len(text)?;
    let mut previous = None;
    for entry in entries {
        if entry.reference_id == 0 {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} has a zero footnote object identifier"
            )));
        }
        if previous.is_some_and(|index| index >= entry.index) {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} footnote positions are not strictly increasing"
            )));
        }
        require_text_boundary(storage_id, entry.index, text)?;
        if entry.index >= text_length
            || utf16_unit_at(text, entry.index) != Some(FOOTNOTE_ANCHOR_UNIT)
        {
            return Err(Error::InvalidFormat(format!(
                "Pages body storage {storage_id} footnote {} is not anchored to U+000E at UTF-16 index {}",
                entry.reference_id, entry.index
            )));
        }
        previous = Some(entry.index);
    }
    Ok(())
}

fn remove_unreferenced_footnote_graph(
    package: &mut IWorkPackage,
    graph: &BodyFootnoteGraph,
) -> Result<[u64; 3]> {
    let reference_id = graph.reference_id;
    remove_unreferenced_footnote_object(
        package,
        reference_id,
        &[FOOTNOTE_REFERENCE_MESSAGE_TYPE],
        "reference attachment",
    )?;
    remove_unreferenced_footnote_object(
        package,
        graph.storage_id,
        STORAGE_MESSAGE_TYPES,
        "storage",
    )?;
    remove_unreferenced_footnote_object(
        package,
        graph.marker_id,
        &[TEXTUAL_ATTACHMENT_MESSAGE_TYPE],
        "marker attachment",
    )?;
    Ok([reference_id, graph.storage_id, graph.marker_id])
}

fn remove_unreferenced_footnote_object(
    package: &mut IWorkPackage,
    identifier: u64,
    message_types: &[u32],
    label: &str,
) -> Result<()> {
    if package_references_object(package, identifier)? {
        return Err(Error::InvalidFormat(format!(
            "Pages footnote {label} object {identifier} remains referenced after body-anchor deletion"
        )));
    }
    remove_component_external_references_to_object(package, DOCUMENT_OBJECT_ID, identifier)?;
    if let Some(component) = component_identifier_for_object_uuid(package, identifier)? {
        remove_component_object_uuids(package, component, &[identifier])?;
    }
    let archive_name = find_object_archive(package, identifier)?;
    package.update_archive(&archive_name, |archive| {
        let object = archive.remove_object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages footnote {label} object {identifier} is missing"
            ))
        })?;
        object_message_data_of_types(&object, message_types, &format!("Pages footnote {label}"))?;
        Ok(())
    })
}

fn storage_at(
    package: &IWorkPackage,
    storage_id: u64,
    label: &str,
) -> Result<(String, tswp::StorageArchive)> {
    // Mutation/template callers intentionally retain the generated Prost
    // value while the source-preserving wire rewrite is staged separately.
    let archive_name = find_object_archive(package, storage_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive
        .object(storage_id)
        .ok_or_else(|| Error::InvalidFormat(format!("{label} storage {storage_id} is missing")))?;
    let message_index = unique_storage_message_index(object, storage_id)?;
    let source = object.messages[message_index].data.as_slice();
    let mut wire_budget = FootnoteGraphBudget::default();
    let storage = decode_storage_for_mutation(source, storage_id, label, &mut wire_budget)?;
    Ok((archive_name, storage))
}

fn with_storage_projection<T, F>(
    package: &IWorkPackage,
    storage_id: u64,
    label: &str,
    wire_budget: &mut FootnoteGraphBudget,
    read: F,
) -> Result<T>
where
    F: for<'source> FnOnce(
        &FootnoteStorageProjection<'source>,
        &mut FootnoteGraphBudget,
    ) -> Result<T>,
{
    let archive_name = find_object_archive(package, storage_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive
        .object(storage_id)
        .ok_or_else(|| Error::InvalidFormat(format!("{label} storage {storage_id} is missing")))?;
    let message_index = unique_storage_message_index(object, storage_id)?;
    let source = object.messages[message_index].data.as_slice();
    let storage = FootnoteStorageProjection::decode(source, storage_id, label, wire_budget)?;
    // The archive is an owned cache clone, so its borrowed projection cannot
    // escape this function. Keep all projection consumers inside this scope.
    read(&storage, wire_budget)
}

/// Admit a legacy storage payload for the mutation/template compatibility path.
///
/// This is deliberately separate from [`FootnoteStorageProjection`]: callers
/// retain the original payload as the mutation authority, while Prost supplies
/// only the temporary template value needed by existing constructors/rewrites.
fn decode_storage_for_mutation(
    source: &[u8],
    storage_id: u64,
    label: &str,
    wire_budget: &mut FootnoteGraphBudget,
) -> Result<tswp::StorageArchive> {
    let validation = litchi_iwa_text_wire::validate_storage_with_limits(
        source,
        litchi_iwa_text_wire::RewriteLimits::default(),
    )
    .map_err(|error| map_storage_wire_error(storage_id, label, error))?;

    // The preflight does not allocate semantic text or generated fields. All
    // aggregate charges therefore happen before Prost can grow its repeated
    // text/table vectors. The extra source-byte charge covers the generated
    // compatibility projection's bounded parse/allocation pass.
    wire_budget.charge_input(source.len())?;
    wire_budget.charge_fields(validation.fields())?;
    wire_budget.charge_work(validation.validation_work())?;
    wire_budget.charge_work(source.len())?;

    tswp::StorageArchive::decode(source).map_err(|error| {
        Error::InvalidFormat(format!(
            "Pages {label} storage {storage_id} failed bounded compatibility decode: {error}"
        ))
    })
}

fn map_storage_wire_error(
    storage_id: u64,
    label: &str,
    error: litchi_iwa_text_wire::RewriteError,
) -> Error {
    match error {
        litchi_iwa_text_wire::RewriteError::Allocation { resource, amount } => {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation { resource, amount })
        },
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind: storage_wire_limit_kind(resource),
            observed,
            limit,
        }),
        error => Error::InvalidFormat(format!(
            "Pages {label} storage {storage_id} failed bounded wire preflight: {error}"
        )),
    }
}

fn storage_wire_limit_kind(resource: &str) -> LimitKind {
    match resource {
        "message bytes" | "text bytes" | "aggregate nested scan bytes" => LimitKind::InputBytes,
        "fields" | "text fragments" | "table entries" | "object references" => LimitKind::Fields,
        "nesting" => LimitKind::Nesting,
        _ => LimitKind::RewriteWork,
    }
}

fn unique_storage_message_index(object: &ArchiveObject, storage_id: u64) -> Result<usize> {
    let mut index = None;
    for (candidate, message) in object.messages.iter().enumerate() {
        if !STORAGE_MESSAGE_TYPES.contains(&message.type_) {
            continue;
        }
        if index.replace(candidate).is_some() {
            return Err(Error::InvalidFormat(format!(
                "Pages text storage {storage_id} must have exactly one writable payload"
            )));
        }
    }
    let Some(index) = index else {
        return Err(Error::InvalidFormat(format!(
            "Pages text storage {storage_id} must have exactly one writable payload"
        )));
    };
    Ok(index)
}

fn object_message_data<'a>(
    object: &'a ArchiveObject,
    message_type: u32,
    label: &str,
) -> Result<&'a [u8]> {
    let mut selected = None;
    for message in &object.messages {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(message).is_some() {
            return Err(Error::InvalidFormat(format!(
                "{label} must contain exactly one message type {message_type}"
            )));
        }
    }
    let Some(message) = selected else {
        return Err(Error::InvalidFormat(format!(
            "{label} must contain exactly one message type {message_type}"
        )));
    };
    Ok(message.data.as_slice())
}

fn object_message_data_of_types<'a>(
    object: &'a ArchiveObject,
    message_types: &[u32],
    label: &str,
) -> Result<&'a [u8]> {
    let mut selected = None;
    for message in &object.messages {
        if !message_types.contains(&message.type_) {
            continue;
        }
        if selected.replace(message).is_some() {
            return Err(Error::InvalidFormat(format!(
                "{label} must contain exactly one supported message payload"
            )));
        }
    }
    let Some(message) = selected else {
        return Err(Error::InvalidFormat(format!(
            "{label} must contain exactly one supported message payload"
        )));
    };
    Ok(message.data.as_slice())
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_type: None,
        deprecated_is_external: None,
    }
}

fn validate_footnote_text(text: &str) -> Result<()> {
    if text.contains(FOOTNOTE_ANCHOR) || text.contains(FOOTNOTE_MARK) {
        return Err(Error::ParseError(
            "Pages footnote text cannot contain native footnote-anchor or attachment markers"
                .to_owned(),
        ));
    }
    Ok(())
}

fn require_text_boundary<F: AsRef<str>>(storage_id: u64, position: u32, text: &[F]) -> Result<()> {
    let mut current = 0u32;
    if position == current {
        return Ok(());
    }
    for fragment in text {
        for character in fragment.as_ref().chars() {
            current = current
                .checked_add(character.len_utf16() as u32)
                .ok_or_else(|| {
                    Error::InvalidFormat("Pages text UTF-16 length overflow".to_owned())
                })?;
            if current == position {
                return Ok(());
            }
            if current > position {
                break;
            }
        }
    }
    Err(Error::InvalidFormat(format!(
        "UTF-16 index {position} is not a scalar boundary in Pages storage {storage_id}"
    )))
}

fn text_utf16_len<F: AsRef<str>>(text: &[F]) -> Result<u32> {
    text.iter().try_fold(0u32, |total, fragment| {
        fragment
            .as_ref()
            .chars()
            .try_fold(total, |total, character| {
                total
                    .checked_add(character.len_utf16() as u32)
                    .ok_or_else(|| {
                        Error::InvalidFormat("Pages text UTF-16 length overflow".to_owned())
                    })
            })
    })
}

fn utf16_unit_at<F: AsRef<str>>(text: &[F], requested: u32) -> Option<u16> {
    let mut index = 0u32;
    for fragment in text {
        for unit in fragment.as_ref().encode_utf16() {
            if index == requested {
                return Some(unit);
            }
            index = index.checked_add(1)?;
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_pages::Package as PagesPackage;
    use litchi_pages::footnote::body::Footnote;

    fn rewrite_body_footnote_text_via_pages(
        editor: &mut PagesEditor,
        selector: Selector,
        text: &str,
    ) -> Result<Footnote> {
        let package = PagesPackage::from_bytes(&editor.to_bytes()?)
            .map_err(|error| Error::InvalidFormat(format!("Pages footnote package: {error}")))?;
        let mut edit = package
            .edit_body_footnote_text(selector)
            .map_err(|error| Error::InvalidFormat(format!("Pages footnote edit: {error}")))?;
        edit.set(text)
            .map_err(|error| Error::InvalidFormat(format!("Pages footnote edit: {error}")))?;
        let commit = edit
            .commit()
            .map_err(|error| Error::InvalidFormat(format!("Pages footnote edit: {error}")))?;
        let mut bytes = Vec::new();
        commit.package().write_to(&mut bytes).map_err(|error| {
            Error::InvalidFormat(format!("Pages footnote package write failed: {error}"))
        })?;
        let reopened = PagesEditor::from_bytes(&bytes)?;
        let updated = body_footnote_by_selector(&reopened, selector)?.footnote;
        *editor = reopened;
        Ok(updated)
    }

    #[test]
    fn footnote_collection_respects_wire_field_limit_before_reserving() {
        let limits = WireLimits::default().with_fields(1).unwrap();
        let mut graphs = Vec::<()>::new();
        let error =
            reserve_footnote_collection(&mut graphs, 2, limits, "Pages body footnote graphs")
                .unwrap_err();
        assert!(matches!(
            error,
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed: 2,
                limit: 1,
            })
        ));
        assert!(graphs.is_empty());
    }

    #[test]
    fn footnote_cleanup_identifier_count_checks_overflow() {
        let error = footnote_cleanup_identifier_count(usize::MAX).unwrap_err();
        assert!(matches!(
            error,
            Error::InvalidFormat(message)
                if message == "Pages removed footnote identifier count overflows usize"
        ));
    }

    #[test]
    fn body_footnote_crud_round_trips_and_restores_a_source_document() {
        let mut editor = PagesEditor::create_with_text("A😀B").unwrap();
        let baseline = editor.to_bytes().unwrap();
        let note = editor
            .insert_body_footnote(Position::from_utf16_index(3).unwrap(), "Initial note")
            .unwrap();
        assert_eq!(editor.body_text().unwrap(), "A😀\u{e}B");
        assert_eq!(note.position, Position::from_utf16_index(3).unwrap());
        assert_eq!(note.text.as_ref(), "Initial note");
        assert_eq!(note.custom_mark, None);
        assert_eq!(editor.body_footnotes().unwrap(), vec![note.clone()]);

        let reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.body_footnotes().unwrap(), vec![note.clone()]);

        let updated =
            rewrite_body_footnote_text_via_pages(&mut editor, Selector::Index(0), "Updated note")
                .unwrap();
        assert_eq!(updated.text.as_ref(), "Updated note");
        assert_eq!(updated.position, note.position);

        let removed = editor
            .remove_body_footnote(Selector::At(note.position))
            .unwrap();
        assert_eq!(removed, updated);
        assert_eq!(editor.body_text().unwrap(), "A😀B");
        assert!(editor.body_footnotes().unwrap().is_empty());
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn removing_adjacent_body_footnote_is_atomic_and_tracks_identity() {
        let mut editor = PagesEditor::create_with_text("AB").unwrap();
        let first = editor
            .insert_body_footnote(Position::from_utf16_index(1).unwrap(), "First")
            .unwrap();
        let second = editor
            .insert_body_footnote(Position::from_utf16_index(2).unwrap(), "Second")
            .unwrap();

        let removed = editor
            .remove_body_footnote(Selector::At(first.position))
            .unwrap();

        assert_eq!(removed, first);
        assert_eq!(editor.body_text().unwrap(), "A\u{e}B");
        assert_eq!(
            editor.body_footnotes().unwrap(),
            vec![Footnote {
                position: Position::from_utf16_index(1).unwrap(),
                text: second.text,
                custom_mark: second.custom_mark,
            }]
        );
    }

    #[test]
    fn ordinary_body_replacement_reclaims_deleted_footnote_graphs() {
        let mut editor = PagesEditor::create_with_text("AB").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(1).unwrap(), "First")
            .unwrap();
        let first_reference_id =
            body_footnote_graphs(editor.package(), editor.body_storage_id.get()).unwrap()[0]
                .reference_id;
        let second = editor
            .insert_body_footnote(Position::from_utf16_index(3).unwrap(), "Second")
            .unwrap();
        assert_eq!(editor.body_text().unwrap(), "A\u{e}B\u{e}");

        editor.replace_body_text(1..2, "").unwrap();
        assert_eq!(editor.body_text().unwrap(), "AB\u{e}");
        assert_eq!(
            editor.body_footnotes().unwrap(),
            vec![Footnote {
                position: Position::from_utf16_index(2).unwrap(),
                text: "Second".into(),
                custom_mark: None,
            }]
        );
        assert!(find_object_archive(editor.package(), first_reference_id).is_err());
        assert_eq!(second.position, Position::from_utf16_index(3).unwrap());
    }

    #[test]
    fn footnote_text_rejects_native_structural_markers_transactionally() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let baseline = editor.to_bytes().unwrap();
        assert!(
            editor
                .insert_body_footnote(Position::ZERO, "Invalid\u{e}")
                .is_err()
        );
        assert!(
            editor
                .insert_body_footnote(Position::ZERO, "Invalid\u{fffc}")
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn native_footnote_reference_without_a_super_payload_is_supported() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        let footnote = editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let reference_id = body_footnote_graphs(editor.package(), editor.body_storage_id.get())
            .unwrap()[0]
            .reference_id;
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, reference_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(reference_id).unwrap();
                let message = &object.messages[0];
                let data = patch_length_delimited_field(
                    message.data.as_slice(),
                    FOOTNOTE_SUPER_FIELD,
                    false,
                    None,
                )?;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: FOOTNOTE_REFERENCE_MESSAGE_TYPE,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let parsed = PagesEditor::from_package(package).unwrap();
        assert_eq!(parsed.body_footnotes().unwrap(), vec![footnote]);
    }

    #[test]
    fn footnote_marker_unknown_fields_survive_an_atomic_text_rewrite() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let graph = body_footnote_graphs(editor.package(), editor.body_storage_id.get())
            .unwrap()
            .pop()
            .unwrap();
        let unknown = [0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, b'o', b'p', b'a'];
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, graph.marker_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(graph.marker_id).unwrap();
                let message = &object.messages[0];
                let mut data = message.data.clone();
                data.extend_from_slice(&unknown);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut edited = PagesEditor::from_package(package).unwrap();
        assert_eq!(edited.body_footnotes().unwrap()[0].text.as_ref(), "Native");
        rewrite_body_footnote_text_via_pages(&mut edited, Selector::Index(0), "Updated").unwrap();
        let marker_archive = find_object_archive(edited.package(), graph.marker_id).unwrap();
        let marker_archive_data = edited.package().archive(&marker_archive).unwrap();
        let marker = marker_archive_data.object(graph.marker_id).unwrap();
        assert!(marker.messages[0].data.ends_with(&unknown));
    }

    #[test]
    fn malformed_footnote_marker_rewrite_is_failure_atomic() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let marker_id = body_footnote_graphs(editor.package(), editor.body_storage_id.get())
            .unwrap()[0]
            .marker_id;
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, marker_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(marker_id).unwrap();
                let message = &object.messages[0];
                let mut data = message.data.clone();
                data.extend_from_slice(&[0xa0, 0x06, 0x80]);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        let mut malformed = PagesEditor::from_package(package).unwrap();
        let baseline = malformed.to_bytes().unwrap();
        assert!(
            rewrite_body_footnote_text_via_pages(&mut malformed, Selector::Index(0), "Updated")
                .is_err()
        );
        assert_eq!(malformed.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn footnote_reference_unknown_fields_survive_an_atomic_text_rewrite() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let reference_id = body_footnote_graphs(editor.package(), editor.body_storage_id.get())
            .unwrap()[0]
            .reference_id;
        let unknown = [0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, b'o', b'p', b'a'];
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, reference_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(reference_id).unwrap();
                let message = &object.messages[0];
                let mut data = message.data.clone();
                data.extend_from_slice(&unknown);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut edited = PagesEditor::from_package(package).unwrap();
        assert_eq!(edited.body_footnotes().unwrap()[0].text.as_ref(), "Native");
        rewrite_body_footnote_text_via_pages(&mut edited, Selector::Index(0), "Updated").unwrap();
        let reference_archive = find_object_archive(edited.package(), reference_id).unwrap();
        let reference_archive_data = edited.package().archive(&reference_archive).unwrap();
        let reference = reference_archive_data.object(reference_id).unwrap();
        assert!(reference.messages[0].data.ends_with(&unknown));
    }

    #[test]
    fn malformed_footnote_reference_rewrite_is_failure_atomic() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let reference_id = body_footnote_graphs(editor.package(), editor.body_storage_id.get())
            .unwrap()[0]
            .reference_id;
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, reference_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(reference_id).unwrap();
                let message = &object.messages[0];
                let mut data = message.data.clone();
                data.extend_from_slice(&[0xa0, 0x06, 0x80]);
                object.replace_message(
                    0,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut malformed = PagesEditor::from_package(package).unwrap();
        let baseline = malformed.to_bytes().unwrap();
        assert!(
            rewrite_body_footnote_text_via_pages(&mut malformed, Selector::Index(0), "Updated")
                .is_err()
        );
        assert_eq!(malformed.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn footnote_body_attachment_unknown_fields_survive_an_atomic_text_rewrite() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let body_storage_id = editor.body_storage_id.get();
        let unknown = [0xa0, 0x06, 0x01, 0xaa, 0x06, 0x03, b'o', b'p', b'a'];
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, body_storage_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(body_storage_id).unwrap();
                let message_index = unique_storage_message_index(object, body_storage_id)?;
                let message = &object.messages[message_index];
                let tables = repeated_length_delimited_payloads(
                    message.data.as_slice(),
                    FOOTNOTE_TABLE_FIELD,
                )?;
                let [table] = tables.as_slice() else {
                    return Err(Error::InvalidFormat(
                        "Pages body footnote test requires one attachment table".to_owned(),
                    ));
                };
                let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
                let [entry] = entries.as_slice() else {
                    return Err(Error::InvalidFormat(
                        "Pages body footnote test requires one attachment entry".to_owned(),
                    ));
                };
                let mut replacement = entry.to_vec();
                replacement.extend_from_slice(&unknown);
                let table = rewrite_repeated_length_delimited_fields(
                    table,
                    TABLE_ENTRIES_FIELD,
                    &[replacement],
                )?;
                let data = patch_length_delimited_field(
                    message.data.as_slice(),
                    FOOTNOTE_TABLE_FIELD,
                    true,
                    Some(&table),
                )?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut edited = PagesEditor::from_package(package).unwrap();
        assert_eq!(edited.body_footnotes().unwrap()[0].text.as_ref(), "Native");
        rewrite_body_footnote_text_via_pages(&mut edited, Selector::Index(0), "Updated").unwrap();
        let body_archive = find_object_archive(edited.package(), body_storage_id).unwrap();
        let body_archive_data = edited.package().archive(&body_archive).unwrap();
        let body = body_archive_data.object(body_storage_id).unwrap();
        let message_index = unique_storage_message_index(body, body_storage_id).unwrap();
        let tables = repeated_length_delimited_payloads(
            body.messages[message_index].data.as_slice(),
            FOOTNOTE_TABLE_FIELD,
        )
        .unwrap();
        let [table] = tables.as_slice() else {
            panic!("updated body has one attachment table");
        };
        let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD).unwrap();
        let [entry] = entries.as_slice() else {
            panic!("updated body has one attachment entry");
        };
        assert!(entry.ends_with(&unknown));
    }

    #[test]
    fn malformed_footnote_body_attachment_rewrite_is_failure_atomic() {
        let mut editor = PagesEditor::create_with_text("Body").unwrap();
        editor
            .insert_body_footnote(Position::from_utf16_index(4).unwrap(), "Native")
            .unwrap();
        let body_storage_id = editor.body_storage_id.get();
        let mut package = editor.package().clone();
        let archive_name = find_object_archive(&package, body_storage_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(body_storage_id).unwrap();
                let message_index = unique_storage_message_index(object, body_storage_id)?;
                let message = &object.messages[message_index];
                let tables = repeated_length_delimited_payloads(
                    message.data.as_slice(),
                    FOOTNOTE_TABLE_FIELD,
                )?;
                let [table] = tables.as_slice() else {
                    return Err(Error::InvalidFormat(
                        "Pages body footnote test requires one attachment table".to_owned(),
                    ));
                };
                let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
                let [entry] = entries.as_slice() else {
                    return Err(Error::InvalidFormat(
                        "Pages body footnote test requires one attachment entry".to_owned(),
                    ));
                };
                let mut replacement = entry.to_vec();
                replacement.extend_from_slice(&[0xa0, 0x06, 0x80]);
                let table = rewrite_repeated_length_delimited_fields(
                    table,
                    TABLE_ENTRIES_FIELD,
                    &[replacement],
                )?;
                let data = patch_length_delimited_field(
                    message.data.as_slice(),
                    FOOTNOTE_TABLE_FIELD,
                    true,
                    Some(&table),
                )?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();

        let mut malformed = PagesEditor::from_package(package).unwrap();
        let baseline = malformed.to_bytes().unwrap();
        assert!(
            rewrite_body_footnote_text_via_pages(&mut malformed, Selector::Index(0), "Updated")
                .is_err()
        );
        assert_eq!(malformed.to_bytes().unwrap(), baseline);
    }
}
