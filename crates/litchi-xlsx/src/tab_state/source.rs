//! Guarded tab-state publication over an immutable positional source.
//!
//! [`SourceBackedWorkbook::edit_tab_state`] stages tab visibility and the
//! active tab; [`Edit::commit`] validates the edit and returns a [`Commit`]
//! holding the new workbook, whose output is the splice the former stream
//! publisher wrote: visibility-only changes overlay the workbook part, an
//! active-tab transition additionally overlays the old and new worksheet
//! parts so `tabSelected` stays synchronized, and every other member is
//! raw-copied. An exact no-op returns the source workbook; a changed edit of a
//! signed source is refused at commit. The workbook's catalog accessors read
//! the committed catalog without materializing the output.
//! [`SourceBackedWorkbook::apply_tab_state`] replays a [`Patch`] onto a
//! workbook holding its before state (record 0360).

use std::io::Write;
use std::sync::Arc;

use litchi_core::Selector as CoreSelector;
use litchi_opc::SourceBackedPackage;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::snapshot::Snapshot;
use crate::error::{Result, TabEditBlock, allocation, invalid};
use crate::{Selector, SourceBackedWorkbook, Visibility, WorksheetKind, raw};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

const MAX_EDIT_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_EDIT_XML_DEPTH: usize = 128;
const MAX_EDIT_XML_EVENTS: usize = 1_000_000;
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

crate::workbook::source::deferred::source_family!(
    "tab-state",
    super::Patch,
    own_limits,
    own_diagnostics
);

crate::record::limit_kinds!("source-backed tab-state", TabState {
    /// XML bytes audited by one tab-state edit.
    XmlBytes => "tab-state XML bytes",
    /// Element depth of one audited XML part.
    Depth => "tab-state XML depth",
    /// XML events visited by one audited XML part.
    Events => "tab-state XML events",
});

/// Resource limits for one source-backed tab-state edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_xml_bytes: usize,
    max_depth: usize,
    max_events: usize,
}

impl Limits {
    /// Construct the default tab-state edit limits.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_xml_bytes: MAX_EDIT_XML_BYTES,
            max_depth: MAX_EDIT_XML_DEPTH,
            max_events: MAX_EDIT_XML_EVENTS,
        }
    }

    /// Maximum bytes audited in one workbook or worksheet XML part.
    #[must_use]
    pub const fn max_xml_bytes(self) -> usize {
        self.max_xml_bytes
    }

    /// Maximum element depth audited in one XML part.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Maximum XML events audited in one XML part.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.max_events
    }

    /// Set the maximum XML bytes audited in one part.
    pub fn with_max_xml_bytes(mut self, value: usize) -> Result<Self> {
        self.max_xml_bytes = checked_limit(LimitKind::XmlBytes, value, MAX_EDIT_XML_BYTES)?;
        Ok(self)
    }

    /// Set the maximum element depth audited in one part.
    pub fn with_max_depth(mut self, value: usize) -> Result<Self> {
        self.max_depth = checked_limit(LimitKind::Depth, value, MAX_EDIT_XML_DEPTH)?;
        Ok(self)
    }

    /// Set the maximum XML events audited in one part.
    pub fn with_max_events(mut self, value: usize) -> Result<Self> {
        self.max_events = checked_limit(LimitKind::Events, value, MAX_EDIT_XML_EVENTS)?;
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::new()
    }
}

fn checked_limit(kind: LimitKind, value: usize, maximum: usize) -> Result<usize> {
    if value == 0 {
        return Err(Error::InvalidLimit {
            kind: kind.into(),
            value: 0,
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        });
    }
    if value > maximum {
        return Err(Error::InvalidLimit {
            kind: kind.into(),
            value: u64::try_from(value).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        });
    }
    Ok(value)
}

/// Options for one source-backed tab-state edit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    /// Bound the XML audit by `limits` instead of the default policy.
    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = Some(limits);
        self
    }

    /// Return the configured limits, if an override was supplied.
    #[must_use]
    pub const fn limits_value(self) -> Option<Limits> {
        self.limits
    }
}

/// What a committed tab-state change replaced.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    touched_parts: usize,
}

impl Diagnostics {
    /// Whether the change replaced any part.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.touched_parts != 0
    }

    /// The replaced parts: the workbook, plus the old and new active
    /// worksheet views for an active-tab transition.
    #[must_use]
    pub const fn touched_parts(&self) -> usize {
        self.touched_parts
    }

    /// The replaced worksheet view parts.
    #[must_use]
    pub const fn touched_worksheets(&self) -> usize {
        self.touched_parts.saturating_sub(1)
    }
}

impl crate::workbook::source::deferred::Report for super::Patch {
    type Diagnostics = Diagnostics;

    fn diagnostics(&self) -> Diagnostics {
        Diagnostics {
            touched_parts: if self.is_empty() {
                0
            } else {
                1 + self.after().touched().len()
            },
        }
    }
}

impl Patch {
    /// The exact workbook tab state the patch requires.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        self.0.before()
    }

    /// The exact workbook tab state the patch produces.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        self.0.after()
    }
}

impl crate::workbook::source::deferred::Change for super::Patch {
    fn is_noop(&self) -> bool {
        self.is_empty()
    }

    fn inverse(&self) -> Self {
        super::Patch::inverse(self)
    }

    fn check(&self, package: &SourceBackedPackage) -> Result<()> {
        // A replay compares content: the workbook catalog, the sheet graph and
        // the touched worksheet views must hold the before state.
        if !self.before().matches_source_backed(package, None)? {
            return Err(Error::PatchConflict {
                part: self.before().workbook_part_name().to_string(),
            });
        }
        Ok(())
    }

    fn write(&self, package: &SourceBackedPackage, writer: &mut dyn Write) -> Result<()> {
        package.check_execution()?;
        crate::workbook::source::deferred::Change::check(self, package)?;
        let target = self.after();
        let mut overlays = Vec::new();
        overlays
            .try_reserve_exact(1 + target.touched().len())
            .map_err(|source| allocation("tab-state source overlay plan", source))?;
        overlays.push((
            target.workbook_part_name().clone(),
            target.workbook_source_arc()?,
        ));
        for part in target.touched() {
            overlays.push((part.part.uri.clone(), part.part.bytes.detached_arc()?));
        }
        package
            .clone()
            .write_part_overlays_shared_to_stream(writer, overlays)?;
        Ok(())
    }

    fn write_authenticated(
        &self,
        package: &SourceBackedPackage,
        writer: &mut dyn Write,
    ) -> Result<()> {
        package.check_execution()?;
        let target = self.after();
        let mut overlays = Vec::new();
        overlays
            .try_reserve_exact(1 + target.touched().len())
            .map_err(|source| allocation("tab-state source overlay plan", source))?;
        overlays.push((
            target.workbook_part_name().clone(),
            target.workbook_source_arc()?,
        ));
        for part in target.touched() {
            overlays.push((part.part.uri.clone(), part.part.bytes.detached_arc()?));
        }
        package
            .clone()
            .write_part_overlays_shared_to_stream(writer, overlays)?;
        Ok(())
    }

    fn catalog(
        &self,
        package: &SourceBackedPackage,
    ) -> Option<crate::workbook::source::deferred::Catalog> {
        let catalog = self.after().catalog_proof().ok()?;
        (package.read_limits() == *catalog.limits)
            .then(|| self.after().deferred_catalog())
            .flatten()
    }

    fn catalog_xml(&self) -> Option<&[u8]> {
        Some(self.after().workbook_xml())
    }
}

impl SourceBackedWorkbook {
    /// Begin an isolated edit of the visibility and active tab of existing
    /// workbook tabs.
    ///
    /// # Errors
    ///
    /// Returns an error when the workbook catalog is invalid or the source
    /// cannot be read.
    pub fn edit_tab_state(&self) -> Result<Edit> {
        self.edit_tab_state_with_options(EditOptions::default())
    }

    /// Begin an isolated tab-state edit with explicit XML audit limits.
    pub fn edit_tab_state_with_options(&self, options: EditOptions) -> Result<Edit> {
        let view = self.view()?;
        let package = view.edit_package()?;
        let source_owner = view.tab_state_source_owner();
        let snapshot_cache = view.tab_state_snapshot();
        let before = match snapshot_cache.get() {
            Some(snapshot) => {
                // The source-owner loader performs this freshness check while
                // creating the closure. Keep the same source fence when a
                // cached immutable closure is reused.
                package.source_version()?;
                (**snapshot).clone()
            },
            None => {
                let snapshot =
                    Snapshot::load_source_backed_from_owner(package, source_owner, Arc::new(()))?;
                // The first edit of an owner keeps its snapshot to itself; a
                // repeated edit caches one shared copy.
                if view.tab_state_snapshot_should_cache() {
                    (**snapshot_cache.get_or_init(|| Arc::new(snapshot))).clone()
                } else {
                    snapshot
                }
            },
        };
        Edit::new(self.clone(), before, options.limits.unwrap_or_default())
    }

    /// Replay a tab-state patch onto this workbook.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the workbook does not hold the
    /// patch's before state, and the signed-source refusal for a changed
    /// patch on a signed source.
    pub fn apply_tab_state(&self, patch: &Patch) -> Result<Commit> {
        apply_change(self, patch)
    }
}

/// An isolated visibility and active-tab edit over one exact source closure
/// of a [`SourceBackedWorkbook`]. Dropping it discards it.
pub struct Edit {
    workbook: SourceBackedWorkbook,
    before: Snapshot,
    limits: Limits,
    staged_visibility: Vec<Visibility>,
    requested_active: Option<usize>,
}

impl Edit {
    fn new(workbook: SourceBackedWorkbook, before: Snapshot, limits: Limits) -> Result<Self> {
        let mut staged_visibility = Vec::new();
        staged_visibility
            .try_reserve_exact(before.tabs().len())
            .map_err(|source| allocation("staged tab visibility", source))?;
        staged_visibility.extend(before.tabs().iter().map(|tab| tab.visibility().clone()));
        Ok(Self {
            workbook,
            before,
            limits,
            staged_visibility,
            requested_active: None,
        })
    }

    /// Exact source state captured when this edit began.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Show an existing tab.
    pub fn show<'s>(&mut self, selector: impl Into<Selector<'s>>) -> Result<bool> {
        self.set_visibility(selector.into(), Visibility::Visible)
    }

    /// Hide an existing tab while retaining it in Excel's Unhide dialog.
    pub fn hide<'s>(&mut self, selector: impl Into<Selector<'s>>) -> Result<bool> {
        self.set_visibility(selector.into(), Visibility::Hidden)
    }

    /// Hide an existing tab from Excel's ordinary Unhide dialog.
    pub fn very_hide<'s>(&mut self, selector: impl Into<Selector<'s>>) -> Result<bool> {
        self.set_visibility(selector.into(), Visibility::VeryHidden)
    }

    /// Make one visible existing worksheet the active tab.
    pub fn activate<'s>(&mut self, selector: impl Into<Selector<'s>>) -> Result<bool> {
        let position = self.resolve(selector.into())?;
        self.require_known_owner(position)?;
        if !self.staged_visibility[position].is_visible() {
            return Err(self.block(position, TabEditBlock::NotVisible));
        }
        let current = self
            .requested_active
            .unwrap_or(self.before.active_position());
        let changed = position != current;
        self.requested_active = (position != self.before.active_position()).then_some(position);
        Ok(changed)
    }

    /// Discard every staged change in place.
    pub fn reset(&mut self) {
        for (staged, before) in self.staged_visibility.iter_mut().zip(self.before.tabs()) {
            *staged = before.visibility().clone();
        }
        self.requested_active = None;
    }

    /// Validate, bind the exact touched closure, and freeze the edit.
    pub fn commit(self) -> Result<Commit> {
        let package = self.workbook.edit_package()?;
        package.check_execution()?;
        // A committed patch may outlive the source owner: capture the
        // catalog proof the snapshot has so far only linked.
        self.before.catalog_proof()?;
        let visibility_changed = self
            .staged_visibility
            .iter()
            .zip(self.before.tabs())
            .enumerate()
            .filter_map(|(position, (after, before))| {
                (after != before.visibility()).then_some(position)
            })
            .collect::<Vec<_>>();
        if visibility_changed.is_empty()
            && self
                .requested_active
                .is_none_or(|position| position == self.before.active_position())
        {
            // A no-op patch may outlive both the edit and its source owner.
            // Capture the exact graph proof before the only owner reference
            // becomes weak, so detached fallback replay remains infallible
            // for the same admitted source closure.
            self.before.materialize_source_graph()?;
            let patch = super::Patch::new(self.before.clone(), self.before.clone());
            return commit_change(&self.workbook, patch);
        }

        if !self.staged_visibility.iter().any(Visibility::is_visible) {
            let position = visibility_changed.first().copied().unwrap_or(0);
            return Err(self.block(position, TabEditBlock::LastVisibleTab));
        }
        let active = if let Some(requested) = self.requested_active {
            if !self.staged_visibility[requested].is_visible() {
                return Err(self.block(requested, TabEditBlock::NotVisible));
            }
            requested
        } else if self.staged_visibility[self.before.active_position()].is_visible() {
            self.before.active_position()
        } else {
            let len = self.staged_visibility.len();
            (1..=len)
                .map(|offset| (self.before.active_position() + offset) % len)
                .find(|position| self.staged_visibility[*position].is_visible())
                .ok_or_else(|| {
                    self.block(self.before.active_position(), TabEditBlock::LastVisibleTab)
                })?
        };
        if active > raw::catalog_edit::MAX_ACTIVE_TAB {
            return Err(self.block(active, TabEditBlock::ActiveTabLimit));
        }

        let audit = audit_xml(self.before.workbook_xml(), XmlOwner::Workbook, self.limits)?;
        let context = visibility_changed.first().copied().unwrap_or(active);
        if audit.protected_structure {
            return Err(self.block(context, TabEditBlock::ProtectedWorkbook));
        }
        if audit.alternate_content {
            return Err(self.block(context, TabEditBlock::MarkupCompatibility));
        }

        let active_changed = active != self.before.active_position();
        let before = if active_changed {
            self.before
                .with_source_touched(package, &[self.before.active_position(), active])?
        } else {
            self.before
        };
        for part in before.touched() {
            let audit = audit_xml(part.part.bytes.as_bytes(), XmlOwner::Sheet, self.limits)?;
            if audit.alternate_content {
                return Err(Error::TabEditBlocked {
                    sheet: before.tabs()[part.position].name().to_owned(),
                    position: part.position,
                    reason: TabEditBlock::MarkupCompatibility,
                });
            }
        }

        let mut tabs = Vec::new();
        tabs.try_reserve_exact(visibility_changed.len())
            .map_err(|source| allocation("tab-state workbook edit plan", source))?;
        let source_catalog = raw::parse_catalog(before.workbook_xml(), &package.read_limits())?;
        for position in visibility_changed {
            let binding = before.binding(position)?;
            let state = raw_visibility(&self.staged_visibility[position]).ok_or_else(|| {
                Error::TabEditBlocked {
                    sheet: before.tabs()[position].name().to_owned(),
                    position,
                    reason: TabEditBlock::MarkupCompatibility,
                }
            })?;
            tabs.push(raw::catalog_edit::Tab {
                sheet: &binding.name,
                position,
                relationship_id: &source_catalog.sheets[position].relationship_id,
                state,
            });
        }
        let active_plan = active_changed.then(|| raw::catalog_edit::Active {
            sheet: before.tabs()[active].name(),
            position: active,
        });
        let workbook = raw::catalog_edit::rewrite(
            before.workbook_xml(),
            raw::catalog_edit::Plan {
                tabs,
                renames: Vec::new(),
                active: active_plan,
                order: None,
            },
        )?;

        let mut touched = Vec::new();
        touched
            .try_reserve_exact(before.touched().len())
            .map_err(|source| allocation("tab-state worksheet rewrite plan", source))?;
        for part in before.touched() {
            let selected = part.position == active;
            let bytes = raw::sheet_view_edit::rewrite(
                part.part.bytes.as_bytes(),
                selected,
                raw::sheet_view_edit::Context {
                    sheet: before.tabs()[part.position].name(),
                    position: part.position,
                },
            )?;
            touched.push((part.position, bytes));
        }
        let after = Snapshot::rewritten(&before, workbook, touched)?;
        package.check_execution()?;
        if after.active_position() != active
            || after
                .tabs()
                .iter()
                .zip(&self.staged_visibility)
                .any(|(actual, expected)| actual.visibility() != expected)
        {
            return Err(invalid("tab-state publication changed staged semantics"));
        }
        commit_change(&self.workbook, super::Patch::new(before, after))
    }

    fn set_visibility(&mut self, selector: Selector<'_>, value: Visibility) -> Result<bool> {
        let position = self.resolve(selector)?;
        self.require_known_owner(position)?;
        if matches!(self.staged_visibility[position], Visibility::Unknown(_)) {
            return Err(self.block(position, TabEditBlock::MarkupCompatibility));
        }
        if self.staged_visibility[position] == value {
            return Ok(false);
        }
        self.staged_visibility[position] = value;
        Ok(true)
    }

    fn resolve(&self, selector: Selector<'_>) -> Result<usize> {
        match selector {
            CoreSelector::Position(position) => self
                .before
                .tabs()
                .get(position.get())
                .map(|_| position.get())
                .ok_or_else(|| invalid("tab selector position did not resolve")),
            CoreSelector::Name(name) => {
                let key = crate::sheet::key(&name);
                self.before
                    .tabs()
                    .iter()
                    .position(|tab| crate::sheet::key(tab.name()) == key)
                    .ok_or_else(|| invalid("tab selector name did not resolve"))
            },
            CoreSelector::Id(never) => match never {},
            _ => Err(Error::UnsupportedSelector),
        }
    }

    fn require_known_owner(&self, position: usize) -> Result<()> {
        let kind = self.before.binding(position)?.kind;
        if kind == WorksheetKind::Unknown {
            Err(self.block(position, TabEditBlock::MarkupCompatibility))
        } else {
            Ok(())
        }
    }

    fn block(&self, position: usize, reason: TabEditBlock) -> Error {
        Error::TabEditBlocked {
            sheet: self
                .before
                .tabs()
                .get(position)
                .map_or_else(|| "<unknown>".to_owned(), |tab| tab.name().to_owned()),
            position,
            reason,
        }
    }
}

#[derive(Clone, Copy)]
enum XmlOwner {
    Workbook,
    Sheet,
}

#[derive(Default)]
struct XmlAudit {
    protected_structure: bool,
    alternate_content: bool,
}

fn audit_xml(content: &[u8], owner: XmlOwner, limits: Limits) -> Result<XmlAudit> {
    if content.len() > limits.max_xml_bytes {
        return Err(Error::limit(
            LimitKind::XmlBytes,
            content.len(),
            limits.max_xml_bytes,
        ));
    }
    let mut reader = NsReader::from_reader(content);
    // Names are recognized on namespace values as written; a declaration that
    // spells the markup-compatibility namespace with references binds it for
    // preprocessing all the same (record 0011).
    let mut audit = XmlAudit {
        alternate_content: litchi_ooxml_common::mce::binds_namespace_by_reference(content),
        ..XmlAudit::default()
    };
    let mut depth = 0usize;
    let mut events = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| Error::limit(LimitKind::Events, usize::MAX, limits.max_events))?;
        if events > limits.max_events {
            return Err(Error::limit(LimitKind::Events, events, limits.max_events));
        }
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::limit(LimitKind::Depth, usize::MAX, limits.max_depth))?;
                if depth > limits.max_depth {
                    return Err(Error::limit(LimitKind::Depth, depth, limits.max_depth));
                }
                observe_element(&mut audit, owner, &namespace, &element, reader.decoder())?;
            },
            Event::Empty(element) => {
                observe_element(&mut audit, owner, &namespace, &element, reader.decoder())?;
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("tab-state XML has an unmatched closing tag"))?;
            },
            Event::PI(_) | Event::DocType(_) => {
                return Err(invalid(
                    "tab-state edits refuse processing instructions and DTDs",
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if depth != 0 {
        return Err(invalid("tab-state XML ended inside an element"));
    }
    Ok(audit)
}

fn observe_element(
    audit: &mut XmlAudit,
    owner: XmlOwner,
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<()> {
    if element.name().local_name().as_ref() == b"AlternateContent"
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
    {
        audit.alternate_content = true;
    }
    if matches!(owner, XmlOwner::Workbook)
        && raw::namespace::is_spreadsheetml_name(namespace, element.name(), b"workbookProtection")
    {
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
            if attribute.key.as_ref() != b"lockStructure" {
                continue;
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| invalid(error.to_string()))?;
            audit.protected_structure |= match value.as_ref() {
                "1" | "true" => true,
                "0" | "false" => false,
                _ => return Err(invalid("invalid workbook lockStructure boolean")),
            };
        }
    }
    Ok(())
}

fn raw_visibility(value: &Visibility) -> Option<raw::catalog_edit::State> {
    match value {
        Visibility::Visible => Some(raw::catalog_edit::State::Visible),
        Visibility::Hidden => Some(raw::catalog_edit::State::Hidden),
        Visibility::VeryHidden => Some(raw::catalog_edit::State::VeryHidden),
        Visibility::Unknown(_) => None,
    }
}
