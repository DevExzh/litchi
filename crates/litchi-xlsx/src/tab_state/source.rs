//! Guarded tab-state publication over an immutable positional source.

use std::io::Write;
use std::sync::Arc;

use litchi_core::{ExecutionContext, ReadAt, Selector as CoreSelector};
use litchi_opc::{ReadLimits, SourceBackedPackage, SourceCacheLimits};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::snapshot::Snapshot;
use super::{Commit, Patch};
use crate::error::{Error, Result, TabEditBlock, allocation, invalid};
use crate::{Selector, Visibility, WorksheetKind, raw};

const MAX_EDIT_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_EDIT_XML_DEPTH: usize = 128;
const MAX_EDIT_XML_EVENTS: usize = 1_000_000;
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

/// An owning source-backed editor for existing workbook tabs.
pub struct SourceBackedEditor {
    package: SourceBackedPackage,
    origin: Arc<()>,
}

/// An isolated visibility and active-tab edit over one exact source closure.
pub struct SourceEdit<'a> {
    package: &'a SourceBackedPackage,
    before: Snapshot,
    staged_visibility: Vec<Visibility>,
    requested_active: Option<usize>,
}

impl SourceBackedEditor {
    /// Open with the standard bounded OPC policy.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open with an explicit bounded OPC policy.
    pub fn from_read_at_with_limits(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_limits(
            source,
            read_limits,
        )?)
    }

    /// Open with an explicit finite deferred-payload cache policy.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_cache_limits(
            source,
            cache_limits,
        )?)
    }

    /// Open with explicit read and finite cache policies.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
                source,
                read_limits,
                cache_limits,
            )?,
        )
    }

    /// Open with an explicit caller-owned execution context and the default cache.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_execution_context(
            source,
            read_limits,
            context,
        )?)
    }

    /// Open with explicit read and execution policies.
    pub fn from_read_at_with_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_read_at_with_execution_context(source, read_limits, context)
    }

    /// Open with explicit read, cache, and caller-owned execution policies.
    pub fn from_read_at_with_limits_and_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        read_limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                read_limits,
                cache_limits,
                context,
            )?,
        )
    }

    /// Build an editor from a validated deferred OPC package.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        package.check_execution()?;
        let editor = Self {
            package,
            origin: Arc::new(()),
        };
        editor.package.check_execution()?;
        Ok(editor)
    }

    /// Capture exact source-bound workbook tab state.
    pub fn snapshot(&self) -> Result<Snapshot> {
        self.package.check_execution()?;
        Snapshot::load_source_backed(&self.package, Arc::clone(&self.origin))
    }

    /// Begin an isolated existing-tab edit.
    pub fn edit(&self) -> Result<SourceEdit<'_>> {
        self.package.check_execution()?;
        SourceEdit::new(&self.package, self.snapshot()?)
    }

    /// Content-free deferred-Part cache diagnostics.
    #[must_use]
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }

    /// Publish a source-checked commit to a sequential sink.
    ///
    /// Visibility-only changes overlay the workbook Part. An active-tab
    /// transition additionally overlays the old and new worksheet Parts so
    /// `tabSelected` stays synchronized. Every other member is raw-copied.
    pub fn publish_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> Result<Snapshot> {
        self.package.check_execution()?;
        let before = commit.patch().before();
        if !before.belongs_to(&self.origin)
            || !before.matches_source_backed(&self.package, Some(&self.origin))?
        {
            return Err(Error::PatchConflict {
                part: before.workbook_part_name().to_string(),
            });
        }
        let target = if commit.patch().is_empty() {
            before.clone()
        } else {
            commit.patch().after().clone()
        };
        let mut overlays = Vec::new();
        overlays
            .try_reserve_exact(1 + target.touched().len())
            .map_err(|source| allocation("tab-state source overlay plan", source))?;
        overlays.push((
            target.workbook_part_name().clone(),
            target.workbook_xml().to_vec(),
        ));
        for part in target.touched() {
            overlays.push((part.part.uri.clone(), part.part.bytes.as_bytes().to_vec()));
        }
        self.package
            .write_part_overlays_to_stream(writer, overlays)?;
        Ok(target)
    }
}

impl<'a> SourceEdit<'a> {
    fn new(package: &'a SourceBackedPackage, before: Snapshot) -> Result<Self> {
        let mut staged_visibility = Vec::new();
        staged_visibility
            .try_reserve_exact(before.tabs().len())
            .map_err(|source| allocation("staged tab visibility", source))?;
        staged_visibility.extend(before.tabs().iter().map(|tab| tab.visibility().clone()));
        Ok(Self {
            package,
            before,
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

    /// Whether staged visibility or explicit activation differs semantically.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.staged_visibility
            .iter()
            .zip(self.before.tabs())
            .any(|(staged, before)| staged != before.visibility())
            || self
                .requested_active
                .is_some_and(|active| active != self.before.active_position())
    }

    /// Validate, bind the exact touched closure, and freeze the edit.
    pub fn commit(self) -> Result<Commit> {
        self.package.check_execution()?;
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
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, 0));
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

        let audit = audit_xml(self.before.workbook_xml(), XmlOwner::Workbook)?;
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
                .with_source_touched(self.package, &[self.before.active_position(), active])?
        } else {
            self.before
        };
        for part in before.touched() {
            let audit = audit_xml(part.part.bytes.as_bytes(), XmlOwner::Sheet)?;
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
        let source_catalog = raw::parse_catalog(before.workbook_xml())?;
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
        self.package.check_execution()?;
        if after.active_position() != active
            || after
                .tabs()
                .iter()
                .zip(&self.staged_visibility)
                .any(|(actual, expected)| actual.visibility() != expected)
        {
            return Err(invalid("tab-state publication changed staged semantics"));
        }
        let touched_parts = 1_u8
            .checked_add(u8::try_from(before.touched().len()).map_err(|_source| {
                invalid("tab-state touched Part count does not fit diagnostics")
            })?)
            .ok_or_else(|| invalid("tab-state touched Part count overflow"))?;
        let patch = Patch::new(before, after.clone());
        Ok(Commit::new(after, patch, touched_parts))
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

fn audit_xml(content: &[u8], owner: XmlOwner) -> Result<XmlAudit> {
    if content.len() > MAX_EDIT_XML_BYTES {
        return Err(invalid("tab-state XML exceeds the 32 MiB edit bound"));
    }
    let mut reader = NsReader::from_reader(content);
    let mut audit = XmlAudit::default();
    let mut depth = 0usize;
    let mut events = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("tab-state XML event count overflow"))?;
        if events > MAX_EDIT_XML_EVENTS {
            return Err(invalid("tab-state XML exceeds the event bound"));
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
                    .ok_or_else(|| invalid("tab-state XML depth overflow"))?;
                if depth > MAX_EDIT_XML_DEPTH {
                    return Err(invalid("tab-state XML exceeds the depth bound"));
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
        for attribute in element.attributes() {
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
