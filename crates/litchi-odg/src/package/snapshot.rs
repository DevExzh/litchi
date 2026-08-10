//! Immutable ODG package snapshots and lossless semantic shape edits.

use crate::model::{
    FormControl,
    group::Group,
    layer::Layer,
    page::Page,
    resource::Resource,
    shape::{Properties as ShapeProperties, Shape, ShapeKind},
    style::Style,
};
use litchi_core::{
    BlobBundle, BlobLimits, CompositionLimits, DiagnosticFingerprint, Error, History,
    HistoryLimits, JoinedSubEdits, Metadata, Patch as CorePatch, PatchLimits, PatchOperation,
    Result, Reversible, ReversibleOperation, SubEdit,
};
use litchi_odf_common::{
    compact_xml,
    core::{
        AuthoredXmlFragment, PackageWriter, XmlSourcePart, XmlSplicePublication, family::Package,
    },
    drawing::Frame,
    media,
};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{collections::BTreeMap, fmt::Write as _, fs, ops::Range, path::Path, sync::Arc};

pub(crate) const MIMETYPE: &str = "application/vnd.oasis.opendocument.graphics";
pub(crate) const TEMPLATE_MIMETYPE: &str = "application/vnd.oasis.opendocument.graphics-template";
const BODY_MARKER: &str = "<office:drawing";
const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const SVG: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FORM: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const FO: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const SCRIPT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XML_EVENTS: &[u8] = b"http://www.w3.org/2001/xml-events";
const XLINK: &[u8] = b"http://www.w3.org/1999/xlink";
const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const XML: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const MAX_DEPTH: usize = 256;
const MAX_PAGES: usize = 16_384;
const MAX_LAYERS: usize = 16_384;
const MAX_FORM_CONTROLS: usize = 65_536;
const MAX_TRANSFER_RESOURCES: usize = 4_096;
const MAX_GROUP_EDITS: usize = 4_096;
const MAX_SHAPES: usize = 1_000_000;
const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;
const DURABLE_FORMAT: &str = "litchi.odg";

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Draw,
    Svg,
    Text,
    Form,
    Style,
    Other,
}

type TextSpans = Vec<Vec<Vec<Option<Range<usize>>>>>;
type NameSpans = Vec<Vec<Option<Range<usize>>>>;
type LayerSpans = Vec<Vec<Option<Range<usize>>>>;
type GeometrySpans = Vec<Vec<[Option<Range<usize>>; 4]>>;
type PathSpans = Vec<Vec<Option<Range<usize>>>>;
type ControlSpans = Vec<Vec<Option<Range<usize>>>>;
type PageAttributeSpans = Vec<[Option<Range<usize>>; 2]>;

struct State {
    package: Package,
    mimetype: &'static str,
    security: SecurityStatus,
    pages: Vec<Page>,
    layers: Vec<Layer>,
    resources: Vec<Resource>,
    form_controls: Vec<FormControl>,
    styles: Vec<Style>,
    active_content: ActiveContentStatus,
}

/// Inert package security state and mutation policy.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SecurityStatus {
    signed: bool,
    encrypted: bool,
}

/// Explicit mutation policy for protected drawing packages.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum SecurityWritePolicy {
    /// Refuse any write that would invalidate signatures or encryption.
    #[default]
    Refuse,
    /// Deliberately remove stale package signatures while preserving all unsigned payloads.
    RemoveSignatures,
}

/// Explicit disposition for inert active-content provenance during writes.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum ActiveContentWritePolicy {
    /// Preserve unmodified active-content markup inertly without executing it.
    #[default]
    PreserveInert,
    /// Refuse publication when the source inventories any active-content surface.
    Refuse,
}

/// Bounded inert inventory of active or externally resolved drawing surfaces.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct ActiveContentStatus {
    scripts: usize,
    events: usize,
    actions: usize,
    dde: usize,
    external_links: usize,
    embedded_objects: usize,
}

impl ActiveContentStatus {
    /// Whether any active-content surface was inventoried.
    #[must_use]
    pub const fn is_present(self) -> bool {
        self.scripts != 0
            || self.events != 0
            || self.actions != 0
            || self.dde != 0
            || self.external_links != 0
            || self.embedded_objects != 0
    }

    /// Script-bearing element count.
    #[must_use]
    pub const fn scripts(self) -> usize {
        self.scripts
    }

    /// XML-event listener count.
    #[must_use]
    pub const fn events(self) -> usize {
        self.events
    }

    /// Presentation action/listener count.
    #[must_use]
    pub const fn actions(self) -> usize {
        self.actions
    }

    /// DDE source count.
    #[must_use]
    pub const fn dde(self) -> usize {
        self.dde
    }

    /// External hyperlink count.
    #[must_use]
    pub const fn external_links(self) -> usize {
        self.external_links
    }

    /// Embedded object/plugin/applet/floating-frame count.
    #[must_use]
    pub const fn embedded_objects(self) -> usize {
        self.embedded_objects
    }
}

impl SecurityStatus {
    /// Whether document or macro signature metadata is present.
    #[must_use]
    pub const fn is_signed(self) -> bool {
        self.signed
    }

    /// Whether any manifest member has encryption metadata.
    #[must_use]
    pub const fn is_encrypted(self) -> bool {
        self.encrypted
    }

    /// Whether ordinary semantic rewrite is allowed without changing security lifecycle.
    #[must_use]
    pub const fn allows_rewrite(self) -> bool {
        !self.signed && !self.encrypted
    }
}

/// An immutable, source-owning ODG package snapshot.
///
/// Unknown package members and unmodeled XML remain in the retained source
/// bytes. Semantic inspection never evaluates controls, scripts, actions, DDE,
/// links, or embedded payloads.
#[derive(Clone)]
pub struct Snapshot(Arc<State>);

impl Snapshot {
    /// Opens a package from a filesystem path.
    ///
    /// # Errors
    ///
    /// Returns an error when the package cannot be read or is not a structurally valid ODG.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(Package::open(path, MIMETYPE, BODY_MARKER, "ODG")?, MIMETYPE)
    }

    /// Opens an `OpenDocument` drawing template from a filesystem path.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is unreadable or is not a structurally valid `OTG`.
    pub fn open_template(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_package(
            Package::open(path, TEMPLATE_MIMETYPE, BODY_MARKER, "OTG")?,
            TEMPLATE_MIMETYPE,
        )
    }

    /// Opens a package from owned bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is not a structurally valid ODG.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_package(
            Package::from_bytes(bytes, MIMETYPE, BODY_MARKER, "ODG")?,
            MIMETYPE,
        )
    }

    /// Opens password-protected ODG bytes for inert inspection.
    ///
    /// Encrypted snapshots remain read-only: semantic commit and durable application refuse to
    /// strip or silently re-encrypt protected entries.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid package, unsupported encryption metadata, or bad password.
    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        Self::from_package(
            Package::from_bytes_with_password(bytes, password, MIMETYPE, BODY_MARKER, "ODG")?,
            MIMETYPE,
        )
    }

    /// Opens a password-protected ODG file for inert inspection.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::from_bytes_with_password`] plus filesystem errors.
    pub fn open_with_password(path: impl AsRef<Path>, password: impl Into<String>) -> Result<Self> {
        Self::from_bytes_with_password(fs::read(path)?, password)
    }

    /// Opens an `OpenDocument` drawing template from owned bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the package is not a structurally valid `OTG`.
    pub fn from_template_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_package(
            Package::from_bytes(bytes, TEMPLATE_MIMETYPE, BODY_MARKER, "OTG")?,
            TEMPLATE_MIMETYPE,
        )
    }

    fn from_package(package: Package, mimetype: &'static str) -> Result<Self> {
        let archive = package.package().package()?;
        let security = SecurityStatus {
            encrypted: archive.manifest().has_encrypted_entries(),
            signed: archive.has_file("META-INF/documentsignatures.xml")
                || archive.has_file("META-INF/macrosignatures.xml"),
        };
        let parsed = parse_content(package.content_xml())?;
        let layers = package
            .styles_xml()
            .map(parse_declared_layers)
            .transpose()?
            .unwrap_or_default();
        if parsed.layer_count.saturating_add(layers.len()) > MAX_LAYERS {
            return invalid("ODG declared layer count exceeds the limit");
        }
        let resources = scan_resources(&package)?;
        let mut styles = parse_style_definitions(package.content_xml())?
            .into_iter()
            .map(|definition| definition.style)
            .collect::<Vec<_>>();
        if let Some(styles_xml) = package.styles_xml() {
            styles.extend(
                parse_style_definitions(styles_xml)?
                    .into_iter()
                    .map(|definition| definition.style),
            );
        }
        styles.sort_unstable_by(|left, right| left.name().cmp(right.name()));
        styles
            .dedup_by(|left, right| left.name() == right.name() && left.family() == right.family());
        let active_content = scan_active_content(package.content_xml(), package.styles_xml())?;
        Ok(Self(Arc::new(State {
            package,
            mimetype,
            security,
            pages: parsed.pages,
            form_controls: parsed.form_controls,
            styles,
            active_content,
            layers,
            resources,
        })))
    }

    /// Returns the exact `content.xml` source.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.0.package.content_xml()
    }

    /// Returns exact `styles.xml`, when present.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.0.package.styles_xml()
    }

    /// Returns common document metadata, when present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.0.package.metadata()
    }

    /// Returns bounded pages in source order.
    #[must_use]
    pub fn pages(&self) -> &[Page] {
        &self.0.pages
    }

    /// Selects one page by exact name or checked position.
    ///
    /// # Errors
    ///
    /// Returns an error when an exact name is ambiguous.
    pub fn page<'selector>(
        &self,
        selector: impl Into<crate::page::Selector<'selector>>,
    ) -> Result<Option<&Page>> {
        let resolved_selector = selector.into();
        match resolved_selector {
            crate::page::Selector::Position(position) => Ok(self.pages().get(position.get())),
            crate::page::Selector::Name(name) => {
                let mut matches = self
                    .pages()
                    .iter()
                    .filter(|page| page.name() == Some(name.as_ref()));
                let selected = matches.next();
                if selected.is_some() && matches.next().is_some() {
                    return invalid("ODG page name selector is ambiguous");
                }
                Ok(selected)
            },
        }
    }

    /// Returns global drawing layers declared by `styles.xml` in source order.
    ///
    /// Page-local declarations are available from [`Page::layers`](crate::page::Page::layers).
    #[must_use]
    pub fn layers(&self) -> &[Layer] {
        &self.0.layers
    }

    /// Returns original package bytes exactly.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.0.package.as_bytes()
    }

    /// Whether this snapshot is an OTG drawing template.
    #[must_use]
    pub fn is_template(&self) -> bool {
        self.0.mimetype == TEMPLATE_MIMETYPE
    }

    /// Inert signature/encryption state and rewrite policy for this snapshot.
    #[must_use]
    pub fn security(&self) -> SecurityStatus {
        self.0.security
    }

    /// Returns an inert, non-executing active-content inventory.
    #[must_use]
    pub fn active_content(&self) -> ActiveContentStatus {
        self.0.active_content
    }

    /// Lists safe package entry names.
    ///
    /// # Errors
    ///
    /// Returns an error when package member validation fails.
    pub fn files(&self) -> Result<Vec<String>> {
        self.0.package.files()
    }

    /// Returns package-local image resources referenced by drawing XML.
    #[must_use]
    pub fn resources(&self) -> &[Resource] {
        &self.0.resources
    }

    /// Returns inert form elements carrying `form:id` in source order.
    #[must_use]
    pub fn form_controls(&self) -> &[FormControl] {
        &self.0.form_controls
    }

    /// Returns inert drawing style definitions from content and styles parts.
    #[must_use]
    pub fn style_definitions(&self) -> &[Style] {
        &self.0.styles
    }

    /// Resolves one group root to its complete flattened nested descendant closure.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, non-group shape, or missing source span.
    pub fn group(&self, page: usize, shape: usize) -> Result<Group> {
        let parsed = parse_content(self.content_xml())?;
        group_selection(&parsed, page, shape)
    }

    /// Reads one inventoried package-local resource without activating it.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or unreadable package member.
    pub fn resource_bytes(&self, resource: usize) -> Result<Option<Vec<u8>>> {
        let selected = self.resources().get(resource).ok_or_else(|| {
            Error::InvalidFormat("ODG resource selector is out of bounds".to_string())
        })?;
        if !selected.is_present() {
            return Ok(None);
        }
        self.0.package.package().get_file(selected.path()).map(Some)
    }

    /// Starts a source-bound semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.edit_with_policies(
            SecurityWritePolicy::Refuse,
            ActiveContentWritePolicy::PreserveInert,
        )
    }

    /// Starts a transaction with an explicit protected-package write policy.
    ///
    /// Encryption is never silently stripped. `RemoveSignatures` permits signed-package edits and
    /// omits the now-stale signature members from the rebuilt package.
    #[must_use]
    pub fn edit_with_security_policy(&self, security_policy: SecurityWritePolicy) -> Transaction {
        self.edit_with_policies(security_policy, ActiveContentWritePolicy::PreserveInert)
    }

    /// Starts a transaction with an explicit inert active-content disposition.
    #[must_use]
    pub fn edit_with_active_content_policy(
        &self,
        active_content_policy: ActiveContentWritePolicy,
    ) -> Transaction {
        self.edit_with_policies(SecurityWritePolicy::Refuse, active_content_policy)
    }

    /// Starts a transaction with explicit security and inert active-content dispositions.
    #[must_use]
    pub fn edit_with_policies(
        &self,
        security_policy: SecurityWritePolicy,
        active_content_policy: ActiveContentWritePolicy,
    ) -> Transaction {
        Transaction {
            source: self.clone(),
            content: self.content_xml().to_string(),
            content_splices: Some(Vec::new()),
            changes: Vec::new(),
            resource_edits: Vec::new(),
            security_policy,
            active_content_policy,
        }
    }

    /// Starts an empty deterministic composition for this exact snapshot.
    #[must_use]
    pub fn joined_edits(&self, limits: CompositionLimits) -> JoinedEdits {
        JoinedEdits::new(Lineage::new(self), limits)
    }

    /// Applies joined disjoint work atomically against this exact base.
    ///
    /// # Errors
    ///
    /// Returns an error for stale lineage, unsupported operations, security refusal, or failed
    /// whole-package readback. No intermediate snapshot is published on failure.
    pub fn apply_joined(&self, joined: JoinedEdits) -> Result<Snapshot> {
        self.apply_joined_with_policies(
            joined,
            SecurityWritePolicy::Refuse,
            ActiveContentWritePolicy::PreserveInert,
        )
    }

    /// Applies joined work under an explicit signature-write policy.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::apply_joined`] plus policy refusal.
    pub fn apply_joined_with_security_policy(
        &self,
        joined: JoinedEdits,
        security_policy: SecurityWritePolicy,
    ) -> Result<Snapshot> {
        self.apply_joined_with_policies(
            joined,
            security_policy,
            ActiveContentWritePolicy::PreserveInert,
        )
    }

    /// Applies joined work under explicit security and inert active-content dispositions.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::apply_joined`] plus either policy refusal.
    pub fn apply_joined_with_policies(
        &self,
        joined: JoinedEdits,
        security_policy: SecurityWritePolicy,
        active_content_policy: ActiveContentWritePolicy,
    ) -> Result<Snapshot> {
        if !joined.lineage().matches(self) {
            return invalid("joined ODG edits do not match the exact source snapshot");
        }
        let mut current = self.clone();
        for edit in joined.into_sub_edits() {
            current = apply_durable_patch(
                &current,
                edit.payload(),
                false,
                security_policy,
                active_content_policy,
            )?;
        }
        Ok(current)
    }

    /// Starts explicit bounded undo/redo history at this snapshot.
    #[must_use]
    pub fn history(&self, limits: HistoryLimits) -> SnapshotHistory {
        History::new(self.clone(), limits)
    }

    /// Prepares one compact shape or complete group subtree for checked cross-drawing transfer.
    ///
    /// The plan retains exact source provenance, the referenced local layer declaration, and
    /// package-local resource bytes. It never evaluates embedded or linked content.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, noncompact source XML, unreadable resources, or
    /// unresolved source dependencies.
    pub fn prepare_shape_transfer(&self, page: usize, shape: usize) -> Result<ShapeTransfer> {
        ensure_compact_rewrite_source(self)?;
        let parsed = parse_content(self.content_xml())?;
        let selected_page = parsed.pages.get(page).ok_or_else(|| {
            Error::InvalidFormat("ODG transfer page selector is out of bounds".into())
        })?;
        let selected_shape = selected_page.shapes().get(shape).cloned().ok_or_else(|| {
            Error::InvalidFormat("ODG transfer shape selector is out of bounds".into())
        })?;
        let span = parsed.shape_spans[page][shape]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG transfer shape span is missing".into()))?;
        let xml = self
            .content_xml()
            .get(span.clone())
            .ok_or_else(|| Error::InvalidFormat("ODG transfer shape span is invalid".into()))?
            .to_owned();
        let dependency_shapes = selected_page
            .shapes()
            .iter()
            .enumerate()
            .filter(|(index, _shape)| {
                parsed.shape_spans[page][*index]
                    .as_ref()
                    .is_some_and(|candidate| {
                        candidate.start >= span.start && candidate.end <= span.end
                    })
            })
            .map(|(_index, shape_value)| shape_value)
            .collect::<Vec<_>>();
        let mut layers = dependency_shapes
            .iter()
            .filter_map(|shape_value| shape_value.layer())
            .map(|name| resolve_transfer_layer(self, selected_page, name))
            .collect::<Result<Vec<_>>>()?;
        layers.sort_unstable_by(|left, right| left.name().cmp(right.name()));
        layers.dedup_by(|left, right| left.name() == right.name());
        let mut required_styles = dependency_shapes
            .iter()
            .flat_map(|shape_value| [shape_value.style_name(), shape_value.text_style_name()])
            .flatten()
            .map(str::to_owned)
            .collect::<Vec<_>>();
        required_styles.sort_unstable();
        required_styles.dedup();
        let mut controls = dependency_shapes
            .iter()
            .filter_map(|shape_value| shape_value.control_reference())
            .map(str::to_owned)
            .collect::<Vec<_>>();
        controls.sort_unstable();
        controls.dedup();
        let mut style_definitions = BTreeMap::new();
        while let Some(style_name) = required_styles.pop() {
            if style_definitions.contains_key(&style_name) {
                continue;
            }
            let definition = find_style_definition(self, &style_name)?.ok_or_else(|| {
                Error::Unsupported(format!(
                    "ODG transfer source has unresolved style '{style_name}'"
                ))
            })?;
            if let Some(parent) = style_parent_name(&definition.xml)? {
                required_styles.push(parent);
            }
            style_definitions.insert(
                style_name.clone(),
                TransferStyle {
                    name: style_name,
                    family: definition.style.family().to_owned(),
                    parent: definition.style.parent().map(str::to_owned),
                    xml: definition.xml,
                },
            );
        }
        let styles = style_definitions.keys().cloned().collect::<Vec<_>>();
        for control in &controls {
            if !self
                .form_controls()
                .iter()
                .any(|declared| declared.id() == control)
            {
                return Err(Error::Unsupported(format!(
                    "ODG transfer source has unresolved form control '{control}'"
                )));
            }
        }
        let dependency_xml = std::iter::once(xml.as_str())
            .chain(style_definitions.values().map(|style| style.xml.as_str()))
            .collect::<String>();
        let mut resources = Vec::new();
        let mut resource_bytes = 0usize;
        for (resource_index, resource) in self
            .resources()
            .iter()
            .enumerate()
            .filter(|(_index, resource)| transfer_xml_references(&dependency_xml, resource.href()))
        {
            if resources
                .iter()
                .any(|existing: &TransferResource| existing.path == resource.path())
            {
                continue;
            }
            if resources.len() >= MAX_TRANSFER_RESOURCES {
                return invalid("ODG transfer resource count exceeds the limit");
            }
            let bytes = self.resource_bytes(resource_index)?.ok_or_else(|| {
                Error::Unsupported(format!(
                    "ODG transfer resource '{}' is missing",
                    resource.path()
                ))
            })?;
            resource_bytes = resource_bytes.checked_add(bytes.len()).ok_or_else(|| {
                Error::InvalidFormat("ODG transfer resource size overflow".to_string())
            })?;
            if resource_bytes > MAX_OUTPUT_BYTES {
                return invalid("ODG transfer resources exceed the byte limit");
            }
            resources.push(TransferResource {
                href: resource.href().to_owned(),
                path: resource.path().to_owned(),
                media_type: resource.media_type().map(str::to_owned),
                bytes: Some(bytes),
            });
        }
        let control_definitions = controls
            .iter()
            .map(|identifier| {
                let position = parsed
                    .form_controls
                    .iter()
                    .position(|control| control.id() == identifier)
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODG transfer form control is missing".into())
                    })?;
                let control_span =
                    parsed.form_control_spans[position]
                        .as_ref()
                        .ok_or_else(|| {
                            Error::InvalidFormat("ODG transfer form-control span is missing".into())
                        })?;
                Ok(TransferControl {
                    control: parsed.form_controls[position].clone(),
                    xml: self.content_xml()[control_span.clone()].to_owned(),
                })
            })
            .collect::<Result<Vec<_>>>()?;
        Ok(ShapeTransfer {
            source: Lineage::new(self),
            shape: selected_shape,
            xml,
            layers,
            styles,
            style_definitions: style_definitions.into_values().collect(),
            controls,
            control_definitions,
            resources,
        })
    }

    /// Consumes the snapshot and returns its source bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        match Arc::try_unwrap(self.0) {
            Ok(state) => state.package.into_bytes(),
            Err(state) => state.package.as_bytes().to_vec(),
        }
    }
}

/// A staged source-bound package shape edit.
pub struct Transaction {
    source: Snapshot,
    content: String,
    content_splices: Option<Vec<ContentSplice>>,
    changes: Vec<Change>,
    resource_edits: Vec<ResourceEdit>,
    security_policy: SecurityWritePolicy,
    active_content_policy: ActiveContentWritePolicy,
}

#[derive(Debug)]
struct ContentSplice {
    source_range: Range<usize>,
    current_range: Range<usize>,
    expected: Vec<u8>,
    replacement: Vec<u8>,
}

impl Transaction {
    fn replace_content_value(&mut self, span: &Range<usize>, replacement: &str) -> Result<()> {
        let escaped = quick_xml::escape::escape(replacement).into_owned();
        let next = replace_xml_value(&self.content, span, replacement)?;
        if let Some(splices) = &mut self.content_splices {
            stage_content_splice(
                self.source.content_xml().as_bytes(),
                self.content.as_bytes(),
                splices,
                span,
                escaped.as_bytes(),
            )?;
        }
        self.content = next;
        Ok(())
    }

    fn replace_content_values(
        &mut self,
        spans: &[&Range<usize>],
        replacements: &[String; 4],
    ) -> Result<()> {
        let mut edits = spans
            .iter()
            .zip(replacements)
            .map(|(span, value)| ((*span).clone(), value.as_str()))
            .collect::<Vec<_>>();
        edits.sort_unstable_by_key(|(span, _)| std::cmp::Reverse(span.start));
        for (span, replacement) in edits {
            self.replace_content_value(&span, replacement)?;
        }
        Ok(())
    }

    fn invalidate_content_splices(&mut self) {
        self.content_splices = None;
    }

    fn require_group_descendant(&self, page: usize, group: usize, descendant: usize) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        if !group_selection(&parsed, page, group)?.contains(descendant) {
            return invalid("ODG shape is not owned by the selected group subtree");
        }
        Ok(())
    }

    /// Replaces geometry on one checked descendant owned by a group subtree.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid ownership or the same geometry errors as
    /// [`Self::set_shape_geometry`].
    pub fn set_group_descendant_geometry(
        &mut self,
        page: usize,
        group: usize,
        descendant: usize,
        x: impl Into<String>,
        y: impl Into<String>,
        width: impl Into<String>,
        height: impl Into<String>,
    ) -> Result<()> {
        self.require_group_descendant(page, group, descendant)?;
        self.set_shape_geometry(page, descendant, x, y, width, height)
    }

    /// Replaces text on one checked descendant owned by a group subtree.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid ownership or a non-lossless text owner.
    pub fn set_group_descendant_text(
        &mut self,
        page: usize,
        group: usize,
        descendant: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        self.require_group_descendant(page, group, descendant)?;
        self.set_shape_text(page, descendant, text)
    }

    /// Assigns one style to every losslessly addressable descendant style owner.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid group ownership, undeclared style, no style owners, or limits.
    pub fn set_group_style_name(
        &mut self,
        page: usize,
        group: usize,
        style_name: impl Into<String>,
    ) -> Result<()> {
        let requested_style = style_name.into();
        validate_bounded_value(&requested_style, "ODG group style name")?;
        let declared_in_content = declares_style_xml(&self.content, &requested_style)?;
        let declared_in_styles = self
            .source
            .styles_xml()
            .map(|styles| declares_style_xml(styles, &requested_style))
            .transpose()?
            .unwrap_or_default();
        if !declared_in_content && !declared_in_styles {
            return invalid("ODG group destination style is not declared");
        }
        let parsed = parse_content(&self.content)?;
        let selection = group_selection(&parsed, page, group)?;
        let targets = selection
            .descendants()
            .iter()
            .copied()
            .filter(|position| parsed.style_name_spans[page][*position].is_some())
            .collect::<Vec<_>>();
        if targets.is_empty() || targets.len() > MAX_GROUP_EDITS {
            return invalid("ODG group style owner count is unsupported");
        }
        for target in targets {
            self.set_shape_style_name(page, target, requested_style.clone())?;
        }
        Ok(())
    }

    /// Replaces every single-span descendant text owner atomically after complete preflight.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid group ownership, no lossless text owners, or limits.
    pub fn set_group_text(
        &mut self,
        page: usize,
        group: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        let replacement_text = text.into();
        if replacement_text.len() > MAX_TEXT_BYTES {
            return invalid("ODG replacement group text exceeds the limit");
        }
        let parsed = parse_content(&self.content)?;
        let selection = group_selection(&parsed, page, group)?;
        let targets = selection
            .descendants()
            .iter()
            .copied()
            .filter(|position| matches!(parsed.text_spans[page][*position].as_slice(), [Some(_)]))
            .collect::<Vec<_>>();
        if targets.is_empty() || targets.len() > MAX_GROUP_EDITS {
            return invalid("ODG group text owner count is unsupported");
        }
        for target in targets {
            self.set_shape_text(page, target, replacement_text.clone())?;
        }
        Ok(())
    }

    /// Changes an inert form reference on one checked control descendant.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid group ownership or control-reference semantics.
    pub fn set_group_descendant_control_reference(
        &mut self,
        page: usize,
        group: usize,
        descendant: usize,
        reference: impl Into<String>,
    ) -> Result<()> {
        self.require_group_descendant(page, group, descendant)?;
        self.set_shape_control_reference(page, descendant, reference)
    }

    /// Renames a page through its existing `draw:name` attribute.
    ///
    /// Page-name references elsewhere in the drawing are dependency checked and cause refusal;
    /// callers must update those owners explicitly rather than leaving dangling references.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, duplicate name, referenced old name, absent
    /// source attribute, or size limit.
    pub fn set_page_name(&mut self, page: usize, name: impl Into<String>) -> Result<()> {
        let after = name.into();
        validate_bounded_value(&after, "ODG page name")?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        let before = selected.name().ok_or_else(|| {
            Error::Unsupported("ODG page rename requires an existing draw:name".into())
        })?;
        if parsed
            .pages
            .iter()
            .enumerate()
            .any(|(index, value)| index != page && value.name() == Some(after.as_str()))
        {
            return invalid("ODG page rename would create a duplicate name");
        }
        if before == after {
            return Ok(());
        }
        if xml_has_attribute(&self.content, DRAW, b"page-name", before)? {
            return Err(Error::Unsupported(
                "ODG page rename is blocked by a draw:page-name dependency".into(),
            ));
        }
        let span = parsed.page_attribute_spans[page][0]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG page name source span is missing".into()))?;
        let before_owned = before.to_owned();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::PageName(PageNameChange {
            page,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Changes a page's existing drawing-page style reference.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, missing source/style declarations, or limits.
    pub fn set_page_style_name(
        &mut self,
        page: usize,
        style_name: impl Into<String>,
    ) -> Result<()> {
        let after = style_name.into();
        validate_bounded_value(&after, "ODG page style name")?;
        if !declares_style(&self.source, &after)? {
            return invalid("ODG destination page style is not declared");
        }
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        let before = selected.style_name().ok_or_else(|| {
            Error::Unsupported("ODG page style edit requires an existing draw:style-name".into())
        })?;
        if before == after {
            return Ok(());
        }
        let span = parsed.page_attribute_spans[page][1]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG page style source span is missing".into()))?;
        let before_owned = before.to_owned();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::PageStyle(PageStyleChange {
            page,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Replaces one shape's sole plain paragraph character-data span.
    ///
    /// Split, mixed, CDATA, and entity-reference text is refused rather than
    /// serialized through a lossy XML model. A transaction owns one edit;
    /// restaging the same selector replaces its pending value.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, unsupported source span, or limit violation.
    pub fn set_shape_text(
        &mut self,
        page: usize,
        shape: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        let after = text.into();
        if after.len() > MAX_TEXT_BYTES {
            return invalid("ODG replacement shape text exceeds the limit");
        }
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|page_value| page_value.shapes().get(shape))
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape selector is out of bounds".to_string())
            })?;
        let spans = parsed
            .text_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape source span is missing".to_string()))?;
        if !matches!(spans.as_slice(), [Some(_)]) {
            return invalid("ODG shape text is not one losslessly replaceable XML span");
        }
        if selected.text() == after {
            return Ok(());
        }
        let span = spans[0].as_ref().ok_or_else(|| {
            Error::InvalidFormat("ODG shape text source span is missing".to_string())
        })?;
        let before = selected.text().to_string();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::Text(TextChange {
            page,
            shape,
            before,
            after,
        }));
        Ok(())
    }

    /// Renames one shape through its existing `draw:name` attribute.
    ///
    /// ODF 1.4 Part 3 §19.197 defines `draw:name` as the reference name for
    /// graphical elements. This preserves the original start tag and attribute
    /// spelling, replacing only the validated attribute-value span.
    ///
    /// # Errors
    ///
    /// Returns an error for an out-of-bounds selector, an unnamed shape, a
    /// name over the bounded size, or a shape whose source attribute cannot be
    /// losslessly addressed.
    pub fn set_shape_name(
        &mut self,
        page: usize,
        shape: usize,
        name: impl Into<String>,
    ) -> Result<()> {
        let after = name.into();
        if after.len() > MAX_TEXT_BYTES {
            return invalid("ODG replacement shape name exceeds the limit");
        }
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|page_value| page_value.shapes().get(shape))
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape selector is out of bounds".to_string())
            })?;
        let before = selected.name().ok_or_else(|| {
            Error::Unsupported(
                "ODG shape rename requires an existing losslessly addressable draw:name"
                    .to_string(),
            )
        })?;
        let span = parsed
            .name_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape name source span is missing".to_string())
            })?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_string();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::Name(NameChange {
            page,
            shape,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Changes a shape's existing layer assignment without normalizing its tag.
    ///
    /// ODF 1.4 Part 3 §§10.2.2-10.2.3 and 19.189 define drawing layers and
    /// their shape assignment. The destination must be one of the declarations
    /// visible through [`Snapshot::layers`].
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, undeclared layer, absent source attribute, or
    /// limit violation.
    pub fn set_shape_layer(
        &mut self,
        page: usize,
        shape: usize,
        layer: impl Into<String>,
    ) -> Result<()> {
        let after = layer.into();
        if after.len() > MAX_TEXT_BYTES {
            return invalid("ODG replacement layer name exceeds the limit");
        }
        let parsed = parse_content(&self.content)?;
        let selected_page = parsed.pages.get(page).ok_or_else(|| {
            Error::InvalidFormat("ODG page selector is out of bounds".to_string())
        })?;
        let selected = selected_page.shapes().get(shape).ok_or_else(|| {
            Error::InvalidFormat("ODG shape selector is out of bounds".to_string())
        })?;
        let visible_layers = if selected_page.has_layer_set() {
            selected_page.layers()
        } else {
            self.source.layers()
        };
        if !visible_layers
            .iter()
            .any(|declared_layer| declared_layer.name() == after)
        {
            return invalid("ODG destination layer is not declared");
        }
        let before = selected.layer().ok_or_else(|| {
            Error::Unsupported(
                "ODG layer change requires an existing losslessly addressable draw:layer"
                    .to_string(),
            )
        })?;
        let span = parsed
            .layer_spans
            .get(page)
            .and_then(|shapes| shapes.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| {
                Error::InvalidFormat("ODG shape layer source span is missing".to_string())
            })?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_string();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::Layer(LayerChange {
            page,
            shape,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Replaces all four existing SVG geometry attributes as one operation.
    ///
    /// # Errors
    ///
    /// Returns an error if the checked selectors fail or the shape does not
    /// own four losslessly addressable geometry attributes.
    pub fn set_shape_geometry(
        &mut self,
        page: usize,
        shape: usize,
        x: impl Into<String>,
        y: impl Into<String>,
        width: impl Into<String>,
        height: impl Into<String>,
    ) -> Result<()> {
        let after = [x.into(), y.into(), width.into(), height.into()];
        validate_geometry(&after)?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        let before = [
            selected.x(),
            selected.y(),
            selected.width(),
            selected.height(),
        ]
        .map(|value| value.map(str::to_owned));
        let spans = parsed
            .geometry_spans
            .get(page)
            .and_then(|values| values.get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape geometry spans are missing".into()))?;
        let ranges = spans
            .iter()
            .map(|span| {
                span.as_ref().ok_or_else(|| {
                    Error::Unsupported(
                        "ODG geometry edit requires existing x, y, width, and height attributes"
                            .into(),
                    )
                })
            })
            .collect::<Result<Vec<_>>>()?;
        if before
            .iter()
            .zip(&after)
            .all(|(source_value, target_value)| {
                source_value.as_deref() == Some(target_value.as_str())
            })
        {
            return Ok(());
        }
        self.replace_content_values(&ranges, &after)?;
        self.changes.push(Change::Geometry(GeometryChange {
            page,
            shape,
            before: before.map(Option::unwrap_or_default),
            after,
        }));
        Ok(())
    }

    /// Changes an existing graphic style reference without normalizing XML.
    ///
    /// # Errors
    ///
    /// Returns an error for a checked-selector failure or missing source attribute.
    pub fn set_shape_style_name(
        &mut self,
        page: usize,
        shape: usize,
        style_name: impl Into<String>,
    ) -> Result<()> {
        let after = style_name.into();
        validate_bounded_value(&after, "ODG shape style name")?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        let before = selected.style_name().ok_or_else(|| {
            Error::Unsupported("ODG style edit requires an existing draw:style-name".into())
        })?;
        let span = parsed
            .style_name_spans
            .get(page)
            .and_then(|values| values.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| Error::InvalidFormat("ODG shape style span is missing".into()))?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_owned();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::Style(StyleChange {
            page,
            shape,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Changes an existing SVG path-data attribute without normalizing XML.
    ///
    /// # Errors
    ///
    /// Returns an error unless the selected shape is a path with an existing,
    /// losslessly addressable `svg:d` attribute.
    pub fn set_shape_path_data(
        &mut self,
        page: usize,
        shape: usize,
        path_data: impl Into<String>,
    ) -> Result<()> {
        let after = path_data.into();
        validate_path_data(&after)?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        if selected.kind() != ShapeKind::Path {
            return Err(Error::Unsupported(
                "ODG path-data edit requires a draw:path shape".into(),
            ));
        }
        let before = selected.path_data().ok_or_else(|| {
            Error::Unsupported("ODG path-data edit requires an existing svg:d".into())
        })?;
        let span = parsed
            .path_spans
            .get(page)
            .and_then(|values| values.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| Error::InvalidFormat("ODG shape path-data span is missing".into()))?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_owned();
        self.replace_content_value(span, &after)?;
        self.changes.push(Change::Path(PathChange {
            page,
            shape,
            before: before_owned,
            after,
        }));
        Ok(())
    }

    /// Changes an existing inert form-control reference without activating it.
    ///
    /// # Errors
    ///
    /// Returns an error unless the selected control shape has an existing,
    /// losslessly addressable `draw:control` attribute.
    pub fn set_shape_control_reference(
        &mut self,
        page: usize,
        shape: usize,
        reference: impl Into<String>,
    ) -> Result<()> {
        let after = reference.into();
        validate_bounded_value(&after, "ODG form-control reference")?;
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        if selected.kind() != ShapeKind::Control {
            return Err(Error::Unsupported(
                "ODG form-control edit requires a draw:control shape".into(),
            ));
        }
        let before = selected.control_reference().ok_or_else(|| {
            Error::Unsupported("ODG form-control edit requires an existing draw:control".into())
        })?;
        let span = parsed
            .control_spans
            .get(page)
            .and_then(|values| values.get(shape))
            .and_then(Option::as_ref)
            .ok_or_else(|| {
                Error::InvalidFormat("ODG form-control source span is missing".into())
            })?;
        if before == after {
            return Ok(());
        }
        let before_owned = before.to_owned();
        self.replace_content_value(span, &after)?;
        self.changes
            .push(Change::ControlReference(ControlReferenceChange {
                page,
                shape,
                before: before_owned,
                after,
            }));
        Ok(())
    }

    /// Inserts a detached page at a checked source-order position.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid position, duplicate page identity, or limit violation.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "detached page values transfer ownership into the transaction"
    )]
    pub fn insert_page(&mut self, position: usize, page: Page) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        if position > parsed.pages.len() {
            return invalid("ODG page insertion position is out of bounds");
        }
        if parsed.pages.len() >= MAX_PAGES {
            return invalid("ODG page count exceeds the limit");
        }
        if let Some(name) = page.name()
            && parsed.pages.iter().any(|value| value.name() == Some(name))
        {
            return invalid("ODG inserted page name is already present");
        }
        let at = if position == parsed.pages.len() {
            parsed.drawing_insert_position
        } else {
            parsed.page_spans[position]
                .as_ref()
                .ok_or_else(|| Error::InvalidFormat("ODG page span is missing".into()))?
                .start
        };
        let xml = serialize_page(&page)?;
        let content = insert_child_xml(&self.content, at, &xml)?;
        self.invalidate_content_splices();
        self.content = content;
        self.changes
            .push(Change::Structure(StructureChange::PageInserted {
                position,
                name: page.name().map(str::to_owned),
            }));
        Ok(())
    }

    /// Appends a detached page.
    ///
    /// # Errors
    ///
    /// Returns an error for duplicate identity or a resource limit.
    pub fn add_page(&mut self, page: Page) -> Result<()> {
        let position = parse_content(&self.content)?.pages.len();
        self.insert_page(position, page)
    }

    /// Removes one page selected by exact name or checked position.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is absent, ambiguous, or unaddressable.
    pub fn remove_page<'selector>(
        &mut self,
        selector: impl Into<crate::page::Selector<'selector>>,
    ) -> Result<Page> {
        let parsed = parse_content(&self.content)?;
        let position = resolve_page_position(&parsed.pages, selector.into())?;
        let page = parsed.pages[position].clone();
        let span = parsed.page_spans[position]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG page span is missing".into()))?;
        let content = remove_xml(&self.content, span)?;
        if let Some(name) = page.name()
            && xml_has_attribute(&content, DRAW, b"page-name", name)?
        {
            return Err(Error::Unsupported(
                "ODG page removal is blocked by a draw:page-name dependency".into(),
            ));
        }
        self.invalidate_content_splices();
        self.content = content;
        self.changes
            .push(Change::Structure(StructureChange::PageRemoved {
                position,
                name: page.name().map(str::to_owned),
            }));
        Ok(page)
    }

    /// Inserts a detached shape at a checked page shape position.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, undeclared layers, or limits.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "detached shape values transfer ownership into the transaction"
    )]
    pub fn insert_shape(&mut self, page: usize, position: usize, shape: Shape) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        let selected_page = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        if position > selected_page.shapes().len() {
            return invalid("ODG shape insertion position is out of bounds");
        }
        validate_shape_layer(selected_page, self.source.layers(), &shape)?;
        if parsed
            .pages
            .iter()
            .map(|value| value.shapes().len())
            .sum::<usize>()
            >= MAX_SHAPES
        {
            return invalid("ODG shape count exceeds the limit");
        }
        let at = if position == selected_page.shapes().len() {
            parsed.page_insert_positions[page]
                .ok_or_else(|| Error::InvalidFormat("ODG page insertion point is missing".into()))?
        } else {
            parsed.shape_spans[page][position]
                .as_ref()
                .ok_or_else(|| Error::InvalidFormat("ODG shape span is missing".into()))?
                .start
        };
        let xml = serialize_shape(&shape)?;
        let content = insert_child_xml(&self.content, at, &xml)?;
        self.invalidate_content_splices();
        self.content = content;
        self.changes
            .push(Change::Structure(StructureChange::ShapeInserted {
                page,
                position,
                kind: shape.kind(),
            }));
        Ok(())
    }

    /// Appends a detached shape to a page.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, dependencies, or limits.
    pub fn add_shape(&mut self, page: usize, shape: Shape) -> Result<()> {
        let position = parse_content(&self.content)?
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?
            .shapes()
            .len();
        self.insert_shape(page, position, shape)
    }

    /// Adds or replaces one bounded inert automatic style definition.
    ///
    /// Arbitrary qualified property attributes are retained as data and never evaluated.
    /// Definitions owned by `styles.xml` are not silently shadowed.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid properties, ambiguous names, external ownership, or limits.
    pub fn put_style(&mut self, style: &Style) -> Result<()> {
        if let Some(parent) = style.parent() {
            let declared_in_content = declares_style_xml(&self.content, parent)?;
            let declared_in_styles = self
                .source
                .styles_xml()
                .map(|styles| declares_style_xml(styles, parent))
                .transpose()?
                .unwrap_or_default();
            if !declared_in_content && !declared_in_styles {
                return invalid("ODG parent style is not declared");
            }
        }
        let xml = serialize_style(style)?;
        let matches = parse_style_definitions(&self.content)?
            .into_iter()
            .filter(|definition| definition.style.name() == style.name())
            .collect::<Vec<_>>();
        if matches.len() > 1 {
            return invalid("ODG automatic style name is ambiguous");
        }
        let externally_owned = self
            .source
            .styles_xml()
            .map(|styles| declares_style_xml(styles, style.name()))
            .transpose()?
            .unwrap_or_default();
        let (content, change) = if let Some(existing) = matches.first() {
            if existing.style == *style {
                return Ok(());
            }
            (
                replace_xml(&self.content, &existing.span, &xml)?,
                StructureChange::StyleReplaced {
                    name: style.name().to_owned(),
                },
            )
        } else {
            if externally_owned {
                return Err(Error::Unsupported(
                    "ODG style owned by styles.xml cannot be shadowed".into(),
                ));
            }
            (
                insert_automatic_style(&self.content, &xml)?,
                StructureChange::StyleInserted {
                    name: style.name().to_owned(),
                },
            )
        };
        self.invalidate_content_splices();
        self.content = content;
        parse_style_definitions(&self.content)?;
        self.changes.push(Change::Structure(change));
        Ok(())
    }

    /// Removes one content-owned automatic style after checking all known references.
    ///
    /// # Errors
    ///
    /// Returns an error for missing/ambiguous/external ownership or a live style dependency.
    pub fn remove_style(&mut self, name: &str) -> Result<Style> {
        let matches = parse_style_definitions(&self.content)?
            .into_iter()
            .filter(|definition| definition.style.name() == name)
            .collect::<Vec<_>>();
        let [existing] = matches.as_slice() else {
            return invalid("ODG content style selector is missing or ambiguous");
        };
        if xml_has_attribute(&self.content, DRAW, b"style-name", name)?
            || xml_has_attribute(&self.content, DRAW, b"text-style-name", name)?
            || xml_has_attribute(&self.content, STYLE, b"parent-style-name", name)?
        {
            return Err(Error::Unsupported(
                "ODG style removal is blocked by a live dependency".into(),
            ));
        }
        let content = remove_xml(&self.content, &existing.span)?;
        let removed = existing.style.clone();
        self.invalidate_content_splices();
        self.content = content;
        self.changes
            .push(Change::Structure(StructureChange::StyleRemoved {
                name: name.to_owned(),
            }));
        Ok(removed)
    }

    /// Appends an empty structural group to a page.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid page selector or limit violation.
    pub fn add_group(&mut self, page: usize, name: impl Into<String>) -> Result<()> {
        self.add_shape(page, Shape::new(ShapeKind::Group).with_name(name))
    }

    /// Adds an inert form-control declaration without activating it.
    ///
    /// # Errors
    ///
    /// Returns an error for duplicate/invalid identifiers or output limits.
    pub fn add_form_control(&mut self, control: &FormControl) -> Result<()> {
        validate_bounded_value(control.id(), "ODG form-control identifier")?;
        if let Some(name) = control.name() {
            validate_bounded_value(name, "ODG form-control name")?;
        }
        let parsed = parse_content(&self.content)?;
        if parsed.form_controls.len() >= MAX_FORM_CONTROLS {
            return invalid("ODG form-control count exceeds the limit");
        }
        if parsed
            .form_controls
            .iter()
            .any(|value| value.id() == control.id())
        {
            return invalid("ODG form-control identifier is already present");
        }
        let control_xml = serialize_form_control(control)?;
        let content = if let Some(at) = parsed.forms_insert_position {
            let form_xml = format!(
                "<form:form xmlns:form=\"{}\" form:name=\"Litchi\">{control_xml}</form:form>",
                std::str::from_utf8(FORM).unwrap_or_default()
            );
            insert_child_xml(&self.content, at, &form_xml)?
        } else {
            let forms_xml = format!(
                "<office:forms xmlns:office=\"{}\" xmlns:form=\"{}\"><form:form form:name=\"Litchi\">{control_xml}</form:form></office:forms>",
                std::str::from_utf8(OFFICE).unwrap_or_default(),
                std::str::from_utf8(FORM).unwrap_or_default()
            );
            insert_xml(&self.content, parsed.drawing_start_position, &forms_xml)?
        };
        self.invalidate_content_splices();
        self.content = content;
        parse_content(&self.content)?;
        self.changes
            .push(Change::Structure(StructureChange::FormControlInserted {
                id: control.id().to_owned(),
            }));
        Ok(())
    }

    /// Replaces one inert form declaration while preserving its referenced `form:id`.
    ///
    /// Arbitrary bounded form attributes remain data only and are never activated.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing identifier, changed identity, invalid attributes, or limits.
    pub fn replace_form_control(&mut self, control: &FormControl) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        let mut matches = parsed
            .form_controls
            .iter()
            .enumerate()
            .filter(|(_index, candidate)| candidate.id() == control.id());
        let (position, _before) = matches.next().ok_or_else(|| {
            Error::InvalidFormat("ODG form-control selector did not match".into())
        })?;
        if matches.next().is_some() {
            return invalid("ODG form-control identifier is ambiguous");
        }
        let span = parsed.form_control_spans[position]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG form-control span is missing".into()))?;
        let xml = serialize_form_control(control)?;
        let content = replace_xml(&self.content, span, &xml)?;
        self.invalidate_content_splices();
        self.content = content;
        parse_content(&self.content)?;
        self.changes
            .push(Change::Structure(StructureChange::FormControlReplaced {
                id: control.id().to_owned(),
            }));
        Ok(())
    }

    /// Removes an inert form declaration only when no drawing shape references it.
    ///
    /// # Errors
    ///
    /// Returns an error for absent/ambiguous identifiers or a live `draw:control` dependency.
    pub fn remove_form_control(&mut self, identifier: &str) -> Result<FormControl> {
        let parsed = parse_content(&self.content)?;
        if parsed.pages.iter().any(|page| {
            page.shapes()
                .iter()
                .any(|shape| shape.control_reference() == Some(identifier))
        }) {
            return Err(Error::Unsupported(
                "ODG form-control removal is blocked by a drawing shape".into(),
            ));
        }
        let mut matches = parsed
            .form_controls
            .iter()
            .enumerate()
            .filter(|(_index, control)| control.id() == identifier);
        let (position, control) = matches.next().ok_or_else(|| {
            Error::InvalidFormat("ODG form-control selector did not match".into())
        })?;
        if matches.next().is_some() {
            return invalid("ODG form-control identifier is ambiguous");
        }
        let removed = control.clone();
        let span = parsed.form_control_spans[position]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG form-control span is missing".into()))?;
        let content = remove_xml(&self.content, span)?;
        self.invalidate_content_splices();
        self.content = content;
        self.changes
            .push(Change::Structure(StructureChange::FormControlRemoved {
                id: identifier.to_owned(),
            }));
        Ok(removed)
    }

    /// Inserts a prepared cross-drawing shape/group and its dependency closure.
    ///
    /// Missing page-local layers and noncolliding package resources are copied. Graphic/text
    /// styles and form controls must already exist in the destination; unresolved dependencies
    /// and differing resource-path collisions are refused before publication.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, unresolved dependencies, collisions, source
    /// policy violations, or output limits.
    pub fn insert_shape_transfer(
        &mut self,
        page: usize,
        position: usize,
        transfer: &ShapeTransfer,
    ) -> Result<()> {
        ShapeTransfer::validate_destination(&self.content, page)?;
        let mut transfer_xml = transfer.xml.clone();
        let mut resource_remaps = BTreeMap::new();
        for resource in &transfer.resources {
            let destination = self.stage_transferred_resource(resource)?;
            if destination != resource.path {
                resource_remaps.insert(resource.href.clone(), destination);
            }
        }

        let mut destination_styles = parse_style_definitions(&self.content)?;
        if let Some(styles) = self.source.styles_xml() {
            destination_styles.extend(parse_style_definitions(styles)?);
        }
        let mut reserved_style_names = destination_styles
            .iter()
            .map(|definition| definition.style.name().to_owned())
            .collect::<Vec<_>>();
        let mut style_remaps = BTreeMap::new();
        for transferred in &transfer.style_definitions {
            let collision = destination_styles
                .iter()
                .find(|definition| definition.style.name() == transferred.name);
            let destination_name = if collision.is_some_and(|definition| {
                definition.xml != transferred.xml || definition.style.family() != transferred.family
            }) {
                unique_collision_name(
                    &transferred.name,
                    transferred.xml.as_bytes(),
                    &reserved_style_names,
                )
            } else {
                transferred.name.clone()
            };
            reserved_style_names.push(destination_name.clone());
            style_remaps.insert(transferred.name.clone(), destination_name);
        }
        for transferred in &transfer.style_definitions {
            let destination_name = style_remaps
                .get(&transferred.name)
                .ok_or_else(|| Error::InvalidFormat("ODG style remap is missing".into()))?;
            let already_present = destination_styles.iter().any(|definition| {
                definition.style.name() == destination_name && definition.xml == transferred.xml
            });
            if already_present {
                continue;
            }
            let mut style_xml = transferred.xml.clone();
            for (before, after) in &style_remaps {
                style_xml = rewrite_qualified_attribute_values(
                    &style_xml,
                    &[b"style:name", b"style:parent-style-name"],
                    before,
                    after,
                )?;
            }
            for (before, after) in &resource_remaps {
                style_xml = rewrite_qualified_attribute_values(
                    &style_xml,
                    &[b"xlink:href"],
                    before,
                    after,
                )?;
            }
            style_xml = ensure_transfer_namespaces(
                &style_xml,
                &[("style", STYLE), ("draw", DRAW), ("svg", SVG), ("fo", FO)],
            )?;
            let content = insert_automatic_style(&self.content, &style_xml)?;
            self.invalidate_content_splices();
            self.content = content;
            self.changes
                .push(Change::Structure(StructureChange::StyleInserted {
                    name: destination_name.clone(),
                }));
        }
        for (before, after) in &style_remaps {
            transfer_xml = rewrite_qualified_attribute_values(
                &transfer_xml,
                &[b"draw:style-name", b"draw:text-style-name"],
                before,
                after,
            )?;
        }
        for (before, after) in &resource_remaps {
            transfer_xml =
                rewrite_qualified_attribute_values(&transfer_xml, &[b"xlink:href"], before, after)?;
        }

        for transferred in &transfer.control_definitions {
            let parsed = parse_content(&self.content)?;
            let collision = parsed
                .form_controls
                .iter()
                .position(|candidate| candidate.id() == transferred.control.id());
            if collision.is_some_and(|control_position| {
                parsed.form_control_spans[control_position]
                    .as_ref()
                    .is_some_and(|span| self.content[span.clone()] == transferred.xml)
            }) {
                continue;
            }
            let mut control_xml = transferred.xml.clone();
            let destination_id = if collision.is_some() {
                let reserved = parsed
                    .form_controls
                    .iter()
                    .map(|candidate| candidate.id().to_owned())
                    .collect::<Vec<_>>();
                let id = unique_collision_name(
                    transferred.control.id(),
                    transferred.xml.as_bytes(),
                    &reserved,
                );
                transfer_xml = rewrite_qualified_attribute_values(
                    &transfer_xml,
                    &[b"draw:control"],
                    transferred.control.id(),
                    &id,
                )?;
                control_xml = rewrite_qualified_attribute_values(
                    &control_xml,
                    &[b"form:id"],
                    transferred.control.id(),
                    &id,
                )?;
                id
            } else {
                transferred.control.id().to_owned()
            };
            control_xml = ensure_transfer_namespaces(
                &control_xml,
                &[("form", FORM), ("xlink", b"http://www.w3.org/1999/xlink")],
            )?;
            self.insert_transferred_form_control(&destination_id, &control_xml)?;
        }
        transfer_xml = ensure_transfer_namespaces(
            &transfer_xml,
            &[
                ("office", OFFICE),
                ("draw", DRAW),
                ("text", TEXT),
                ("svg", SVG),
                ("style", STYLE),
                ("form", FORM),
                ("fo", FO),
                ("xlink", XLINK),
                ("presentation", PRESENTATION),
            ],
        )?;
        let initial = parse_content(&self.content)?;
        let initial_page = initial.pages.get(page).ok_or_else(|| {
            Error::InvalidFormat("ODG transfer page selector is out of bounds".into())
        })?;
        let needs_local_layer_set = !initial_page.has_layer_set()
            && transfer.layers.iter().any(|required| {
                !self
                    .source
                    .layers()
                    .iter()
                    .any(|global| global.name() == required.name())
            });
        if needs_local_layer_set {
            let mut closure = initial_page
                .shapes()
                .iter()
                .filter_map(|shape| shape.layer())
                .map(|name| resolve_transfer_layer(&self.source, initial_page, name))
                .collect::<Result<Vec<_>>>()?;
            closure.extend(transfer.layers.iter().cloned());
            closure.sort_unstable_by(|left, right| left.name().cmp(right.name()));
            closure.dedup_by(|left, right| left.name() == right.name());
            for layer in closure {
                self.add_layer(page, layer)?;
            }
        }
        for layer in &transfer.layers {
            let parsed = parse_content(&self.content)?;
            let destination_page = parsed.pages.get(page).ok_or_else(|| {
                Error::InvalidFormat("ODG transfer page selector is out of bounds".into())
            })?;
            let visible = if destination_page.has_layer_set() {
                destination_page.layers()
            } else {
                self.source.layers()
            };
            if !visible.iter().any(|value| value.name() == layer.name()) {
                self.add_layer(page, layer.clone())?;
            }
        }
        let parsed = parse_content(&self.content)?;
        let destination_page = parsed.pages.get(page).ok_or_else(|| {
            Error::InvalidFormat("ODG transfer page selector is out of bounds".into())
        })?;
        if position > destination_page.shapes().len() {
            return invalid("ODG transfer shape position is out of bounds");
        }
        if parsed
            .pages
            .iter()
            .map(|value| value.shapes().len())
            .sum::<usize>()
            >= MAX_SHAPES
        {
            return invalid("ODG shape count exceeds the limit");
        }
        let at = if position == destination_page.shapes().len() {
            parsed.page_insert_positions[page]
                .ok_or_else(|| Error::InvalidFormat("ODG page insertion point is missing".into()))?
        } else {
            parsed.shape_spans[page][position]
                .as_ref()
                .ok_or_else(|| Error::InvalidFormat("ODG shape span is missing".into()))?
                .start
        };
        let content = insert_child_xml(&self.content, at, &transfer_xml)?;
        self.invalidate_content_splices();
        self.content = content;
        parse_content(&self.content)?;
        self.changes
            .push(Change::Structure(StructureChange::ShapeInserted {
                page,
                position,
                kind: transfer.shape.kind(),
            }));
        Ok(())
    }

    /// Removes one shape; removing a group owns and removes its complete subtree.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid checked selector or missing source span.
    pub fn remove_shape(&mut self, page: usize, shape: usize) -> Result<Shape> {
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .and_then(|value| value.shapes().get(shape))
            .cloned()
            .ok_or_else(|| Error::InvalidFormat("ODG shape selector is out of bounds".into()))?;
        let span = parsed.shape_spans[page][shape]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG shape span is missing".into()))?;
        let content = remove_xml(&self.content, span)?;
        self.invalidate_content_splices();
        self.content = content;
        self.changes
            .push(Change::Structure(StructureChange::ShapeRemoved {
                page,
                position: shape,
                kind: selected.kind(),
            }));
        Ok(selected)
    }

    /// Adds a page-local layer declaration.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid selectors, duplicate names, or limits.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "detached layer values transfer ownership into the transaction"
    )]
    pub fn add_layer(&mut self, page: usize, layer: Layer) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        if selected
            .layers()
            .iter()
            .any(|value| value.name() == layer.name())
        {
            return invalid("ODG page-local layer name is already present");
        }
        let layer_xml = serialize_layer(&layer)?;
        let content = if selected.has_layer_set() {
            let at = parsed.layer_set_insert_positions[page].ok_or_else(|| {
                Error::InvalidFormat("ODG page-local layer-set insertion point is missing".into())
            })?;
            insert_child_xml(&self.content, at, &layer_xml)?
        } else {
            let page_span = parsed.page_spans[page]
                .as_ref()
                .ok_or_else(|| Error::InvalidFormat("ODG page span is missing".into()))?;
            let xml = format!(
                "<draw:layer-set xmlns:draw=\"{}\">{layer_xml}</draw:layer-set>",
                std::str::from_utf8(DRAW).unwrap_or_default()
            );
            let empty_at = page_span.end.saturating_sub(2);
            if self.content.as_bytes().get(empty_at..page_span.end) == Some(b"/>") {
                insert_child_xml(&self.content, empty_at, &xml)?
            } else {
                let at = start_tag_end(&self.content, page_span.start)?;
                insert_xml(&self.content, at, &xml)?
            }
        };
        self.invalidate_content_splices();
        self.content = content;
        self.changes
            .push(Change::Structure(StructureChange::LayerInserted {
                page,
                name: layer.name().to_owned(),
            }));
        Ok(())
    }

    /// Removes an unreferenced page-local layer by exact name.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/ambiguous name or a live shape dependency.
    pub fn remove_layer(&mut self, page: usize, name: &str) -> Result<Layer> {
        let parsed = parse_content(&self.content)?;
        let selected = parsed
            .pages
            .get(page)
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into()))?;
        if selected
            .shapes()
            .iter()
            .any(|shape| shape.layer() == Some(name))
        {
            return Err(Error::Unsupported(
                "ODG layer removal is blocked by a shape assignment".into(),
            ));
        }
        let mut matches = selected
            .layers()
            .iter()
            .enumerate()
            .filter(|(_, layer)| layer.name() == name);
        let (position, matched_layer) = matches
            .next()
            .ok_or_else(|| Error::InvalidFormat("ODG layer selector did not match".into()))?;
        if matches.next().is_some() {
            return invalid("ODG layer name selector is ambiguous");
        }
        let removed_layer = matched_layer.clone();
        let span = parsed.layer_element_spans[page][position]
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("ODG layer span is missing".into()))?;
        let content = remove_xml(&self.content, span)?;
        self.invalidate_content_splices();
        self.content = content;
        self.changes
            .push(Change::Structure(StructureChange::LayerRemoved {
                page,
                name: name.to_owned(),
            }));
        Ok(removed_layer)
    }

    /// Adds or replaces one referenced package-local resource and manifest entry.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, media type, or size limit.
    pub fn set_resource(
        &mut self,
        resource: usize,
        media_type: impl Into<String>,
        bytes: Vec<u8>,
    ) -> Result<()> {
        let target_media_type = media_type.into();
        validate_media_type(&target_media_type)?;
        if bytes.len() > MAX_OUTPUT_BYTES {
            return invalid("ODG resource exceeds the output limit");
        }
        self.stage_resource(resource, Some(target_media_type), Some(bytes))
    }

    /// Adds a noncolliding package-local media/resource member.
    ///
    /// The member remains inert until drawing XML references it. Existing paths and unsafe package
    /// paths are refused rather than overwritten.
    ///
    /// # Errors
    ///
    /// Returns an error for unsafe/colliding paths, invalid media types, or size limits.
    pub fn add_resource(
        &mut self,
        path: impl Into<String>,
        media_type: impl Into<String>,
        bytes: Vec<u8>,
    ) -> Result<()> {
        let owned_path = path.into();
        let owned_media_type = media_type.into();
        validate_resource_path(&owned_path)?;
        validate_media_type(&owned_media_type)?;
        if bytes.len() > MAX_OUTPUT_BYTES {
            return invalid("ODG resource exceeds the output limit");
        }
        if self
            .source
            .files()?
            .iter()
            .any(|value| value == &owned_path)
            || self
                .resource_edits
                .iter()
                .any(|edit| edit.path == owned_path)
        {
            return invalid("ODG resource path is already present");
        }
        self.resource_edits.push(ResourceEdit {
            resource: self
                .source
                .resources()
                .len()
                .saturating_add(self.resource_edits.len()),
            path: owned_path,
            before_media_type: None,
            after_media_type: Some(owned_media_type),
            before_bytes: None,
            after_bytes: Some(bytes),
        });
        Ok(())
    }

    /// Removes a package member only when drawing XML has no live reference to its path.
    ///
    /// # Errors
    ///
    /// Returns an error for unsafe/absent paths or a live drawing resource dependency.
    pub fn remove_unreferenced_resource(&mut self, path: &str) -> Result<()> {
        validate_resource_path(path)?;
        if self
            .source
            .resources()
            .iter()
            .any(|resource| resource.path() == path)
        {
            return Err(Error::Unsupported(
                "ODG resource removal is blocked by a drawing reference".into(),
            ));
        }
        let archive = self.source.0.package.package();
        if !archive.has_file(path)? {
            return invalid("ODG resource path is absent");
        }
        let package = archive.package()?;
        self.resource_edits.push(ResourceEdit {
            resource: self.source.resources().len(),
            path: path.to_owned(),
            before_media_type: package.manifest().get_media_type(path).map(str::to_owned),
            after_media_type: None,
            before_bytes: Some(package.get_file(path)?),
            after_bytes: None,
        });
        Ok(())
    }

    /// Removes one package-local resource while retaining its inert reference.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid resource selector.
    pub fn remove_resource(&mut self, resource: usize) -> Result<()> {
        self.stage_resource(resource, None, None)
    }

    fn stage_resource(
        &mut self,
        resource: usize,
        after_media_type: Option<String>,
        after_bytes: Option<Vec<u8>>,
    ) -> Result<()> {
        let selected =
            self.source.resources().get(resource).ok_or_else(|| {
                Error::InvalidFormat("ODG resource selector is out of bounds".into())
            })?;
        let before_bytes = self.source.resource_bytes(resource)?;
        let before_media_type = selected.media_type().map(str::to_owned);
        if let Some(edit) = self
            .resource_edits
            .iter_mut()
            .find(|edit| edit.resource == resource)
        {
            edit.after_media_type = after_media_type;
            edit.after_bytes = after_bytes;
        } else {
            self.resource_edits.push(ResourceEdit {
                resource,
                path: selected.path().to_owned(),
                before_media_type: before_media_type.clone(),
                after_media_type,
                before_bytes: before_bytes.clone(),
                after_bytes,
            });
        }
        self.resource_edits.retain(|edit| {
            edit.before_media_type != edit.after_media_type || edit.before_bytes != edit.after_bytes
        });
        Ok(())
    }

    fn stage_transferred_resource(&mut self, resource: &TransferResource) -> Result<String> {
        validate_resource_path(&resource.path)?;
        let Some(bytes) = &resource.bytes else {
            return Ok(resource.path.clone());
        };
        if bytes.len() > MAX_OUTPUT_BYTES {
            return invalid("ODG transferred resource exceeds the output limit");
        }
        if let Some(staged) = self
            .resource_edits
            .iter()
            .find(|edit| edit.path == resource.path)
            && staged.after_bytes.as_ref() == Some(bytes)
            && staged.after_media_type == resource.media_type
        {
            return Ok(resource.path.clone());
        }
        let archive = self.source.0.package.package();
        if archive.has_file(&resource.path)? {
            let existing = archive.get_file(&resource.path)?;
            let existing_media_type = archive
                .package()?
                .manifest()
                .get_media_type(&resource.path)
                .map(str::to_owned);
            if existing == *bytes && existing_media_type == resource.media_type {
                return Ok(resource.path.clone());
            }
        }
        let destination = if self
            .resource_edits
            .iter()
            .any(|edit| edit.path == resource.path)
            || archive.has_file(&resource.path)?
        {
            unique_resource_path(&self.source, &self.resource_edits, &resource.path, bytes)?
        } else {
            resource.path.clone()
        };
        let media_type = resource
            .media_type
            .clone()
            .unwrap_or_else(|| "application/octet-stream".to_string());
        validate_media_type(&media_type)?;
        self.resource_edits.push(ResourceEdit {
            resource: self
                .source
                .resources()
                .len()
                .saturating_add(self.resource_edits.len()),
            path: destination.clone(),
            before_media_type: None,
            after_media_type: Some(media_type),
            before_bytes: None,
            after_bytes: Some(bytes.clone()),
        });
        Ok(destination)
    }

    fn insert_transferred_form_control(&mut self, id: &str, xml: &str) -> Result<()> {
        let parsed = parse_content(&self.content)?;
        if parsed.form_controls.len() >= MAX_FORM_CONTROLS {
            return invalid("ODG form-control count exceeds the limit");
        }
        if parsed
            .form_controls
            .iter()
            .any(|control| control.id() == id)
        {
            return invalid("ODG transferred form-control identifier is already present");
        }
        let content = if let Some(at) = parsed.forms_insert_position {
            let form_xml = format!(
                "<form:form xmlns:form=\"{}\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" form:name=\"Litchi\">{xml}</form:form>",
                std::str::from_utf8(FORM).unwrap_or_default()
            );
            insert_child_xml(&self.content, at, &form_xml)?
        } else {
            let forms_xml = format!(
                "<office:forms xmlns:office=\"{}\" xmlns:form=\"{}\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"><form:form form:name=\"Litchi\">{xml}</form:form></office:forms>",
                std::str::from_utf8(OFFICE).unwrap_or_default(),
                std::str::from_utf8(FORM).unwrap_or_default()
            );
            insert_xml(&self.content, parsed.drawing_start_position, &forms_xml)?
        };
        self.invalidate_content_splices();
        self.content = content;
        parse_content(&self.content)?;
        self.changes
            .push(Change::Structure(StructureChange::FormControlInserted {
                id: id.to_owned(),
            }));
        Ok(())
    }

    /// Atomically validates, rebuilds, and publishes the edited package.
    ///
    /// # Errors
    ///
    /// Returns an error when source policy, rebuilding, parsing, or typed readback fails.
    pub fn commit(self) -> Result<Commit> {
        if self.content == self.source.content_xml() && self.resource_edits.is_empty() {
            return Ok(Commit::unchanged(self.source));
        }
        enforce_security_policy(&self.source, self.security_policy)?;
        enforce_active_content_policy(&self.source, self.active_content_policy)?;
        let replacements = self
            .resource_edits
            .iter()
            .map(|edit| ResourceReplacement {
                path: &edit.path,
                media_type: edit.after_media_type.as_deref().unwrap_or_default(),
                bytes: edit.after_bytes.as_deref(),
            })
            .collect::<Vec<_>>();
        let requires_package_projection;
        let rebuilt = if let Some(splices) = &self.content_splices {
            requires_package_projection = self.source.security().is_signed()
                || ensure_compact_rewrite_source(&self.source).is_err();
            let publication = content_splice_publication(&self.source, splices)?;
            rebuild_spliced(
                &self.source,
                publication,
                &replacements,
                self.security_policy,
            )?
        } else {
            requires_package_projection = false;
            ensure_compact_rewrite_source(&self.source)?;
            compact_xml::validate(self.content.as_bytes()).map_err(Error::from)?;
            rebuild(
                &self.source,
                &self.content,
                &replacements,
                self.security_policy,
            )?
        };
        let snapshot = if self.source.is_template() {
            Snapshot::from_template_bytes(rebuilt)?
        } else {
            Snapshot::from_bytes(rebuilt)?
        };
        if snapshot.content_xml() != self.content {
            return invalid("ODG package edit failed exact content readback");
        }
        for edit in &self.resource_edits {
            let archive = snapshot.0.package.package().package()?;
            if archive.manifest().get_media_type(&edit.path) != edit.after_media_type.as_deref() {
                return invalid("ODG resource edit failed manifest readback");
            }
            let actual = if snapshot.0.package.package().has_file(&edit.path)? {
                Some(snapshot.0.package.package().get_file(&edit.path)?)
            } else {
                None
            };
            if actual != edit.after_bytes {
                return invalid("ODG resource edit failed byte readback");
            }
        }
        let resource_changes = self
            .resource_edits
            .iter()
            .map(ResourceEdit::change)
            .collect::<Vec<_>>();
        Ok(Commit {
            patch: Patch {
                source: self.source,
                target: snapshot.clone(),
                changes: self.changes,
                resource_changes,
                requires_package_projection,
            },
            snapshot,
            changed: true,
        })
    }
}

/// One semantic operation published by a unified ODG package transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Change {
    ControlReference(ControlReferenceChange),
    Text(TextChange),
    Name(NameChange),
    Layer(LayerChange),
    Geometry(GeometryChange),
    Style(StyleChange),
    Path(PathChange),
    PageName(PageNameChange),
    PageStyle(PageStyleChange),
    Structure(StructureChange),
}

/// One reversible page-name change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PageNameChange {
    page: usize,
    before: String,
    after: String,
}

impl PageNameChange {
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible drawing-page style-reference change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PageStyleChange {
    page: usize,
    before: String,
    after: String,
}

impl PageStyleChange {
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible inert form-control reference change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ControlReferenceChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl ControlReferenceChange {
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible semantic shape-text operation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl TextChange {
    /// The zero-based source-order page position.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// The zero-based source-order shape position.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// Text expected before application.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Text produced after application.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible `draw:name` change for a shape.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NameChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl NameChange {
    /// The zero-based source-order page position.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// The zero-based source-order shape position.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// Name expected before application.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// Name produced after application.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible drawing-layer assignment change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LayerChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl LayerChange {
    /// The zero-based source-order page position.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// The zero-based source-order shape position.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// The layer name expected before application.
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    /// The layer name produced after application.
    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// One reversible four-attribute geometry change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GeometryChange {
    page: usize,
    shape: usize,
    before: [String; 4],
    after: [String; 4],
}

impl GeometryChange {
    /// Page position at the time of this operation.
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    /// Shape position at the time of this operation.
    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    /// Source `[x, y, width, height]` lexical values.
    #[must_use]
    pub fn before(&self) -> &[String; 4] {
        &self.before
    }

    /// Target `[x, y, width, height]` lexical values.
    #[must_use]
    pub fn after(&self) -> &[String; 4] {
        &self.after
    }
}

/// One reversible graphic-style reference change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct StyleChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

/// One reversible SVG path-data change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PathChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl PathChange {
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

impl StyleChange {
    #[must_use]
    pub const fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub const fn shape(&self) -> usize {
        self.shape
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

/// A structural page, layer, shape, or group operation.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum StructureChange {
    PageInserted {
        position: usize,
        name: Option<String>,
    },
    PageRemoved {
        position: usize,
        name: Option<String>,
    },
    LayerInserted {
        page: usize,
        name: String,
    },
    LayerRemoved {
        page: usize,
        name: String,
    },
    ShapeInserted {
        page: usize,
        position: usize,
        kind: ShapeKind,
    },
    ShapeRemoved {
        page: usize,
        position: usize,
        kind: ShapeKind,
    },
    FormControlInserted {
        id: String,
    },
    FormControlRemoved {
        id: String,
    },
    FormControlReplaced {
        id: String,
    },
    StyleInserted {
        name: String,
    },
    StyleRemoved {
        name: String,
    },
    StyleReplaced {
        name: String,
    },
}

/// One package-local resource replacement or removal.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ResourceChange {
    resource: usize,
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_size: Option<usize>,
    after_size: Option<usize>,
}

impl ResourceChange {
    #[must_use]
    pub const fn resource(&self) -> usize {
        self.resource
    }

    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    #[must_use]
    pub fn before_media_type(&self) -> Option<&str> {
        self.before_media_type.as_deref()
    }

    #[must_use]
    pub fn after_media_type(&self) -> Option<&str> {
        self.after_media_type.as_deref()
    }

    #[must_use]
    pub const fn before_size(&self) -> Option<usize> {
        self.before_size
    }

    #[must_use]
    pub const fn after_size(&self) -> Option<usize> {
        self.after_size
    }
}

struct ResourceEdit {
    resource: usize,
    path: String,
    before_media_type: Option<String>,
    after_media_type: Option<String>,
    before_bytes: Option<Vec<u8>>,
    after_bytes: Option<Vec<u8>>,
}

impl ResourceEdit {
    fn change(&self) -> ResourceChange {
        ResourceChange {
            resource: self.resource,
            path: self.path.clone(),
            before_media_type: self.before_media_type.clone(),
            after_media_type: self.after_media_type.clone(),
            before_size: self.before_bytes.as_ref().map(Vec::len),
            after_size: self.after_bytes.as_ref().map(Vec::len),
        }
    }
}

struct ResourceReplacement<'a> {
    path: &'a str,
    media_type: &'a str,
    bytes: Option<&'a [u8]>,
}

/// A committed package publication and its exact-source patch.
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn unchanged(snapshot: Snapshot) -> Self {
        Self {
            patch: Patch {
                source: snapshot.clone(),
                target: snapshot.clone(),
                changes: Vec::new(),
                resource_changes: Vec::new(),
                requires_package_projection: false,
            },
            snapshot,
            changed: false,
        }
    }

    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// The published immutable snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The reversible exact-source patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the commit into its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

/// An exact-source-checked reversible package patch.
#[derive(Clone)]
pub struct Patch {
    source: Snapshot,
    target: Snapshot,
    changes: Vec<Change>,
    resource_changes: Vec<ResourceChange>,
    requires_package_projection: bool,
}

impl Patch {
    /// Whether this patch authorizes the supplied exact source bytes.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Snapshot) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Applies this patch only to its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` is not the exact source artifact.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if !self.is_applicable_to(source) {
            return invalid("ODG package patch source does not match");
        }
        Ok(self.target.clone())
    }

    /// The semantic change represented by this patch.
    #[must_use]
    pub fn change(&self) -> Option<&TextChange> {
        self.changes.iter().find_map(|change| match change {
            Change::Text(value) => Some(value),
            Change::ControlReference(_)
            | Change::Name(_)
            | Change::Layer(_)
            | Change::Geometry(_)
            | Change::Style(_)
            | Change::Path(_)
            | Change::PageName(_)
            | Change::PageStyle(_)
            | Change::Structure(_) => None,
        })
    }

    /// The semantic `draw:name` change, when this is a name patch.
    #[must_use]
    pub fn name_change(&self) -> Option<&NameChange> {
        self.changes.iter().find_map(|change| match change {
            Change::Name(value) => Some(value),
            Change::ControlReference(_)
            | Change::Text(_)
            | Change::Layer(_)
            | Change::Geometry(_)
            | Change::Style(_)
            | Change::Path(_)
            | Change::PageName(_)
            | Change::PageStyle(_)
            | Change::Structure(_) => None,
        })
    }

    /// The semantic drawing-layer change, when present.
    #[must_use]
    pub fn layer_change(&self) -> Option<&LayerChange> {
        self.changes.iter().find_map(|change| match change {
            Change::Layer(value) => Some(value),
            Change::ControlReference(_)
            | Change::Text(_)
            | Change::Name(_)
            | Change::Geometry(_)
            | Change::Style(_)
            | Change::Path(_)
            | Change::PageName(_)
            | Change::PageStyle(_)
            | Change::Structure(_) => None,
        })
    }

    /// All semantic operations in transaction order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Package-local resource changes in source selector order.
    #[must_use]
    pub fn resource_changes(&self) -> &[ResourceChange] {
        &self.resource_changes
    }

    /// Composes adjacent exact-lineage patches.
    ///
    /// # Errors
    ///
    /// Returns an error unless this target is byte-identical to `next`'s source.
    pub fn then(&self, next: &Self) -> Result<Self> {
        if self.target.as_bytes() != next.source.as_bytes() {
            return invalid("ODG patch composition lineage does not match");
        }
        let mut changes = self.changes.clone();
        changes.extend_from_slice(&next.changes);
        let mut resource_changes = self.resource_changes.clone();
        resource_changes.extend_from_slice(&next.resource_changes);
        Ok(Self {
            source: self.source.clone(),
            target: next.target.clone(),
            changes,
            resource_changes,
            requires_package_projection: self.requires_package_projection
                || next.requires_package_projection,
        })
    }

    /// An exact-source patch restoring the original package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self.changes.iter().rev().map(inverse_change).collect(),
            resource_changes: self
                .resource_changes
                .iter()
                .rev()
                .map(inverse_resource_change)
                .collect(),
            requires_package_projection: self.requires_package_projection,
        }
    }

    /// Projects this exact-source patch into the shared durable semantic wire format.
    ///
    /// # Errors
    ///
    /// Returns an error when a semantic operation exceeds the durable patch limits.
    pub fn durable(&self) -> Result<DurablePatch> {
        let limits = durable_limits();
        let source = fingerprint(self.source.as_bytes());
        let target = fingerprint(self.target.as_bytes());
        if !self.resource_changes.is_empty()
            || self.requires_package_projection
            || self
                .changes
                .iter()
                .any(|change| matches!(change, Change::Structure(_)))
        {
            return package_replacement_patch(self, limits, &source, &target);
        }
        let mut operations = Vec::new();
        for change in &self.changes {
            operations.push(change_operation(change, limits, &source, &target)?);
        }
        CorePatch::<Reversible>::new(
            limits,
            DURABLE_FORMAT,
            operations,
            BlobBundle::new(limits.blobs()),
            BlobBundle::new(limits.blobs()),
        )
        .map(|inner| DurablePatch { inner })
        .map_err(durable_error)
    }

    /// Prepares this patch as an independently joinable sub-edit.
    ///
    /// # Errors
    ///
    /// Returns an error when durable projection or bounded effect construction fails.
    pub fn prepare(
        &self,
        identifier: impl Into<String>,
        limits: CompositionLimits,
    ) -> Result<PreparedEdit> {
        let durable = self.durable()?;
        let writes = if durable
            .operations()
            .iter()
            .any(|operation| operation.op == "package.replace")
        {
            vec!["package".to_string()]
        } else {
            durable
                .operations()
                .iter()
                .map(|operation| format!("{}#{}", operation.target, operation.op))
                .collect::<Vec<_>>()
        };
        SubEdit::new(
            Lineage::new(&self.source),
            limits,
            identifier,
            Vec::<String>::new(),
            writes,
            durable,
        )
        .map_err(|error| Error::InvalidFormat(format!("invalid ODG sub-edit: {error}")))
    }
}

/// Exact ODG source lineage used by deterministic sub-edit composition.
#[derive(Clone, PartialEq, Eq)]
pub struct Lineage(Arc<[u8]>);

impl Lineage {
    fn new(snapshot: &Snapshot) -> Self {
        Self(Arc::from(snapshot.as_bytes()))
    }

    fn matches(&self, snapshot: &Snapshot) -> bool {
        self.0.as_ref() == snapshot.as_bytes()
    }
}

/// One independently prepared ODG semantic patch.
pub type PreparedEdit = SubEdit<Lineage, DurablePatch>;

/// Deterministically ordered, provably disjoint ODG sub-edits.
pub type JoinedEdits = JoinedSubEdits<Lineage, DurablePatch>;

/// Non-mutating three-way ODG merge plan.
pub type MergePlan = litchi_core::ThreeWayMergePlan<Lineage, DurablePatch>;

/// Explicit bounded ODG undo/redo history.
pub type SnapshotHistory = History<Snapshot>;

/// A bounded, provenance-bound shape or complete group-subtree transfer plan.
#[derive(Clone)]
pub struct ShapeTransfer {
    source: Lineage,
    shape: Shape,
    xml: String,
    layers: Vec<Layer>,
    styles: Vec<String>,
    style_definitions: Vec<TransferStyle>,
    controls: Vec<String>,
    control_definitions: Vec<TransferControl>,
    resources: Vec<TransferResource>,
}

impl ShapeTransfer {
    /// Root shape semantics retained by the transfer.
    #[must_use]
    pub const fn shape(&self) -> &Shape {
        &self.shape
    }

    /// Required declared layers in stable name order.
    #[must_use]
    pub fn layers(&self) -> &[Layer] {
        &self.layers
    }

    /// Required graphic/text style names in stable order.
    #[must_use]
    pub fn styles(&self) -> &[String] {
        &self.styles
    }

    /// Exact compact style-definition closure in stable name order.
    #[must_use]
    pub fn style_definitions(&self) -> &[TransferStyle] {
        &self.style_definitions
    }

    /// Required inert form-control identifiers in stable order.
    #[must_use]
    pub fn controls(&self) -> &[String] {
        &self.controls
    }

    /// Inert form-control declaration closure in stable identifier order.
    #[must_use]
    pub fn control_definitions(&self) -> &[TransferControl] {
        &self.control_definitions
    }

    /// Package-local resource closure in source occurrence order.
    #[must_use]
    pub fn resources(&self) -> &[TransferResource] {
        &self.resources
    }

    /// Content-free fingerprint of the exact source artifact.
    #[must_use]
    pub fn source_fingerprint(&self) -> String {
        DiagnosticFingerprint::of(self.source.0.as_ref()).as_hex()
    }

    fn validate_destination(content: &str, page: usize) -> Result<()> {
        if parse_content(content)?.pages.get(page).is_none() {
            return invalid("ODG transfer destination page is out of bounds");
        }
        Ok(())
    }
}

/// One exact compact style dependency retained by a transfer plan.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TransferStyle {
    name: String,
    family: String,
    parent: Option<String>,
    xml: String,
}

impl TransferStyle {
    /// Source style name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Source style family.
    #[must_use]
    pub fn family(&self) -> &str {
        &self.family
    }

    /// Optional parent-style dependency.
    #[must_use]
    pub fn parent(&self) -> Option<&str> {
        self.parent.as_deref()
    }
}

/// One exact inert form-control dependency retained by a transfer plan.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TransferControl {
    control: FormControl,
    xml: String,
}

impl TransferControl {
    /// Parsed inert form semantics.
    #[must_use]
    pub const fn control(&self) -> &FormControl {
        &self.control
    }
}

/// One inert package-local resource retained by a shape transfer plan.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TransferResource {
    href: String,
    path: String,
    media_type: Option<String>,
    bytes: Option<Vec<u8>>,
}

impl TransferResource {
    /// Exact source hyperlink spelling retained for collision rewriting.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Safe package-member path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Declared manifest media type, when present.
    #[must_use]
    pub fn media_type(&self) -> Option<&str> {
        self.media_type.as_deref()
    }

    /// Retained inert bytes, or `None` for a missing source reference.
    #[must_use]
    pub fn bytes(&self) -> Option<&[u8]> {
        self.bytes.as_deref()
    }
}

/// A durable, versioned ODG semantic patch.
#[derive(Clone)]
pub struct DurablePatch {
    inner: CorePatch<Reversible>,
}

impl DurablePatch {
    /// Parses canonical deterministic JSON under the ODG patch limits.
    ///
    /// # Errors
    ///
    /// Returns an error for non-canonical, malformed, over-limit, or wrong-format input.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        let inner = CorePatch::<Reversible>::from_deterministic_json(bytes, durable_limits())
            .map_err(durable_error)?;
        if inner.format() != DURABLE_FORMAT {
            return invalid("durable patch is not an ODG patch");
        }
        validate_durable_patch(&inner)?;
        validate_durable_patch(&inner.inverse())?;
        Ok(Self { inner })
    }

    /// Applies every operation after exact source precondition checks.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source, unsupported operations, security refusal, or failed
    /// whole-package readback.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        self.apply_with_policies(
            source,
            SecurityWritePolicy::Refuse,
            ActiveContentWritePolicy::PreserveInert,
        )
    }

    /// Applies this patch under an explicit signature-write policy.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source, invalid policy, unsupported encryption, or readback.
    pub fn apply_with_security_policy(
        &self,
        source: &Snapshot,
        security_policy: SecurityWritePolicy,
    ) -> Result<Snapshot> {
        self.apply_with_policies(
            source,
            security_policy,
            ActiveContentWritePolicy::PreserveInert,
        )
    }

    /// Applies this patch under explicit security and inert active-content dispositions.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source, policy refusal, unsupported operations, or readback.
    pub fn apply_with_policies(
        &self,
        source: &Snapshot,
        security_policy: SecurityWritePolicy,
        active_content_policy: ActiveContentWritePolicy,
    ) -> Result<Snapshot> {
        apply_durable_patch(source, self, true, security_policy, active_content_policy)
    }

    /// Returns the inverse durable patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            inner: self.inner.inverse(),
        }
    }

    /// Forward semantic operations in deterministic order.
    #[must_use]
    pub fn operations(&self) -> &[PatchOperation] {
        self.inner.operations()
    }

    /// Serializes canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch exceeds its retained serialization limits.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.inner.to_deterministic_json().map_err(durable_error)
    }

    /// Content-free diagnostic fingerprint of the canonical wire envelope.
    ///
    /// # Errors
    ///
    /// Returns an error when bounded canonical serialization fails.
    pub fn fingerprint(&self) -> Result<DiagnosticFingerprint> {
        self.inner.fingerprint().map_err(durable_error)
    }
}

fn package_replacement_patch(
    patch: &Patch,
    limits: PatchLimits,
    source: &str,
    target: &str,
) -> Result<DurablePatch> {
    let mut forward_blobs = BlobBundle::new(limits.blobs());
    let forward_id = forward_blobs
        .insert(patch.target.as_bytes())
        .map_err(durable_error)?;
    let mut reverse_blobs = BlobBundle::new(limits.blobs());
    let reverse_id = reverse_blobs
        .insert(patch.source.as_bytes())
        .map_err(durable_error)?;
    let operation = reversible_operation(
        limits,
        "package.replace",
        "package",
        source,
        target,
        serde_json::Value::String(reverse_id.as_hex()),
        serde_json::Value::String(forward_id.as_hex()),
    )?;
    CorePatch::<Reversible>::new(
        limits,
        DURABLE_FORMAT,
        [operation],
        forward_blobs,
        reverse_blobs,
    )
    .map(|inner| DurablePatch { inner })
    .map_err(durable_error)
}

fn change_operation(
    change: &Change,
    limits: PatchLimits,
    source: &str,
    target: &str,
) -> Result<ReversibleOperation> {
    let (name, semantic_target, before, after) = match change {
        Change::ControlReference(value) => (
            "shape.control.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::Text(value) => (
            "shape.text.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::Name(value) => (
            "shape.name.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::Layer(value) => (
            "shape.layer.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::Geometry(value) => (
            "shape.geometry.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::json!(value.before),
            serde_json::json!(value.after),
        ),
        Change::Style(value) => (
            "shape.style.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::Path(value) => (
            "shape.path.set",
            format!("page/{}/shape/{}", value.page, value.shape),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::PageName(value) => (
            "page.name.set",
            format!("page/{}", value.page),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::PageStyle(value) => (
            "page.style.set",
            format!("page/{}", value.page),
            serde_json::Value::String(value.before.clone()),
            serde_json::Value::String(value.after.clone()),
        ),
        Change::Structure(_) => {
            return invalid("structural ODG change requires package replacement projection");
        },
    };
    reversible_operation(
        limits,
        name,
        &semantic_target,
        source,
        target,
        before,
        after,
    )
}

fn string_value(operation: &PatchOperation) -> Result<&str> {
    operation
        .value
        .as_str()
        .ok_or_else(|| Error::InvalidFormat("ODG durable patch value is not a string".to_string()))
}

fn geometry_value(value: &serde_json::Value) -> Result<[String; 4]> {
    let values = value.as_array().ok_or_else(|| {
        Error::InvalidFormat("ODG durable geometry value is not an array".to_string())
    })?;
    if values.len() != 4 {
        return invalid("ODG durable geometry value must contain four attributes");
    }
    let parsed = values
        .iter()
        .map(|attribute| {
            attribute.as_str().map(str::to_owned).ok_or_else(|| {
                Error::InvalidFormat("ODG durable geometry attribute is not a string".to_string())
            })
        })
        .collect::<Result<Vec<_>>>()?;
    parsed
        .try_into()
        .map_err(|_values| Error::InvalidFormat("ODG durable geometry is invalid".to_string()))
}

fn durable_blob_id(value: &serde_json::Value) -> Result<&str> {
    let identifier = value.as_str().ok_or_else(|| {
        Error::InvalidFormat("ODG durable package blob identifier is not a string".to_string())
    })?;
    if identifier.len() != 64
        || !identifier
            .bytes()
            .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
    {
        return invalid("ODG durable package blob identifier is invalid");
    }
    Ok(identifier)
}

fn durable_blob<'patch>(
    patch: &'patch DurablePatch,
    value: &serde_json::Value,
) -> Result<&'patch [u8]> {
    let identifier = durable_blob_id(value)?;
    let blob_id = patch
        .inner
        .blobs()
        .ids()
        .find(|candidate| candidate.as_hex() == identifier)
        .ok_or_else(|| Error::InvalidFormat("ODG durable package blob is missing".to_string()))?;
    patch
        .inner
        .blobs()
        .get(blob_id)
        .ok_or_else(|| Error::InvalidFormat("ODG durable package blob is missing".to_string()))
}

fn apply_durable_patch(
    source: &Snapshot,
    patch: &DurablePatch,
    check_source: bool,
    security_policy: SecurityWritePolicy,
    active_content_policy: ActiveContentWritePolicy,
) -> Result<Snapshot> {
    validate_durable_patch(&patch.inner)?;
    let source_fingerprint = fingerprint(source.as_bytes());
    if check_source
        && patch.operations().iter().any(|operation| {
            operation
                .preconditions
                .get("source")
                .and_then(serde_json::Value::as_str)
                != Some(source_fingerprint.as_str())
        })
    {
        return invalid("ODG durable patch source does not match");
    }
    let mut current = source.clone();
    for operation in patch.operations() {
        if operation.op == "package.replace" {
            enforce_security_policy(&current, security_policy)?;
            enforce_active_content_policy(&current, active_content_policy)?;
            let blob = durable_blob(patch, &operation.value)?;
            current = if current.is_template() {
                Snapshot::from_template_bytes(blob.to_vec())?
            } else {
                Snapshot::from_bytes(blob.to_vec())?
            };
            continue;
        }
        let mut edit = current.edit_with_policies(security_policy, active_content_policy);
        match operation.op.as_str() {
            "page.name.set" => {
                edit.set_page_name(
                    parse_page_target(&operation.target)?,
                    string_value(operation)?,
                )?;
            },
            "page.style.set" => {
                edit.set_page_style_name(
                    parse_page_target(&operation.target)?,
                    string_value(operation)?,
                )?;
            },
            "shape.geometry.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                let values = geometry_value(&operation.value)?;
                edit.set_shape_geometry(
                    page, shape, &values[0], &values[1], &values[2], &values[3],
                )?;
            },
            "shape.control.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                edit.set_shape_control_reference(page, shape, string_value(operation)?)?;
            },
            "shape.layer.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                edit.set_shape_layer(page, shape, string_value(operation)?)?;
            },
            "shape.name.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                edit.set_shape_name(page, shape, string_value(operation)?)?;
            },
            "shape.path.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                edit.set_shape_path_data(page, shape, string_value(operation)?)?;
            },
            "shape.style.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                edit.set_shape_style_name(page, shape, string_value(operation)?)?;
            },
            "shape.text.set" => {
                let (page, shape) = parse_shape_target(&operation.target)?;
                edit.set_shape_text(page, shape, string_value(operation)?)?;
            },
            _ => return invalid("ODG durable patch operation is unsupported"),
        }
        current = edit.commit()?.into_snapshot();
    }
    Ok(current)
}

fn durable_error(error: litchi_core::PatchError) -> Error {
    let message = error.to_string();
    drop(error);
    Error::InvalidFormat(format!("invalid ODG durable patch: {message}"))
}

fn durable_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(8, MAX_OUTPUT_BYTES, MAX_OUTPUT_BYTES),
        4 * 1024 * 1024,
        1_024,
        32,
        MAX_TEXT_BYTES,
        32 * 1024 * 1024,
    )
}

fn fingerprint(bytes: &[u8]) -> String {
    DiagnosticFingerprint::of(bytes).as_hex()
}

fn reversible_operation(
    limits: PatchLimits,
    name: &str,
    semantic_target: &str,
    source: &str,
    target: &str,
    before: serde_json::Value,
    after: serde_json::Value,
) -> Result<ReversibleOperation> {
    let forward = PatchOperation::new(
        limits,
        name,
        semantic_target,
        source_precondition(source),
        after,
    )
    .map_err(durable_error)?;
    let inverse = PatchOperation::new(
        limits,
        name,
        semantic_target,
        source_precondition(target),
        before,
    )
    .map_err(durable_error)?;
    Ok(ReversibleOperation::new(forward, inverse))
}

fn source_precondition(source: &str) -> BTreeMap<String, serde_json::Value> {
    BTreeMap::from([(
        "source".to_string(),
        serde_json::Value::String(source.to_string()),
    )])
}

fn parse_shape_target(target: &str) -> Result<(usize, usize)> {
    let Some(remainder) = target.strip_prefix("page/") else {
        return invalid("ODG durable patch target is invalid");
    };
    let Some((page_text, shape_text)) = remainder.split_once("/shape/") else {
        return invalid("ODG durable patch target is invalid");
    };
    if shape_text.contains('/') {
        return invalid("ODG durable patch target is invalid");
    }
    let page = page_text
        .parse::<usize>()
        .map_err(|_error| Error::InvalidFormat("ODG durable page target is invalid".to_string()))?;
    let shape = shape_text.parse::<usize>().map_err(|_error| {
        Error::InvalidFormat("ODG durable shape target is invalid".to_string())
    })?;
    Ok((page, shape))
}

fn parse_page_target(target: &str) -> Result<usize> {
    let page_text = target
        .strip_prefix("page/")
        .filter(|value| !value.contains('/'))
        .ok_or_else(|| Error::InvalidFormat("ODG durable page target is invalid".to_string()))?;
    page_text
        .parse::<usize>()
        .map_err(|_error| Error::InvalidFormat("ODG durable page target is invalid".to_string()))
}

fn validate_durable_patch(patch: &CorePatch<Reversible>) -> Result<()> {
    for operation in patch.operations() {
        if !matches!(
            operation.op.as_str(),
            "package.replace"
                | "page.name.set"
                | "page.style.set"
                | "shape.geometry.set"
                | "shape.control.set"
                | "shape.layer.set"
                | "shape.name.set"
                | "shape.path.set"
                | "shape.style.set"
                | "shape.text.set"
        ) {
            return invalid("ODG durable patch operation is unsupported");
        }
        if operation.preconditions.len() != 1
            || operation
                .preconditions
                .get("source")
                .and_then(serde_json::Value::as_str)
                .is_none_or(|value| value.len() != 64 || !value.is_ascii())
        {
            return invalid("ODG durable patch source precondition is invalid");
        }
        match operation.op.as_str() {
            "package.replace" => {
                if operation.target != "package" {
                    return invalid("ODG durable package target is invalid");
                }
                let identifier = durable_blob_id(&operation.value)?;
                if !patch
                    .blobs()
                    .ids()
                    .any(|candidate| candidate.as_hex() == identifier)
                {
                    return invalid("ODG durable package blob is missing");
                }
            },
            "shape.geometry.set" => {
                parse_shape_target(&operation.target)?;
                geometry_value(&operation.value)?;
            },
            "page.name.set" | "page.style.set" => {
                parse_page_target(&operation.target)?;
                string_value(operation)?;
            },
            _ => {
                parse_shape_target(&operation.target)?;
                string_value(operation)?;
            },
        }
    }
    Ok(())
}

struct Parsed {
    control_spans: ControlSpans,
    form_controls: Vec<FormControl>,
    form_control_spans: Vec<Option<Range<usize>>>,
    forms_insert_position: Option<usize>,
    drawing_start_position: usize,
    pages: Vec<Page>,
    text_spans: TextSpans,
    name_spans: NameSpans,
    layer_spans: LayerSpans,
    layer_count: usize,
    geometry_spans: GeometrySpans,
    path_spans: PathSpans,
    style_name_spans: Vec<Vec<Option<Range<usize>>>>,
    page_spans: Vec<Option<Range<usize>>>,
    page_attribute_spans: PageAttributeSpans,
    page_insert_positions: Vec<Option<usize>>,
    shape_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_element_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_set_insert_positions: Vec<Option<usize>>,
    drawing_insert_position: usize,
}

fn group_selection(parsed: &Parsed, page: usize, shape: usize) -> Result<Group> {
    let selected = parsed
        .pages
        .get(page)
        .and_then(|value| value.shapes().get(shape))
        .ok_or_else(|| Error::InvalidFormat("ODG group selector is out of bounds".into()))?;
    if selected.kind() != ShapeKind::Group {
        return Err(Error::Unsupported(
            "ODG group operation requires a draw:g root".into(),
        ));
    }
    let root = parsed.shape_spans[page][shape]
        .as_ref()
        .ok_or_else(|| Error::InvalidFormat("ODG group source span is missing".into()))?;
    let descendants = parsed.shape_spans[page]
        .iter()
        .enumerate()
        .filter(|(position, span)| {
            *position != shape
                && span.as_ref().is_some_and(|candidate| {
                    candidate.start >= root.start && candidate.end <= root.end
                })
        })
        .map(|(position, _span)| position)
        .collect();
    Ok(Group::parsed(page, shape, descendants))
}

struct ActiveShape {
    depth: usize,
    page: usize,
    shape: usize,
    start: usize,
}

struct ActiveFormControl {
    depth: usize,
    control: usize,
    start: usize,
}

#[derive(Clone, Copy)]
enum AccessibilityKind {
    Description,
    Title,
}

struct ActiveAccessibility {
    depth: usize,
    page: usize,
    shape: usize,
    kind: AccessibilityKind,
}

struct Scanner {
    depth: usize,
    root_seen: bool,
    body_seen: bool,
    drawing_seen: bool,
    body_depth: Option<usize>,
    drawing_depth: Option<usize>,
    drawing_start_position: Option<usize>,
    forms_depth: Option<usize>,
    forms_insert_position: Option<usize>,
    form_controls: Vec<FormControl>,
    form_control_spans: Vec<Option<Range<usize>>>,
    active_form_controls: Vec<ActiveFormControl>,
    pages: Vec<Page>,
    page_depths: Vec<usize>,
    page_starts: Vec<usize>,
    layer_sets: Vec<(usize, Option<usize>)>,
    active_shapes: Vec<ActiveShape>,
    active_accessibility: Option<ActiveAccessibility>,
    control_spans: ControlSpans,
    paragraph_depths: Vec<usize>,
    text_spans: TextSpans,
    name_spans: NameSpans,
    layer_spans: LayerSpans,
    geometry_spans: GeometrySpans,
    path_spans: PathSpans,
    style_name_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_count: usize,
    shape_count: usize,
    text_bytes: usize,
    page_spans: Vec<Option<Range<usize>>>,
    page_attribute_spans: PageAttributeSpans,
    page_insert_positions: Vec<Option<usize>>,
    shape_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_element_spans: Vec<Vec<Option<Range<usize>>>>,
    layer_set_starts: Vec<(usize, usize, usize)>,
    layer_set_insert_positions: Vec<Option<usize>>,
    active_layers: Vec<(usize, usize, usize, usize)>,
    drawing_insert_position: Option<usize>,
}

impl Scanner {
    fn new() -> Self {
        Self {
            depth: 0,
            root_seen: false,
            body_seen: false,
            drawing_seen: false,
            body_depth: None,
            drawing_depth: None,
            drawing_start_position: None,
            forms_depth: None,
            forms_insert_position: None,
            form_controls: Vec::new(),
            form_control_spans: Vec::new(),
            active_form_controls: Vec::new(),
            pages: Vec::new(),
            page_depths: Vec::new(),
            page_starts: Vec::new(),
            layer_sets: Vec::new(),
            active_shapes: Vec::new(),
            active_accessibility: None,
            control_spans: Vec::new(),
            paragraph_depths: Vec::new(),
            text_spans: Vec::new(),
            name_spans: Vec::new(),
            layer_spans: Vec::new(),
            geometry_spans: Vec::new(),
            path_spans: Vec::new(),
            style_name_spans: Vec::new(),
            layer_count: 0,
            shape_count: 0,
            text_bytes: 0,
            page_spans: Vec::new(),
            page_attribute_spans: Vec::new(),
            page_insert_positions: Vec::new(),
            shape_spans: Vec::new(),
            layer_element_spans: Vec::new(),
            layer_set_starts: Vec::new(),
            layer_set_insert_positions: Vec::new(),
            active_layers: Vec::new(),
            drawing_insert_position: None,
        }
    }

    fn start(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: NamespaceKind,
        element: &BytesStart<'_>,
        tag: &[u8],
        tag_start: usize,
        empty: bool,
    ) -> Result<()> {
        self.depth = self
            .depth
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODG XML depth overflow".to_string()))?;
        if self.depth > MAX_DEPTH {
            return invalid("ODG XML nesting exceeds the limit");
        }
        self.observe(reader, namespace, element, tag, tag_start, empty)?;
        if empty {
            self.depth = self.depth.saturating_sub(1);
        }
        Ok(())
    }

    fn observe(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: NamespaceKind,
        element: &BytesStart<'_>,
        tag: &[u8],
        tag_start: usize,
        empty: bool,
    ) -> Result<()> {
        let local_name = element.local_name();
        let local = local_name.as_ref();
        if self.depth == 1 {
            if self.root_seen
                || namespace != NamespaceKind::Office
                || local != b"document-content"
                || empty
            {
                return invalid("ODG content.xml requires one office:document-content root");
            }
            self.root_seen = true;
            return Ok(());
        }
        if namespace == NamespaceKind::Office && local == b"body" {
            if self.body_seen || self.depth != 2 || empty {
                return invalid("ODG content.xml requires one non-empty office:body");
            }
            self.body_seen = true;
            self.body_depth = Some(self.depth);
            return Ok(());
        }
        if namespace == NamespaceKind::Office && local == b"forms" {
            if self.forms_depth.is_some()
                || self.forms_insert_position.is_some()
                || self.drawing_depth != Some(self.depth - 1)
            {
                return invalid("ODG office:forms is misplaced or duplicated");
            }
            if empty {
                self.forms_insert_position = Some(tag_start + tag.len() - 2);
            } else {
                self.forms_depth = Some(self.depth);
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Office && local == b"drawing" {
            if self.drawing_seen || self.body_depth != Some(self.depth - 1) {
                return invalid("ODG office:drawing is misplaced or duplicated");
            }
            self.drawing_seen = true;
            self.drawing_start_position = Some(if empty {
                tag_start + tag.len() - 2
            } else {
                tag_start + tag.len()
            });
            if empty {
                self.drawing_insert_position = Some(tag_start + tag.len() - 2);
            } else {
                self.drawing_depth = Some(self.depth);
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Form
            && self
                .forms_depth
                .is_some_and(|forms_depth| self.depth > forms_depth)
            && let Some(identifier) = attribute(reader, element, FORM, b"id")?
        {
            if self.form_controls.len() >= MAX_FORM_CONTROLS {
                return invalid("ODG form-control count exceeds the limit");
            }
            if self
                .form_controls
                .iter()
                .any(|control| control.id() == identifier)
            {
                return invalid("ODG form-control identifier is duplicated");
            }
            let control = self.form_controls.len();
            self.form_controls.push(FormControl::parsed(
                identifier,
                attribute(reader, element, FORM, b"name")?,
                String::from_utf8_lossy(local).into_owned(),
                arbitrary_attributes(
                    reader,
                    element,
                    &[(FORM, b"id".as_slice()), (FORM, b"name".as_slice())],
                )?,
            ));
            self.form_control_spans
                .push(empty.then_some(tag_start..tag_start + tag.len()));
            if !empty {
                self.active_form_controls.push(ActiveFormControl {
                    depth: self.depth,
                    control,
                    start: tag_start,
                });
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Draw && local == b"layer-set" {
            let page = self.current_page().ok_or_else(|| {
                Error::InvalidFormat("ODG layer-set is outside draw:page".to_string())
            })?;
            self.pages[page].mark_layer_set();
            if empty {
                self.layer_set_insert_positions[page] = Some(tag_start + tag.len() - 2);
            } else {
                self.layer_sets.push((self.depth, Some(page)));
                self.layer_set_starts.push((self.depth, page, tag_start));
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Draw
            && local == b"layer"
            && self
                .layer_sets
                .last()
                .is_some_and(|(depth, _)| *depth + 1 == self.depth)
        {
            self.add_layer(reader, element)?;
            let page = self.current_page().ok_or_else(|| {
                Error::InvalidFormat("ODG layer is outside draw:page".to_string())
            })?;
            let layer = self.pages[page].layers().len() - 1;
            if empty {
                self.layer_element_spans[page][layer] = Some(tag_start..tag_start + tag.len());
            } else {
                self.active_layers
                    .push((self.depth, page, layer, tag_start));
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Draw && local == b"page" {
            if self.drawing_depth != Some(self.depth - 1) {
                return invalid("ODG draw:page is outside office:drawing");
            }
            if self.pages.len() >= MAX_PAGES {
                return invalid("ODG page count exceeds the limit");
            }
            self.pages.push(Page::parsed(
                attribute(reader, element, DRAW, b"name")?,
                attribute(reader, element, XML, b"id")?,
                attribute(reader, element, DRAW, b"style-name")?,
                attribute(reader, element, DRAW, b"master-page-name")?,
            ));
            self.text_spans.push(Vec::new());
            self.control_spans.push(Vec::new());
            self.name_spans.push(Vec::new());
            self.layer_spans.push(Vec::new());
            self.geometry_spans.push(Vec::new());
            self.path_spans.push(Vec::new());
            self.style_name_spans.push(Vec::new());
            self.page_spans.push(None);
            self.page_attribute_spans.push([
                attribute_source_span(reader, element, tag, tag_start, DRAW, b"name")?,
                attribute_source_span(reader, element, tag, tag_start, DRAW, b"style-name")?,
            ]);
            self.page_insert_positions.push(None);
            self.shape_spans.push(Vec::new());
            self.layer_element_spans.push(Vec::new());
            self.layer_set_insert_positions.push(None);
            if empty {
                let page = self.pages.len() - 1;
                self.page_spans[page] = Some(tag_start..tag_start + tag.len());
                self.page_insert_positions[page] = Some(tag_start + tag.len() - 2);
            } else {
                self.page_depths.push(self.depth);
                self.page_starts.push(tag_start);
            }
            return Ok(());
        }
        if let Some(kind) = shape_kind(namespace, local) {
            let page = self.current_page().ok_or_else(|| {
                Error::InvalidFormat("ODG drawing shape is outside draw:page".to_string())
            })?;
            if self.shape_count >= MAX_SHAPES {
                return invalid("ODG shape count exceeds the limit");
            }
            self.shape_count += 1;
            let name = attribute(reader, element, DRAW, b"name")?;
            let page_name = self.pages[page].name().map(str::to_string);
            let frame = if kind == ShapeKind::Frame {
                Some(frame(reader, element, name.clone(), page_name)?)
            } else {
                None
            };
            let z_index = optional_u32_attribute(reader, element, DRAW, b"z-index")?;
            let geometry = [
                attribute(reader, element, SVG, b"x")?,
                attribute(reader, element, SVG, b"y")?,
                attribute(reader, element, SVG, b"width")?,
                attribute(reader, element, SVG, b"height")?,
            ];
            let shape = self.pages[page].shapes().len();
            self.pages[page].push_shape(Shape::parsed(
                ShapeProperties {
                    control_reference: attribute(reader, element, DRAW, b"control")?,
                    geometry,
                    layer: attribute(reader, element, DRAW, b"layer")?,
                    name,
                    path_data: attribute(reader, element, SVG, b"d")?,
                    style_name: attribute(reader, element, DRAW, b"style-name")?,
                    text_style_name: attribute(reader, element, DRAW, b"text-style-name")?,
                    z_index,
                },
                kind,
                frame,
            ));
            self.text_spans[page].push(Vec::new());
            self.control_spans[page].push(attribute_source_span(
                reader, element, tag, tag_start, DRAW, b"control",
            )?);
            self.name_spans[page].push(shape_name_span(reader, element, tag, tag_start)?);
            self.layer_spans[page].push(attribute_source_span(
                reader, element, tag, tag_start, DRAW, b"layer",
            )?);
            self.geometry_spans[page].push([
                attribute_source_span(reader, element, tag, tag_start, SVG, b"x")?,
                attribute_source_span(reader, element, tag, tag_start, SVG, b"y")?,
                attribute_source_span(reader, element, tag, tag_start, SVG, b"width")?,
                attribute_source_span(reader, element, tag, tag_start, SVG, b"height")?,
            ]);
            self.path_spans[page].push(attribute_source_span(
                reader, element, tag, tag_start, SVG, b"d",
            )?);
            self.style_name_spans[page].push(attribute_source_span(
                reader,
                element,
                tag,
                tag_start,
                DRAW,
                b"style-name",
            )?);
            self.shape_spans[page].push(empty.then_some(tag_start..tag_start + tag.len()));
            if !empty {
                self.active_shapes.push(ActiveShape {
                    depth: self.depth,
                    page,
                    shape,
                    start: tag_start,
                });
            }
            return Ok(());
        }
        if namespace == NamespaceKind::Svg
            && matches!(local, b"title" | b"desc")
            && !empty
            && let Some(active) = self.active_shapes.last()
            && active.depth + 1 == self.depth
        {
            self.active_accessibility = Some(ActiveAccessibility {
                depth: self.depth,
                page: active.page,
                shape: active.shape,
                kind: if local == b"title" {
                    AccessibilityKind::Title
                } else {
                    AccessibilityKind::Description
                },
            });
            return Ok(());
        }
        if !self.active_shapes.is_empty()
            && namespace == NamespaceKind::Text
            && local == b"p"
            && !empty
        {
            self.paragraph_depths.push(self.depth);
        }
        Ok(())
    }

    fn add_layer(&mut self, reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
        if self.layer_count >= MAX_LAYERS {
            return invalid("ODG declared layer count exceeds the limit");
        }
        self.layer_count += 1;
        let name = required_attribute(reader, element, DRAW, b"name", "draw:layer")?;
        let protected = optional_bool_attribute(reader, element, DRAW, b"protected")?;
        let layer = Layer::parsed(
            name,
            attribute(reader, element, DRAW, b"display")?,
            protected,
        );
        if let Some((_, Some(page))) = self.layer_sets.last() {
            self.pages[*page].push_layer(layer.clone());
            self.layer_element_spans[*page].push(None);
        }
        Ok(())
    }

    fn text(&mut self, span: Option<Range<usize>>, value: &str) -> Result<()> {
        if let Some(accessibility) = &self.active_accessibility {
            self.text_bytes = self.text_bytes.checked_add(value.len()).ok_or_else(|| {
                Error::InvalidFormat("ODG text extraction size overflow".to_string())
            })?;
            if self.text_bytes > MAX_TEXT_BYTES {
                return invalid("ODG text extraction exceeds the limit");
            }
            let shape = self.pages[accessibility.page]
                .shape_mut(accessibility.shape)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODG active accessibility shape disappeared".to_string())
                })?;
            match accessibility.kind {
                AccessibilityKind::Description => shape.push_description(value),
                AccessibilityKind::Title => shape.push_title(value),
            }
            return Ok(());
        }
        if self.paragraph_depths.is_empty() {
            return Ok(());
        }
        let Some(active) = self.active_shapes.last() else {
            return Ok(());
        };
        let (page, shape_index) = (active.page, active.shape);
        self.text_bytes = self
            .text_bytes
            .checked_add(value.len())
            .ok_or_else(|| Error::InvalidFormat("ODG text extraction size overflow".to_string()))?;
        if self.text_bytes > MAX_TEXT_BYTES {
            return invalid("ODG text extraction exceeds the limit");
        }
        let shape = self.pages[page]
            .shape_mut(shape_index)
            .ok_or_else(|| Error::InvalidFormat("ODG active shape disappeared".to_string()))?;
        shape.push_text(value);
        self.text_spans[page][shape_index].push(span);
        Ok(())
    }

    fn end(
        &mut self,
        namespace: NamespaceKind,
        local: &[u8],
        tag_start: usize,
        tag_end: usize,
    ) -> Result<()> {
        if self
            .active_form_controls
            .last()
            .is_some_and(|control| control.depth == self.depth)
        {
            let active = self.active_form_controls.pop().ok_or_else(|| {
                Error::InvalidFormat("ODG active form control disappeared".into())
            })?;
            self.form_control_spans[active.control] = Some(active.start..tag_end);
        }
        if self
            .active_accessibility
            .as_ref()
            .is_some_and(|active| active.depth == self.depth)
        {
            self.active_accessibility = None;
        }
        if self.paragraph_depths.last() == Some(&self.depth)
            && namespace == NamespaceKind::Text
            && local == b"p"
        {
            self.paragraph_depths.pop();
        }
        if self
            .active_shapes
            .last()
            .is_some_and(|shape| shape.depth == self.depth)
        {
            let active = self
                .active_shapes
                .pop()
                .ok_or_else(|| Error::InvalidFormat("ODG active shape disappeared".to_string()))?;
            self.shape_spans[active.page][active.shape] = Some(active.start..tag_end);
        }
        if self
            .active_layers
            .last()
            .is_some_and(|layer| layer.0 == self.depth)
        {
            let (_, page, layer, start) = self
                .active_layers
                .pop()
                .ok_or_else(|| Error::InvalidFormat("ODG active layer disappeared".to_string()))?;
            self.layer_element_spans[page][layer] = Some(start..tag_end);
        }
        if namespace == NamespaceKind::Draw
            && local == b"layer-set"
            && self
                .layer_sets
                .last()
                .is_some_and(|set| set.0 == self.depth)
        {
            self.layer_sets.pop();
            let (_, page, _) = self.layer_set_starts.pop().ok_or_else(|| {
                Error::InvalidFormat("ODG active layer-set disappeared".to_string())
            })?;
            self.layer_set_insert_positions[page] = Some(tag_start);
        }
        if namespace == NamespaceKind::Draw
            && local == b"page"
            && self.page_depths.last() == Some(&self.depth)
        {
            self.page_depths.pop();
            let start = self
                .page_starts
                .pop()
                .ok_or_else(|| Error::InvalidFormat("ODG active page disappeared".to_string()))?;
            let page = self.pages.len() - 1;
            self.page_spans[page] = Some(start..tag_end);
            self.page_insert_positions[page] = Some(tag_start);
        }
        if self.drawing_depth == Some(self.depth) {
            self.drawing_depth = None;
            self.drawing_insert_position = Some(tag_start);
        }
        if namespace == NamespaceKind::Office
            && local == b"forms"
            && self.forms_depth == Some(self.depth)
        {
            self.forms_depth = None;
            self.forms_insert_position = Some(tag_start);
        }
        if self.body_depth == Some(self.depth) {
            self.body_depth = None;
        }
        self.depth = self
            .depth
            .checked_sub(1)
            .ok_or_else(|| Error::InvalidFormat("ODG XML depth underflow".to_string()))?;
        Ok(())
    }

    fn current_page(&self) -> Option<usize> {
        self.page_depths
            .last()
            .filter(|depth| self.depth > **depth)
            .map(|_| self.pages.len() - 1)
    }

    fn finish(self) -> Result<Parsed> {
        if self.depth != 0
            || !self.root_seen
            || !self.body_seen
            || !self.drawing_seen
            || self.body_depth.is_some()
            || self.drawing_depth.is_some()
            || self.forms_depth.is_some()
            || !self.page_depths.is_empty()
            || !self.layer_sets.is_empty()
            || !self.active_form_controls.is_empty()
            || self.active_accessibility.is_some()
        {
            return invalid("ODG content.xml has an incomplete drawing structure");
        }
        Ok(Parsed {
            control_spans: self.control_spans,
            form_controls: self.form_controls,
            form_control_spans: self.form_control_spans,
            forms_insert_position: self.forms_insert_position,
            drawing_start_position: self.drawing_start_position.ok_or_else(|| {
                Error::InvalidFormat("ODG drawing start position is missing".to_string())
            })?,
            pages: self.pages,
            text_spans: self.text_spans,
            name_spans: self.name_spans,
            layer_spans: self.layer_spans,
            geometry_spans: self.geometry_spans,
            path_spans: self.path_spans,
            style_name_spans: self.style_name_spans,
            layer_count: self.layer_count,
            page_spans: self.page_spans,
            page_attribute_spans: self.page_attribute_spans,
            page_insert_positions: self.page_insert_positions,
            shape_spans: self.shape_spans,
            layer_element_spans: self.layer_element_spans,
            layer_set_insert_positions: self.layer_set_insert_positions,
            drawing_insert_position: self.drawing_insert_position.ok_or_else(|| {
                Error::InvalidFormat("ODG drawing insertion point is missing".to_string())
            })?,
        })
    }
}

fn parse_content(xml: &str) -> Result<Parsed> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut scanner = Scanner::new();
    loop {
        let start = position(&reader)?;
        let (resolved_namespace, borrowed_event) = reader
            .read_resolved_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG content.xml: {error}")))?;
        let namespace = classify(&resolved_namespace);
        let event = borrowed_event.into_owned();
        let end = position(&reader)?;
        match event {
            Event::Start(element) => scanner.start(
                &reader,
                namespace,
                &element,
                xml.as_bytes().get(start..end).ok_or_else(|| {
                    Error::InvalidFormat("ODG XML event span is invalid".to_string())
                })?,
                start,
                false,
            )?,
            Event::Empty(element) => scanner.start(
                &reader,
                namespace,
                &element,
                xml.as_bytes().get(start..end).ok_or_else(|| {
                    Error::InvalidFormat("ODG XML event span is invalid".to_string())
                })?,
                start,
                true,
            )?,
            Event::End(element) => {
                scanner.end(namespace, element.local_name().as_ref(), start, end)?;
            },
            Event::Text(text) => {
                let value = text_value(&text)?;
                scanner.text(Some(start..end), &value)?;
            },
            Event::CData(text) => {
                let value = text
                    .decode()
                    .map_err(|error| Error::InvalidFormat(format!("invalid ODG CDATA: {error}")))?;
                scanner.text(None, &value)?;
            },
            Event::GeneralRef(reference) => {
                let value = reference_value(&reference)?;
                scanner.text(None, &value)?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODG content.xml"),
            Event::Eof => return scanner.finish(),
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
        }
    }
}

fn parse_declared_layers(xml: &str) -> Result<Vec<Layer>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut layer_sets = Vec::<usize>::new();
    let mut layers = Vec::new();
    loop {
        let (resolved_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG styles.xml: {error}")))?;
        let namespace = classify(&resolved_namespace);
        match event {
            Event::Start(element) => {
                depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Draw && local.as_ref() == b"layer-set" {
                    layer_sets.push(depth);
                } else if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"layer"
                    && layer_sets.last().is_some_and(|set| *set + 1 == depth)
                {
                    push_declared_layer(&reader, &element, &mut layers)?;
                }
            },
            Event::Empty(element) => {
                let virtual_depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Draw
                    && local.as_ref() == b"layer"
                    && layer_sets
                        .last()
                        .is_some_and(|set| *set + 1 == virtual_depth)
                {
                    push_declared_layer(&reader, &element, &mut layers)?;
                }
            },
            Event::End(element) => {
                if namespace == NamespaceKind::Draw
                    && element.local_name().as_ref() == b"layer-set"
                    && layer_sets.last() == Some(&depth)
                {
                    layer_sets.pop();
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("ODG styles XML depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODG styles.xml"),
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    if depth != 0 || !layer_sets.is_empty() {
        return invalid("ODG styles.xml has incomplete layer declarations");
    }
    Ok(layers)
}

fn scan_resources(package: &Package) -> Result<Vec<Resource>> {
    let archive = package.package().package()?;
    let images = media::scan_package(package.content_xml(), package.styles_xml(), &archive)?;
    let mut resources = Vec::new();
    for (occurrence, image) in images.into_iter().enumerate() {
        match image.source {
            media::Source::PackagePart {
                href,
                path,
                manifest_media_type,
            } => resources.push(Resource::new(
                occurrence,
                href,
                path,
                manifest_media_type,
                true,
            )),
            media::Source::MissingPackagePart {
                href,
                resolved_path,
            } => resources.push(Resource::new(occurrence, href, resolved_path, None, false)),
            media::Source::Inline { .. }
            | media::Source::Linked { .. }
            | media::Source::Missing
            | _ => {},
        }
    }
    Ok(resources)
}

fn scan_active_content(content: &str, styles: Option<&str>) -> Result<ActiveContentStatus> {
    let mut status = ActiveContentStatus::default();
    scan_active_xml(content, &mut status)?;
    if let Some(style_xml) = styles {
        scan_active_xml(style_xml, &mut status)?;
    }
    Ok(status)
}

fn scan_active_xml(xml: &str, status: &mut ActiveContentStatus) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    loop {
        let (resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG active-content inventory XML: {error}"))
        })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if resolved_bound(&resolved_namespace, SCRIPT)
                    || resolved_bound(&resolved_namespace, OFFICE) && local == b"scripts"
                {
                    status.scripts = checked_active_count(status.scripts)?;
                }
                if resolved_bound(&resolved_namespace, XML_EVENTS) {
                    status.events = checked_active_count(status.events)?;
                }
                if resolved_bound(&resolved_namespace, PRESENTATION)
                    && matches!(local, b"event-listener" | b"show")
                {
                    status.actions = checked_active_count(status.actions)?;
                }
                if resolved_bound(&resolved_namespace, OFFICE) && local == b"dde-source" {
                    status.dde = checked_active_count(status.dde)?;
                }
                if resolved_bound(&resolved_namespace, DRAW)
                    && matches!(
                        local,
                        b"object" | b"object-ole" | b"plugin" | b"applet" | b"floating-frame"
                    )
                {
                    status.embedded_objects = checked_active_count(status.embedded_objects)?;
                }
                if attribute(&reader, &element, XLINK, b"href")?
                    .as_deref()
                    .is_some_and(is_external_href)
                {
                    status.external_links = checked_active_count(status.external_links)?;
                }
            },
            Event::DocType(_) => return invalid("DTD XML is prohibited in ODG active content"),
            Event::GeneralRef(reference) => {
                reference_value(&reference)?;
            },
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::End(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    Ok(())
}

fn checked_active_count(count: usize) -> Result<usize> {
    let next_count = count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("ODG active-content count overflow".into()))?;
    if next_count > MAX_SHAPES {
        return invalid("ODG active-content count exceeds the limit");
    }
    Ok(next_count)
}

fn is_external_href(href: &str) -> bool {
    href.contains("://")
        || href.starts_with("file:")
        || href.starts_with("mailto:")
        || href.starts_with("data:")
}

#[derive(Clone)]
struct ParsedStyleDefinition {
    style: Style,
    xml: String,
    span: Range<usize>,
}

struct ActiveStyleDefinition {
    depth: usize,
    start: usize,
    name: String,
    family: String,
    parent: Option<String>,
    properties: BTreeMap<String, String>,
}

fn parse_style_definitions(xml: &str) -> Result<Vec<ParsedStyleDefinition>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut active: Option<ActiveStyleDefinition> = None;
    let mut definitions = Vec::new();
    loop {
        let start = position(&reader)?;
        let (resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG style catalog XML: {error}"))
        })?;
        let namespace = classify(&resolved_namespace);
        let active_namespace = resolved_bound(&resolved_namespace, SCRIPT)
            || resolved_bound(&resolved_namespace, XML_EVENTS);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if active.is_some()
                    && (active_namespace
                        || namespace == NamespaceKind::Office && local.as_ref() == b"scripts")
                {
                    return invalid("active XML is prohibited in ODG style definitions");
                }
                if namespace == NamespaceKind::Style && local.as_ref() == b"style" {
                    if active.is_some() {
                        return invalid("ODG style definitions cannot be nested");
                    }
                    active = Some(ActiveStyleDefinition {
                        depth,
                        start,
                        name: required_attribute(&reader, &element, STYLE, b"name", "style")?,
                        family: required_attribute(&reader, &element, STYLE, b"family", "style")?,
                        parent: attribute(&reader, &element, STYLE, b"parent-style-name")?,
                        properties: BTreeMap::new(),
                    });
                } else if namespace == NamespaceKind::Style
                    && local.as_ref().ends_with(b"-properties")
                    && let Some(value) = &mut active
                    && value.depth.checked_add(1) == Some(depth)
                {
                    let owner = String::from_utf8_lossy(local.as_ref());
                    for (name, property) in arbitrary_attributes(&reader, &element, &[])? {
                        if value
                            .properties
                            .insert(format!("style:{owner}/{name}"), property)
                            .is_some()
                        {
                            return invalid("ODG style property attribute is duplicated");
                        }
                    }
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if active.is_some()
                    && (active_namespace
                        || namespace == NamespaceKind::Office && local.as_ref() == b"scripts")
                {
                    return invalid("active XML is prohibited in ODG style definitions");
                }
                if namespace == NamespaceKind::Style && local.as_ref() == b"style" {
                    let name = required_attribute(&reader, &element, STYLE, b"name", "style")?;
                    let family = required_attribute(&reader, &element, STYLE, b"family", "style")?;
                    let parent = attribute(&reader, &element, STYLE, b"parent-style-name")?;
                    definitions.push(ParsedStyleDefinition {
                        style: Style::parsed(name, family, parent, BTreeMap::new()),
                        xml: xml[start..end].to_owned(),
                        span: start..end,
                    });
                } else if namespace == NamespaceKind::Style
                    && local.as_ref().ends_with(b"-properties")
                    && let Some(value) = &mut active
                    && value.depth == depth
                {
                    let owner = String::from_utf8_lossy(local.as_ref());
                    for (name, property) in arbitrary_attributes(&reader, &element, &[])? {
                        if value
                            .properties
                            .insert(format!("style:{owner}/{name}"), property)
                            .is_some()
                        {
                            return invalid("ODG style property attribute is duplicated");
                        }
                    }
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Style
                    && local.as_ref() == b"style"
                    && active.as_ref().is_some_and(|value| value.depth == depth)
                {
                    let value = active.take().ok_or_else(|| {
                        Error::InvalidFormat("ODG style source span is missing".into())
                    })?;
                    definitions.push(ParsedStyleDefinition {
                        style: Style::parsed(
                            value.name,
                            value.family,
                            value.parent,
                            value.properties,
                        ),
                        xml: xml[value.start..end].to_owned(),
                        span: value.start..end,
                    });
                }
                depth = depth.saturating_sub(1);
            },
            Event::GeneralRef(reference) => {
                reference_value(&reference)?;
            },
            Event::DocType(_) => {
                return invalid("DTD XML is prohibited in ODG style catalogs");
            },
            Event::PI(_) if active.is_some() => {
                return invalid("processing instructions are prohibited in ODG style definitions");
            },
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    if depth != 0 || active.is_some() {
        return invalid("ODG style catalog XML is incomplete");
    }
    Ok(definitions)
}

fn checked_xml_depth(depth: usize) -> Result<usize> {
    let next_depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("ODG XML depth overflow".to_string()))?;
    if next_depth > MAX_DEPTH {
        return invalid("ODG XML nesting exceeds the limit");
    }
    Ok(next_depth)
}

fn push_declared_layer(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    layers: &mut Vec<Layer>,
) -> Result<()> {
    if layers.len() >= MAX_LAYERS {
        return invalid("ODG declared layer count exceeds the limit");
    }
    layers.push(Layer::parsed(
        required_attribute(reader, element, DRAW, b"name", "draw:layer")?,
        attribute(reader, element, DRAW, b"display")?,
        optional_bool_attribute(reader, element, DRAW, b"protected")?,
    ));
    Ok(())
}

fn text_value(text: &quick_xml::events::BytesText<'_>) -> Result<String> {
    let decoded = text
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODG text: {error}")))?;
    quick_xml::escape::unescape(&decoded)
        .map(std::borrow::Cow::into_owned)
        .map_err(|error| Error::InvalidFormat(format!("invalid ODG text escape: {error}")))
}

fn reference_value(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(value) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid ODG character reference: {error}"))
    })? {
        return Ok(value.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid ODG entity reference: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "apos" => Ok("'".to_string()),
        "quot" => Ok("\"".to_string()),
        _ => invalid("ODG custom entities are not allowed"),
    }
}

fn content_splice_publication(
    source: &Snapshot,
    splices: &[ContentSplice],
) -> Result<XmlSplicePublication> {
    let source_part = XmlSourcePart::load(source.0.package.package(), "content.xml")?;
    if source_part.bytes() != source.content_xml().as_bytes() {
        return invalid("ODG content splice has different package provenance");
    }
    let mut publication = XmlSplicePublication::new(source_part.clone());
    for splice in splices {
        let proof = source_part.checked_range(splice.source_range.clone(), &splice.expected)?;
        let fragment = if splice.replacement.is_empty() {
            AuthoredXmlFragment::deletion()
        } else {
            AuthoredXmlFragment::text(splice.replacement.clone())?
        };
        publication.replace(proof, fragment)?;
    }
    Ok(publication)
}

fn rebuild_spliced(
    source: &Snapshot,
    content: XmlSplicePublication,
    replacements: &[ResourceReplacement<'_>],
    security_policy: SecurityWritePolicy,
) -> Result<Vec<u8>> {
    enforce_security_policy(source, security_policy)?;
    let archive = source.0.package.package();
    let mut writer = PackageWriter::new_bounded(MAX_OUTPUT_BYTES);
    writer.set_mimetype(source.0.mimetype)?;
    content.publish(&mut writer)?;
    for path in ["styles.xml", "meta.xml", "settings.xml"] {
        if archive.has_file(path)? {
            XmlSplicePublication::new(XmlSourcePart::load(archive, path)?).publish(&mut writer)?;
        }
    }
    let mut excluded = replacements
        .iter()
        .map(|replacement| replacement.path.to_owned())
        .collect::<Vec<_>>();
    excluded.push("settings.xml".to_string());
    writer.copy_auxiliary_files_from_except(archive, &excluded, &[])?;
    for replacement in replacements {
        if let Some(bytes) = replacement.bytes {
            writer.add_file_with_media_type(replacement.path, bytes, replacement.media_type)?;
        }
    }
    writer.finish_to_bounded_bytes()
}

fn rebuild(
    source: &Snapshot,
    content: &str,
    replacements: &[ResourceReplacement<'_>],
    security_policy: SecurityWritePolicy,
) -> Result<Vec<u8>> {
    enforce_security_policy(source, security_policy)?;
    let archive = source.0.package.package();
    let mut writer = PackageWriter::new_bounded(MAX_OUTPUT_BYTES);
    writer.set_mimetype(source.0.mimetype)?;
    writer.add_file("content.xml", content.as_bytes())?;
    for path in ["styles.xml", "meta.xml", "settings.xml"] {
        if archive.has_file(path)? {
            writer.add_file(path, &archive.get_file(path)?)?;
        }
    }
    let mut excluded = replacements
        .iter()
        .map(|replacement| replacement.path.to_owned())
        .collect::<Vec<_>>();
    excluded.push("settings.xml".to_string());
    writer.copy_auxiliary_files_from_except(archive, &excluded, &[])?;
    for replacement in replacements {
        if let Some(bytes) = replacement.bytes {
            writer.add_file_with_media_type(replacement.path, bytes, replacement.media_type)?;
        }
    }
    writer.finish_to_bounded_bytes()
}

fn enforce_security_policy(source: &Snapshot, policy: SecurityWritePolicy) -> Result<()> {
    if source.security().is_signed() && policy == SecurityWritePolicy::Refuse {
        return invalid("ODG package edits refuse signed packages");
    }
    if source.security().is_encrypted() {
        return invalid("ODG package edits refuse encrypted packages");
    }
    Ok(())
}

fn enforce_active_content_policy(
    source: &Snapshot,
    policy: ActiveContentWritePolicy,
) -> Result<()> {
    if policy == ActiveContentWritePolicy::Refuse && source.active_content().is_present() {
        return Err(Error::Unsupported(
            "ODG active-content write policy refuses the source inventory".into(),
        ));
    }
    Ok(())
}

fn ensure_compact_rewrite_source(source: &Snapshot) -> Result<()> {
    let archive = source.0.package.package();
    for path in source.files()? {
        let is_xml = Path::new(&path)
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"));
        if is_xml && path != "META-INF/manifest.xml" {
            compact_xml::validate(&archive.get_file(&path)?).map_err(Error::from)?;
        }
    }
    Ok(())
}

fn replace_xml_value(source: &str, span: &Range<usize>, replacement: &str) -> Result<String> {
    if span.start > span.end || span.end > source.len() {
        return invalid("ODG text source span is invalid");
    }
    let escaped_replacement = quick_xml::escape::escape(replacement);
    let capacity = source
        .len()
        .checked_sub(span.end - span.start)
        .and_then(|size| size.checked_add(escaped_replacement.len()))
        .ok_or_else(|| Error::InvalidFormat("ODG edited content size overflow".to_string()))?;
    if capacity > MAX_OUTPUT_BYTES {
        return invalid("ODG edited content exceeds the output limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG edited content",
            source: allocation_error,
        })?;
    output.push_str(&source[..span.start]);
    output.push_str(&escaped_replacement);
    output.push_str(&source[span.end..]);
    Ok(output)
}

fn stage_content_splice(
    source: &[u8],
    current: &[u8],
    splices: &mut Vec<ContentSplice>,
    range: &Range<usize>,
    replacement: &[u8],
) -> Result<()> {
    let actual = current
        .get(range.clone())
        .ok_or_else(|| Error::InvalidFormat("ODG content splice range is invalid".into()))?;
    if let Some(index) = splices
        .iter()
        .position(|splice| ranges_overlap_or_conflict(&splice.current_range, range))
    {
        if splices[index].current_range != *range || actual != splices[index].replacement {
            return invalid("ODG content splice overlaps an earlier semantic edit");
        }
        let old_end = splices[index].current_range.end;
        let new_end = range
            .start
            .checked_add(replacement.len())
            .ok_or_else(|| Error::InvalidFormat("ODG content splice size overflow".into()))?;
        splices[index].current_range.end = new_end;
        splices[index].replacement = replacement.to_vec();
        shift_current_splices(splices, index, old_end, new_end)?;
        return Ok(());
    }

    let (removed_before, added_before) = splices
        .iter()
        .filter(|splice| splice.current_range.end <= range.start)
        .try_fold((0usize, 0usize), |(removed, added), splice| {
            Ok::<_, Error>((
                removed
                    .checked_add(splice.source_range.len())
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODG content splice size overflow".into())
                    })?,
                added
                    .checked_add(splice.current_range.len())
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODG content splice size overflow".into())
                    })?,
            ))
        })?;
    let source_start = range
        .start
        .checked_sub(added_before)
        .and_then(|value| value.checked_add(removed_before))
        .ok_or_else(|| Error::InvalidFormat("ODG content splice mapping is invalid".into()))?;
    let source_end = source_start
        .checked_add(range.len())
        .ok_or_else(|| Error::InvalidFormat("ODG content splice size overflow".into()))?;
    let source_range = source_start..source_end;
    let expected = source
        .get(source_range.clone())
        .ok_or_else(|| Error::InvalidFormat("ODG content splice source range is invalid".into()))?;
    if expected != actual {
        return invalid("ODG content splice lost exact source provenance");
    }
    let new_end = range
        .start
        .checked_add(replacement.len())
        .ok_or_else(|| Error::InvalidFormat("ODG content splice size overflow".into()))?;
    let insertion_index = splices.len();
    splices.push(ContentSplice {
        source_range,
        current_range: range.start..new_end,
        expected: expected.to_vec(),
        replacement: replacement.to_vec(),
    });
    shift_current_splices(splices, insertion_index, range.end, new_end)?;
    splices.sort_unstable_by_key(|splice| splice.current_range.start);
    Ok(())
}

fn shift_current_splices(
    splices: &mut [ContentSplice],
    changed: usize,
    old_end: usize,
    new_end: usize,
) -> Result<()> {
    for (index, splice) in splices.iter_mut().enumerate() {
        if index == changed || splice.current_range.start < old_end {
            continue;
        }
        splice.current_range.start = splice
            .current_range
            .start
            .checked_sub(old_end)
            .and_then(|offset| new_end.checked_add(offset))
            .ok_or_else(|| Error::InvalidFormat("ODG content splice shift is invalid".into()))?;
        splice.current_range.end = splice
            .current_range
            .end
            .checked_sub(old_end)
            .and_then(|offset| new_end.checked_add(offset))
            .ok_or_else(|| Error::InvalidFormat("ODG content splice shift is invalid".into()))?;
    }
    Ok(())
}

fn ranges_overlap_or_conflict(left: &Range<usize>, right: &Range<usize>) -> bool {
    left.start < right.end && right.start < left.end
        || (left.start == left.end && right.start == right.end && left.start == right.start)
}

fn insert_xml(source: &str, at: usize, xml: &str) -> Result<String> {
    if at > source.len() || !source.is_char_boundary(at) {
        return invalid("ODG XML insertion point is invalid");
    }
    let capacity = source
        .len()
        .checked_add(xml.len())
        .ok_or_else(|| Error::InvalidFormat("ODG edited content size overflow".into()))?;
    if capacity > MAX_OUTPUT_BYTES {
        return invalid("ODG edited content exceeds the output limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|allocation_error| Error::Allocation {
            resource: "ODG edited content",
            source: allocation_error,
        })?;
    output.push_str(&source[..at]);
    output.push_str(xml);
    output.push_str(&source[at..]);
    Ok(output)
}

fn insert_automatic_style(source: &str, style: &str) -> Result<String> {
    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut automatic_depth = None;
    let mut insertion = None;
    let mut body_start = None;
    loop {
        let start = position(&reader)?;
        let (resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG automatic-style owner XML: {error}"))
        })?;
        let namespace = classify(&resolved_namespace);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                depth = checked_xml_depth(depth)?;
                let local = element.local_name();
                if namespace == NamespaceKind::Office && local.as_ref() == b"automatic-styles" {
                    if automatic_depth.replace(depth).is_some() {
                        return invalid("ODG automatic-styles owner is duplicated");
                    }
                } else if namespace == NamespaceKind::Office && local.as_ref() == b"body" {
                    body_start = Some(start);
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Office
                    && local.as_ref() == b"automatic-styles"
                    && insertion.replace(end.saturating_sub(2)).is_some()
                {
                    return invalid("ODG automatic-styles owner is duplicated");
                }
            },
            Event::End(element) => {
                let local = element.local_name();
                if namespace == NamespaceKind::Office
                    && local.as_ref() == b"automatic-styles"
                    && automatic_depth == Some(depth)
                {
                    insertion = Some(start);
                    automatic_depth = None;
                }
                depth = depth.saturating_sub(1);
            },
            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in ODG automatic styles");
            },
            Event::Eof => break,
            Event::CData(_) | Event::Comment(_) | Event::Decl(_) | Event::Text(_) => {},
        }
    }
    if let Some(at) = insertion {
        return insert_child_xml(source, at, style);
    }
    let at = body_start
        .ok_or_else(|| Error::InvalidFormat("ODG office:body source span is missing".into()))?;
    let owner = format!(
        "<office:automatic-styles xmlns:office=\"{}\">{style}</office:automatic-styles>",
        std::str::from_utf8(OFFICE).unwrap_or_default()
    );
    insert_xml(source, at, &owner)
}

fn insert_child_xml(source: &str, at: usize, child: &str) -> Result<String> {
    if source.as_bytes().get(at..at.saturating_add(2)) != Some(b"/>") {
        return insert_xml(source, at, child);
    }
    let element_start = source
        .get(..at)
        .and_then(|prefix| prefix.rfind('<'))
        .ok_or_else(|| Error::InvalidFormat("ODG empty owner start is missing".into()))?;
    let name_start = element_start + 1;
    let name_end = source
        .as_bytes()
        .get(name_start..at)
        .and_then(|bytes| {
            bytes
                .iter()
                .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
                .map(|offset| name_start + offset)
        })
        .unwrap_or(at);
    let name = source
        .get(name_start..name_end)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| Error::InvalidFormat("ODG empty owner name is missing".into()))?;
    let replacement = format!(">{child}</{name}>");
    let mut output = String::with_capacity(
        source
            .len()
            .saturating_sub(2)
            .saturating_add(replacement.len()),
    );
    output.push_str(&source[..at]);
    output.push_str(&replacement);
    output.push_str(&source[at + 2..]);
    if output.len() > MAX_OUTPUT_BYTES {
        return invalid("ODG edited content exceeds the output limit");
    }
    Ok(output)
}

fn remove_xml(source: &str, span: &Range<usize>) -> Result<String> {
    if span.start > span.end
        || span.end > source.len()
        || !source.is_char_boundary(span.start)
        || !source.is_char_boundary(span.end)
    {
        return invalid("ODG XML removal span is invalid");
    }
    let mut output = String::with_capacity(source.len() - (span.end - span.start));
    output.push_str(&source[..span.start]);
    output.push_str(&source[span.end..]);
    Ok(output)
}

fn replace_xml(source: &str, span: &Range<usize>, replacement: &str) -> Result<String> {
    let without = remove_xml(source, span)?;
    insert_xml(&without, span.start, replacement)
}

fn start_tag_end(source: &str, start: usize) -> Result<usize> {
    source
        .get(start..)
        .and_then(|tail| tail.find('>').map(|offset| start + offset + 1))
        .ok_or_else(|| Error::InvalidFormat("ODG element start tag is unterminated".into()))
}

fn serialize_page(page: &Page) -> Result<String> {
    let mut xml = format!(
        "<draw:page xmlns:draw=\"{}\"",
        std::str::from_utf8(DRAW).unwrap_or_default()
    );
    push_attribute(&mut xml, "draw:name", page.name())?;
    push_attribute(&mut xml, "xml:id", page.xml_id())?;
    push_attribute(&mut xml, "draw:style-name", page.style_name())?;
    push_attribute(&mut xml, "draw:master-page-name", page.master_page_name())?;
    xml.push_str("></draw:page>");
    Ok(xml)
}

fn serialize_layer(layer: &Layer) -> Result<String> {
    validate_bounded_value(layer.name(), "ODG layer name")?;
    let mut xml = String::from("<draw:layer");
    push_attribute(&mut xml, "draw:name", Some(layer.name()))?;
    push_attribute(&mut xml, "draw:display", layer.display())?;
    if let Some(protected) = layer.protected() {
        push_attribute(
            &mut xml,
            "draw:protected",
            Some(if protected { "true" } else { "false" }),
        )?;
    }
    xml.push_str("/>");
    Ok(xml)
}

fn serialize_form_control(control: &FormControl) -> Result<String> {
    validate_bounded_value(control.id(), "ODG form-control identifier")?;
    validate_xml_local_name(control.element(), "ODG form-control element")?;
    let mut xml = format!(
        "<form:{} xmlns:form=\"{}\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"",
        control.element(),
        std::str::from_utf8(FORM).unwrap_or_default()
    );
    push_attribute(&mut xml, "form:id", Some(control.id()))?;
    push_attribute(&mut xml, "form:name", control.name())?;
    for (name, value) in control.attributes() {
        validate_xml_qualified_name(name, "ODG form-control attribute")?;
        if matches!(name.as_str(), "form:id" | "form:name") {
            return invalid("ODG form-control arbitrary attributes duplicate identity");
        }
        push_attribute(&mut xml, name, Some(value))?;
    }
    xml.push_str("/>");
    Ok(xml)
}

fn serialize_style(style: &Style) -> Result<String> {
    validate_bounded_value(style.name(), "ODG style name")?;
    validate_bounded_value(style.family(), "ODG style family")?;
    let mut xml = format!(
        "<style:style xmlns:style=\"{}\" xmlns:draw=\"{}\" xmlns:svg=\"{}\" xmlns:fo=\"{}\"",
        std::str::from_utf8(STYLE).unwrap_or_default(),
        std::str::from_utf8(DRAW).unwrap_or_default(),
        std::str::from_utf8(SVG).unwrap_or_default(),
        std::str::from_utf8(FO).unwrap_or_default()
    );
    push_attribute(&mut xml, "style:name", Some(style.name()))?;
    push_attribute(&mut xml, "style:family", Some(style.family()))?;
    push_attribute(&mut xml, "style:parent-style-name", style.parent())?;
    if style.properties().is_empty() {
        xml.push_str("/>");
        return Ok(xml);
    }
    let mut owners: BTreeMap<&str, Vec<(&str, &str)>> = BTreeMap::new();
    for (path, value) in style.properties() {
        let (owner, name) = path.split_once('/').ok_or_else(|| {
            Error::InvalidFormat("ODG style property owner path is invalid".into())
        })?;
        validate_style_property_owner(owner)?;
        validate_style_property_name(name)?;
        owners
            .entry(owner)
            .or_default()
            .push((name, value.as_str()));
    }
    xml.push('>');
    for (owner, properties) in owners {
        xml.push('<');
        xml.push_str(owner);
        for (name, value) in properties {
            push_attribute(&mut xml, name, Some(value))?;
        }
        xml.push_str("/>");
    }
    xml.push_str("</style:style>");
    Ok(xml)
}

fn validate_style_property_owner(value: &str) -> Result<()> {
    let Some((prefix, local)) = value.split_once(':') else {
        return invalid("ODG style property owner requires a qualified name");
    };
    if prefix != "style" || !local.ends_with("-properties") {
        return invalid("ODG style property owner is unsupported");
    }
    validate_xml_local_name(local, "ODG style property owner")
}

fn validate_style_property_name(value: &str) -> Result<()> {
    let Some((prefix, local)) = value.split_once(':') else {
        return invalid("ODG style property requires a qualified name");
    };
    if !matches!(prefix, "style" | "draw" | "svg" | "fo") {
        return invalid("ODG style property uses an unsupported namespace prefix");
    }
    validate_xml_local_name(local, "ODG style property")
}

fn validate_xml_local_name(value: &str, context: &str) -> Result<()> {
    if value.is_empty()
        || value.len() > 256
        || !value.bytes().enumerate().all(|(index, byte)| {
            byte.is_ascii_alphanumeric() || byte == b'_' || index > 0 && byte == b'-'
        })
    {
        return invalid(format!("{context} is invalid"));
    }
    Ok(())
}

fn validate_xml_qualified_name(value: &str, context: &str) -> Result<()> {
    let Some((prefix, local)) = value.split_once(':') else {
        return invalid(format!("{context} requires a qualified name"));
    };
    if !matches!(prefix, "form" | "xlink") {
        return invalid(format!("{context} uses an unsupported namespace prefix"));
    }
    validate_xml_local_name(local, context)
}

fn serialize_shape(shape: &Shape) -> Result<String> {
    let element = shape.kind().element_name();
    let mut xml = format!(
        "<draw:{element} xmlns:draw=\"{}\" xmlns:svg=\"{}\" xmlns:text=\"{}\"",
        std::str::from_utf8(DRAW).unwrap_or_default(),
        std::str::from_utf8(SVG).unwrap_or_default(),
        std::str::from_utf8(TEXT).unwrap_or_default()
    );
    push_attribute(&mut xml, "draw:name", shape.name())?;
    push_attribute(&mut xml, "draw:layer", shape.layer())?;
    if shape.control_reference().is_some() && shape.kind() != ShapeKind::Control {
        return invalid("ODG draw:control is only supported on detached control shapes");
    }
    push_attribute(&mut xml, "draw:control", shape.control_reference())?;
    push_attribute(&mut xml, "draw:style-name", shape.style_name())?;
    push_attribute(&mut xml, "draw:text-style-name", shape.text_style_name())?;
    if let Some(z_index) = shape.z_index() {
        push_attribute(&mut xml, "draw:z-index", Some(&z_index.to_string()))?;
    }
    push_attribute(&mut xml, "svg:x", shape.x())?;
    push_attribute(&mut xml, "svg:y", shape.y())?;
    push_attribute(&mut xml, "svg:width", shape.width())?;
    push_attribute(&mut xml, "svg:height", shape.height())?;
    if shape.path_data().is_some() && shape.kind() != ShapeKind::Path {
        return invalid("ODG svg:d is only supported on detached path shapes");
    }
    if let Some(path_data) = shape.path_data() {
        validate_path_data(path_data)?;
    }
    push_attribute(&mut xml, "svg:d", shape.path_data())?;
    if shape.title().is_none() && shape.description().is_none() && shape.text().is_empty() {
        xml.push_str("/>");
        return Ok(xml);
    }
    xml.push('>');
    if let Some(title) = shape.title() {
        xml.push_str("<svg:title>");
        xml.push_str(&quick_xml::escape::escape(title));
        xml.push_str("</svg:title>");
    }
    if let Some(description) = shape.description() {
        xml.push_str("<svg:desc>");
        xml.push_str(&quick_xml::escape::escape(description));
        xml.push_str("</svg:desc>");
    }
    if !shape.text().is_empty() {
        xml.push_str("<text:p>");
        xml.push_str(&quick_xml::escape::escape(shape.text()));
        xml.push_str("</text:p>");
    }
    xml.push_str("</draw:");
    xml.push_str(element);
    xml.push('>');
    Ok(xml)
}

fn push_attribute(output: &mut String, name: &str, value: Option<&str>) -> Result<()> {
    let Some(attribute_value) = value else {
        return Ok(());
    };
    validate_bounded_value(attribute_value, "ODG XML attribute value")?;
    output.push(' ');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(&quick_xml::escape::escape(attribute_value));
    output.push('"');
    Ok(())
}

fn validate_bounded_value(value: &str, owner: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_TEXT_BYTES || value.contains('\0') {
        return Err(Error::InvalidFormat(format!("{owner} is invalid")));
    }
    Ok(())
}

fn validate_geometry(values: &[String; 4]) -> Result<()> {
    for value in values {
        validate_bounded_value(value, "ODG geometry value")?;
        if value.bytes().any(|byte| byte.is_ascii_whitespace()) {
            return invalid("ODG geometry value contains whitespace");
        }
    }
    Ok(())
}

fn validate_path_data(value: &str) -> Result<()> {
    validate_bounded_value(value, "ODG path data")?;
    if value.chars().any(char::is_control) {
        return invalid("ODG path data contains a control character");
    }
    Ok(())
}

fn validate_shape_layer(page: &Page, global_layers: &[Layer], shape: &Shape) -> Result<()> {
    let Some(layer) = shape.layer() else {
        return Ok(());
    };
    let visible = if page.has_layer_set() {
        page.layers()
    } else {
        global_layers
    };
    if !visible.iter().any(|value| value.name() == layer) {
        return invalid("ODG inserted shape references an undeclared layer");
    }
    Ok(())
}

fn resolve_transfer_layer(snapshot: &Snapshot, page: &Page, name: &str) -> Result<Layer> {
    let visible = if page.has_layer_set() {
        page.layers()
    } else {
        snapshot.layers()
    };
    visible
        .iter()
        .find(|layer| layer.name() == name)
        .cloned()
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "ODG transfer shape references undeclared layer '{name}'"
            ))
        })
}

fn transfer_xml_references(xml: &str, href: &str) -> bool {
    let escaped = quick_xml::escape::escape(href);
    xml.contains(&format!("xlink:href=\"{escaped}\""))
        || xml.contains(&format!("xlink:href='{escaped}'"))
}

fn declares_style(snapshot: &Snapshot, name: &str) -> Result<bool> {
    if declares_style_xml(snapshot.content_xml(), name)? {
        return Ok(true);
    }
    snapshot
        .styles_xml()
        .map(|xml| declares_style_xml(xml, name))
        .transpose()
        .map(Option::unwrap_or_default)
}

fn declares_style_xml(xml: &str, name: &str) -> Result<bool> {
    parse_style_definitions(xml).map(|definitions| {
        definitions
            .iter()
            .any(|definition| definition.style.name() == name)
    })
}

fn find_style_definition(snapshot: &Snapshot, name: &str) -> Result<Option<ParsedStyleDefinition>> {
    let mut matches = parse_style_definitions(snapshot.content_xml())?
        .into_iter()
        .filter(|definition| definition.style.name() == name)
        .collect::<Vec<_>>();
    if matches.len() > 1 {
        return invalid("ODG style definition is ambiguous");
    }
    if let Some(definition) = matches.pop() {
        return Ok(Some(definition));
    }
    if let Some(styles) = snapshot.styles_xml() {
        let mut style_matches = parse_style_definitions(styles)?
            .into_iter()
            .filter(|definition| definition.style.name() == name)
            .collect::<Vec<_>>();
        if style_matches.len() > 1 {
            return invalid("ODG style definition is ambiguous");
        }
        return Ok(style_matches.pop());
    }
    Ok(None)
}

fn style_parent_name(xml: &str) -> Result<Option<String>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    loop {
        let (namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG transferred style XML: {error}"))
        })?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if classify(&namespace) == NamespaceKind::Style
                    && element.local_name().as_ref() == b"style" =>
            {
                return attribute(&reader, &element, STYLE, b"parent-style-name");
            },
            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in ODG transferred styles");
            },
            Event::Eof => return Ok(None),
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::Start(_)
            | Event::Empty(_) => {},
        }
    }
}

fn xml_has_attribute(xml: &str, namespace: &[u8], local: &[u8], value: &str) -> Result<bool> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    loop {
        let (_resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG dependency XML: {error}"))
        })?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if attribute(&reader, &element, namespace, local)?.as_deref() == Some(value) {
                    return Ok(true);
                }
            },
            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in ODG dependencies");
            },
            Event::Eof => return Ok(false),
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::End(_)
            | Event::Text(_) => {},
        }
    }
}

fn resolve_page_position(pages: &[Page], selector: crate::page::Selector<'_>) -> Result<usize> {
    match selector {
        crate::page::Selector::Position(position) => pages
            .get(position.get())
            .map(|_| position.get())
            .ok_or_else(|| Error::InvalidFormat("ODG page selector is out of bounds".into())),
        crate::page::Selector::Name(name) => {
            let mut matches = pages
                .iter()
                .enumerate()
                .filter(|(_, page)| page.name() == Some(name.as_ref()));
            let selected = matches
                .next()
                .ok_or_else(|| Error::InvalidFormat("ODG page selector did not match".into()))?;
            if matches.next().is_some() {
                return invalid("ODG page name selector is ambiguous");
            }
            Ok(selected.0)
        },
    }
}

fn validate_media_type(media_type: &str) -> Result<()> {
    if media_type.is_empty()
        || media_type.len() > 1_024
        || !media_type.is_ascii()
        || !media_type.contains('/')
        || media_type
            .bytes()
            .any(|byte| byte.is_ascii_control() || byte.is_ascii_whitespace())
    {
        return invalid("ODG resource media type is invalid");
    }
    Ok(())
}

fn validate_resource_path(path: &str) -> Result<()> {
    if path.is_empty()
        || path.len() > 4_096
        || path.starts_with('/')
        || path.ends_with('/')
        || path.contains('\\')
        || path.contains(':')
        || path
            .split('/')
            .any(|segment| segment.is_empty() || matches!(segment, "." | ".."))
        || path.bytes().any(|byte| byte.is_ascii_control())
    {
        return invalid("ODG resource path is unsafe");
    }
    Ok(())
}

fn unique_collision_name(base: &str, bytes: &[u8], reserved: &[String]) -> String {
    let fingerprint = DiagnosticFingerprint::of(bytes).as_hex();
    let suffix = &fingerprint[..12];
    let candidate = format!("{base}_litchi_{suffix}");
    if !reserved.iter().any(|name| name == &candidate) {
        return candidate;
    }
    for ordinal in 2usize..=MAX_TRANSFER_RESOURCES {
        let numbered_candidate = format!("{base}_litchi_{suffix}_{ordinal}");
        if !reserved.iter().any(|name| name == &numbered_candidate) {
            return numbered_candidate;
        }
    }
    format!("{base}_litchi_{fingerprint}")
}

fn unique_resource_path(
    source: &Snapshot,
    staged: &[ResourceEdit],
    path: &str,
    bytes: &[u8],
) -> Result<String> {
    let (stem, extension) = path
        .rsplit_once('.')
        .map_or((path, ""), |(stem, extension)| (stem, extension));
    let fingerprint = DiagnosticFingerprint::of(bytes).as_hex();
    for width in [12usize, 24, 64] {
        let suffix = &fingerprint[..width];
        let candidate = if extension.is_empty() {
            format!("{stem}_litchi_{suffix}")
        } else {
            format!("{stem}_litchi_{suffix}.{extension}")
        };
        validate_resource_path(&candidate)?;
        if !source.files()?.iter().any(|value| value == &candidate)
            && !staged.iter().any(|edit| edit.path == candidate)
        {
            return Ok(candidate);
        }
    }
    invalid("ODG transferred resource collision could not be remapped")
}

fn rewrite_qualified_attribute_values(
    xml: &str,
    qualified_names: &[&[u8]],
    before: &str,
    after: &str,
) -> Result<String> {
    if before == after {
        return Ok(xml.to_owned());
    }
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut spans = Vec::new();
    loop {
        let start = position_reader(&reader)?;
        let event = reader.read_event().map_err(|error| {
            Error::InvalidFormat(format!("invalid ODG transfer fragment XML: {error}"))
        })?;
        let end = position_reader(&reader)?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let tag = xml.as_bytes().get(start..end).ok_or_else(|| {
                    Error::InvalidFormat("ODG transfer start tag span is invalid".into())
                })?;
                for raw_attribute in element.attributes() {
                    let parsed = raw_attribute.map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODG transfer attribute: {error}"))
                    })?;
                    if !qualified_names
                        .iter()
                        .any(|name| parsed.key.as_ref() == *name)
                    {
                        continue;
                    }
                    let value = parsed
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| {
                            Error::InvalidFormat(format!(
                                "invalid ODG transfer attribute value: {error}"
                            ))
                        })?;
                    if value == before {
                        let (value_start, value_end) =
                            attribute_value_span(tag, parsed.key.as_ref())?;
                        spans.push(start + value_start..start + value_end);
                    }
                }
            },
            Event::DocType(_) | Event::GeneralRef(_) | Event::PI(_) => {
                return invalid("active XML is prohibited in ODG transfer fragments");
            },
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::End(_)
            | Event::Text(_) => {},
        }
    }
    let mut output = xml.to_owned();
    spans.sort_unstable_by_key(|span| std::cmp::Reverse(span.start));
    for span in spans {
        output = replace_xml_value(&output, &span, after)?;
    }
    Ok(output)
}

fn ensure_transfer_namespaces(xml: &str, namespaces: &[(&str, &[u8])]) -> Result<String> {
    let tag_end = xml
        .find('>')
        .ok_or_else(|| Error::InvalidFormat("ODG transfer fragment start tag is missing".into()))?;
    let tag = &xml[..tag_end];
    let name_end = xml
        .bytes()
        .enumerate()
        .skip(1)
        .find(|(_index, byte)| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
        .map(|(index, _byte)| index)
        .ok_or_else(|| Error::InvalidFormat("ODG transfer fragment name is invalid".into()))?;
    let mut declarations = String::new();
    for (prefix, namespace_uri) in namespaces {
        if !tag.contains(&format!("xmlns:{prefix}=")) {
            let namespace_text = std::str::from_utf8(namespace_uri).map_err(|error| {
                Error::InvalidFormat(format!("ODG transfer namespace is invalid: {error}"))
            })?;
            write!(declarations, " xmlns:{prefix}=\"{namespace_text}\"").map_err(|error| {
                Error::InvalidFormat(format!("ODG transfer namespace write failed: {error}"))
            })?;
        }
    }
    insert_xml(xml, name_end, &declarations)
}

fn inverse_change(change: &Change) -> Change {
    match change {
        Change::ControlReference(value) => Change::ControlReference(ControlReferenceChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Text(value) => Change::Text(TextChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Name(value) => Change::Name(NameChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Layer(value) => Change::Layer(LayerChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Geometry(value) => Change::Geometry(GeometryChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Style(value) => Change::Style(StyleChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Path(value) => Change::Path(PathChange {
            page: value.page,
            shape: value.shape,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::PageName(value) => Change::PageName(PageNameChange {
            page: value.page,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::PageStyle(value) => Change::PageStyle(PageStyleChange {
            page: value.page,
            before: value.after.clone(),
            after: value.before.clone(),
        }),
        Change::Structure(value) => Change::Structure(match value {
            StructureChange::PageInserted { position, name } => StructureChange::PageRemoved {
                position: *position,
                name: name.clone(),
            },
            StructureChange::PageRemoved { position, name } => StructureChange::PageInserted {
                position: *position,
                name: name.clone(),
            },
            StructureChange::LayerInserted { page, name } => StructureChange::LayerRemoved {
                page: *page,
                name: name.clone(),
            },
            StructureChange::LayerRemoved { page, name } => StructureChange::LayerInserted {
                page: *page,
                name: name.clone(),
            },
            StructureChange::ShapeInserted {
                page,
                position,
                kind,
            } => StructureChange::ShapeRemoved {
                page: *page,
                position: *position,
                kind: *kind,
            },
            StructureChange::ShapeRemoved {
                page,
                position,
                kind,
            } => StructureChange::ShapeInserted {
                page: *page,
                position: *position,
                kind: *kind,
            },
            StructureChange::FormControlInserted { id } => {
                StructureChange::FormControlRemoved { id: id.clone() }
            },
            StructureChange::FormControlRemoved { id } => {
                StructureChange::FormControlInserted { id: id.clone() }
            },
            StructureChange::FormControlReplaced { id } => {
                StructureChange::FormControlReplaced { id: id.clone() }
            },
            StructureChange::StyleInserted { name } => {
                StructureChange::StyleRemoved { name: name.clone() }
            },
            StructureChange::StyleRemoved { name } => {
                StructureChange::StyleInserted { name: name.clone() }
            },
            StructureChange::StyleReplaced { name } => {
                StructureChange::StyleReplaced { name: name.clone() }
            },
        }),
    }
}

fn inverse_resource_change(change: &ResourceChange) -> ResourceChange {
    ResourceChange {
        resource: change.resource,
        path: change.path.clone(),
        before_media_type: change.after_media_type.clone(),
        after_media_type: change.before_media_type.clone(),
        before_size: change.after_size,
        after_size: change.before_size,
    }
}

fn frame(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: Option<String>,
    page_name: Option<String>,
) -> Result<Frame> {
    Ok(Frame {
        name,
        xml_id: attribute(reader, element, XML, b"id")?,
        title: None,
        description: None,
        anchor_type: attribute(reader, element, TEXT, b"anchor-type")?,
        x: attribute(reader, element, SVG, b"x")?,
        y: attribute(reader, element, SVG, b"y")?,
        width: attribute(reader, element, SVG, b"width")?,
        height: attribute(reader, element, SVG, b"height")?,
        end_cell_address: None,
        page_name,
        sheet_name: None,
        sheet_shape: false,
    })
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    let mut value = None;
    for raw_attribute in element.attributes() {
        let parsed_attribute = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(parsed_attribute.key);
        if resolved_bound(&namespace, expected) && name.as_ref() == local {
            let decoded = parsed_attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODG attribute value: {error}"))
                })?
                .into_owned();
            if value.replace(decoded).is_some() {
                return invalid("ODG element has a duplicate namespaced attribute");
            }
        }
    }
    Ok(value)
}

fn arbitrary_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    excluded: &[(&[u8], &[u8])],
) -> Result<BTreeMap<String, String>> {
    let mut values = BTreeMap::new();
    for raw_attribute in element.attributes() {
        let parsed = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG attribute: {error}")))?;
        let raw_name = parsed.key.as_ref();
        if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(parsed.key);
        if excluded.iter().any(|(expected, wanted)| {
            resolved_bound(&namespace, expected) && local.as_ref() == *wanted
        }) {
            continue;
        }
        let name = std::str::from_utf8(raw_name)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG attribute name: {error}")))?
            .to_owned();
        let value = parsed
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG attribute value: {error}")))?
            .into_owned();
        if values.insert(name, value).is_some() {
            return invalid("ODG element has a duplicate arbitrary attribute");
        }
    }
    Ok(values)
}

fn required_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
    owner: &str,
) -> Result<String> {
    attribute(reader, element, expected, local)?
        .ok_or_else(|| Error::InvalidFormat(format!("ODG {owner} requires a namespaced attribute")))
}

fn optional_bool_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
) -> Result<Option<bool>> {
    attribute(reader, element, expected, local)?
        .map(|value| match value.as_str() {
            "false" => Ok(false),
            "true" => Ok(true),
            _ => invalid("ODG Boolean attribute is not true or false"),
        })
        .transpose()
}

fn optional_u32_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected: &[u8],
    local: &[u8],
) -> Result<Option<u32>> {
    attribute(reader, element, expected, local)?
        .map(|value| {
            value.parse::<u32>().map_err(|_error| {
                Error::InvalidFormat("ODG integer attribute is invalid".to_string())
            })
        })
        .transpose()
}

fn shape_name_span(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: &[u8],
    tag_start: usize,
) -> Result<Option<Range<usize>>> {
    attribute_source_span(reader, element, tag, tag_start, DRAW, b"name")
}

fn attribute_source_span(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: &[u8],
    tag_start: usize,
    expected: &[u8],
    wanted_local: &[u8],
) -> Result<Option<Range<usize>>> {
    let mut key = None;
    for raw_attribute in element.attributes() {
        let parsed_attribute = raw_attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid ODG attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(parsed_attribute.key);
        if resolved_bound(&namespace, expected)
            && local.as_ref() == wanted_local
            && key
                .replace(parsed_attribute.key.as_ref().to_vec())
                .is_some()
        {
            return invalid("ODG element has duplicate namespaced attributes");
        }
    }
    let Some(name_key) = key else {
        return Ok(None);
    };
    let (start, end) = attribute_value_span(tag, &name_key)?;
    Ok(Some(tag_start + start..tag_start + end))
}

fn attribute_value_span(tag: &[u8], wanted: &[u8]) -> Result<(usize, usize)> {
    let mut cursor = 1usize;
    while cursor < tag.len() && !tag[cursor].is_ascii_whitespace() && tag[cursor] != b'>' {
        cursor += 1;
    }
    while cursor < tag.len() {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return invalid("ODG shape attribute is missing '='");
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'\"'))
            .ok_or_else(|| Error::InvalidFormat("ODG shape attribute is not quoted".to_string()))?;
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= tag.len() {
            return invalid("ODG shape attribute is unterminated");
        }
        cursor += 1;
        if &tag[name_start..name_end] == wanted {
            return Ok((value_start, value_end));
        }
    }
    invalid("ODG shape name span was not found")
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_error| {
        Error::InvalidFormat("ODG XML position exceeds platform limits".to_string())
    })
}

fn position_reader(reader: &quick_xml::Reader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_error| {
        Error::InvalidFormat("ODG XML position exceeds platform limits".to_string())
    })
}

fn shape_kind(namespace: NamespaceKind, local: &[u8]) -> Option<ShapeKind> {
    (namespace == NamespaceKind::Draw).then_some(match local {
        b"caption" => ShapeKind::Caption,
        b"circle" => ShapeKind::Circle,
        b"connector" => ShapeKind::Connector,
        b"control" => ShapeKind::Control,
        b"custom-shape" => ShapeKind::Custom,
        b"ellipse" => ShapeKind::Ellipse,
        b"frame" => ShapeKind::Frame,
        b"g" => ShapeKind::Group,
        b"line" => ShapeKind::Line,
        b"measure" => ShapeKind::Measure,
        b"path" => ShapeKind::Path,
        b"page-thumbnail" => ShapeKind::PageThumbnail,
        b"polygon" => ShapeKind::Polygon,
        b"polyline" => ShapeKind::Polyline,
        b"rect" => ShapeKind::Rectangle,
        b"regular-polygon" => ShapeKind::RegularPolygon,
        _ => return None,
    })
}

fn classify(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == DRAW => NamespaceKind::Draw,
        ResolveResult::Bound(Namespace(uri)) if *uri == TEXT => NamespaceKind::Text,
        ResolveResult::Bound(Namespace(uri)) if *uri == SVG => NamespaceKind::Svg,
        ResolveResult::Bound(Namespace(uri)) if *uri == FORM => NamespaceKind::Form,
        ResolveResult::Bound(Namespace(uri)) if *uri == STYLE => NamespaceKind::Style,
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn resolved_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
