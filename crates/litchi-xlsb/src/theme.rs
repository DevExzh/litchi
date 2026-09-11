//! Typed XLSB DrawingML theme ownership.
//!
//! An XLSB workbook stores its theme in the OPC Theme part described by
//! ISO/IEC 29500.  The BIFF12 streams do not contain a second theme grammar;
//! this module therefore owns only the XLSB package edge and delegates the
//! color/font vocabulary to `litchi-drawingml`.
//!
//! Theme snapshots retain the exact source XML.  A detached transaction can
//! replace the typed color or font schemes while preserving unrelated source
//! markup.  If the scheme being changed contains markup outside the shared
//! typed vocabulary, the edit is refused before any package mutation.

use std::ops::Range;
use std::sync::Arc;

use litchi_core::{SourceVersion, xml::escape_xml};
use litchi_drawingml::theme::codec;
use litchi_drawingml::theme::family;
pub use litchi_drawingml::theme::family::Family;
pub use litchi_drawingml::theme::{Color, Face, FontSet, Palette, Slot, Theme};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::part::Part;
use litchi_opc::{
    OpcPackage, PackURI, PartData, PartView, Relationships, SourceBackedPackage, TargetMode,
};
use quick_xml::Reader;
use quick_xml::events::Event;

use crate::package::error::{Error, Result};

/// Transitional DrawingML Theme relationship type.
pub const RELATIONSHIP_TYPE: &str = relationship_type::THEME;
/// Theme part content type shared by Transitional and Strict packages.
pub const CONTENT_TYPE: &str = content_type::OFC_THEME;

const STRICT_THEME_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";
const STRICT_IMAGE_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image";

/// The DrawingML 2012 extension URI that owns `themeFamily`.
pub const THEME_FAMILY_EXTENSION_URI: &str = family::part::NATIVE_EXTENSION_URI;
/// The normative namespace-shaped extension identifier for `themeFamily`.
pub const THEME_FAMILY_NAMESPACE_EXTENSION_URI: &str = family::part::EXTENSION_URI;

/// Hard ceiling inherited from the shared DrawingML theme codec.
pub const MAX_XML_BYTES: usize = codec::MAX_XML_BYTES;
const HARD_MAX_NODES: usize = 100_000;
const HARD_MAX_DEPTH: usize = 128;

/// Finite resource policy for one Theme part and its XML grammar.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    /// Maximum decoded Theme XML bytes.
    pub max_xml_bytes: usize,
    /// Maximum XML element events inspected by the bounded preflight.
    pub max_nodes: usize,
    /// Maximum XML nesting depth.
    pub max_depth: usize,
}

impl Limits {
    /// Conservative finite Theme limits.
    pub const DEFAULT: Self = Self {
        max_xml_bytes: MAX_XML_BYTES,
        max_nodes: HARD_MAX_NODES,
        max_depth: HARD_MAX_DEPTH,
    };

    /// Construct explicit finite limits.
    #[must_use]
    pub const fn new(max_xml_bytes: usize, max_nodes: usize, max_depth: usize) -> Self {
        Self {
            max_xml_bytes,
            max_nodes,
            max_depth,
        }
    }

    /// Validate a caller policy before reading any Theme payload.
    pub(crate) fn validate(self) -> Result<Self> {
        if self.max_xml_bytes == 0 || self.max_xml_bytes > MAX_XML_BYTES {
            return Err(invalid(format!(
                "Theme XML byte limit must be between 1 and {MAX_XML_BYTES}"
            )));
        }
        if self.max_nodes == 0 || self.max_nodes > HARD_MAX_NODES {
            return Err(invalid(format!(
                "Theme node limit must be between 1 and {HARD_MAX_NODES}"
            )));
        }
        if self.max_depth == 0 || self.max_depth > HARD_MAX_DEPTH {
            return Err(invalid(format!(
                "Theme depth limit must be between 1 and {HARD_MAX_DEPTH}"
            )));
        }
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Descriptive alias for callers that prefer the format-specific name.
/// Immutable, exact-source Theme snapshot from an eager [`OpcPackage`].
#[derive(Clone, Debug)]
pub struct Snapshot {
    theme: Arc<Theme>,
    family: Option<Family>,
    source_xml: Arc<Vec<u8>>,
    graph: GraphState,
    limits: Limits,
}

impl Snapshot {
    /// Borrow the typed color/font metadata projected from the DrawingML theme.
    #[must_use]
    pub fn theme(&self) -> &Theme {
        self.theme.as_ref()
    }

    /// Borrow the optional DrawingML 2012 applied-theme family metadata.
    #[must_use]
    pub fn family(&self) -> Option<&Family> {
        self.family.as_ref()
    }

    /// Borrow the exact source XML captured from the Theme part.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Return the finite policy used to parse this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Return the absolute Theme part name for diagnostics and package-owned
    /// integrations.  Ordinary edits do not require this identifier.
    #[must_use]
    #[allow(dead_code, reason = "used by Workbook publication integration")]
    pub(crate) fn part_name(&self) -> &str {
        &self.graph.theme_name
    }

    /// Start a detached source-checked edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            before: self.clone(),
            staged: self.theme.as_ref().clone(),
            staged_family: self.family.clone(),
        }
    }

    fn same_source(&self, other: &Self) -> bool {
        self.graph == other.graph && self.source_xml == other.source_xml
    }

    #[allow(dead_code, reason = "used by Workbook publication integration")]
    fn same_state(&self, other: &Self) -> bool {
        self.same_source(other) && self.theme == other.theme && self.family == other.family
    }
}

/// Source-backed Theme view retaining the managed [`PartData`] reservation.
///
/// The view owns the parsed typed model and the source payload handle.  It
/// does not detach a managed allocation into an unreserved `Arc`; callers who
/// need an edit use [`Self::edit`], which intentionally creates a detached
/// snapshot.
#[derive(Clone, Debug)]
pub struct View {
    theme: Arc<Theme>,
    family: Option<Family>,
    source: PartData,
    graph: GraphState,
    version: SourceVersion,
    limits: Limits,
}

impl View {
    /// Borrow the typed color/font metadata projected from the DrawingML theme.
    #[must_use]
    pub fn theme(&self) -> &Theme {
        self.theme.as_ref()
    }

    /// Borrow the optional DrawingML 2012 applied-theme family metadata.
    #[must_use]
    pub fn family(&self) -> Option<&Family> {
        self.family.as_ref()
    }

    /// Borrow the exact Theme XML captured from the source package.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source.as_bytes()
    }

    /// Return the Theme part name without exposing OPC URI types in the
    /// ordinary semantic API.
    #[must_use]
    pub fn part_name(&self) -> &str {
        &self.graph.theme_name
    }

    /// Return the source version captured while opening this view.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.version
    }

    /// Return the finite policy used to parse this view.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Start a detached edit over the captured source.
    ///
    /// The detached transaction copies only this bounded Theme XML part;
    /// the source-backed view itself continues to retain its managed payload
    /// reservation.
    pub fn edit(&self) -> Result<Transaction> {
        let source_xml = copy_bytes(self.source.as_bytes(), "Theme transaction source")?;
        Ok(Transaction {
            before: Snapshot {
                theme: Arc::clone(&self.theme),
                family: self.family.clone(),
                source_xml: Arc::new(source_xml),
                graph: self.graph.clone(),
                limits: self.limits,
            },
            staged: self.theme.as_ref().clone(),
            staged_family: self.family.clone(),
        })
    }
}

/// A detached, source-checked Theme draft.
#[derive(Clone, Debug)]
pub struct Transaction {
    before: Snapshot,
    staged: Theme,
    staged_family: Option<Family>,
}

impl Transaction {
    /// Snapshot used for stale-source checks.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged typed color/font metadata.
    #[must_use]
    pub fn theme(&self) -> &Theme {
        &self.staged
    }

    /// Borrow the currently staged applied-theme family metadata.
    #[must_use]
    pub fn family(&self) -> Option<&Family> {
        self.staged_family.as_ref()
    }

    /// Replace the typed color/font metadata after bounded validation.
    pub fn replace(&mut self, theme: Theme) -> Result<bool> {
        if self.staged == theme {
            return Ok(false);
        }
        validate_model(&theme, self.before.limits)?;
        self.staged = theme;
        Ok(true)
    }

    /// Change the Theme display name.
    pub fn set_name(&mut self, name: impl Into<String>) -> Result<bool> {
        let mut candidate = self.staged.clone();
        candidate.name = name.into();
        self.replace(candidate)
    }

    /// Replace the complete twelve-slot color palette.
    pub fn set_palette(&mut self, colors: Palette) -> Result<bool> {
        let mut candidate = self.staged.clone();
        candidate.colors = colors;
        self.replace(candidate)
    }

    /// Replace the complete major/minor font set.
    pub fn set_fonts(&mut self, fonts: FontSet) -> Result<bool> {
        let mut candidate = self.staged.clone();
        candidate.fonts = fonts;
        self.replace(candidate)
    }

    /// Set or replace the applied-theme family metadata.
    pub fn set_family(&mut self, family: Family) -> Result<bool> {
        let family = preserve_family_source(self.before.family.as_ref(), family)?;
        if self.staged_family.as_ref() == Some(&family) {
            return Ok(false);
        }
        self.staged_family = Some(family);
        Ok(true)
    }

    /// Remove the applied-theme family metadata while retaining the Theme.
    pub fn remove_family(&mut self) -> Result<bool> {
        if self.staged_family.is_none() {
            return Ok(false);
        }
        self.staged_family = None;
        Ok(true)
    }

    /// Commit the detached draft into an exact, reversible XML patch.
    pub fn commit(self) -> Result<Commit> {
        if self.staged == *self.before.theme && self.staged_family == self.before.family {
            let before = self.before;
            return Ok(Commit::new(Patch::new(before.clone(), before), false));
        }
        validate_model(&self.staged, self.before.limits)?;
        let source = rewrite_source_with_family(
            self.before.source_xml(),
            self.before.theme(),
            &self.staged,
            self.before.family.as_ref(),
            self.staged_family.as_ref(),
            self.before.limits,
        )?;
        let (parsed, parsed_family) = parse_theme_content(&source, self.before.limits)?;
        if parsed != self.staged {
            return Err(invalid(
                "Theme transaction read-back did not match the staged typed model",
            ));
        }
        if parsed_family.as_ref() != self.staged_family.as_ref() {
            return Err(invalid(
                "Theme transaction read-back did not match the staged family metadata",
            ));
        }
        let after = Snapshot {
            theme: Arc::new(parsed),
            family: parsed_family,
            source_xml: Arc::new(source),
            graph: self.before.graph.clone(),
            limits: self.before.limits,
        };
        Ok(Commit::new(Patch::new(self.before, after), true))
    }
}

/// A detached Theme transaction result.
#[derive(Clone, Debug)]
pub struct Commit {
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn new(patch: Patch, changed: bool) -> Self {
        Self { patch, changed }
    }

    /// Whether the Theme XML bytes change.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Planned resulting snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        self.patch.after()
    }

    /// Reversible source-checked patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its planned snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        let snapshot = self.patch.after().clone();
        (snapshot, self.patch)
    }
}

/// A source-checked, reversible Theme-part replacement.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
    changed: bool,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        let changed = !before.same_source(&after);
        Self {
            before,
            after,
            changed,
        }
    }

    /// Source state required before publication.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Exact state produced by publication.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether applying this patch changes no Theme bytes or graph metadata.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        !self.changed
    }

    /// Return an exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            changed: self.changed,
        }
    }

    /// Apply atomically after checking the complete Theme source closure.
    #[allow(dead_code, reason = "called by Workbook publication integration")]
    pub(crate) fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        let current = read(package, self.before.limits)?;
        let Some(current) = current else {
            return Err(invalid("Theme patch source is missing its Theme part"));
        };
        if !current.same_source(&self.before) {
            return Err(invalid("Theme patch source is stale"));
        }
        if self.is_empty() {
            return Ok(current);
        }
        if package.is_signed() || package.requires_signature_edit_policy() {
            return Err(Error::Opc(
                litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy,
            ));
        }
        let mut candidate = package.clone();
        let part_name = PackURI::new(self.after.part_name().to_owned())
            .map_err(|error| Error::InvalidUri(error.to_string()))?;
        {
            let part = candidate.get_part(&part_name)?;
            if part.content_type() != CONTENT_TYPE {
                return Err(Error::InvalidContentType {
                    expected: CONTENT_TYPE.to_owned(),
                    got: part.content_type().to_owned(),
                });
            }
        }
        candidate.unsign();
        replace_source_xml_part(
            &mut candidate,
            &part_name,
            self.before.source_xml(),
            self.after.source_xml(),
        )?;
        let resulting = read(&candidate, self.before.limits)?
            .ok_or_else(|| invalid("Theme patch read-back lost the Theme part"))?;
        if !resulting.same_state(&self.after) {
            return Err(invalid(
                "Theme patch read-back changed the planned Theme state",
            ));
        }
        *package = candidate;
        Ok(resulting)
    }
}

/// Read the optional Theme owner from an eager OPC package.
pub fn read(package: &OpcPackage, limits: Limits) -> Result<Option<Snapshot>> {
    let limits = limits.validate()?;
    let Some((graph, part)) = locate_eager(package)? else {
        return Ok(None);
    };
    let xml = part.blob();
    let (theme, family) = parse_theme_content(xml, limits)?;
    validate_conformance(xml, &graph.workbook_relationship.reltype)?;
    Ok(Some(Snapshot {
        theme: Arc::new(theme),
        family,
        source_xml: part.blob_arc(),
        graph,
        limits,
    }))
}

/// Read the optional Theme owner from a deferred source-backed OPC package.
///
/// This is crate-visible for `SourceBackedWorkbook`; callers should use that
/// facade so its normal source/cancellation preflight surrounds the call.
pub(crate) fn read_source(package: &SourceBackedPackage, limits: Limits) -> Result<Option<View>> {
    let limits = limits.validate()?;
    package.check_execution()?;
    let version = package.source_version()?;
    let workbook = package.main_document_part()?;
    let Some((graph, part)) = locate_source(package, &workbook)? else {
        package.check_execution()?;
        package.source_version()?;
        return Ok(None);
    };
    let declared = part.declared_uncompressed_size()?;
    let maximum =
        u64::try_from(limits.max_xml_bytes).map_err(|_error| Error::CapacityOverflow {
            resource: "Theme XML limit",
        })?;
    if declared > maximum {
        return Err(Error::LimitExceeded {
            resource: "Theme XML bytes",
            actual: usize::try_from(declared).unwrap_or(usize::MAX),
            maximum: limits.max_xml_bytes,
        });
    }
    let source = part.data()?;
    let (theme, family) = parse_theme_content(source.as_bytes(), limits)?;
    validate_conformance(source.as_bytes(), &graph.workbook_relationship.reltype)?;
    let view = View {
        theme: Arc::new(theme),
        family,
        source,
        graph,
        version,
        limits,
    };
    package.check_execution()?;
    package.source_version()?;
    Ok(Some(view))
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct GraphState {
    workbook_name: String,
    theme_name: String,
    workbook_relationship: RelationshipState,
    theme_relationships: Vec<RelationshipState>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct RelationshipState {
    id: String,
    reltype: String,
    target_ref: String,
    target_mode: TargetMode,
    target_name: Option<String>,
}

fn locate_eager(package: &OpcPackage) -> Result<Option<(GraphState, &dyn Part)>> {
    let workbook = package.main_document_part()?;
    if workbook.content_type() != content_type::XLSB_BIN {
        return Err(Error::InvalidContentType {
            expected: content_type::XLSB_BIN.to_owned(),
            got: workbook.content_type().to_owned(),
        });
    }
    let mut theme_parts = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE);
    let theme_part = theme_parts.next();
    if theme_parts.next().is_some() {
        return Err(Error::InvalidRelationship(
            "XLSB package contains multiple Theme parts".to_string(),
        ));
    }
    let Some(relationship) = find_theme_relationship(workbook.rels())? else {
        if theme_part.is_some() {
            return Err(Error::InvalidRelationship(
                "Theme part is orphaned from the workbook".to_string(),
            ));
        }
        return Ok(None);
    };
    if relationship.is_external() {
        return Err(Error::InvalidRelationship(
            "workbook Theme relationship must be internal".to_string(),
        ));
    }
    let theme_name = relationship.target_partname()?;
    let part = package.get_part(&theme_name)?;
    if theme_part.is_none_or(|candidate| !candidate.partname().is_equivalent_to(&theme_name)) {
        return Err(Error::InvalidRelationship(
            "workbook Theme relationship does not identify the unique Theme part".to_string(),
        ));
    }
    if part.content_type() != CONTENT_TYPE {
        return Err(Error::InvalidContentType {
            expected: CONTENT_TYPE.to_owned(),
            got: part.content_type().to_owned(),
        });
    }
    let theme_relationships = validate_theme_relationships(
        part.rels(),
        relationship.reltype() == STRICT_THEME_RELATIONSHIP,
        |target| {
            let image = package.get_part(target)?;
            validate_image_target(image.content_type(), image.rels().is_empty())
        },
    )?;
    validate_single_theme_inbound(package, workbook.partname(), &theme_name)?;
    let graph = graph_state(
        workbook.partname(),
        relationship,
        &theme_name,
        theme_relationships,
    );
    Ok(Some((graph, part)))
}

fn locate_source<'a>(
    package: &'a SourceBackedPackage,
    workbook: &PartView<'a>,
) -> Result<Option<(GraphState, PartView<'a>)>> {
    if workbook.content_type() != content_type::XLSB_BIN {
        return Err(Error::InvalidContentType {
            expected: content_type::XLSB_BIN.to_owned(),
            got: workbook.content_type().to_owned(),
        });
    }
    let mut theme_parts = package
        .iter_parts()
        .filter(|part| part.content_type() == CONTENT_TYPE);
    let theme_part = theme_parts.next();
    if theme_parts.next().is_some() {
        return Err(Error::InvalidRelationship(
            "XLSB package contains multiple Theme parts".to_string(),
        ));
    }
    let Some(relationship) = find_theme_relationship(workbook.rels())? else {
        if theme_part.is_some() {
            return Err(Error::InvalidRelationship(
                "Theme part is orphaned from the workbook".to_string(),
            ));
        }
        return Ok(None);
    };
    if relationship.is_external() {
        return Err(Error::InvalidRelationship(
            "workbook Theme relationship must be internal".to_string(),
        ));
    }
    let theme_name = relationship.target_partname()?;
    let part = package.part(&theme_name)?;
    if theme_part.is_none_or(|candidate| !candidate.partname().is_equivalent_to(&theme_name)) {
        return Err(Error::InvalidRelationship(
            "workbook Theme relationship does not identify the unique Theme part".to_string(),
        ));
    }
    if part.content_type() != CONTENT_TYPE {
        return Err(Error::InvalidContentType {
            expected: CONTENT_TYPE.to_owned(),
            got: part.content_type().to_owned(),
        });
    }
    let theme_relationships = validate_theme_relationships(
        part.rels(),
        relationship.reltype() == STRICT_THEME_RELATIONSHIP,
        |target| {
            let image = package.part(target)?;
            validate_image_target(image.content_type(), image.rels().is_empty())
        },
    )?;
    validate_single_theme_inbound_source(package, workbook.partname(), &theme_name)?;
    let graph = graph_state(
        workbook.partname(),
        relationship,
        &theme_name,
        theme_relationships,
    );
    Ok(Some((graph, part)))
}

fn find_theme_relationship(
    relationships: &Relationships,
) -> Result<Option<&litchi_opc::Relationship>> {
    let mut matching = relationships
        .iter()
        .filter(|relationship| is_theme_relationship(relationship.reltype()));
    let Some(first) = matching.next() else {
        return Ok(None);
    };
    if matching.next().is_some() {
        return Err(Error::InvalidRelationship(
            "workbook has multiple Theme relationships".to_string(),
        ));
    }
    Ok(Some(first))
}

fn graph_state(
    workbook: &PackURI,
    workbook_relationship: &litchi_opc::Relationship,
    theme: &PackURI,
    mut theme_relationships: Vec<RelationshipState>,
) -> GraphState {
    theme_relationships.sort_by(|left, right| left.id.cmp(&right.id));
    GraphState {
        workbook_name: workbook.as_str().to_owned(),
        theme_name: theme.as_str().to_owned(),
        workbook_relationship: relationship_state(workbook_relationship),
        theme_relationships,
    }
}

fn relationship_state(relationship: &litchi_opc::Relationship) -> RelationshipState {
    RelationshipState {
        id: relationship.r_id().to_owned(),
        reltype: relationship.reltype().to_owned(),
        target_ref: relationship.target_ref().to_owned(),
        target_mode: relationship.target_mode(),
        target_name: (!relationship.is_external())
            .then(|| relationship.target_partname().ok())
            .flatten()
            .map(|uri| uri.as_str().to_owned()),
    }
}

fn validate_theme_relationships(
    relationships: &Relationships,
    strict: bool,
    mut target_exists: impl FnMut(&PackURI) -> Result<()>,
) -> Result<Vec<RelationshipState>> {
    let mut states = Vec::new();
    states
        .try_reserve(relationships.len())
        .map_err(|source| Error::Allocation {
            resource: "Theme relationship metadata",
            source,
        })?;
    for relationship in relationships.iter() {
        let expected = if strict {
            STRICT_IMAGE_RELATIONSHIP
        } else {
            relationship_type::IMAGE
        };
        if relationship.reltype() != expected {
            return Err(Error::InvalidRelationship(format!(
                "Theme part relationship {:?} is not a {expected} Image relationship",
                relationship.r_id(),
            )));
        }
        if !relationship.is_external() {
            let target = relationship.target_partname()?;
            target_exists(&target)?;
        }
        states.push(relationship_state(relationship));
    }
    Ok(states)
}

fn validate_image_target(content_type_value: &str, is_leaf: bool) -> Result<()> {
    if !content_type_value.starts_with("image/") {
        return Err(Error::InvalidContentType {
            expected: "an image/* Theme relationship target".to_string(),
            got: content_type_value.to_string(),
        });
    }
    if !is_leaf {
        return Err(Error::InvalidRelationship(
            "Theme Image relationship target must be a leaf part".to_string(),
        ));
    }
    Ok(())
}

fn validate_single_theme_inbound(
    package: &OpcPackage,
    workbook_name: &PackURI,
    theme_name: &PackURI,
) -> Result<()> {
    let mut inbound = 0usize;
    for relationship in package.rels().iter() {
        if !relationship.is_external()
            && relationship.target_partname()?.is_equivalent_to(theme_name)
        {
            return Err(Error::InvalidRelationship(
                "Theme part has an inbound package-root relationship".to_string(),
            ));
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            if target.is_equivalent_to(theme_name) {
                inbound = inbound.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "Theme inbound relationship count",
                })?;
                if !part.partname().is_equivalent_to(workbook_name)
                    || !is_theme_relationship(relationship.reltype())
                {
                    return Err(Error::InvalidRelationship(
                        "Theme part has an inbound relationship other than the workbook Theme edge"
                            .to_string(),
                    ));
                }
            }
        }
    }
    if inbound != 1 {
        return Err(Error::InvalidRelationship(format!(
            "Theme part must have exactly one workbook inbound relationship, found {inbound}"
        )));
    }
    Ok(())
}

fn validate_single_theme_inbound_source(
    package: &SourceBackedPackage,
    workbook_name: &PackURI,
    theme_name: &PackURI,
) -> Result<()> {
    let mut inbound = 0usize;
    for relationship in package.rels().iter() {
        if !relationship.is_external()
            && relationship.target_partname()?.is_equivalent_to(theme_name)
        {
            return Err(Error::InvalidRelationship(
                "Theme part has an inbound package-root relationship".to_string(),
            ));
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            if target.is_equivalent_to(theme_name) {
                inbound = inbound.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "Theme inbound relationship count",
                })?;
                if !part.partname().is_equivalent_to(workbook_name)
                    || !is_theme_relationship(relationship.reltype())
                {
                    return Err(Error::InvalidRelationship(
                        "Theme part has an inbound relationship other than the workbook Theme edge"
                            .to_string(),
                    ));
                }
            }
        }
    }
    if inbound != 1 {
        return Err(Error::InvalidRelationship(format!(
            "Theme part must have exactly one workbook inbound relationship, found {inbound}"
        )));
    }
    Ok(())
}

fn validate_conformance(xml: &[u8], relationship_type_value: &str) -> Result<()> {
    let namespace = source_namespace(xml)?
        .ok_or_else(|| invalid("Theme XML root namespace is missing or unsupported"))?;
    let expected = if relationship_type_value == STRICT_THEME_RELATIONSHIP {
        codec::STRICT_NAMESPACE
    } else {
        codec::NAMESPACE
    };
    if namespace != expected {
        return Err(Error::InvalidRelationship(format!(
            "Theme relationship conformance does not match Theme XML namespace: expected {expected}"
        )));
    }
    let mut reader = quick_xml::reader::NsReader::from_reader(xml);
    reader.resolver_mut().set_max_declarations_per_element(256);
    loop {
        let (resolved, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid Theme XML: {error}")))?;
        if matches!(event, Event::Start(_) | Event::Empty(_))
            && let quick_xml::name::ResolveResult::Bound(value) = resolved
            && (value.as_ref() == codec::NAMESPACE.as_bytes()
                || value.as_ref() == codec::STRICT_NAMESPACE.as_bytes())
            && value.as_ref() != expected.as_bytes()
        {
            return Err(invalid(
                "Theme XML mixes Strict and Transitional DrawingML elements",
            ));
        }
        if matches!(event, Event::Eof) {
            break;
        }
    }
    Ok(())
}

fn is_theme_relationship(value: &str) -> bool {
    matches!(value, RELATIONSHIP_TYPE | STRICT_THEME_RELATIONSHIP)
}

fn is_image_relationship(value: &str) -> bool {
    matches!(value, relationship_type::IMAGE | STRICT_IMAGE_RELATIONSHIP)
}

fn parse_theme(xml: &[u8], limits: Limits) -> Result<Theme> {
    bounded_xml_preflight(xml, limits)?;
    Ok(codec::read(xml)?)
}

pub(crate) fn parse_theme_content(xml: &[u8], limits: Limits) -> Result<(Theme, Option<Family>)> {
    let theme = parse_theme(xml, limits)?;
    let family = family::part::read_family(xml)?;
    Ok((theme, family))
}

fn validate_model(theme: &Theme, limits: Limits) -> Result<()> {
    // `theme@name` is optional in the DrawingML schema.  The shared authoring
    // helper requires a non-empty name, so use a validation-only placeholder
    // when editing a source that omitted the attribute; rewrite_source keeps
    // that source spelling intact unless the caller explicitly changes it.
    let name = if theme.name.is_empty() {
        "Theme"
    } else {
        theme.name.as_str()
    };
    let xml = codec::encode_part(name, &theme.colors, &theme.fonts)?;
    if xml.len() > limits.max_xml_bytes {
        return Err(Error::LimitExceeded {
            resource: "generated Theme XML bytes",
            actual: xml.len(),
            maximum: limits.max_xml_bytes,
        });
    }
    Ok(())
}

fn bounded_xml_preflight(xml: &[u8], limits: Limits) -> Result<()> {
    if xml.len() > limits.max_xml_bytes {
        return Err(Error::LimitExceeded {
            resource: "Theme XML bytes",
            actual: xml.len(),
            maximum: limits.max_xml_bytes,
        });
    }
    let mut reader = Reader::from_reader(xml);
    let mut stack = Vec::<Vec<u8>>::new();
    let mut nodes = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid Theme XML: {error}")))?;
        match event {
            Event::Start(element) => {
                nodes = nodes.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "Theme XML node count",
                })?;
                if nodes > limits.max_nodes {
                    return Err(Error::LimitExceeded {
                        resource: "Theme XML nodes",
                        actual: nodes,
                        maximum: limits.max_nodes,
                    });
                }
                if stack.len() >= limits.max_depth {
                    return Err(Error::LimitExceeded {
                        resource: "Theme XML depth",
                        actual: stack.len() + 1,
                        maximum: limits.max_depth,
                    });
                }
                stack.push(element.local_name().as_ref().to_vec());
            },
            Event::Empty(_) => {
                if stack.len() >= limits.max_depth {
                    return Err(Error::LimitExceeded {
                        resource: "Theme XML depth",
                        actual: stack.len() + 1,
                        maximum: limits.max_depth,
                    });
                }
                nodes = nodes.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "Theme XML node count",
                })?;
                if nodes > limits.max_nodes {
                    return Err(Error::LimitExceeded {
                        resource: "Theme XML nodes",
                        actual: nodes,
                        maximum: limits.max_nodes,
                    });
                }
            },
            Event::End(element) => {
                let Some(open) = stack.pop() else {
                    return Err(invalid("Theme XML has an unexpected closing element"));
                };
                if open.as_slice() != element.local_name().as_ref() {
                    return Err(invalid("Theme XML closing element does not match"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("Theme XML contains forbidden markup"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("Theme XML is unterminated"));
    }
    Ok(())
}

#[cfg(test)]
fn rewrite_source(source: &[u8], before: &Theme, after: &Theme, limits: Limits) -> Result<Vec<u8>> {
    rewrite_source_with_family(source, before, after, None, None, limits)
}

pub(crate) fn rewrite_source_with_family(
    source: &[u8],
    before: &Theme,
    after: &Theme,
    before_family: Option<&Family>,
    after_family: Option<&Family>,
    limits: Limits,
) -> Result<Vec<u8>> {
    if before_family != after_family {
        let patched = rewrite_family_source(source, before_family, after_family, limits)?;
        return rewrite_base_source(&patched, before, after, limits);
    }
    rewrite_base_source(source, before, after, limits)
}

fn rewrite_base_source(
    source: &[u8],
    before: &Theme,
    after: &Theme,
    limits: Limits,
) -> Result<Vec<u8>> {
    let strict = source_namespace(source)? == Some(codec::STRICT_NAMESPACE);
    let mut replacements = Vec::<Replacement>::new();
    if before.name != after.name {
        replacements.push(root_name_replacement(source, &after.name)?);
    }
    if before.colors != after.colors {
        let range = codec::scheme_replacement_range(source, b"clrScheme")?;
        let fragment = strict_fragment(codec::encode_palette_fragment(&after.colors)?, strict)?;
        replacements.push(Replacement {
            range,
            bytes: fragment,
        });
    }
    if before.fonts != after.fonts {
        let range = codec::scheme_replacement_range(source, b"fontScheme")?;
        let fragment = strict_fragment(codec::encode_fonts_fragment(&after.fonts)?, strict)?;
        replacements.push(Replacement {
            range,
            bytes: fragment,
        });
    }
    replacements.sort_by_key(|replacement| replacement.range.start);
    let mut output_len = source.len();
    let mut previous_end = 0usize;
    for replacement in &replacements {
        if replacement.range.start < previous_end
            || replacement.range.start > replacement.range.end
            || replacement.range.end > source.len()
        {
            return Err(invalid(
                "Theme source replacement ranges overlap or exceed the source",
            ));
        }
        output_len = output_len
            .checked_sub(replacement.range.len())
            .and_then(|length| length.checked_add(replacement.bytes.len()))
            .ok_or(Error::CapacityOverflow {
                resource: "patched Theme XML bytes",
            })?;
        previous_end = replacement.range.end;
    }
    if output_len > limits.max_xml_bytes {
        return Err(Error::LimitExceeded {
            resource: "patched Theme XML bytes",
            actual: output_len,
            maximum: limits.max_xml_bytes,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "patched Theme XML bytes",
            source,
        })?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&source[cursor..replacement.range.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.range.end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

fn rewrite_family_source(
    source: &[u8],
    before: Option<&Family>,
    after: Option<&Family>,
    limits: Limits,
) -> Result<Vec<u8>> {
    let output = match (before, after) {
        (Some(_), Some(family)) => {
            family::part::replace_family_with_limit(source, family, limits.max_xml_bytes)
        },
        (Some(_), None) => family::part::remove_family_with_limit(source, limits.max_xml_bytes),
        (None, Some(family)) => family::part::add_family_with_uri_limit(
            source,
            family,
            family::part::NATIVE_EXTENSION_URI,
            limits.max_xml_bytes,
        ),
        (None, None) => Ok(source.to_vec()),
    }?;
    Ok(output)
}

struct Replacement {
    range: Range<usize>,
    bytes: Vec<u8>,
}

fn strict_fragment(mut fragment: Vec<u8>, strict: bool) -> Result<Vec<u8>> {
    if strict {
        // Only the encoder's root namespace declaration changes. A font or
        // scheme name may itself contain the Transitional namespace URI.
        let range = raw_attribute_value_range(&fragment, b"xmlns:a")?;
        if &fragment[range.clone()] != codec::NAMESPACE.as_bytes() {
            return Err(invalid(
                "generated Theme fragment has an unexpected namespace",
            ));
        }
        fragment.splice(range, codec::STRICT_NAMESPACE.bytes());
    }
    Ok(fragment)
}

fn source_namespace(xml: &[u8]) -> Result<Option<&'static str>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::LimitExceeded {
            resource: "Theme XML bytes",
            actual: xml.len(),
            maximum: MAX_XML_BYTES,
        });
    }
    let mut reader = quick_xml::reader::NsReader::from_reader(xml);
    reader.resolver_mut().set_max_declarations_per_element(256);
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid Theme XML: {error}")))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if element.local_name().as_ref() != b"theme" {
                    return Err(invalid("Theme XML has an unexpected root"));
                }
                return Ok(match namespace {
                    quick_xml::name::ResolveResult::Bound(value)
                        if value.as_ref() == codec::STRICT_NAMESPACE.as_bytes() =>
                    {
                        Some(codec::STRICT_NAMESPACE)
                    },
                    quick_xml::name::ResolveResult::Bound(value)
                        if value.as_ref() == codec::NAMESPACE.as_bytes() =>
                    {
                        Some(codec::NAMESPACE)
                    },
                    quick_xml::name::ResolveResult::Bound(_)
                    | quick_xml::name::ResolveResult::Unbound
                    | quick_xml::name::ResolveResult::Unknown(_) => None,
                });
            },
            Event::Eof => return Ok(None),
            Event::DocType(_) | Event::PI(_) | Event::End(_) => {
                return Err(invalid("Theme XML has forbidden markup before its root"));
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn root_name_replacement(xml: &[u8], name: &str) -> Result<Replacement> {
    let encoded = escape_xml(name)
        .replace('\r', "&#13;")
        .replace('\n', "&#10;")
        .replace('\t', "&#9;");
    let mut reader = Reader::from_reader(xml);
    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|_error| {
            Error::CapacityOverflow {
                resource: "Theme XML offset",
            }
        })?;
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid Theme XML: {error}")))?;
        let end = usize::try_from(reader.buffer_position()).map_err(|_error| {
            Error::CapacityOverflow {
                resource: "Theme XML offset",
            }
        })?;
        if let Event::Start(element) = event {
            if element.local_name().as_ref() != b"theme" {
                return Err(invalid("Theme XML has an unexpected root"));
            }
            if let Some(range) = optional_attribute_value_range(&xml[start..end], b"name")? {
                return Ok(Replacement {
                    range: (start + range.start)..(start + range.end),
                    bytes: encoded.into_bytes(),
                });
            }
            let insertion = end
                .checked_sub(1)
                .ok_or_else(|| invalid("Theme root has no closing delimiter"))?;
            return Ok(Replacement {
                range: insertion..insertion,
                bytes: format!(" name=\"{encoded}\"").into_bytes(),
            });
        }
        if matches!(event, Event::Eof) {
            return Err(invalid("Theme XML has no root"));
        }
    }
}

fn raw_attribute_value_range(raw: &[u8], wanted: &[u8]) -> Result<Range<usize>> {
    optional_attribute_value_range(raw, wanted)?
        .ok_or_else(|| invalid("expected Theme root attribute is missing"))
}

fn optional_attribute_value_range(raw: &[u8], wanted: &[u8]) -> Result<Option<Range<usize>>> {
    let mut cursor = 1usize;
    while cursor < raw.len()
        && !raw[cursor].is_ascii_whitespace()
        && !matches!(raw[cursor], b'>' | b'/')
    {
        cursor += 1;
    }
    while cursor < raw.len() {
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= raw.len() || matches!(raw[cursor], b'>' | b'/') {
            break;
        }
        let key_start = cursor;
        while cursor < raw.len()
            && !raw[cursor].is_ascii_whitespace()
            && !matches!(raw[cursor], b'=' | b'>' | b'/')
        {
            cursor += 1;
        }
        let key_end = cursor;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= raw.len() || raw[cursor] != b'=' {
            return Err(invalid("Theme root attribute has no value"));
        }
        cursor += 1;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *raw
            .get(cursor)
            .ok_or_else(|| invalid("Theme root attribute value is missing"))?;
        if !matches!(quote, b'"' | b'\'') {
            return Err(invalid("Theme root attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < raw.len() && raw[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= raw.len() {
            return Err(invalid("Theme root attribute value is unterminated"));
        }
        if &raw[key_start..key_end] == wanted {
            return Ok(Some(value_start..value_end));
        }
        cursor += 1;
    }
    Ok(None)
}

fn copy_bytes(bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    output.extend_from_slice(bytes);
    Ok(output)
}

/// Replace one Theme payload while retaining OPC source provenance for the
/// resulting XML. Native Theme parts commonly carry an XML declaration and
/// line endings, which are valid source bytes but fail the compact authored
/// XML audit after a raw `set_blob`; the owned XML splice seam records the new
/// source token atomically and therefore permits save/reopen to preserve those
/// bytes exactly.
pub(crate) fn replace_source_xml_part(
    package: &mut OpcPackage,
    part_name: &PackURI,
    expected: &[u8],
    replacement: &[u8],
) -> Result<()> {
    let source = package.source_xml_part(part_name)?;
    if source.bytes() != expected {
        return Err(invalid("Theme XML source provenance is stale"));
    }
    let (old_open, _) = root_element_ranges(source.bytes())?;
    let (_, new_full) = root_element_ranges(replacement)?;
    let replacement_element = replacement
        .get(new_full)
        .ok_or_else(|| invalid("Theme XML replacement root range is invalid"))?;
    let owned = source.replace_element(old_open, replacement_element)?;
    if owned.bytes() != replacement {
        return Err(invalid(
            "Theme XML replacement changed bytes outside the root element",
        ));
    }
    package.try_replace_owned_xml_part(expected, owned)?;
    Ok(())
}

fn root_element_ranges(xml: &[u8]) -> Result<(Range<usize>, Range<usize>)> {
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut open = None;
    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|_error| {
            Error::CapacityOverflow {
                resource: "Theme XML root offset",
            }
        })?;
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("invalid Theme XML: {error}")))?;
        let end = usize::try_from(reader.buffer_position()).map_err(|_error| {
            Error::CapacityOverflow {
                resource: "Theme XML root offset",
            }
        })?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if element.local_name().as_ref() != b"theme" {
                        return Err(invalid("Theme XML has an unexpected root"));
                    }
                    open = Some(start..end);
                }
                depth = depth.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "Theme XML root depth",
                })?;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if element.local_name().as_ref() != b"theme" {
                        return Err(invalid("Theme XML has an unexpected root"));
                    }
                    return Ok((start..end, start..end));
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("Theme XML has an unexpected closing element"))?;
                if depth == 0 {
                    let open = open
                        .take()
                        .ok_or_else(|| invalid("Theme XML root opening tag is missing"))?;
                    let full = open.start..end;
                    return Ok((open, full));
                }
            },
            Event::Eof => return Err(invalid("Theme XML has no complete root")),
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("Theme XML contains forbidden markup"));
            },
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

/// Retain the source-backed family fragment when a caller supplies a detached
/// value for an existing owner.  The typed scalar values are the transaction's
/// requested state; unknown attributes, namespace declarations, comments, and
/// extension children remain owned by the existing source fragment.
fn preserve_family_source(existing: Option<&Family>, incoming: Family) -> Result<Family> {
    let Some(existing) = existing else {
        return Ok(incoming);
    };
    let mut value = existing.clone();
    value.set_name(incoming.name())?;
    value.set_id(incoming.id().as_str())?;
    value.set_variant_id(incoming.variant_id().as_str())?;
    Ok(value)
}

mod lifecycle;

pub(crate) use lifecycle::read_owner;
pub use lifecycle::{OwnerCommit, OwnerPatch, OwnerSnapshot, OwnerTransaction};

#[cfg(test)]
mod splice_tests;

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "theme fixture assertions use immediate extraction for compact unit tests"
)]
mod tests {
    use super::*;
    use litchi_drawingml::theme::{Color, Face};

    fn palette() -> Palette {
        Slot::ALL
            .into_iter()
            .fold(Palette::new("Office"), |palette, slot| {
                palette.with(slot, Color::rgb("4F81BD").expect("valid color"))
            })
    }

    fn theme() -> Theme {
        Theme {
            name: "Office".to_string(),
            colors: palette(),
            fonts: FontSet::new("Office", Face::new("Aptos"), Face::new("Aptos")),
        }
    }

    #[test]
    fn generated_theme_reads_and_noop_is_exact() {
        let value = theme();
        let xml = codec::encode_part(&value.name, &value.colors, &value.fonts).expect("encode");
        let parsed = parse_theme(&xml, Limits::DEFAULT).expect("read");
        assert_eq!(parsed, value);
        let snapshot = Snapshot {
            theme: Arc::new(parsed.clone()),
            family: None,
            source_xml: Arc::new(xml.clone()),
            graph: GraphState {
                workbook_name: "/xl/workbook.bin".to_string(),
                theme_name: "/xl/theme/theme1.xml".to_string(),
                workbook_relationship: RelationshipState {
                    id: "rId1".to_string(),
                    reltype: RELATIONSHIP_TYPE.to_string(),
                    target_ref: "theme/theme1.xml".to_string(),
                    target_mode: TargetMode::Internal,
                    target_name: Some("/xl/theme/theme1.xml".to_string()),
                },
                theme_relationships: Vec::new(),
            },
            limits: Limits::DEFAULT,
        };
        let commit = snapshot.edit().commit().expect("commit");
        assert!(!commit.changed());
        assert_eq!(commit.snapshot().source_xml(), xml.as_slice());
    }

    #[test]
    fn palette_edit_preserves_root_and_font_source() {
        let value = theme();
        let xml = codec::encode_part(&value.name, &value.colors, &value.fonts).expect("encode");
        let snapshot = Snapshot {
            theme: Arc::new(value.clone()),
            family: None,
            source_xml: Arc::new(xml.clone()),
            graph: GraphState {
                workbook_name: "/xl/workbook.bin".to_string(),
                theme_name: "/xl/theme/theme1.xml".to_string(),
                workbook_relationship: RelationshipState {
                    id: "rId1".to_string(),
                    reltype: RELATIONSHIP_TYPE.to_string(),
                    target_ref: "theme/theme1.xml".to_string(),
                    target_mode: TargetMode::Internal,
                    target_name: Some("/xl/theme/theme1.xml".to_string()),
                },
                theme_relationships: Vec::new(),
            },
            limits: Limits::DEFAULT,
        };
        let mut replacement = palette();
        replacement = replacement.with(Slot::Accent1, Color::rgb("FF0000").expect("color"));
        let mut edit = snapshot.edit();
        assert!(edit.set_palette(replacement).expect("set palette"));
        let commit = edit.commit().expect("commit");
        assert!(commit.changed());
        assert!(
            commit
                .snapshot()
                .source_xml()
                .windows(b"Aptos".len())
                .any(|window| window == b"Aptos")
        );
        assert_eq!(
            commit.snapshot().theme().colors.color(Slot::Accent1),
            Some(&Color::Rgb("FF0000".to_string()))
        );
        assert_eq!(
            commit.patch().inverse().after().source_xml(),
            xml.as_slice()
        );
    }
}
