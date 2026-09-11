//! Optional Theme-part lifecycle and package publication.
//!
//! [`super::Snapshot`] deliberately represents a present Theme part because it
//! offers the typed XML projection used by the source-backed read API.  The
//! package owner, however, is optional.  This module supplies the owner-level
//! snapshot and transaction used when a caller must create or remove that
//! part.  The owner transaction keeps the present-only snapshot API stable
//! while making topology changes source-checked and reversible.

use std::sync::Arc;

use litchi_opc::constants::{content_type, relationship_type as opc_relationship_type};
use litchi_opc::part::Part;
use litchi_opc::{
    BlobPart, OpcPackage, OwnedContentTypes, OwnedRelationships, OwnedXmlPart, PackURI, TargetMode,
};

use super::{
    CONTENT_TYPE, Family, GraphState, Limits, RELATIONSHIP_TYPE, RelationshipState,
    STRICT_THEME_RELATIONSHIP, Snapshot, Theme, codec, invalid, is_image_relationship,
    parse_theme_content, read, relationship_state, rewrite_source_with_family,
    validate_image_target, validate_model,
};
use crate::package::error::{Error, Result};

const DEFAULT_THEME_DIRECTORY: &str = "/xl/theme";
const DEFAULT_THEME_STEM: &str = "theme";
const DEFAULT_THEME_EXTENSION: &str = "xml";

fn relationship_output_limit() -> usize {
    litchi_opc::ReadLimits::default().max_relationship_xml_bytes()
}

fn content_types_output_limit() -> usize {
    litchi_opc::ReadLimits::default().max_content_types_bytes()
}

/// An eager snapshot of the optional Workbook-owned Theme part.
///
/// `OwnerSnapshot::theme` is `None` when the package has no Theme owner.  A
/// present value is the ordinary [`super::Snapshot`] and retains exact source
/// XML plus all validated Theme/Image graph metadata.  The owner snapshot also
/// retains the Workbook relationship list so a create/remove patch detects
/// unrelated graph changes that could alter the selected relationship ID.
#[derive(Clone, Debug)]
pub struct OwnerSnapshot {
    theme: Option<Snapshot>,
    source: OwnerSource,
    plan: CreationPlan,
    limits: Limits,
}

impl OwnerSnapshot {
    /// Borrow the typed Theme, or `None` when the owner is absent.
    #[must_use]
    pub fn theme(&self) -> Option<&Theme> {
        self.theme.as_ref().map(Snapshot::theme)
    }

    /// Borrow the optional applied-theme family metadata.
    #[must_use]
    pub fn family(&self) -> Option<&Family> {
        self.theme.as_ref().and_then(Snapshot::family)
    }

    /// Borrow the present Theme snapshot, or `None` when the owner is absent.
    #[must_use]
    pub fn snapshot(&self) -> Option<&Snapshot> {
        self.theme.as_ref()
    }

    /// Borrow the exact Theme XML when the owner is present.
    #[must_use]
    pub fn source_xml(&self) -> Option<&[u8]> {
        self.theme.as_ref().map(Snapshot::source_xml)
    }

    /// Whether the Workbook owns a Theme part.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.theme.is_some()
    }

    /// Return the finite limits used to inspect and stage this owner.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Return the absolute Theme part name used by the current or planned
    /// owner.  This is a diagnostic string rather than an OPC URI type.
    #[must_use]
    pub fn part_name(&self) -> &str {
        &self.plan.theme_name
    }

    /// Start a detached create, replace, or remove transaction.
    #[must_use]
    pub fn edit(&self) -> OwnerTransaction {
        OwnerTransaction {
            before: self.clone(),
            staged: self.theme.as_ref().map(|snapshot| snapshot.theme().clone()),
            staged_family: self.family().cloned(),
        }
    }

    fn same_source(&self, other: &Self) -> bool {
        if !self.source.same_source(&other.source) || self.plan != other.plan {
            return false;
        }
        match (&self.theme, &other.theme) {
            (None, None) => true,
            (Some(left), Some(right)) => left.same_source(right),
            _ => false,
        }
    }

    fn same_state(&self, other: &Self) -> bool {
        if !self.same_source(other) {
            return false;
        }
        match (&self.theme, &other.theme) {
            (None, None) => true,
            (Some(left), Some(right)) => left.same_state(right),
            _ => false,
        }
    }

    fn with_present(
        &self,
        theme: Theme,
        family: Option<Family>,
        source_xml: Vec<u8>,
    ) -> Result<Self> {
        let graph = GraphState {
            workbook_name: self.plan.workbook_name.clone(),
            theme_name: self.plan.theme_name.clone(),
            workbook_relationship: self.plan.relationship.clone(),
            theme_relationships: self
                .theme
                .as_ref()
                .map(|snapshot| snapshot.graph.theme_relationships.clone())
                .unwrap_or_default(),
        };
        let snapshot = Snapshot {
            theme: Arc::new(theme),
            family,
            source_xml: Arc::new(source_xml),
            graph,
            limits: self.limits,
        };
        let source = self.source.with_present(&self.plan, self.theme.is_some())?;
        Ok(Self {
            theme: Some(snapshot),
            source,
            plan: self.plan.clone(),
            limits: self.limits,
        })
    }

    fn with_absent(&self) -> Result<Self> {
        let source = self.source.with_absent(&self.plan)?;
        let plan = plan_from_absent_source(&source, &self.plan.relationship.reltype)?;
        Ok(Self {
            theme: None,
            source,
            plan,
            limits: self.limits,
        })
    }
}

/// Detached optional Theme owner edit.
#[derive(Clone, Debug)]
pub struct OwnerTransaction {
    before: OwnerSnapshot,
    staged: Option<Theme>,
    staged_family: Option<Family>,
}

impl OwnerTransaction {
    /// Immutable owner source used for stale checks.
    #[must_use]
    pub fn before(&self) -> &OwnerSnapshot {
        &self.before
    }

    /// Borrow the currently staged Theme, or `None` when removal is staged.
    #[must_use]
    pub fn theme(&self) -> Option<&Theme> {
        self.staged.as_ref()
    }

    /// Borrow the currently staged applied-theme family metadata.
    #[must_use]
    pub fn family(&self) -> Option<&Family> {
        self.staged_family.as_ref()
    }

    /// Replace or remove the complete typed Theme owner.
    pub fn replace(&mut self, theme: Option<Theme>) -> Result<bool> {
        if let Some(value) = theme.as_ref() {
            validate_model(value, self.before.limits)?;
            if !self.before.is_present() && value.name.is_empty() {
                return Err(invalid(
                    "a newly created Theme requires a non-empty display name",
                ));
            }
        }
        if self.staged == theme {
            if theme.is_none() && self.staged_family.is_some() {
                self.staged_family = None;
                return Ok(true);
            }
            return Ok(false);
        }
        self.staged = theme;
        if self.staged.is_none() {
            self.staged_family = None;
        }
        Ok(true)
    }

    /// Set or create the complete typed Theme owner.
    pub fn set_theme(&mut self, theme: Theme) -> Result<bool> {
        self.replace(Some(theme))
    }

    /// Set or replace the applied-theme family metadata.
    pub fn set_family(&mut self, family: Family) -> Result<bool> {
        let family = super::preserve_family_source(self.before.family(), family)?;
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

    /// Remove the optional Theme part and its Workbook owner edge.
    pub fn remove(&mut self) -> Result<bool> {
        self.replace(None)
    }

    /// Change the Theme display name while retaining its other typed fields.
    pub fn set_name(&mut self, name: impl Into<String>) -> Result<bool> {
        let mut candidate = self
            .staged
            .clone()
            .ok_or_else(|| invalid("cannot set a Theme name while the owner is absent"))?;
        candidate.name = name.into();
        self.replace(Some(candidate))
    }

    /// Replace the complete twelve-slot color palette.
    pub fn set_palette(&mut self, colors: super::Palette) -> Result<bool> {
        let mut candidate = self
            .staged
            .clone()
            .ok_or_else(|| invalid("cannot set a Theme palette while the owner is absent"))?;
        candidate.colors = colors;
        self.replace(Some(candidate))
    }

    /// Replace the complete major/minor font set.
    pub fn set_fonts(&mut self, fonts: super::FontSet) -> Result<bool> {
        let mut candidate = self
            .staged
            .clone()
            .ok_or_else(|| invalid("cannot set Theme fonts while the owner is absent"))?;
        candidate.fonts = fonts;
        self.replace(Some(candidate))
    }

    /// Commit the detached create, replace, or remove operation.
    pub fn commit(self) -> Result<OwnerCommit> {
        let before_theme = self
            .before
            .theme
            .as_ref()
            .map(|snapshot| snapshot.theme().clone());
        let before_family = self.before.family().cloned();
        if self.staged == before_theme && self.staged_family == before_family {
            let before = self.before;
            return Ok(OwnerCommit::new(
                OwnerPatch::new(before.clone(), before),
                false,
            ));
        }

        let after = match self.staged {
            Some(theme) => {
                validate_model(&theme, self.before.limits)?;
                if let Some(before) = self.before.theme.as_ref() {
                    let source = rewrite_source_with_family(
                        before.source_xml(),
                        before.theme(),
                        &theme,
                        before.family(),
                        self.staged_family.as_ref(),
                        self.before.limits,
                    )?;
                    let (parsed, parsed_family) = parse_theme_content(&source, self.before.limits)?;
                    if parsed != theme {
                        return Err(invalid(
                            "Theme owner transaction read-back did not match the staged model",
                        ));
                    }
                    if parsed_family.as_ref() != self.staged_family.as_ref() {
                        return Err(invalid(
                            "Theme owner transaction read-back did not match the staged family metadata",
                        ));
                    }
                    self.before.with_present(parsed, parsed_family, source)?
                } else {
                    let source = encode_new_theme(
                        &theme,
                        self.staged_family.as_ref(),
                        &self.before.plan,
                        self.before.limits,
                    )?;
                    let (parsed, parsed_family) = parse_theme_content(&source, self.before.limits)?;
                    if parsed != theme {
                        return Err(invalid(
                            "new Theme owner read-back did not match the staged model",
                        ));
                    }
                    if parsed_family.as_ref() != self.staged_family.as_ref() {
                        return Err(invalid(
                            "new Theme owner read-back did not match the staged family metadata",
                        ));
                    }
                    self.before.with_present(parsed, parsed_family, source)?
                }
            },
            None => {
                if self.staged_family.is_some() {
                    return Err(invalid(
                        "cannot commit Theme family metadata while the owner is absent",
                    ));
                }
                self.before.with_absent()?
            },
        };
        Ok(OwnerCommit::new(OwnerPatch::new(self.before, after), true))
    }
}

/// Result of an optional Theme owner transaction.
#[derive(Clone, Debug)]
pub struct OwnerCommit {
    patch: OwnerPatch,
    changed: bool,
}

impl OwnerCommit {
    fn new(patch: OwnerPatch, changed: bool) -> Self {
        Self { patch, changed }
    }

    /// Whether the Theme payload or owner topology changes.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Planned resulting owner snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &OwnerSnapshot {
        self.patch.after()
    }

    /// Reversible source-checked publication patch.
    #[must_use]
    pub fn patch(&self) -> &OwnerPatch {
        &self.patch
    }

    /// Consume the commit into its planned snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (OwnerSnapshot, OwnerPatch) {
        let snapshot = self.patch.after().clone();
        (snapshot, self.patch)
    }
}

/// A source-checked, reversible create/replace/remove Theme owner patch.
#[derive(Clone, Debug)]
pub struct OwnerPatch {
    before: OwnerSnapshot,
    after: OwnerSnapshot,
    changed: bool,
}

impl OwnerPatch {
    fn new(before: OwnerSnapshot, after: OwnerSnapshot) -> Self {
        let changed = !before.same_source(&after);
        Self {
            before,
            after,
            changed,
        }
    }

    /// Source state required before publication.
    #[must_use]
    pub fn before(&self) -> &OwnerSnapshot {
        &self.before
    }

    /// Exact state produced by publication.
    #[must_use]
    pub fn after(&self) -> &OwnerSnapshot {
        &self.after
    }

    /// Whether applying this patch changes bytes or owner topology.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        !self.changed
    }

    /// Return a patch that restores the source state represented by `before`.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            changed: self.changed,
        }
    }

    /// Apply atomically after checking the complete owner source closure.
    pub(crate) fn apply(&self, package: &mut OpcPackage) -> Result<OwnerSnapshot> {
        let current = read_owner(package, self.before.limits)?;
        if !current.same_source(&self.before) {
            return Err(invalid("Theme owner patch source is stale"));
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
        candidate.unsign();
        materialize(&mut candidate, &self.before, &self.after)?;
        let resulting = read_owner(&candidate, self.before.limits)?;
        if !resulting.same_state(&self.after) {
            return Err(invalid(
                "Theme owner publication changed the planned semantic or graph state",
            ));
        }
        *package = candidate;
        Ok(resulting)
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct CreationPlan {
    workbook_name: String,
    theme_name: String,
    relationship: RelationshipState,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct OwnerSource {
    workbook_name: String,
    workbook_relationships: Vec<RelationshipState>,
    part_names: Vec<String>,
    workbook_relationship_xml: OwnedRelationships,
    theme_relationship_xml: ThemeRelationships,
    /// Exact source XML for a present Theme, retained so a remove/inverse
    /// cycle can restore a non-compact source part with its publication
    /// provenance.  This token is restoration data; semantic stale checks
    /// remain anchored by `Snapshot::source_xml` and the package graph.
    theme_xml: Option<OwnedXmlPart>,
    content_types: OwnedContentTypes,
}

/// Physical relationship-member state for the optional Theme part.
///
/// A newly authored Theme with no outgoing edges has no `.rels` member.  That
/// is distinct from an existing source member containing zero edges: the
/// latter must be restored byte for byte, while the former must reject a
/// newly introduced physical member during read-back.
#[derive(Clone, Debug, PartialEq, Eq)]
enum ThemeRelationships {
    Absent,
    Exact(OwnedRelationships),
    FreshAbsent { owner: String },
}

impl OwnerSource {
    fn same_source(&self, other: &Self) -> bool {
        if self.workbook_name != other.workbook_name
            || self.workbook_relationships != other.workbook_relationships
            || self.part_names != other.part_names
            || self.workbook_relationship_xml != other.workbook_relationship_xml
            || self.content_types != other.content_types
        {
            return false;
        }
        self.theme_relationship_xml == other.theme_relationship_xml
    }

    fn with_present(&self, plan: &CreationPlan, owner_was_present: bool) -> Result<Self> {
        let mut workbook_relationships = self.workbook_relationships.clone();
        let mut workbook_relationship_xml = self.workbook_relationship_xml.clone();
        if !workbook_relationships
            .iter()
            .any(|relationship| relationship.id == plan.relationship.id)
        {
            workbook_relationships.push(plan.relationship.clone());
            workbook_relationship_xml = workbook_relationship_xml.with_relationship(
                &plan.relationship.reltype,
                &plan.relationship.target_ref,
                &plan.relationship.id,
                plan.relationship.target_mode,
                relationship_output_limit(),
            )?;
        }
        workbook_relationships.sort_by(|left, right| left.id.cmp(&right.id));
        let mut part_names = self.part_names.clone();
        let mut content_types = self.content_types.clone();
        if !part_names.iter().any(|name| name == &plan.theme_name) {
            part_names.push(plan.theme_name.clone());
            let theme_name = PackURI::new(&plan.theme_name)
                .map_err(|error| Error::InvalidUri(error.to_string()))?;
            content_types = content_types.with_part_overrides(
                &[(&theme_name, CONTENT_TYPE)],
                content_types_output_limit(),
            )?;
        }
        part_names.sort();
        let theme_relationship_xml = if owner_was_present {
            self.theme_relationship_xml.clone()
        } else {
            ThemeRelationships::FreshAbsent {
                owner: plan.theme_name.clone(),
            }
        };
        Ok(Self {
            workbook_name: self.workbook_name.clone(),
            workbook_relationships,
            part_names,
            workbook_relationship_xml,
            theme_relationship_xml,
            theme_xml: self.theme_xml.clone(),
            content_types,
        })
    }

    fn with_absent(&self, plan: &CreationPlan) -> Result<Self> {
        let mut workbook_relationships = Vec::new();
        workbook_relationships
            .try_reserve_exact(self.workbook_relationships.len())
            .map_err(|source| Error::Allocation {
                resource: "Theme owner Workbook relationships",
                source,
            })?;
        for relationship in &self.workbook_relationships {
            if relationship.id != plan.relationship.id {
                workbook_relationships.push(relationship.clone());
            }
        }
        let mut part_names = Vec::new();
        part_names
            .try_reserve_exact(self.part_names.len())
            .map_err(|source| Error::Allocation {
                resource: "Theme owner package part names",
                source,
            })?;
        for name in &self.part_names {
            if name != &plan.theme_name {
                part_names.push(name.clone());
            }
        }
        let workbook_relationship_xml = self
            .workbook_relationship_xml
            .without_relationship(&plan.relationship.id, relationship_output_limit())?;
        let theme_name =
            PackURI::new(&plan.theme_name).map_err(|error| Error::InvalidUri(error.to_string()))?;
        let content_types = self.content_types.without_parts(
            std::slice::from_ref(&theme_name),
            content_types_output_limit(),
        )?;
        Ok(Self {
            workbook_name: self.workbook_name.clone(),
            workbook_relationships,
            part_names,
            workbook_relationship_xml,
            theme_relationship_xml: ThemeRelationships::Absent,
            theme_xml: self.theme_xml.clone(),
            content_types,
        })
    }
}

/// Read the optional Theme owner, retaining enough Workbook graph metadata for
/// source-checked create/remove publication.
pub(crate) fn read_owner(package: &OpcPackage, limits: Limits) -> Result<OwnerSnapshot> {
    let limits = limits.validate()?;
    let theme = read(package, limits)?;
    let workbook = package.main_document_part()?;
    if workbook.content_type() != content_type::XLSB_BIN {
        return Err(Error::InvalidContentType {
            expected: content_type::XLSB_BIN.to_owned(),
            got: workbook.content_type().to_owned(),
        });
    }
    let mut workbook_relationships = Vec::new();
    workbook_relationships
        .try_reserve(workbook.rels().len())
        .map_err(|source| Error::Allocation {
            resource: "Theme owner Workbook relationships",
            source,
        })?;
    for relationship in workbook.rels().iter() {
        workbook_relationships.push(relationship_state(relationship));
    }
    workbook_relationships.sort_by(|left, right| left.id.cmp(&right.id));
    let part_names = collect_part_names(package)?;
    let opc_limits = litchi_opc::ReadLimits::default();
    let workbook_relationship_xml =
        package.source_relationships_with_limits(workbook.partname(), opc_limits)?;
    let content_types = package.source_content_types_with_limits(opc_limits)?;
    let theme_relationship_xml = theme
        .as_ref()
        .map(|snapshot| -> Result<ThemeRelationships> {
            let name = PackURI::new(&snapshot.graph.theme_name)
                .map_err(|error| Error::InvalidUri(error.to_string()))?;
            let relationships = package.source_relationships_with_limits(&name, opc_limits)?;
            if relationships.member_present() {
                Ok(ThemeRelationships::Exact(relationships))
            } else {
                Ok(ThemeRelationships::FreshAbsent {
                    owner: snapshot.graph.theme_name.clone(),
                })
            }
        })
        .transpose()?
        .unwrap_or(ThemeRelationships::Absent);
    let theme_xml = theme
        .as_ref()
        .map(|snapshot| {
            let name = PackURI::new(&snapshot.graph.theme_name)
                .map_err(|error| Error::InvalidUri(error.to_string()))?;
            package.source_xml_part(&name).map_err(Error::from)
        })
        .transpose()?;
    let plan = match theme.as_ref() {
        Some(snapshot) => CreationPlan {
            workbook_name: snapshot.graph.workbook_name.clone(),
            theme_name: snapshot.graph.theme_name.clone(),
            relationship: snapshot.graph.workbook_relationship.clone(),
        },
        None => plan_absent(package, workbook, &workbook_relationships)?,
    };
    let source = OwnerSource {
        workbook_name: workbook.partname().as_str().to_owned(),
        workbook_relationships,
        part_names,
        workbook_relationship_xml,
        theme_relationship_xml,
        theme_xml,
        content_types,
    };
    Ok(OwnerSnapshot {
        theme,
        source,
        plan,
        limits,
    })
}

fn collect_part_names(package: &OpcPackage) -> Result<Vec<String>> {
    let mut names = Vec::new();
    names
        .try_reserve_exact(package.part_count())
        .map_err(|source| Error::Allocation {
            resource: "Theme owner package part names",
            source,
        })?;
    for part in package.iter_parts() {
        names.push(part.partname().as_str().to_owned());
    }
    names.sort();
    Ok(names)
}

fn plan_absent(
    package: &OpcPackage,
    workbook: &dyn Part,
    workbook_relationships: &[RelationshipState],
) -> Result<CreationPlan> {
    let theme_name = next_theme_name(package)?;
    let relationship_id = next_relationship_id(workbook_relationships)?;
    let relationship_type = package
        .rels()
        .iter()
        .find(|relationship| {
            matches!(
                relationship.reltype(),
                opc_relationship_type::OFFICE_DOCUMENT
                    | opc_relationship_type::STRICT_OFFICE_DOCUMENT
            )
        })
        .map(|relationship| {
            if relationship.reltype() == opc_relationship_type::STRICT_OFFICE_DOCUMENT {
                STRICT_THEME_RELATIONSHIP
            } else {
                RELATIONSHIP_TYPE
            }
        })
        .ok_or_else(|| invalid("XLSB package has no OfficeDocument relationship"))?;
    let target_ref = theme_name.relative_ref(workbook.partname().base_uri());
    Ok(CreationPlan {
        workbook_name: workbook.partname().as_str().to_owned(),
        theme_name: theme_name.as_str().to_owned(),
        relationship: RelationshipState {
            id: relationship_id,
            reltype: relationship_type.to_owned(),
            target_ref,
            target_mode: TargetMode::Internal,
            target_name: Some(theme_name.as_str().to_owned()),
        },
    })
}

fn next_theme_name(package: &OpcPackage) -> Result<PackURI> {
    let maximum = package
        .part_count()
        .checked_add(1)
        .ok_or(Error::CapacityOverflow {
            resource: "Theme part name candidates",
        })?;
    for index in 1..=maximum {
        let candidate = PackURI::new(format!(
            "{DEFAULT_THEME_DIRECTORY}/{DEFAULT_THEME_STEM}{index}.{DEFAULT_THEME_EXTENSION}"
        ))
        .map_err(|error| Error::InvalidUri(error.to_string()))?;
        if package
            .iter_parts()
            .all(|part| !part.partname().is_equivalent_to(&candidate))
        {
            return Ok(candidate);
        }
    }
    Err(invalid("no bounded Theme part name is available"))
}

fn next_relationship_id(relationships: &[RelationshipState]) -> Result<String> {
    let maximum = relationships
        .len()
        .checked_add(1)
        .ok_or(Error::CapacityOverflow {
            resource: "Theme relationship ID candidates",
        })?;
    for index in 1..=maximum {
        let candidate = format!("rId{index}");
        if !relationships
            .iter()
            .any(|relationship| relationship.id == candidate)
        {
            return Ok(candidate);
        }
    }
    Err(invalid("no bounded Theme relationship ID is available"))
}

fn plan_from_absent_source(source: &OwnerSource, relationship_type: &str) -> Result<CreationPlan> {
    let maximum = source
        .part_names
        .len()
        .checked_add(1)
        .ok_or(Error::CapacityOverflow {
            resource: "Theme part name candidates",
        })?;
    let theme_name = (1..=maximum)
        .find_map(|index| {
            let candidate = PackURI::new(format!(
                "{DEFAULT_THEME_DIRECTORY}/{DEFAULT_THEME_STEM}{index}.{DEFAULT_THEME_EXTENSION}"
            ))
            .ok()?;
            source
                .part_names
                .iter()
                .all(|name| {
                    PackURI::new(name)
                        .map(|existing| !existing.is_equivalent_to(&candidate))
                        .unwrap_or(false)
                })
                .then_some(candidate)
        })
        .ok_or_else(|| invalid("no bounded Theme part name is available"))?;
    let maximum =
        source
            .workbook_relationships
            .len()
            .checked_add(1)
            .ok_or(Error::CapacityOverflow {
                resource: "Theme relationship ID candidates",
            })?;
    let relationship_id = (1..=maximum)
        .map(|index| format!("rId{index}"))
        .find(|candidate| {
            !source
                .workbook_relationships
                .iter()
                .any(|relationship| relationship.id == *candidate)
        })
        .ok_or_else(|| invalid("no bounded Theme relationship ID is available"))?;
    let workbook = PackURI::new(&source.workbook_name)
        .map_err(|error| Error::InvalidUri(error.to_string()))?;
    let target_ref = theme_name.relative_ref(workbook.base_uri());
    Ok(CreationPlan {
        workbook_name: source.workbook_name.clone(),
        theme_name: theme_name.as_str().to_owned(),
        relationship: RelationshipState {
            id: relationship_id,
            reltype: relationship_type.to_owned(),
            target_ref,
            target_mode: TargetMode::Internal,
            target_name: Some(theme_name.as_str().to_owned()),
        },
    })
}

fn encode_new_theme(
    theme: &Theme,
    family: Option<&Family>,
    plan: &CreationPlan,
    limits: Limits,
) -> Result<Vec<u8>> {
    let name = if theme.name.is_empty() {
        return Err(invalid(
            "a newly created Theme requires a non-empty display name",
        ));
    } else {
        theme.name.as_str()
    };
    let mut source = codec::encode_part(name, &theme.colors, &theme.fonts)?;
    if plan.relationship.reltype == STRICT_THEME_RELATIONSHIP {
        let range = super::raw_attribute_value_range(&source, b"xmlns:a")?;
        if &source[range.clone()] != codec::NAMESPACE.as_bytes() {
            return Err(invalid("generated Theme has an unexpected namespace"));
        }
        source.splice(range, codec::STRICT_NAMESPACE.bytes());
    }
    if let Some(family) = family {
        source = super::family::part::add_family_with_uri_limit(
            &source,
            family,
            super::family::part::NATIVE_EXTENSION_URI,
            limits.max_xml_bytes,
        )?;
    }
    if source.len() > limits.max_xml_bytes {
        return Err(Error::LimitExceeded {
            resource: "generated Theme XML bytes",
            actual: source.len(),
            maximum: limits.max_xml_bytes,
        });
    }
    Ok(source)
}

fn materialize(
    package: &mut OpcPackage,
    before: &OwnerSnapshot,
    after: &OwnerSnapshot,
) -> Result<()> {
    let workbook_name = package.main_document_part()?.partname().clone();
    let expected_workbook = PackURI::new(&before.source.workbook_name)
        .map_err(|error| Error::InvalidUri(error.to_string()))?;
    if !workbook_name.is_equivalent_to(&expected_workbook) {
        return Err(invalid("Theme owner Workbook changed during publication"));
    }
    match (before.theme.as_ref(), after.theme.as_ref()) {
        (Some(before), Some(after)) => {
            if before.graph != after.graph {
                return Err(invalid(
                    "Theme replacement cannot change its validated package graph",
                ));
            }
            let part_name = PackURI::new(&before.graph.theme_name)
                .map_err(|error| Error::InvalidUri(error.to_string()))?;
            let part = package.get_part(&part_name)?;
            if part.content_type() != CONTENT_TYPE
                || part.rels().iter().count() != before.graph.theme_relationships.len()
            {
                return Err(invalid("Theme part graph changed during replacement"));
            }
            super::replace_source_xml_part(
                package,
                &part_name,
                before.source_xml(),
                after.source_xml(),
            )?;
        },
        (None, Some(_)) => add_owner(package, &workbook_name, after)?,
        (Some(before), None) => remove_owner(package, &workbook_name, before)?,
        (None, None) => return Err(invalid("changed Theme owner patch has no resulting owner")),
    }
    restore_source_tokens(package, &after.source)?;
    Ok(())
}

fn restore_source_tokens(package: &mut OpcPackage, source: &OwnerSource) -> Result<()> {
    let limits = litchi_opc::ReadLimits::default();
    let workbook_name = PackURI::new(&source.workbook_name)
        .map_err(|error| Error::InvalidUri(error.to_string()))?;
    let current_workbook = package.source_relationships_with_limits(&workbook_name, limits)?;
    if current_workbook != source.workbook_relationship_xml {
        package.try_replace_relationships_with_limits(
            &current_workbook,
            &source.workbook_relationship_xml,
            limits,
        )?;
    }
    match &source.theme_relationship_xml {
        ThemeRelationships::Absent => {},
        ThemeRelationships::FreshAbsent { owner } => {
            let owner =
                PackURI::new(owner).map_err(|error| Error::InvalidUri(error.to_string()))?;
            let current = package.source_relationships_with_limits(&owner, limits)?;
            if current.member_present() {
                return Err(invalid(
                    "new Theme unexpectedly acquired a physical relationship member",
                ));
            }
        },
        ThemeRelationships::Exact(theme_relationship_xml) => {
            let current_theme =
                package.source_relationships_with_limits(theme_relationship_xml.owner(), limits)?;
            if current_theme != *theme_relationship_xml {
                package.try_replace_relationships_with_limits(
                    &current_theme,
                    theme_relationship_xml,
                    limits,
                )?;
            }
        },
    }
    let current_content_types = package.source_content_types_with_limits(limits)?;
    if current_content_types.bytes() != source.content_types.bytes() {
        package.try_replace_content_types_with_limits(
            current_content_types.bytes(),
            &source.content_types,
            limits,
        )?;
    }
    Ok(())
}

fn add_owner(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    after: &OwnerSnapshot,
) -> Result<()> {
    let snapshot = after
        .theme
        .as_ref()
        .ok_or_else(|| invalid("Theme creation is missing its resulting snapshot"))?;
    let theme_name = PackURI::new(&snapshot.graph.theme_name)
        .map_err(|error| Error::InvalidUri(error.to_string()))?;
    package.validate_new_part_name(&theme_name)?;
    let source_xml = Arc::clone(&snapshot.source_xml);
    let source_token = after.source.theme_xml.clone();
    if let Some(source) = source_token {
        package.try_add_owned_xml_part(source)?;
    } else {
        package.try_add_part(Box::new(BlobPart::new_shared(
            theme_name.clone(),
            CONTENT_TYPE.to_owned(),
            Arc::clone(&source_xml),
        )))?;
    }
    for relationship in &snapshot.graph.theme_relationships {
        if !is_image_relationship(&relationship.reltype) {
            return Err(invalid("Theme inverse contains a non-Image relationship"));
        }
        if relationship.target_mode == TargetMode::Internal {
            let target_name = relationship
                .target_name
                .as_ref()
                .ok_or_else(|| invalid("Theme inverse is missing an internal Image target"))?;
            let target =
                PackURI::new(target_name).map_err(|error| Error::InvalidUri(error.to_string()))?;
            let image = package.get_part(&target)?;
            validate_image_target(image.content_type(), image.rels().is_empty())?;
        }
    }
    let theme_part = package.get_part_mut(&theme_name)?;
    for relationship in &snapshot.graph.theme_relationships {
        theme_part.rels_mut().try_add_relationship(
            relationship.reltype.clone(),
            relationship.target_ref.clone(),
            relationship.id.clone(),
            relationship.target_mode,
        )?;
    }
    let relationship = &snapshot.graph.workbook_relationship;
    let workbook = package.get_part_mut(workbook_name)?;
    if workbook.rels().get(&relationship.id).is_some() {
        return Err(invalid("Theme creation relationship ID is already present"));
    }
    workbook.rels_mut().try_add_relationship(
        relationship.reltype.clone(),
        relationship.target_ref.clone(),
        relationship.id.clone(),
        TargetMode::Internal,
    )?;
    Ok(())
}

fn remove_owner(
    package: &mut OpcPackage,
    workbook_name: &PackURI,
    before: &Snapshot,
) -> Result<()> {
    let theme_name = PackURI::new(&before.graph.theme_name)
        .map_err(|error| Error::InvalidUri(error.to_string()))?;
    let relationship_id = &before.graph.workbook_relationship.id;
    let workbook = package.get_part_mut(workbook_name)?;
    let relationship = workbook
        .rels()
        .get(relationship_id)
        .ok_or_else(|| invalid("Theme Workbook relationship is missing during removal"))?;
    if relationship.is_external() || !is_theme_relationship(relationship.reltype()) {
        return Err(invalid(
            "Theme Workbook relationship changed during removal",
        ));
    }
    let relationship_target = relationship.target_partname()?;
    if !relationship_target.is_equivalent_to(&theme_name) {
        return Err(invalid(
            "Theme Workbook relationship changed during removal",
        ));
    }
    workbook.rels_mut().remove(relationship_id);
    if !package.remove_part(&theme_name) {
        return Err(invalid("Theme part disappeared during removal"));
    }
    Ok(())
}

fn is_theme_relationship(value: &str) -> bool {
    matches!(value, RELATIONSHIP_TYPE | STRICT_THEME_RELATIONSHIP)
}
