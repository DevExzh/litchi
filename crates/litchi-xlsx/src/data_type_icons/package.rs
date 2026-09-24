//! Source-bound package ownership for data-type-icon metadata.

use std::sync::Arc;

use litchi_opc::{OpcPackage, OwnedRelationships, PackURI, Part as OpcPart};

use crate::error::{Error, Result};
use crate::named_sheet_view::{CONTENT_TYPE, RELATIONSHIP};
use crate::sheet_view::parse_worksheet_views;

use super::codec::{self, OwnerKind};
use super::model::ShowDataTypeIcons;
use super::{MAX_PART_BYTES, invalid};

const MAX_INCOMING_RELATIONSHIPS: usize = 65_536;

/// The typed, source-bound state of one worksheet's two data-type-icon owners.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    worksheet: PackURI,
    worksheet_xml: Arc<Vec<u8>>,
    worksheet_source: SourcePart,
    named_source: Option<SourcePart>,
    worksheet_views: Vec<Option<ShowDataTypeIcons>>,
    custom_sheet_views: Vec<Option<ShowDataTypeIcons>>,
}

impl Snapshot {
    /// Load one worksheet and its optional Named Sheet Views part.
    ///
    /// The reader captures metadata only. It does not apply a view, evaluate
    /// formulas, display an icon, refresh a value, or follow an external URI.
    pub fn load(package: &OpcPackage, worksheet: &PackURI) -> Result<Self> {
        let part = package.get_part(worksheet)?;
        if part.content_type() != litchi_opc::constants::content_type::SML_WORKSHEET {
            return Err(invalid(format!("part '{}' is not a worksheet", worksheet)));
        }
        if part.blob().len() > MAX_PART_BYTES {
            return Err(invalid(
                "worksheet data-type-icon source part exceeds its byte limit",
            ));
        }
        let views = parse_worksheet_views(part.blob())?;
        let worksheet_views = views
            .as_ref()
            .map(|collection| {
                collection
                    .entries()
                    .iter()
                    .map(|entry| entry.show_data_type_icons())
                    .collect()
            })
            .unwrap_or_default();

        // Reuse the Named Sheet Views owner validator so orphan parts,
        // duplicate worksheet edges, wrong content types, and outbound
        // relationships are rejected before a source-bound edit is staged.
        let _ = crate::named_sheet_view::load_worksheet_named_sheet_views(package, worksheet)?;
        let worksheet_source = SourcePart::capture(package, part)?;
        let named_source = named_part(package, worksheet)?
            .map(|part| SourcePart::capture(package, part))
            .transpose()?;
        let custom_sheet_views = named_source
            .as_ref()
            .map(|source| {
                crate::named_sheet_view::parse_named_sheet_views(source.bytes.as_slice()).map(
                    |views| {
                        views
                            .views()
                            .iter()
                            .map(|view| view.show_data_type_icons_custom_sheet_view())
                            .collect::<Vec<_>>()
                    },
                )
            })
            .transpose()?
            .unwrap_or_default();

        Ok(Self {
            worksheet: worksheet.clone(),
            worksheet_xml: Arc::clone(&worksheet_source.bytes),
            worksheet_source,
            named_source,
            worksheet_views,
            custom_sheet_views,
        })
    }

    /// Alias emphasizing that the result is tied to exact source bytes.
    pub fn read(package: &OpcPackage, worksheet: &PackURI) -> Result<Self> {
        Self::load(package, worksheet)
    }

    #[must_use]
    pub fn worksheet_part(&self) -> &PackURI {
        &self.worksheet
    }

    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.worksheet_xml.as_slice()
    }

    #[must_use]
    pub fn source_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.worksheet_xml)
    }

    /// The ordinary worksheet `sheetView` values in source order.
    #[must_use]
    pub fn worksheet_views(&self) -> &[Option<ShowDataTypeIcons>] {
        &self.worksheet_views
    }

    /// The Named Sheet View values in source order.
    #[must_use]
    pub fn custom_sheet_views(&self) -> &[Option<ShowDataTypeIcons>] {
        &self.custom_sheet_views
    }

    #[must_use]
    pub fn named_sheet_views_part(&self) -> Option<&PackURI> {
        self.named_source.as_ref().map(|source| &source.name)
    }

    #[must_use]
    pub fn named_sheet_views_source_arc(&self) -> Option<Arc<Vec<u8>>> {
        self.named_source
            .as_ref()
            .map(|source| Arc::clone(&source.bytes))
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.worksheet == other.worksheet
            && self.worksheet_source == other.worksheet_source
            && self.named_source == other.named_source
    }

    fn same_closure(&self, other: &Self) -> bool {
        self.worksheet_source.same_closure(&other.worksheet_source)
            && match (&self.named_source, &other.named_source) {
                (None, None) => true,
                (Some(left), Some(right)) => left.same_closure(right),
                _ => false,
            }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourcePart {
    name: PackURI,
    content_type: String,
    bytes: Arc<Vec<u8>>,
    relationships: OwnedRelationships,
    incoming: Vec<RelationshipState>,
}

impl SourcePart {
    fn capture(package: &OpcPackage, part: &dyn OpcPart) -> Result<Self> {
        let name = part.partname().clone();
        let relationships = package.source_relationships(&name)?;
        let incoming = capture_incoming(package, &name)?;
        Ok(Self {
            name: name.clone(),
            content_type: part.content_type().to_owned(),
            bytes: part.blob_arc(),
            relationships,
            incoming,
        })
    }

    fn same_closure(&self, other: &Self) -> bool {
        self.name == other.name
            && self.content_type == other.content_type
            && self.relationships == other.relationships
            && self.incoming == other.incoming
    }
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct RelationshipState {
    source: String,
    id: String,
    reltype: String,
    target: String,
    external: bool,
}

fn capture_incoming(package: &OpcPackage, target: &PackURI) -> Result<Vec<RelationshipState>> {
    let mut incoming = Vec::new();
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()?.is_equivalent_to(target) {
            push_incoming(
                &mut incoming,
                RelationshipState {
                    source: "/".to_owned(),
                    id: relationship.r_id().to_owned(),
                    reltype: relationship.reltype().to_owned(),
                    target: relationship.target_ref().to_owned(),
                    external: relationship.is_external(),
                },
            )?;
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if !relationship.is_external()
                && relationship.target_partname()?.is_equivalent_to(target)
            {
                push_incoming(
                    &mut incoming,
                    RelationshipState {
                        source: part.partname().as_str().to_owned(),
                        id: relationship.r_id().to_owned(),
                        reltype: relationship.reltype().to_owned(),
                        target: relationship.target_ref().to_owned(),
                        external: relationship.is_external(),
                    },
                )?;
            }
        }
    }
    incoming.sort_unstable();
    Ok(incoming)
}

fn push_incoming(values: &mut Vec<RelationshipState>, value: RelationshipState) -> Result<()> {
    if values.len() >= MAX_INCOMING_RELATIONSHIPS {
        return Err(invalid(
            "data-type-icon incoming relationship count exceeds limit",
        ));
    }
    values.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "data-type-icon incoming relationship closure",
        source,
    })?;
    values.push(value);
    Ok(())
}

fn named_part<'a>(package: &'a OpcPackage, worksheet: &PackURI) -> Result<Option<&'a dyn OpcPart>> {
    let worksheet_part = package.get_part(worksheet)?;
    let mut found = None;
    for relationship in worksheet_part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == RELATIONSHIP)
    {
        if found.is_some() {
            return Err(invalid(
                "worksheet has multiple Named Sheet Views relationships",
            ));
        }
        if relationship.is_external() {
            return Err(invalid("Named Sheet Views relationship cannot be external"));
        }
        let name = relationship.target_partname()?;
        let part = package.get_part(&name)?;
        if part.content_type() != CONTENT_TYPE {
            return Err(invalid(
                "Named Sheet Views relationship targets the wrong content type",
            ));
        }
        if !part.rels().is_empty() {
            return Err(invalid(
                "Named Sheet Views part must not have relationships",
            ));
        }
        found = Some(part);
    }
    Ok(found)
}

/// Failure-atomic edits over one worksheet's source-bound icon metadata.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    worksheet_views: Vec<Option<ShowDataTypeIcons>>,
    custom_sheet_views: Vec<Option<ShowDataTypeIcons>>,
}

impl<'a> Transaction<'a> {
    pub fn new(target: &'a mut OpcPackage, worksheet: &PackURI) -> Result<Self> {
        let before = Snapshot::load(target, worksheet)?;
        Ok(Self {
            worksheet_views: before.worksheet_views.clone(),
            custom_sheet_views: before.custom_sheet_views.clone(),
            target,
            before,
        })
    }

    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    #[must_use]
    pub fn worksheet_views(&self) -> &[Option<ShowDataTypeIcons>] {
        &self.worksheet_views
    }

    #[must_use]
    pub fn custom_sheet_views(&self) -> &[Option<ShowDataTypeIcons>] {
        &self.custom_sheet_views
    }

    pub fn set_worksheet_view(
        &mut self,
        index: usize,
        value: Option<ShowDataTypeIcons>,
    ) -> Result<bool> {
        let slot = self
            .worksheet_views
            .get_mut(index)
            .ok_or_else(|| invalid("worksheet sheetView index is out of range"))?;
        if *slot == value {
            return Ok(false);
        }
        *slot = value;
        Ok(true)
    }

    pub fn set_custom_sheet_view(
        &mut self,
        index: usize,
        value: Option<ShowDataTypeIcons>,
    ) -> Result<bool> {
        let slot = self
            .custom_sheet_views
            .get_mut(index)
            .ok_or_else(|| invalid("Named Sheet View index is out of range"))?;
        if *slot == value {
            return Ok(false);
        }
        *slot = value;
        Ok(true)
    }

    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.worksheet_views != self.worksheet_views
            || self.before.custom_sheet_views != self.custom_sheet_views
    }

    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            return Ok(Commit::new(
                self.before.clone(),
                Patch::new(self.before.clone(), self.before.clone()),
                false,
            ));
        }
        if self.target.is_signed() {
            return Err(Error::Signed);
        }
        let current = Snapshot::load(self.target, &self.before.worksheet)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "data-type-icon source closure".to_owned(),
            });
        }
        let mut candidate = self.target.clone();
        if self.before.worksheet_views != self.worksheet_views {
            let output = rewrite_worksheet(&self.before, &self.worksheet_views)?;
            candidate
                .get_part_mut(&self.before.worksheet)?
                .set_blob_shared(Arc::new(output));
        }
        if self.before.custom_sheet_views != self.custom_sheet_views {
            let source =
                self.before.named_source.as_ref().ok_or_else(|| {
                    invalid("Named Sheet Views part is required for a changed edit")
                })?;
            let output = rewrite_named(
                source.bytes.as_slice(),
                &self.before.custom_sheet_views,
                &self.custom_sheet_views,
            )?;
            candidate
                .get_part_mut(&source.name)?
                .set_blob_shared(Arc::new(output));
        }
        let after = Snapshot::load(&candidate, &self.before.worksheet)?;
        if after.worksheet_views != self.worksheet_views
            || after.custom_sheet_views != self.custom_sheet_views
            || !after.same_closure(&self.before)
        {
            return Err(invalid(
                "data-type-icon publication changed staged metadata or source closure",
            ));
        }
        let patch = Patch::new(self.before, after.clone());
        *self.target = candidate;
        Ok(Commit::new(after, patch, true))
    }
}

fn rewrite_worksheet(before: &Snapshot, values: &[Option<ShowDataTypeIcons>]) -> Result<Vec<u8>> {
    let mut output = before.worksheet_xml.as_slice().to_vec();
    for (index, value) in values.iter().copied().enumerate() {
        if before.worksheet_views.get(index).copied() == Some(value) {
            continue;
        }
        output = codec::rewrite(&output, OwnerKind::WorksheetView, index, value)?;
        if output.len() > MAX_PART_BYTES {
            return Err(invalid(
                "worksheet data-type-icon output exceeds part limit",
            ));
        }
    }
    Ok(output)
}

fn rewrite_named(
    xml: &[u8],
    before: &[Option<ShowDataTypeIcons>],
    values: &[Option<ShowDataTypeIcons>],
) -> Result<Vec<u8>> {
    let mut output = xml.to_vec();
    for (index, value) in values.iter().copied().enumerate() {
        if before.get(index).copied() == Some(value) {
            continue;
        }
        output = codec::rewrite(&output, OwnerKind::CustomSheetView, index, value)?;
        if output.len() > MAX_PART_BYTES {
            return Err(invalid(
                "Named Sheet Views data-type-icon output exceeds part limit",
            ));
        }
    }
    Ok(output)
}

/// Start a source-bound transaction for one worksheet.
pub fn edit<'a>(target: &'a mut OpcPackage, worksheet: &PackURI) -> Result<Transaction<'a>> {
    Transaction::new(target, worksheet)
}

/// Load the source-bound metadata for one worksheet.
pub fn load(target: &OpcPackage, worksheet: &PackURI) -> Result<Snapshot> {
    Snapshot::load(target, worksheet)
}

/// Apply an exact source-checked patch atomically.
pub fn apply_patch(patch: &Patch, target: &mut OpcPackage) -> Result<()> {
    patch.apply(target)
}

/// An exact source-checked reversible edit.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    pub fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target, &self.before.worksheet)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "data-type-icon source closure".to_owned(),
            });
        }
        if self.is_empty() {
            return Ok(());
        }
        if target.is_signed() {
            return Err(Error::Signed);
        }
        let mut candidate = target.clone();
        candidate
            .get_part_mut(&self.after.worksheet)?
            .set_blob_shared(Arc::clone(&self.after.worksheet_xml));
        if let Some(source) = &self.after.named_source {
            candidate
                .get_part_mut(&source.name)?
                .set_blob_shared(Arc::clone(&source.bytes));
        }
        let resulting = Snapshot::load(&candidate, &self.after.worksheet)?;
        if resulting.worksheet_views != self.after.worksheet_views
            || resulting.custom_sheet_views != self.after.custom_sheet_views
            || !resulting.same_source(&self.after)
        {
            return Err(invalid("data-type-icon patch verification failed"));
        }
        *target = candidate;
        Ok(())
    }
}

/// A committed transaction and its reversible patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn new(snapshot: Snapshot, patch: Patch, changed: bool) -> Self {
        Self {
            snapshot,
            patch,
            changed,
        }
    }

    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}
