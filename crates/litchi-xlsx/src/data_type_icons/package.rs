//! Source-bound package ownership for data-type-icon metadata.

use std::sync::Arc;

use litchi_opc::{OpcPackage, OwnedRelationships, PackURI, Part as OpcPart};

use crate::error::Result;
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
            return Err(Error::limit(
                LimitKind::PartBytes,
                part.blob().len(),
                MAX_PART_BYTES,
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
        let relationships = package.source_relationships(&name, package.read_limits())?;
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
        return Err(Error::limit(
            LimitKind::IncomingRelationships,
            values.len().saturating_add(1),
            MAX_INCOMING_RELATIONSHIPS,
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

crate::record::record_family!("data-type-icon");

crate::record::limit_kinds!("data-type-icon", DataTypeIcons {
    /// Bytes in an extension element.
    ExtensionBytes => "data-type-icon extension bytes",
    /// Bytes in one source or output XML part.
    PartBytes => "data-type-icon part bytes",
    /// XML events in one bounded traversal.
    XmlNodes => "data-type-icon XML nodes",
    /// Nested XML elements.
    XmlDepth => "data-type-icon XML depth",
    /// Relationships that target an owned part.
    IncomingRelationships => "data-type-icon incoming relationships",
});

/// A staged, failure-atomic edit of one worksheet's source-bound
/// data-type-icon metadata of one [`crate::Package`]. Dropping it discards
/// it; [`Self::commit`] publishes a new package.
#[derive(Debug, Clone)]
pub struct Edit {
    target: crate::record::Source,
    worksheet: PackURI,
    before: Option<Snapshot>,
    worksheet_views: Option<Vec<Option<ShowDataTypeIcons>>>,
    custom_sheet_views: Option<Vec<Option<ShowDataTypeIcons>>>,
}

impl Edit {
    fn start(target: crate::record::Source, worksheet: &PackURI) -> Result<Self> {
        Ok(Self {
            target,
            worksheet: worksheet.clone(),
            before: None,
            worksheet_views: None,
            custom_sheet_views: None,
        })
    }

    /// Materialize the selected worksheet and its optional named-view part on
    /// first use. A failed load leaves no partial draft, so callers may retry
    /// without retaining source-sized state.
    fn ensure_prepared(&mut self) -> Result<()> {
        if self.before.is_some() {
            return Ok(());
        }
        let before = Snapshot::load(&self.target, &self.worksheet)?;
        self.worksheet_views = Some(before.worksheet_views.clone());
        self.custom_sheet_views = Some(before.custom_sheet_views.clone());
        self.before = Some(before);
        Ok(())
    }

    fn snapshot(&self) -> Result<&Snapshot> {
        self.before
            .as_ref()
            .ok_or_else(|| invalid("data-type-icon edit is not prepared"))
    }

    #[must_use = "use the source snapshot or handle the error"]
    pub fn before(&mut self) -> Result<&Snapshot> {
        self.ensure_prepared()?;
        self.snapshot()
    }

    #[must_use = "use the worksheet views or handle the error"]
    pub fn worksheet_views(&mut self) -> Result<&[Option<ShowDataTypeIcons>]> {
        self.ensure_prepared()?;
        self.worksheet_views
            .as_deref()
            .ok_or_else(|| invalid("data-type-icon edit is not prepared"))
    }

    #[must_use = "use the custom sheet views or handle the error"]
    pub fn custom_sheet_views(&mut self) -> Result<&[Option<ShowDataTypeIcons>]> {
        self.ensure_prepared()?;
        self.custom_sheet_views
            .as_deref()
            .ok_or_else(|| invalid("data-type-icon edit is not prepared"))
    }

    pub fn set_worksheet_view(
        &mut self,
        index: usize,
        value: Option<ShowDataTypeIcons>,
    ) -> Result<bool> {
        self.ensure_prepared()?;
        let slot = self
            .worksheet_views
            .as_mut()
            .ok_or_else(|| invalid("data-type-icon edit is not prepared"))?
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
        self.ensure_prepared()?;
        let slot = self
            .custom_sheet_views
            .as_mut()
            .ok_or_else(|| invalid("data-type-icon edit is not prepared"))?
            .get_mut(index)
            .ok_or_else(|| invalid("Named Sheet View index is out of range"))?;
        if *slot == value {
            return Ok(false);
        }
        *slot = value;
        Ok(true)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        self.ensure_prepared()?;
        let worksheet_views = self.snapshot()?.worksheet_views.clone();
        let custom_sheet_views = self.snapshot()?.custom_sheet_views.clone();
        self.worksheet_views = Some(worksheet_views);
        self.custom_sheet_views = Some(custom_sheet_views);
        Ok(())
    }

    /// Validate and publish the staged metadata as a new package. An
    /// unchanged edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`], [`Error::PatchConflict`] and the
    /// publication's validation errors.
    pub fn commit(self) -> Result<Commit> {
        self.publish().map(|(commit, _)| commit)
    }

    /// [`Self::commit`], keeping the published snapshot.
    fn publish(mut self) -> Result<(Commit, Snapshot)> {
        self.ensure_prepared()?;
        let before = self
            .before
            .take()
            .ok_or_else(|| invalid("data-type-icon edit is not prepared"))?;
        let worksheet_views = self
            .worksheet_views
            .take()
            .ok_or_else(|| invalid("data-type-icon edit is not prepared"))?;
        let custom_sheet_views = self
            .custom_sheet_views
            .take()
            .ok_or_else(|| invalid("data-type-icon edit is not prepared"))?;
        if before.worksheet_views == worksheet_views
            && before.custom_sheet_views == custom_sheet_views
        {
            let patch = Patch::new(before.clone(), before.clone());
            return Ok((
                Commit::new(self.target.into_package(), patch, Diagnostics::new(false)),
                before,
            ));
        }
        if self.target.is_signed() {
            return Err(Error::Signed);
        }
        let current = Snapshot::load(&self.target, &before.worksheet)?;
        if !current.same_source(&before) {
            return Err(Error::PatchConflict {
                part: "data-type-icon source closure".to_owned(),
            });
        }
        let mut candidate = OpcPackage::clone(&self.target);
        if before.worksheet_views != worksheet_views {
            let output = rewrite_worksheet(&before, &worksheet_views)?;
            candidate
                .get_part_mut(&before.worksheet)?
                .set_blob_shared(Arc::new(output));
        }
        if before.custom_sheet_views != custom_sheet_views {
            let source = before
                .named_source
                .as_ref()
                .ok_or_else(|| invalid("Named Sheet Views part is required for a changed edit"))?;
            let output = rewrite_named(
                source.bytes.as_slice(),
                &before.custom_sheet_views,
                &custom_sheet_views,
            )?;
            candidate
                .get_part_mut(&source.name)?
                .set_blob_shared(Arc::new(output));
        }
        let after = Snapshot::load(&candidate, &before.worksheet)?;
        if after.worksheet_views != worksheet_views
            || after.custom_sheet_views != custom_sheet_views
            || !after.same_closure(&before)
        {
            return Err(invalid(
                "data-type-icon publication changed staged metadata or source closure",
            ));
        }
        let patch = Patch::new(before, after.clone());
        let package = self.target.publish(candidate)?;
        Ok((Commit::new(package, patch, Diagnostics::new(true)), after))
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
            return Err(Error::limit(
                LimitKind::PartBytes,
                output.len(),
                MAX_PART_BYTES,
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
            return Err(Error::limit(
                LimitKind::PartBytes,
                output.len(),
                MAX_PART_BYTES,
            ));
        }
    }
    Ok(output)
}

impl crate::Package {
    /// Start an edit of one worksheet's data-type-icon metadata.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused, a `NotFound` refusal for a selector that matches no sheet,
    /// [`Error::NotWorksheet`] for a chart or dialog sheet, and the
    /// validation errors of the metadata.
    pub fn edit_data_type_icons<'a>(
        &self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
    ) -> Result<Edit> {
        let target = crate::record::Source::new(self, "edit_data_type_icons")?;
        let worksheet = self.worksheet_part(sheet, "edit_data_type_icons")?;
        Edit::start(target, &worksheet)
    }

    /// Apply an exact data-type-icon patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_data_type_icons(&self, patch: &Patch) -> Result<Commit> {
        crate::record::replay(self, patch, "apply_data_type_icons")
    }
}

crate::record::raw_test_route!(edit_data_type_icons, worksheet: &PackURI => crate::record::sheet_of);

/// The raw-graph test route for one worksheet.
#[cfg(test)]
pub(crate) fn edit<'a>(target: &'a mut OpcPackage, worksheet: &PackURI) -> Result<Transaction<'a>> {
    Transaction::new(target, worksheet)
}

/// Load the source-bound metadata for one worksheet.
pub fn load(target: &OpcPackage, worksheet: &PackURI) -> Result<Snapshot> {
    Snapshot::load(target, worksheet)
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

    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    pub(crate) fn apply(&self, target: &mut OpcPackage) -> Result<()> {
        let current = Snapshot::load(target, &self.before.worksheet)?;
        if !current.same_source(&self.before) {
            return Err(Error::PatchConflict {
                part: "data-type-icon source closure".to_owned(),
            });
        }
        if self.is_noop() {
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

impl crate::record::Record for Patch {
    fn is_noop(&self) -> bool {
        Self::is_noop(self)
    }

    fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        self.apply(package)
    }
}

#[cfg(test)]
mod lazy_tests {
    use super::*;

    #[test]
    fn edit_entry_stays_cold_until_first_projection() {
        let package = crate::Package::new().expect("minimal XLSX");
        let mut edit = package
            .edit_data_type_icons(0usize)
            .expect("worksheet selector");
        assert!(edit.before.is_none());
        assert!(edit.worksheet_views.is_none());
        assert!(edit.custom_sheet_views.is_none());

        let _ = edit.worksheet_views().expect("first projection");
        assert!(edit.before.is_some());
        assert!(edit.worksheet_views.is_some());
    }

    #[test]
    fn failed_projection_is_retryable_without_partial_state() {
        let source = crate::Package::new().expect("minimal XLSX");
        let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").expect("worksheet URI");
        let mut package = source.opc().clone();
        package
            .get_part_mut(&worksheet)
            .expect("worksheet")
            .set_blob(b"<worksheet".to_vec());
        let package = crate::Package::from_opc(package).expect("defer worksheet XML");
        let mut edit = package
            .edit_data_type_icons(0usize)
            .expect("worksheet selector");
        assert!(edit.worksheet_views().is_err());
        assert!(edit.before.is_none());
        assert!(edit.worksheet_views().is_err());
        assert!(edit.before.is_none());
    }
}
