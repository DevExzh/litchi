//! Immutable source-bound workbook calculation metadata.

use std::sync::Arc;

use litchi_opc::constants::content_type as ct;
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{OpcPackage, PackURI, Part, PartView, SourceBackedPackage, TargetMode};

use super::{Features, Limits, Properties, inspect};
use crate::error::{Error, Result, invalid};
use crate::source_provenance::{SourceBinding, SourceProvenance};

/// The semantic and exact physical state of workbook calculation metadata.
#[derive(Clone, Debug)]
pub struct Snapshot {
    properties: Option<Properties>,
    features: Option<Features>,
    source: SourceState,
    binding: SourceBinding,
    limits: Limits,
}

impl Snapshot {
    /// Load calculation metadata with the default resource policy.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        Self::load_with_limits(package, &Limits::default())
    }

    /// Load calculation metadata from the package's resolved workbook part.
    pub fn load_with_limits(package: &OpcPackage, limits: &Limits) -> Result<Self> {
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;

        // Reject oversized input before retaining its Arc or allocating owner
        // graph state.
        if workbook.blob().len() > limits.max_raw_bytes() {
            return Err(invalid(
                "workbook calculation metadata exceeds raw byte limit",
            ));
        }
        let inspection = inspect(workbook.blob(), limits)?;
        let properties = inspection.properties;
        let features = inspection.features;
        let source = SourceState::capture(package, workbook)?;
        Ok(Self {
            properties,
            features,
            source,
            binding: SourceBinding::default(),
            limits: *limits,
        })
    }

    pub(super) fn load_source_backed_with_limits(
        package: &SourceBackedPackage,
        limits: &Limits,
    ) -> Result<Self> {
        let workbook = package.main_document_part()?;
        require_workbook_content_type(workbook.content_type())?;
        let bytes = workbook.data()?.into_arc()?;
        if bytes.len() > limits.max_raw_bytes() {
            return Err(invalid(
                "workbook calculation metadata exceeds raw byte limit",
            ));
        }
        let (properties, features) = {
            let inspection = inspect(bytes.as_slice(), limits)?;
            (inspection.properties, inspection.features)
        };
        let source = SourceState::capture_source_backed(package, &workbook, bytes)?;
        let mut snapshot = Self {
            properties,
            features,
            source,
            binding: SourceBinding::default(),
            limits: *limits,
        };
        snapshot.binding = SourceBinding::capture(package)?;
        Ok(snapshot)
    }

    pub(super) fn from_rewritten_source(source: &Self, bytes: Vec<u8>) -> Result<Self> {
        if bytes.len() > source.limits.max_raw_bytes() {
            return Err(invalid(
                "workbook calculation metadata exceeds raw byte limit",
            ));
        }
        let (properties, features) = {
            let inspection = inspect(bytes.as_slice(), &source.limits)?;
            (inspection.properties, inspection.features)
        };
        let mut state = source.source.clone();
        state.bytes = Arc::new(bytes);
        Ok(Self {
            properties,
            features,
            source: state,
            binding: source.binding.clone(),
            limits: source.limits,
        })
    }

    /// Alias for [`Self::load`] emphasizing the source-bound result.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    /// Alias for [`Self::load_with_limits`].
    pub fn read_with_limits(package: &OpcPackage, limits: &Limits) -> Result<Self> {
        Self::load_with_limits(package, limits)
    }

    /// Exact authored `calcPr` state, if present.
    #[must_use]
    pub fn properties(&self) -> Option<&Properties> {
        self.properties.as_ref()
    }

    /// Ordered calculation-feature occurrences, if present.
    #[must_use]
    pub fn features(&self) -> Option<&Features> {
        self.features.as_ref()
    }

    /// Exact source workbook XML.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source.bytes.as_slice()
    }

    /// Shared ownership of the exact source workbook XML.
    #[must_use]
    pub fn source_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.source.bytes)
    }

    /// Resolved package part containing the workbook root.
    #[must_use]
    pub fn workbook_part_name(&self) -> &PackURI {
        &self.source.part_name
    }

    /// Exact source content type of the workbook part.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.source.content_type
    }

    /// Resource policy retained for publication and patch application.
    #[must_use]
    pub fn limits(&self) -> Limits {
        self.limits
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source && self.binding.same_or_unavailable(&other.binding)
    }

    /// Check the retained source lineage and revision without reloading XML.
    pub(crate) fn matches_source_backed(
        &self,
        package: &SourceBackedPackage,
    ) -> Result<SourceProvenance> {
        self.binding.check(package)
    }

    /// Compare the retained owner and workbook part without processing XML.
    pub(crate) fn matches_current_source(&self, package: &OpcPackage) -> bool {
        let Ok(workbook) = package.main_document_part() else {
            return false;
        };
        workbook.partname() == &self.source.part_name
            && workbook.content_type() == self.source.content_type
            && workbook.blob() == self.source.bytes.as_slice()
            && current_owner_relationship(package.rels())
                .is_some_and(|relationship| self.source.owner_relationship.matches(relationship))
    }

    pub(crate) fn same_semantics(
        &self,
        properties: Option<&Properties>,
        features: Option<&Features>,
    ) -> bool {
        same_properties(self.properties(), properties) && self.features() == features
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source && self.limits == other.limits
    }
}

impl Eq for Snapshot {}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SourceState {
    part_name: PackURI,
    content_type: String,
    bytes: Arc<Vec<u8>>,
    owner_relationship: SourceRelationship,
}

impl SourceState {
    fn capture(package: &OpcPackage, workbook: &dyn Part) -> Result<Self> {
        Ok(Self {
            part_name: workbook.partname().clone(),
            content_type: copy_string(
                workbook.content_type(),
                "calculation metadata workbook content type",
            )?,
            bytes: workbook.blob_arc(),
            owner_relationship: SourceRelationship::capture(
                current_owner_relationship(package.rels())
                    .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?,
            )?,
        })
    }

    fn capture_source_backed(
        package: &SourceBackedPackage,
        workbook: &PartView<'_>,
        bytes: Arc<Vec<u8>>,
    ) -> Result<Self> {
        Ok(Self {
            part_name: workbook.partname().clone(),
            content_type: copy_string(
                workbook.content_type(),
                "calculation metadata workbook content type",
            )?,
            bytes,
            owner_relationship: SourceRelationship::capture(
                current_owner_relationship(package.rels())
                    .ok_or_else(|| invalid("workbook has no unique officeDocument owner"))?,
            )?,
        })
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourceRelationship {
    id: String,
    relationship_type: String,
    target: String,
    mode: TargetMode,
}

impl SourceRelationship {
    fn capture(relationship: &litchi_opc::Relationship) -> Result<Self> {
        Ok(Self {
            id: copy_string(
                relationship.r_id(),
                "calculation metadata owner relationship ID",
            )?,
            relationship_type: copy_string(
                relationship.reltype(),
                "calculation metadata owner relationship type",
            )?,
            target: copy_string(
                relationship.target_ref(),
                "calculation metadata owner relationship target",
            )?,
            mode: relationship.target_mode(),
        })
    }

    fn matches(&self, relationship: &litchi_opc::Relationship) -> bool {
        relationship.r_id() == self.id
            && relationship.reltype() == self.relationship_type
            && relationship.target_ref() == self.target
            && relationship.target_mode() == self.mode
    }
}

fn current_owner_relationship(
    relationships: &litchi_opc::Relationships,
) -> Option<&litchi_opc::Relationship> {
    let mut owners = relationships.iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    });
    let owner = owners.next()?;
    if owners.next().is_some() || owner.target_mode() != TargetMode::Internal {
        return None;
    }
    Some(owner)
}

fn copy_string(value: &str, resource: &'static str) -> Result<String> {
    let mut copied = String::new();
    copied
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copied.push_str(value);
    Ok(copied)
}

pub(crate) fn same_properties(left: Option<&Properties>, right: Option<&Properties>) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => left.same_specification(right),
        (None, None) => true,
        _ => false,
    }
}

fn require_workbook_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        ct::SML_SHEET_MAIN
            | ct::SML_TEMPLATE_MAIN
            | ct::SML_SHEET_MACRO_MAIN
            | ct::SML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "main part has non-XLSX content type '{content_type}'"
        )))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::BlobPart;
    use litchi_opc::constants::relationship_type as rt;

    #[test]
    fn resolves_noncanonical_main_part_for_all_workbook_types_and_conformance_modes() {
        let content_types = [
            ct::SML_SHEET_MAIN,
            ct::SML_TEMPLATE_MAIN,
            ct::SML_SHEET_MACRO_MAIN,
            ct::SML_TEMPLATE_MACRO_MAIN,
        ];
        for (index, content_type) in content_types.into_iter().enumerate() {
            let strict = index % 2 == 1;
            let namespace = if strict {
                "http://purl.oclc.org/ooxml/spreadsheetml/main"
            } else {
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            };
            let mut package = OpcPackage::new();
            let part_name = PackURI::new(format!("/custom/book{index}.xml")).unwrap();
            package.add_part(Box::new(BlobPart::new(
                part_name.clone(),
                content_type.to_owned(),
                format!(r#"<workbook xmlns="{namespace}"><calcPr calcId="7"/></workbook>"#)
                    .into_bytes(),
            )));
            package
                .rels_mut()
                .try_add_relationship(
                    if strict {
                        rt::STRICT_OFFICE_DOCUMENT
                    } else {
                        rt::OFFICE_DOCUMENT
                    }
                    .to_owned(),
                    format!("custom/book{index}.xml"),
                    "rIdMain".to_owned(),
                    TargetMode::Internal,
                )
                .unwrap();

            let snapshot = Snapshot::load(&package).unwrap();
            assert_eq!(snapshot.workbook_part_name(), &part_name);
            assert_eq!(snapshot.content_type(), content_type);
            assert_eq!(snapshot.properties().unwrap().calculation_id(), 7);
        }
    }

    #[test]
    fn enforces_raw_limit_before_retaining_source_arc() {
        let package = crate::package::build_minimal_package().unwrap();
        let limits = Limits::new().with_max_raw_bytes(8).unwrap();
        assert!(Snapshot::load_with_limits(&package, &limits).is_err());
    }

    #[test]
    fn source_binding_ignores_unrelated_workbook_relationships() {
        let mut package = crate::package::build_minimal_package().unwrap();
        let snapshot = Snapshot::load(&package).unwrap();
        let workbook_name = package.main_document_part().unwrap().partname().clone();
        package
            .get_part_mut(&workbook_name)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                rt::HYPERLINK.to_owned(),
                "https://example.invalid/".to_owned(),
                "rIdUnrelated".to_owned(),
                TargetMode::External,
            )
            .unwrap();

        assert!(snapshot.matches_current_source(&package));
        assert!(snapshot.same_source(&Snapshot::load(&package).unwrap()));
    }
}
