#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::match_same_arms,
    reason = "separate arms document distinct OOXML grammar cases"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
#![expect(
    clippy::similar_names,
    reason = "domain names mirror distinct OOXML roles"
)]
#![expect(
    clippy::unnecessary_wraps,
    reason = "the Result signature preserves a uniform fallible codec API"
)]
//! Low-level glossary OPC graph vocabulary and relationship closure validation.

use super::codec::{
    Content, Node, bounded, invalid, read, valid_ncname, validate_physical_part, validate_raw_part,
    write,
};
use super::model::{Catalog, Conformance};
use super::package::Owner;
use super::{
    ACTIVE_X_BINARY_CT, ACTIVE_X_BINARY_REL, ACTIVE_X_DESCRIPTOR_CT, ATTACHED_TOOLBARS_CT,
    ATTACHED_TOOLBARS_REL, Arc, BlobPart, CHART_COLOR_STYLE_CT, CHART_COLOR_STYLE_REL,
    CHART_COLOR_STYLE_REL_2012, CHART_STYLE_CT, CHART_STYLE_REL, CHART_STYLE_REL_2012, CT,
    CUSTOMIZATIONS_CT, CUSTOMIZATIONS_REL, ContentType, DIAGRAM_DRAWING_REL, Error, FONT_DATA_CT,
    FONT_TTF_CT, HashMap, HashSet, MAX_GRAPH_BYTES, MAX_GRAPH_METADATA_BYTES, MAX_VALUES,
    OBFUSCATED_FONT_CT, OpcPackage, PRINTER_SETTINGS_CT, PackURI, Part, R, RECIPIENT_DATA_CT, REL,
    RS, Result, STRICT_REL, STYLES_EFFECTS_CT, STYLES_EFFECTS_REL, VecDeque, ct,
};
pub mod raw {
    use super::{Catalog, Conformance};
    use std::sync::Arc;

    #[derive(Clone, Debug, PartialEq, Eq)]
    pub struct Rel {
        pub id: String,
        pub kind: String,
        pub target: String,
        pub external: bool,
    }

    #[derive(Clone, Debug, PartialEq, Eq)]
    pub struct Part {
        pub name: String,
        pub content_type: String,
        pub(in crate::glossary) data: Arc<Vec<u8>>,
        pub rels: Vec<Rel>,
    }

    impl Part {
        /// Validate basic physical metadata and take ownership of a payload.
        ///
        /// # Errors
        ///
        /// Returns an error if the operation cannot be completed.
        pub fn new(
            name: impl Into<String>,
            content_type: impl Into<String>,
            data: Vec<u8>,
        ) -> super::Result<Self> {
            let name = name.into();
            let content_type = content_type.into();
            super::validate_raw_part(&name, &content_type, data.len())?;
            Ok(Self {
                name,
                content_type,
                data: Arc::new(data),
                rels: Vec::new(),
            })
        }

        #[must_use]
        pub fn data(&self) -> &[u8] {
            self.data.as_slice()
        }

        /// Absolute OPC part name.
        #[must_use]
        pub fn name(&self) -> &str {
            &self.name
        }

        /// OPC content type retained for this opaque part.
        #[must_use]
        pub fn content_type(&self) -> &str {
            &self.content_type
        }

        /// Relationships owned by this opaque part.
        #[must_use]
        pub fn relationships(&self) -> &[Rel] {
            &self.rels
        }

        /// Replace the opaque payload after checking its physical bounds.
        ///
        /// # Errors
        ///
        /// Returns an error if the operation cannot be completed.
        pub fn replace_data(&mut self, data: Vec<u8>) -> super::Result<()> {
            super::validate_raw_part(&self.name, &self.content_type, data.len())?;
            self.data = Arc::new(data);
            Ok(())
        }

        /// Replace the opaque relationship list after bounded metadata checks.
        ///
        /// # Errors
        ///
        /// Returns an error if the operation cannot be completed.
        pub fn set_relationships(&mut self, relationships: Vec<Rel>) -> super::Result<()> {
            if relationships.len() > super::MAX_VALUES {
                return Err(super::invalid("glossary relationship limit exceeded"));
            }
            let mut ids = std::collections::HashSet::new();
            for relationship in &relationships {
                super::validate_relationship_metadata(
                    &relationship.id,
                    &relationship.kind,
                    &relationship.target,
                )?;
                if !ids.insert(&relationship.id) {
                    return Err(super::invalid("duplicate glossary relationship ID"));
                }
            }
            self.rels = relationships;
            Ok(())
        }

        /// Consume the part and retain shared ownership of its payload.
        #[must_use]
        pub fn into_data(self) -> Arc<Vec<u8>> {
            self.data
        }

        pub(in crate::glossary) fn shared_data(&self) -> Arc<Vec<u8>> {
            Arc::clone(&self.data)
        }

        pub(in crate::glossary) fn from_shared(
            name: String,
            content_type: String,
            data: Arc<Vec<u8>>,
            rels: Vec<Rel>,
        ) -> Self {
            Self {
                name,
                content_type,
                data,
                rels,
            }
        }
    }

    #[derive(Clone, Debug, PartialEq, Eq)]
    pub struct Graph {
        pub catalog: Catalog,
        pub conformance: Conformance,
        pub rels: Vec<Rel>,
        pub parts: Vec<Part>,
        pub(in crate::glossary) root_name: String,
        pub(in crate::glossary) root_xml: Option<Arc<Vec<u8>>>,
        pub(in crate::glossary) owner_main: Option<String>,
        pub(in crate::glossary) owner_id: Option<String>,
        pub(in crate::glossary) owner_target: Option<String>,
    }

    impl Graph {
        #[must_use]
        pub fn new(catalog: Catalog, conformance: Conformance) -> Self {
            Self {
                catalog,
                conformance,
                rels: Vec::new(),
                parts: Vec::new(),
                root_name: "/word/glossary/document.xml".to_owned(),
                root_xml: None,
                owner_main: None,
                owner_id: None,
                owner_target: None,
            }
        }

        /// Producer-selected glossary root part name.
        #[must_use]
        pub fn root_name(&self) -> &str {
            &self.root_name
        }

        /// Original producer XML, when this graph was loaded from a package.
        pub fn root_xml(&self) -> Option<&[u8]> {
            self.root_xml.as_deref().map(Vec::as_slice)
        }

        /// Producer-selected ID of the main-document owner relationship.
        #[must_use]
        pub fn owner_relationship_id(&self) -> Option<&str> {
            self.owner_id.as_deref()
        }

        /// Producer-selected main-document part that owns this graph.
        #[must_use]
        pub fn owner_main_part(&self) -> Option<&str> {
            self.owner_main.as_deref()
        }

        /// Producer-selected literal target of the main-document owner relationship.
        #[must_use]
        pub fn owner_target(&self) -> Option<&str> {
            self.owner_target.as_deref()
        }
    }

    /// Load the complete glossary-owned OPC graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn load(package: &litchi_opc::OpcPackage) -> super::Result<Option<Graph>> {
        super::load_graph(package)
    }

    /// Publish a complete graph without consuming the caller's recovery copy.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn put(package: &mut litchi_opc::OpcPackage, graph: &Graph) -> super::Result<bool> {
        super::put_graph(package, graph)
    }

    /// Remove and return the complete graph for graph-preserving transfer.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn remove(package: &mut litchi_opc::OpcPackage) -> super::Result<Option<Graph>> {
        super::remove_graph(package)
    }
}

pub(in crate::glossary) fn load_graph(package: &OpcPackage) -> Result<Option<raw::Graph>> {
    let Some(owner) = locate(package)? else {
        return Ok(None);
    };
    let root = package.get_part(&owner.root)?;
    let (catalog, conformance) = read(root.blob())?;
    if conformance != owner.conformance {
        return Err(invalid(
            "glossary relationship and XML use different conformance families",
        ));
    }
    validate_relationship_integrity(&catalog, root, conformance)?;
    let owned = glossary_owned_parts(package, &owner.root, conformance)?;
    validate_exclusive_ownership(package, &owner, &owned)?;
    let rels = copy_relationships(root)?;
    let mut names = owned
        .into_iter()
        .filter(|uri| uri != &owner.root)
        .collect::<Vec<_>>();
    names.sort_by(|left, right| left.as_str().cmp(right.as_str()));
    let mut parts = Vec::new();
    parts
        .try_reserve_exact(names.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary auxiliary graph",
            source,
        })?;
    let mut total_bytes = 0usize;
    for uri in names {
        let part = package.get_part(&uri)?;
        total_bytes = total_bytes
            .checked_add(part.blob().len())
            .ok_or_else(|| invalid("glossary auxiliary payload size overflow"))?;
        if total_bytes > MAX_GRAPH_BYTES {
            return Err(invalid(
                "glossary auxiliary graph exceeds the 256 MiB aggregate limit",
            ));
        }
        parts.push(raw::Part::from_shared(
            uri.as_str().to_owned(),
            part.content_type().to_owned(),
            part.blob_arc(),
            copy_relationships(part)?,
        ));
    }
    Ok(Some(raw::Graph {
        catalog,
        conformance,
        rels,
        parts,
        root_name: owner.root.as_str().to_owned(),
        root_xml: Some(root.blob_arc()),
        owner_main: Some(owner.main.as_str().to_owned()),
        owner_id: Some(owner.relationship_id),
        owner_target: Some(owner.relationship_target),
    }))
}

pub(in crate::glossary) fn copy_relationships(part: &dyn Part) -> Result<Vec<raw::Rel>> {
    let relationship_count = part.rels().iter().count();
    let mut rels = Vec::new();
    rels.try_reserve_exact(relationship_count)
        .map_err(|source| Error::Allocation {
            resource: "glossary relationships",
            source,
        })?;
    for relationship in part.rels().iter() {
        rels.push(raw::Rel {
            id: relationship.r_id().to_owned(),
            kind: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        });
    }
    rels.sort_by(|left, right| left.id.cmp(&right.id));
    Ok(rels)
}

pub(in crate::glossary) fn put_graph(package: &mut OpcPackage, value: &raw::Graph) -> Result<bool> {
    validate_package_conformance(package, value.conformance)?;
    validate_raw_graph_metadata(value)?;
    let preserve_root = if let Some(xml) = &value.root_xml {
        let (source, source_conformance) = read(xml.as_slice())?;
        source_conformance == value.conformance && source == value.catalog
    } else {
        false
    };
    let root_uri = PackURI::new(&value.root_name).map_err(Error::Uri)?;
    if is_signature_part(&root_uri) || is_reserved_physical_part(&root_uri) {
        return Err(invalid(
            "glossary root cannot use reserved OPC package infrastructure",
        ));
    }
    let mut root = if preserve_root {
        BlobPart::new_shared(
            root_uri.clone(),
            CT.to_owned(),
            value
                .root_xml
                .as_ref()
                .ok_or_else(|| invalid("preserved glossary XML is missing"))?
                .clone(),
        )
    } else {
        BlobPart::new(
            root_uri.clone(),
            CT.to_owned(),
            write(&value.catalog, value.conformance)?,
        )
    };
    let canonical_catalog;
    let effective_catalog = if preserve_root {
        &value.catalog
    } else {
        let (canonical, canonical_conformance) = read(root.blob())?;
        if canonical_conformance != value.conformance {
            return Err(invalid("canonical glossary changed conformance"));
        }
        canonical_catalog = canonical;
        &canonical_catalog
    };
    add_relationships(&mut root, &value.rels, value.conformance)?;
    validate_relationship_integrity(effective_catalog, &root, value.conformance)?;
    let staged_root_xml = root.blob_arc();
    if value.parts.len() > MAX_VALUES {
        return Err(invalid("glossary auxiliary part limit exceeded"));
    }
    let mut staged = HashMap::new();
    staged
        .try_reserve(value.parts.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary graph staging",
            source,
        })?;
    let mut total_bytes = 0usize;
    for auxiliary in &value.parts {
        total_bytes = total_bytes
            .checked_add(auxiliary.data().len())
            .ok_or_else(|| invalid("glossary auxiliary payload size overflow"))?;
        if total_bytes > MAX_GRAPH_BYTES {
            return Err(invalid(
                "glossary auxiliary graph exceeds the 256 MiB aggregate limit",
            ));
        }
        let uri = validate_physical_part(
            &auxiliary.name,
            &auxiliary.content_type,
            auxiliary.data().len(),
        )?;
        if uri == root_uri {
            return Err(invalid(format!(
                "glossary auxiliary part '{}' conflicts with the root",
                uri.as_str()
            )));
        }
        let mut part = BlobPart::new_shared(
            uri.clone(),
            auxiliary.content_type.clone(),
            auxiliary.shared_data(),
        );
        add_relationships(&mut part, &auxiliary.rels, value.conformance)?;
        if staged.insert(uri.clone(), part).is_some() {
            return Err(invalid(format!(
                "duplicate glossary auxiliary part '{uri}'"
            )));
        }
    }

    if let Some(current) = load_graph(package)?
        && graph_matches_catalog(&current, value, effective_catalog, Some(&staged_root_xml))?
    {
        return Ok(false);
    }

    validate_all_internal_targets(package)?;
    let owner = locate(package)?;
    let old_owned = if let Some(owner) = &owner {
        let owned = glossary_owned_parts(package, &owner.root, owner.conformance)?;
        validate_exclusive_ownership(package, owner, &owned)?;
        owned
    } else {
        HashSet::new()
    };

    let main = package.main_document_part()?.partname().clone();
    let mut candidate = package.clone();
    candidate.unsign();
    if let Some(owner) = &owner {
        candidate
            .get_part_mut(&main)?
            .rels_mut()
            .remove(&owner.relationship_id);
    }
    for uri in &old_owned {
        candidate.remove_part(uri);
    }
    for (_, part) in staged {
        candidate.try_add_part(Box::new(part))?;
    }
    candidate.try_add_part(Box::new(root))?;
    let generated_target = root_uri.relative_ref(main.base_uri());
    let owner_target = if value.owner_main.as_deref() == Some(main.as_str()) {
        value.owner_target.as_deref().unwrap_or(&generated_target)
    } else {
        &generated_target
    };
    let owner_relationships = candidate.get_part_mut(&main)?.rels_mut();
    let preserve_owner_id = value
        .owner_id
        .as_ref()
        .filter(|owner_id| owner_relationships.get(owner_id).is_none());
    if let Some(owner_id) = preserve_owner_id {
        owner_relationships.try_add_relationship(
            value.conformance.glossary_relationship().to_owned(),
            owner_target.to_owned(),
            owner_id.clone(),
            litchi_opc::TargetMode::Internal,
        )?;
    } else {
        owner_relationships.get_or_add(value.conformance.glossary_relationship(), owner_target);
    }

    validate_all_internal_targets(&candidate)?;
    let round_trip =
        load_graph(&candidate)?.ok_or_else(|| invalid("staged glossary graph is missing"))?;
    if !graph_matches_catalog(
        &round_trip,
        value,
        effective_catalog,
        Some(&staged_root_xml),
    )? {
        return Err(invalid("staged glossary graph did not round-trip"));
    }
    *package = candidate;
    Ok(true)
}

pub(in crate::glossary) fn remove_graph(package: &mut OpcPackage) -> Result<Option<raw::Graph>> {
    let Some(owner) = locate(package)? else {
        return Ok(None);
    };
    let graph = load_graph(package)?.ok_or_else(|| invalid("missing glossary"))?;
    let owned = glossary_owned_parts(package, &owner.root, owner.conformance)?;
    validate_exclusive_ownership(package, &owner, &owned)?;
    validate_all_internal_targets(package)?;

    let mut candidate = package.clone();
    candidate.unsign();
    candidate
        .get_part_mut(&owner.main)?
        .rels_mut()
        .remove(&owner.relationship_id);
    for uri in owned {
        candidate.remove_part(&uri);
    }
    validate_all_internal_targets(&candidate)?;
    *package = candidate;
    Ok(Some(graph))
}

pub(in crate::glossary) fn graph_matches_catalog(
    actual: &raw::Graph,
    expected: &raw::Graph,
    expected_catalog: &Catalog,
    expected_root_xml: Option<&Arc<Vec<u8>>>,
) -> Result<bool> {
    Ok(actual.catalog == *expected_catalog
        && actual.conformance == expected.conformance
        && actual.root_name == expected.root_name
        && expected_root_xml.is_none_or(|xml| {
            actual.root_xml.as_ref().is_some_and(|actual| {
                Arc::ptr_eq(actual, xml) || actual.as_slice() == xml.as_slice()
            })
        })
        && keyed_rels_match(&actual.rels, &expected.rels)?
        && keyed_parts_match(&actual.parts, &expected.parts)?)
}

pub(in crate::glossary) fn keyed_rels_match(left: &[raw::Rel], right: &[raw::Rel]) -> Result<bool> {
    if left.len() != right.len() {
        return Ok(false);
    }
    let mut by_id = HashMap::new();
    by_id
        .try_reserve(right.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary relationship comparison",
            source,
        })?;
    for relationship in right {
        if by_id
            .insert(relationship.id.as_str(), relationship)
            .is_some()
        {
            return Ok(false);
        }
    }
    Ok(left
        .iter()
        .all(|relationship| by_id.get(relationship.id.as_str()) == Some(&relationship)))
}

pub(in crate::glossary) fn keyed_parts_match(
    left: &[raw::Part],
    right: &[raw::Part],
) -> Result<bool> {
    if left.len() != right.len() {
        return Ok(false);
    }
    let mut by_name = HashMap::new();
    by_name
        .try_reserve(right.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary part comparison",
            source,
        })?;
    for part in right {
        if by_name.insert(part.name.as_str(), part).is_some() {
            return Ok(false);
        }
    }
    for part in left {
        let Some(candidate) = by_name.get(part.name.as_str()) else {
            return Ok(false);
        };
        if part.content_type != candidate.content_type
            || (!Arc::ptr_eq(&part.data, &candidate.data) && part.data() != candidate.data())
            || !keyed_rels_match(&part.rels, &candidate.rels)?
        {
            return Ok(false);
        }
    }
    Ok(true)
}

pub(in crate::glossary) fn is_signature_part(uri: &PackURI) -> bool {
    const PREFIX: &str = "/_xmlsignatures/";
    uri.as_str()
        .get(..PREFIX.len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case(PREFIX))
}

pub(in crate::glossary) fn is_reserved_physical_part(uri: &PackURI) -> bool {
    let value = uri.as_str();
    if value == "/" || value.eq_ignore_ascii_case("/[Content_Types].xml") {
        return true;
    }
    let Some((directory, filename)) = value.rsplit_once('/') else {
        return false;
    };
    filename
        .get(filename.len().saturating_sub(5)..)
        .is_some_and(|suffix| suffix.eq_ignore_ascii_case(".rels"))
        && directory
            .rsplit('/')
            .next()
            .is_some_and(|segment| segment.eq_ignore_ascii_case("_rels"))
}

pub(in crate::glossary) fn locate(package: &OpcPackage) -> Result<Option<Owner>> {
    let package_conformance = package_conformance(package)?;
    let main_part = package.main_document_part()?;
    if !matches!(
        main_part.content_type(),
        ct::WML_DOCUMENT_MAIN
            | ct::WML_TEMPLATE_MAIN
            | ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
    ) {
        return Err(Error::ContentType {
            expected: format!(
                "{}, {}, {}, or {}",
                ct::WML_DOCUMENT_MAIN,
                ct::WML_TEMPLATE_MAIN,
                ct::WML_DOCUMENT_MACRO_MAIN,
                ct::WML_TEMPLATE_MACRO_MAIN,
            ),
            actual: main_part.content_type().to_owned(),
        });
    }
    let main = main_part.partname().clone();
    if package
        .rels()
        .iter()
        .any(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
    {
        return Err(invalid(
            "package root cannot source a glossary-document relationship",
        ));
    }

    let mut found = None;
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
        {
            validate_relationship_metadata(
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
            )?;
            if part.partname() != &main {
                return Err(invalid(format!(
                    "glossary-document relationship has invalid source '{}'",
                    part.partname()
                )));
            }
            if found.is_some() {
                return Err(invalid("main document has multiple glossary relationships"));
            }
            if relationship.is_external() {
                return Err(invalid("glossary relationship cannot be external"));
            }
            let conformance = if relationship.reltype() == STRICT_REL {
                Conformance::Strict
            } else {
                Conformance::Transitional
            };
            if conformance != package_conformance {
                return Err(invalid(
                    "glossary relationship does not match package conformance",
                ));
            }
            let requested = relationship.target_partname()?;
            let target = package.get_part(&requested)?.partname().clone();
            let target_part = package.get_part(&target)?;
            if target_part.content_type() != CT {
                return Err(Error::ContentType {
                    expected: CT.to_owned(),
                    actual: target_part.content_type().to_owned(),
                });
            }
            found = Some(Owner {
                main: main.clone(),
                root: target,
                relationship_id: relationship.r_id().to_owned(),
                relationship_target: relationship.target_ref().to_owned(),
                conformance,
            });
        }
    }
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == CT)
    {
        match &found {
            Some(owner) if part.partname() == &owner.root => {},
            Some(_) => {
                return Err(invalid(format!(
                    "orphan glossary content-type part '{}' exists beside the owned root",
                    part.partname()
                )));
            },
            None => {
                return Err(invalid(format!(
                    "orphan glossary content-type part '{}' has no main-document relationship",
                    part.partname()
                )));
            },
        }
    }
    Ok(found)
}

pub(in crate::glossary) fn package_conformance(package: &OpcPackage) -> Result<Conformance> {
    use litchi_opc::constants::relationship_type::{OFFICE_DOCUMENT, STRICT_OFFICE_DOCUMENT};
    let mut relationships = package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            OFFICE_DOCUMENT | STRICT_OFFICE_DOCUMENT
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| invalid("main-document relationship is missing"))?;
    if relationships.next().is_some() {
        return Err(invalid("package has multiple main-document relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("main-document relationship cannot be external"));
    }
    Ok(if relationship.reltype() == STRICT_OFFICE_DOCUMENT {
        Conformance::Strict
    } else {
        Conformance::Transitional
    })
}

pub(in crate::glossary) fn validate_package_conformance(
    package: &OpcPackage,
    requested: Conformance,
) -> Result<()> {
    if package_conformance(package)? == requested {
        Ok(())
    } else {
        Err(invalid(
            "requested glossary conformance does not match the document package",
        ))
    }
}

pub(in crate::glossary) fn add_relationships(
    part: &mut BlobPart,
    relationships: &[raw::Rel],
    conformance: Conformance,
) -> Result<()> {
    if relationships.len() > MAX_VALUES {
        return Err(invalid("glossary relationship limit exceeded"));
    }
    let mut ids = HashSet::new();
    for relationship in relationships {
        validate_relationship_metadata(&relationship.id, &relationship.kind, &relationship.target)?;
        if !ids.insert(relationship.id.clone()) {
            return Err(invalid("duplicate glossary relationship ID"));
        }
        relationship_kind(conformance, &relationship.kind).ok_or_else(|| {
            invalid(format!(
                "unsupported glossary relationship type '{}' for {conformance:?} conformance",
                relationship.kind
            ))
        })?;
        part.rels_mut().try_add_relationship(
            relationship.kind.clone(),
            relationship.target.clone(),
            relationship.id.clone(),
            if relationship.external {
                litchi_opc::TargetMode::External
            } else {
                litchi_opc::TargetMode::Internal
            },
        )?;
    }
    Ok(())
}

pub(in crate::glossary) fn validate_relationship_metadata(
    id: &str,
    kind: &str,
    target: &str,
) -> Result<()> {
    bounded(id)?;
    bounded(kind)?;
    bounded(target)?;
    if !valid_ncname(id) {
        return Err(invalid(format!(
            "glossary relationship ID '{id}' is not an XML NCName"
        )));
    }
    if kind.is_empty()
        || kind
            .chars()
            .any(|character| character.is_whitespace() || character.is_control())
    {
        return Err(invalid("glossary relationship type is not a valid URI"));
    }
    if target.is_empty() || target.chars().any(char::is_control) {
        return Err(invalid(
            "glossary relationship target is empty or contains a control character",
        ));
    }
    Ok(())
}

pub(in crate::glossary) fn validate_raw_graph_metadata(graph: &raw::Graph) -> Result<()> {
    if graph.parts.len() > MAX_VALUES {
        return Err(invalid("glossary auxiliary part limit exceeded"));
    }
    let mut relationship_count = 0usize;
    let mut metadata_bytes = 0usize;
    add_graph_metadata(&mut metadata_bytes, &graph.root_name)?;
    if let Some(owner_main) = &graph.owner_main {
        bounded(owner_main)?;
        PackURI::new(owner_main).map_err(Error::Uri)?;
        add_graph_metadata(&mut metadata_bytes, owner_main)?;
    }
    if let Some(owner_id) = &graph.owner_id {
        bounded(owner_id)?;
        if !valid_ncname(owner_id) {
            return Err(invalid(
                "glossary owner relationship ID is not an XML NCName",
            ));
        }
        add_graph_metadata(&mut metadata_bytes, owner_id)?;
    }
    if let Some(owner_target) = &graph.owner_target {
        bounded(owner_target)?;
        if owner_target.is_empty() || owner_target.chars().any(char::is_control) {
            return Err(invalid("glossary owner relationship target is invalid"));
        }
        add_graph_metadata(&mut metadata_bytes, owner_target)?;
    }
    validate_raw_relationship_set(&graph.rels, &mut relationship_count, &mut metadata_bytes)?;
    for part in &graph.parts {
        bounded(&part.name)?;
        bounded(&part.content_type)?;
        add_graph_metadata(&mut metadata_bytes, &part.name)?;
        add_graph_metadata(&mut metadata_bytes, &part.content_type)?;
        validate_raw_relationship_set(&part.rels, &mut relationship_count, &mut metadata_bytes)?;
    }
    Ok(())
}

pub(in crate::glossary) fn validate_raw_relationship_set(
    relationships: &[raw::Rel],
    total_count: &mut usize,
    metadata_bytes: &mut usize,
) -> Result<()> {
    if relationships.len() > MAX_VALUES {
        return Err(invalid("glossary relationship limit exceeded"));
    }
    *total_count = total_count
        .checked_add(relationships.len())
        .ok_or_else(|| invalid("glossary relationship count overflow"))?;
    if *total_count > MAX_VALUES {
        return Err(invalid("glossary graph-wide relationship limit exceeded"));
    }
    for relationship in relationships {
        validate_relationship_metadata(&relationship.id, &relationship.kind, &relationship.target)?;
        add_graph_metadata(metadata_bytes, &relationship.id)?;
        add_graph_metadata(metadata_bytes, &relationship.kind)?;
        add_graph_metadata(metadata_bytes, &relationship.target)?;
    }
    Ok(())
}

pub(in crate::glossary) fn add_graph_metadata(total: &mut usize, value: &str) -> Result<()> {
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| invalid("glossary graph metadata size overflow"))?;
    if *total > MAX_GRAPH_METADATA_BYTES {
        return Err(invalid(
            "glossary graph metadata exceeds the 32 MiB aggregate limit",
        ));
    }
    Ok(())
}

pub(in crate::glossary) fn relationship_kind(
    conformance: Conformance,
    value: &str,
) -> Option<&str> {
    if value == STYLES_EFFECTS_REL {
        return Some("stylesWithEffects");
    }
    if value == CUSTOMIZATIONS_REL {
        return Some("keyMapCustomizations");
    }
    if value == ATTACHED_TOOLBARS_REL {
        return Some("attachedToolbars");
    }
    if value == DIAGRAM_DRAWING_REL {
        return Some("diagramDrawing");
    }
    if matches!(value, CHART_STYLE_REL | CHART_STYLE_REL_2012) {
        return Some("chartStyle");
    }
    if matches!(value, CHART_COLOR_STYLE_REL | CHART_COLOR_STYLE_REL_2012) {
        return Some("chartColorStyle");
    }
    if value == ACTIVE_X_BINARY_REL {
        return Some("activeXControlBinary");
    }
    value
        .strip_prefix(conformance.relationships())
        .and_then(|kind| kind.strip_prefix('/'))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(in crate::glossary) enum GraphRole {
    Glossary,
    RichStory,
    Settings,
    FontTable,
    Numbering,
    WebSettings,
    Chart,
    ChartDrawing,
    DiagramData,
    DiagramLayout,
    EmbeddedObject,
    EmbeddedPackage,
    Control,
    ActiveX,
    CustomXml,
    Customizations,
    Leaf,
}

#[derive(Clone, Copy)]
pub(in crate::glossary) enum EdgeMode {
    Internal,
    External,
    Either,
}

#[derive(Clone, Copy)]
pub(in crate::glossary) enum TargetProfile {
    Exact(&'static str),
    Image,
    Video,
    Xml,
    Font,
    Any,
}

#[derive(Clone, Copy)]
pub(in crate::glossary) struct EdgeSpec {
    pub(in crate::glossary) mode: EdgeMode,
    pub(in crate::glossary) target: TargetProfile,
    pub(in crate::glossary) role: Option<GraphRole>,
    pub(in crate::glossary) owned: bool,
}

pub(in crate::glossary) fn edge_spec(
    conformance: Conformance,
    role: GraphRole,
    value: &str,
) -> Result<EdgeSpec> {
    let kind = relationship_kind(conformance, value).ok_or_else(|| {
        invalid(format!(
            "unsupported glossary relationship type '{value}' for {conformance:?} conformance"
        ))
    })?;
    let internal = |target, role| EdgeSpec {
        mode: EdgeMode::Internal,
        target,
        role: Some(role),
        owned: true,
    };
    let either = |target, role| EdgeSpec {
        mode: EdgeMode::Either,
        target,
        role: Some(role),
        owned: true,
    };
    let reference = EdgeSpec {
        mode: EdgeMode::Either,
        target: TargetProfile::Any,
        role: None,
        owned: false,
    };
    let external = EdgeSpec {
        mode: EdgeMode::External,
        target: TargetProfile::Any,
        role: None,
        owned: false,
    };
    let spec = match (role, kind) {
        (GraphRole::Glossary, "comments") => {
            internal(TargetProfile::Exact(ct::WML_COMMENTS), GraphRole::RichStory)
        },
        (GraphRole::Glossary, "settings") => {
            internal(TargetProfile::Exact(ct::WML_SETTINGS), GraphRole::Settings)
        },
        (GraphRole::Glossary, "endnotes") => {
            internal(TargetProfile::Exact(ct::WML_ENDNOTES), GraphRole::RichStory)
        },
        (GraphRole::Glossary, "fontTable") => internal(
            TargetProfile::Exact(ct::WML_FONT_TABLE),
            GraphRole::FontTable,
        ),
        (GraphRole::Glossary, "footnotes") => internal(
            TargetProfile::Exact(ct::WML_FOOTNOTES),
            GraphRole::RichStory,
        ),
        (GraphRole::Glossary, "numbering") => internal(
            TargetProfile::Exact(ct::WML_NUMBERING),
            GraphRole::Numbering,
        ),
        (GraphRole::Glossary, "styles") => {
            internal(TargetProfile::Exact(ct::WML_STYLES), GraphRole::Leaf)
        },
        (GraphRole::Glossary, "webSettings") => internal(
            TargetProfile::Exact(ct::WML_WEB_SETTINGS),
            GraphRole::WebSettings,
        ),
        (GraphRole::Glossary, "aFChunk") => internal(TargetProfile::Any, GraphRole::Leaf),
        (GraphRole::Glossary, "chart") => {
            internal(TargetProfile::Exact(ct::DML_CHART), GraphRole::Chart)
        },
        (GraphRole::Glossary, "customXml") => internal(TargetProfile::Xml, GraphRole::CustomXml),
        (GraphRole::Glossary, "diagramColors") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_COLORS),
            GraphRole::Leaf,
        ),
        (GraphRole::Glossary, "diagramData") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_DATA),
            GraphRole::DiagramData,
        ),
        (GraphRole::Glossary, "diagramLayout") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_LAYOUT),
            GraphRole::DiagramLayout,
        ),
        (GraphRole::Glossary, "diagramQuickStyle") => {
            internal(TargetProfile::Exact(ct::DML_DIAGRAM_STYLE), GraphRole::Leaf)
        },
        (GraphRole::Glossary, "control") => internal(TargetProfile::Any, GraphRole::Control),
        (GraphRole::Glossary, "oleObject") => either(TargetProfile::Any, GraphRole::EmbeddedObject),
        (GraphRole::Glossary, "package") => either(TargetProfile::Any, GraphRole::EmbeddedPackage),
        (GraphRole::Glossary, "footer") => {
            internal(TargetProfile::Exact(ct::WML_FOOTER), GraphRole::RichStory)
        },
        (GraphRole::Glossary, "header") => {
            internal(TargetProfile::Exact(ct::WML_HEADER), GraphRole::RichStory)
        },
        (GraphRole::Glossary, "hyperlink") => reference,
        (GraphRole::Glossary, "image") => either(TargetProfile::Image, GraphRole::Leaf),
        (GraphRole::Glossary, "printerSettings") => {
            internal(TargetProfile::Exact(PRINTER_SETTINGS_CT), GraphRole::Leaf)
        },
        (GraphRole::Glossary, "video") => either(TargetProfile::Video, GraphRole::Leaf),
        (GraphRole::Glossary, "stylesWithEffects") => {
            internal(TargetProfile::Exact(STYLES_EFFECTS_CT), GraphRole::Leaf)
        },
        (GraphRole::Glossary, "keyMapCustomizations") => internal(
            TargetProfile::Exact(CUSTOMIZATIONS_CT),
            GraphRole::Customizations,
        ),
        (GraphRole::RichStory, "aFChunk") => internal(TargetProfile::Any, GraphRole::Leaf),
        (GraphRole::RichStory, "control") => internal(TargetProfile::Any, GraphRole::Control),
        (GraphRole::RichStory, "oleObject") => {
            either(TargetProfile::Any, GraphRole::EmbeddedObject)
        },
        (GraphRole::RichStory, "package") => either(TargetProfile::Any, GraphRole::EmbeddedPackage),
        (GraphRole::RichStory, "hyperlink") => reference,
        (GraphRole::RichStory, "image") => either(TargetProfile::Image, GraphRole::Leaf),
        (GraphRole::RichStory, "video") => either(TargetProfile::Video, GraphRole::Leaf),
        (GraphRole::RichStory, "chart") => {
            internal(TargetProfile::Exact(ct::DML_CHART), GraphRole::Chart)
        },
        (GraphRole::RichStory, "customXml") => internal(TargetProfile::Xml, GraphRole::CustomXml),
        (GraphRole::RichStory, "diagramColors") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_COLORS),
            GraphRole::Leaf,
        ),
        (GraphRole::RichStory, "diagramData") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_DATA),
            GraphRole::DiagramData,
        ),
        (GraphRole::RichStory, "diagramLayout") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_LAYOUT),
            GraphRole::DiagramLayout,
        ),
        (GraphRole::RichStory, "diagramQuickStyle") => {
            internal(TargetProfile::Exact(ct::DML_DIAGRAM_STYLE), GraphRole::Leaf)
        },
        (
            GraphRole::Settings,
            "attachedTemplate" | "mailMergeSource" | "mailMergeHeaderSource" | "transform",
        ) => external,
        (GraphRole::Settings, "recipientData") => {
            internal(TargetProfile::Exact(RECIPIENT_DATA_CT), GraphRole::Leaf)
        },
        (GraphRole::FontTable, "font") => internal(TargetProfile::Font, GraphRole::Leaf),
        (GraphRole::Numbering, "image") => either(TargetProfile::Image, GraphRole::Leaf),
        (GraphRole::WebSettings, "frame") => external,
        (GraphRole::Chart, "chartUserShapes") => internal(
            TargetProfile::Exact(ct::DML_CHARTSHAPES),
            GraphRole::ChartDrawing,
        ),
        (GraphRole::Chart, "chartStyle") => {
            internal(TargetProfile::Exact(CHART_STYLE_CT), GraphRole::Leaf)
        },
        (GraphRole::Chart, "chartColorStyle") => {
            internal(TargetProfile::Exact(CHART_COLOR_STYLE_CT), GraphRole::Leaf)
        },
        (GraphRole::Chart, "themeOverride") => internal(
            TargetProfile::Exact(ct::OFC_THEME_OVERRIDE),
            GraphRole::Leaf,
        ),
        (GraphRole::Chart, "package") => either(TargetProfile::Any, GraphRole::EmbeddedPackage),
        (GraphRole::ChartDrawing, "image") => either(TargetProfile::Image, GraphRole::Leaf),
        (GraphRole::ChartDrawing, "chart") => {
            internal(TargetProfile::Exact(ct::DML_CHART), GraphRole::Chart)
        },
        (GraphRole::ChartDrawing, "customXml") => {
            internal(TargetProfile::Xml, GraphRole::CustomXml)
        },
        (GraphRole::DiagramData | GraphRole::DiagramLayout, "image") => {
            either(TargetProfile::Image, GraphRole::Leaf)
        },
        (GraphRole::DiagramData, "diagramDrawing") => internal(
            TargetProfile::Exact(ct::DML_DIAGRAM_DRAWING),
            GraphRole::Leaf,
        ),
        (GraphRole::DiagramData, "hyperlink") => reference,
        (GraphRole::EmbeddedObject | GraphRole::EmbeddedPackage, "hyperlink") => reference,
        (GraphRole::ActiveX, "activeXControlBinary") => {
            internal(TargetProfile::Exact(ACTIVE_X_BINARY_CT), GraphRole::Leaf)
        },
        (GraphRole::CustomXml, "customXmlProps") => internal(
            TargetProfile::Exact(litchi_ooxml_common::custom_xml::PROPS_CONTENT_TYPE),
            GraphRole::Leaf,
        ),
        (GraphRole::Customizations, "attachedToolbars") => {
            internal(TargetProfile::Exact(ATTACHED_TOOLBARS_CT), GraphRole::Leaf)
        },
        _ => {
            return Err(invalid(format!(
                "relationship type '{value}' is invalid for glossary role {role:?}"
            )));
        },
    };
    Ok(spec)
}

pub(in crate::glossary) fn validate_edge_mode(
    kind: &str,
    mode: EdgeMode,
    external: bool,
) -> Result<()> {
    let valid = match mode {
        EdgeMode::Internal => !external,
        EdgeMode::External => external,
        EdgeMode::Either => true,
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(format!(
            "glossary relationship kind '{kind}' has an invalid target mode"
        )))
    }
}

pub(in crate::glossary) fn validate_target_profile(
    kind: &str,
    profile: TargetProfile,
    content_type: &str,
) -> Result<()> {
    ContentType::new(content_type.to_owned())?;
    let media_type = content_type.split(';').next().unwrap_or_default();
    let valid = match profile {
        TargetProfile::Exact(expected) => media_type.eq_ignore_ascii_case(expected),
        TargetProfile::Image => starts_with_ascii_case_insensitive(media_type, "image/"),
        TargetProfile::Video => starts_with_ascii_case_insensitive(media_type, "video/"),
        TargetProfile::Xml => {
            media_type.eq_ignore_ascii_case("application/xml")
                || media_type.eq_ignore_ascii_case("text/xml")
                || ends_with_ascii_case_insensitive(media_type, "+xml")
        },
        TargetProfile::Font => [FONT_DATA_CT, FONT_TTF_CT, OBFUSCATED_FONT_CT]
            .iter()
            .any(|expected| media_type.eq_ignore_ascii_case(expected)),
        TargetProfile::Any => true,
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(format!(
            "glossary relationship kind '{kind}' cannot target content type '{content_type}'"
        )))
    }
}

pub(in crate::glossary) fn target_graph_role(role: GraphRole, content_type: &str) -> GraphRole {
    if role == GraphRole::Control
        && content_type
            .split(';')
            .next()
            .is_some_and(|value| value.eq_ignore_ascii_case(ACTIVE_X_DESCRIPTOR_CT))
    {
        GraphRole::ActiveX
    } else {
        role
    }
}

pub(in crate::glossary) fn starts_with_ascii_case_insensitive(value: &str, prefix: &str) -> bool {
    value
        .get(..prefix.len())
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case(prefix))
}

pub(in crate::glossary) fn ends_with_ascii_case_insensitive(value: &str, suffix: &str) -> bool {
    value
        .get(value.len().saturating_sub(suffix.len())..)
        .is_some_and(|candidate| candidate.eq_ignore_ascii_case(suffix))
}

pub(in crate::glossary) fn validate_relationship_integrity(
    document: &Catalog,
    part: &dyn Part,
    conformance: Conformance,
) -> Result<()> {
    for id in catalog_relationship_references(document, conformance)? {
        if part.rels().get(id).is_none() {
            return Err(invalid(format!(
                "glossary XML references missing relationship '{id}'"
            )));
        }
    }
    for relationship in part.rels().iter() {
        relationship_kind(conformance, relationship.reltype()).ok_or_else(|| {
            invalid(format!(
                "unsupported glossary relationship type '{}' for {conformance:?} conformance",
                relationship.reltype()
            ))
        })?;
    }
    Ok(())
}

pub(in crate::glossary) fn catalog_relationship_references(
    document: &Catalog,
    conformance: Conformance,
) -> Result<HashSet<&str>> {
    let mut reference_count = document.background_refs.len();
    for entry in &document.entries {
        let refs = entry
            .producer
            .as_deref()
            .filter(|producer| producer.conformance == conformance)
            .map_or(entry.refs.as_ref(), |producer| producer.refs.as_ref());
        reference_count = reference_count
            .checked_add(refs.len())
            .ok_or_else(|| invalid("glossary relationship reference count overflow"))?;
    }
    let mut referenced = HashSet::new();
    referenced
        .try_reserve(reference_count)
        .map_err(|source| Error::Allocation {
            resource: "glossary relationship reference index",
            source,
        })?;
    referenced.extend(document.background_refs.iter().map(String::as_str));
    for entry in &document.entries {
        let refs = entry
            .producer
            .as_deref()
            .filter(|producer| producer.conformance == conformance)
            .map_or(entry.refs.as_ref(), |producer| producer.refs.as_ref());
        referenced.extend(refs.iter().map(String::as_str));
    }
    Ok(referenced)
}

pub(in crate::glossary) fn collect_relationship_references(
    node: &Node,
    output: &mut HashSet<String>,
) {
    for attribute in &node.attrs {
        if matches!(attribute.ns.as_ref(), R | RS) {
            output.insert(attribute.v.clone());
        }
    }
    for content in &node.content {
        if let Content::Node(child) = content {
            collect_relationship_references(child, output);
        }
    }
}

pub(in crate::glossary) fn relationship_references(node: &Node) -> Result<Arc<[String]>> {
    let mut refs = HashSet::new();
    collect_relationship_references(node, &mut refs);
    let mut refs = refs.into_iter().collect::<Vec<_>>();
    refs.sort();
    Ok(Arc::from(refs))
}

pub(in crate::glossary) fn glossary_owned_parts(
    package: &OpcPackage,
    root: &PackURI,
    conformance: Conformance,
) -> Result<HashSet<PackURI>> {
    if is_signature_part(root) || is_reserved_physical_part(root) {
        return Err(invalid(
            "glossary root cannot use reserved OPC package infrastructure",
        ));
    }
    let root = package.get_part(root)?.partname().clone();
    let mut owned = HashSet::from([root.clone()]);
    let mut roles = HashMap::from([(root.clone(), GraphRole::Glossary)]);
    let mut queue = VecDeque::from([(root, GraphRole::Glossary)]);
    let mut relationship_count = 0usize;
    let mut metadata_bytes = 0usize;
    while let Some((uri, role)) = queue.pop_front() {
        let part = package.get_part(&uri)?;
        validate_physical_part(uri.as_str(), part.content_type(), part.blob().len())?;
        add_graph_metadata(&mut metadata_bytes, uri.as_str())?;
        add_graph_metadata(&mut metadata_bytes, part.content_type())?;
        let part_relationship_count = part.rels().iter().count();
        if part_relationship_count > MAX_VALUES {
            return Err(invalid("glossary relationship limit exceeded"));
        }
        relationship_count = relationship_count
            .checked_add(part_relationship_count)
            .ok_or_else(|| invalid("glossary relationship count overflow"))?;
        if relationship_count > MAX_VALUES {
            return Err(invalid("glossary graph-wide relationship limit exceeded"));
        }
        for relationship in part.rels().iter() {
            validate_relationship_metadata(
                relationship.r_id(),
                relationship.reltype(),
                relationship.target_ref(),
            )?;
            add_graph_metadata(&mut metadata_bytes, relationship.r_id())?;
            add_graph_metadata(&mut metadata_bytes, relationship.reltype())?;
            add_graph_metadata(&mut metadata_bytes, relationship.target_ref())?;
        }
        for relationship in part.rels().iter() {
            let kind = relationship_kind(conformance, relationship.reltype()).ok_or_else(|| {
                invalid(format!(
                    "unsupported glossary relationship type '{}' for {conformance:?} conformance",
                    relationship.reltype()
                ))
            })?;
            let spec = edge_spec(conformance, role, relationship.reltype())?;
            validate_edge_mode(kind, spec.mode, relationship.is_external())?;
            if relationship.is_external() {
                continue;
            }
            let requested = relationship.target_partname()?;
            let target = package.get_part(&requested)?.partname().clone();
            if is_signature_part(&target) || is_reserved_physical_part(&target) {
                return Err(invalid(format!(
                    "glossary relationship '{}' targets reserved OPC package infrastructure",
                    relationship.r_id()
                )));
            }
            let target_part = package.get_part(&target)?;
            validate_target_profile(kind, spec.target, target_part.content_type())?;
            if !spec.owned {
                continue;
            }
            let target_role = spec
                .role
                .ok_or_else(|| invalid("owned glossary relationship is missing a target role"))?;
            let target_role = target_graph_role(target_role, target_part.content_type());
            if let Some(existing_role) = roles.get(&target) {
                if *existing_role != target_role {
                    return Err(invalid(format!(
                        "glossary-owned part '{target}' has conflicting graph roles"
                    )));
                }
                continue;
            }
            roles.insert(target.clone(), target_role);
            if owned.insert(target.clone()) {
                if owned.len() > MAX_VALUES + 1 {
                    return Err(invalid("glossary owned-part limit exceeded"));
                }
                queue.push_back((target, target_role));
            }
        }
    }
    Ok(owned)
}

pub(in crate::glossary) fn validate_exclusive_ownership(
    package: &OpcPackage,
    owner: &Owner,
    owned: &HashSet<PackURI>,
) -> Result<()> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        let target = package
            .get_part(&relationship.target_partname()?)?
            .partname();
        if owned.contains(target) {
            return Err(invalid(format!(
                "glossary-owned part '{target}' has an inbound package-root relationship"
            )));
        }
    }
    for source in package.iter_parts() {
        for relationship in source
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            let target = package
                .get_part(&relationship.target_partname()?)?
                .partname();
            if owned.contains(target)
                && !owned.contains(source.partname())
                && !(source.partname() == &owner.main
                    && relationship.r_id() == owner.relationship_id)
            {
                return Err(invalid(format!(
                    "glossary-owned part '{target}' is shared by '{}'",
                    source.partname()
                )));
            }
        }
    }
    Ok(())
}

pub(in crate::glossary) fn validate_all_internal_targets(package: &OpcPackage) -> Result<()> {
    for relationship in package
        .rels()
        .iter()
        .filter(|relationship| !relationship.is_external())
    {
        package.get_part(&relationship.target_partname()?)?;
    }
    for part in package.iter_parts() {
        for relationship in part
            .rels()
            .iter()
            .filter(|relationship| !relationship.is_external())
        {
            package.get_part(&relationship.target_partname()?)?;
        }
    }
    Ok(())
}
