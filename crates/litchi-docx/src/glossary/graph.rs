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
    Content, Node, bounded, invalid, plan_write_with_limits, read_with_limits, valid_ncname,
    validate_catalog_limits, validate_physical_part, write_plan_with_limits,
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

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(in crate::glossary) struct ClosureLedger {
    pub(in crate::glossary) total_bytes: usize,
    pub(in crate::glossary) metadata_bytes: usize,
    pub(in crate::glossary) relationship_count: usize,
    pub(in crate::glossary) part_count: usize,
}

pub mod raw {
    pub use super::super::limit::Limits;
    use super::super::model::Catalog as TypedCatalog;
    use super::super::{Error, OpcPackage, Result};
    use super::{Catalog, Conformance};
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    use quick_xml::Reader;
    use quick_xml::events::Event;
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
        /// Validate physical metadata under `limits` before retaining a payload.
        ///
        /// # Errors
        ///
        /// Returns an error if the operation cannot be completed.
        pub fn new(
            name: impl AsRef<str>,
            content_type: impl AsRef<str>,
            data: Vec<u8>,
            limits: &Limits,
        ) -> Result<Self> {
            admit_part(name.as_ref(), content_type.as_ref(), data.len(), limits)?;
            Ok(Self {
                name: name.as_ref().to_owned(),
                content_type: content_type.as_ref().to_owned(),
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

        /// Replace the opaque payload after checking its physical bounds and policy.
        ///
        /// # Errors
        ///
        /// Returns an error if the operation cannot be completed.
        pub fn replace_data(&mut self, data: Vec<u8>, limits: &Limits) -> Result<()> {
            admit_part(&self.name, &self.content_type, data.len(), limits)?;
            self.data = Arc::new(data);
            Ok(())
        }

        /// Replace the opaque relationship list after bounded metadata checks.
        ///
        /// # Errors
        ///
        /// Returns an error if the operation cannot be completed.
        pub fn set_relationships(
            &mut self,
            relationships: Vec<Rel>,
            limits: &Limits,
        ) -> Result<()> {
            admit_relationships(&relationships, limits)?;
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

        /// Borrow every opaque auxiliary part retained by this graph.
        #[must_use]
        pub fn parts(&self) -> &[Part] {
            &self.parts
        }
    }

    /// Load the complete glossary-owned OPC graph.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn load(package: &OpcPackage, limits: &Limits) -> Result<Option<Graph>> {
        super::load_graph_with_limits(package, limits)
    }

    /// Publish a complete graph without consuming the caller's recovery copy.
    ///
    /// An identical graph publishes nothing and keeps a signature; a different
    /// one is refused on a signed package with
    /// [`litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy`] (ADR 0006)
    /// until [`litchi_opc::OpcPackage::unsign`] removes it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn put(package: &mut OpcPackage, graph: &Graph, limits: &Limits) -> Result<bool> {
        super::put_graph_checked(package, graph, limits)
    }

    /// Remove and return the complete graph for graph-preserving transfer.
    ///
    /// Absence is a signature-preserving no-op; removing a graph from a signed
    /// package is refused with
    /// [`litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy`] (ADR 0006)
    /// until [`litchi_opc::OpcPackage::unsign`] removes the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn remove(package: &mut OpcPackage, limits: &Limits) -> Result<Option<Graph>> {
        super::remove_graph_with_limits(package, limits)
    }

    /// Parse one detached glossary XML payload under `limits`.
    pub fn read(xml: &[u8], limits: &Limits) -> Result<(TypedCatalog, Conformance)> {
        super::super::codec::read_with_limits(xml, limits)
    }

    /// Serialize one detached glossary catalog under `limits`.
    pub fn write(
        catalog: &TypedCatalog,
        conformance: Conformance,
        limits: &Limits,
    ) -> Result<Vec<u8>> {
        admit_catalog(catalog, limits)?;
        super::super::codec::write_with_limits(catalog, conformance, limits)
    }

    pub(in crate::glossary) fn admit_catalog(
        catalog: &TypedCatalog,
        limits: &Limits,
    ) -> Result<()> {
        if catalog.entries.len() > limits.values() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::Entry,
                actual: catalog.entries.len(),
                maximum: limits.values(),
            });
        }
        for entry in &catalog.entries {
            if let Some(body) = &entry.body {
                admit_xml(body, limits)?;
            }
        }
        if let Some(background) = &catalog.background {
            admit_xml(background, limits)?;
        }
        Ok(())
    }

    pub(in crate::glossary) fn admit_part(
        name: &str,
        content_type: &str,
        data_bytes: usize,
        limits: &Limits,
    ) -> Result<super::PackURI> {
        if name.len() > limits.name_bytes() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::StringBytes,
                actual: name.len(),
                maximum: limits.name_bytes(),
            });
        }
        if content_type.len() > limits.string_bytes() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::StringBytes,
                actual: content_type.len(),
                maximum: limits.string_bytes(),
            });
        }
        if data_bytes > limits.graph_bytes() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::GraphBytes,
                actual: data_bytes,
                maximum: limits.graph_bytes(),
            });
        }
        let metadata = name
            .len()
            .checked_add(content_type.len())
            .ok_or(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::GraphMetadataBytes,
                actual: usize::MAX,
                maximum: limits.metadata_bytes(),
            })?;
        if metadata > limits.metadata_bytes() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::GraphMetadataBytes,
                actual: metadata,
                maximum: limits.metadata_bytes(),
            });
        }
        let uri = super::validate_physical_part(name, content_type, data_bytes)?;
        if uri.as_str() == "/word/glossary/document.xml" {
            return Err(super::invalid(
                "glossary auxiliary part conflicts with the default root",
            ));
        }
        Ok(uri)
    }

    pub(in crate::glossary) fn admit_relationships(
        relationships: &[Rel],
        limits: &Limits,
    ) -> Result<()> {
        if relationships.len() > limits.values() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::Relationship,
                actual: relationships.len(),
                maximum: limits.values(),
            });
        }
        let mut metadata = 0usize;
        for relationship in relationships {
            for value in [
                relationship.id.as_str(),
                relationship.kind.as_str(),
                relationship.target.as_str(),
            ] {
                if value.len() > limits.string_bytes() {
                    return Err(Error::GlossaryLimit {
                        kind: crate::glossary::LimitKind::StringBytes,
                        actual: value.len(),
                        maximum: limits.string_bytes(),
                    });
                }
            }
            super::validate_relationship_metadata(
                &relationship.id,
                &relationship.kind,
                &relationship.target,
            )?;
            metadata = metadata
                .checked_add(relationship.id.len())
                .and_then(|value| value.checked_add(relationship.kind.len()))
                .and_then(|value| value.checked_add(relationship.target.len()))
                .ok_or(Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::GraphMetadataBytes,
                    actual: usize::MAX,
                    maximum: limits.metadata_bytes(),
                })?;
        }
        if metadata > limits.metadata_bytes() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::GraphMetadataBytes,
                actual: metadata,
                maximum: limits.metadata_bytes(),
            });
        }
        Ok(())
    }

    pub(in crate::glossary) fn admit_graph(graph: &Graph, limits: &Limits) -> Result<()> {
        super::validate_catalog_limits(&graph.catalog, limits)?;
        admit_catalog(&graph.catalog, limits)?;
        let generated_plan = if graph.root_xml().is_none() {
            Some(super::super::codec::plan_write_with_limits(
                &graph.catalog,
                graph.conformance,
                limits,
            )?)
        } else {
            None
        };
        let owner = graph
            .owner_main
            .as_deref()
            .zip(graph.owner_id.as_deref())
            .zip(graph.owner_target.as_deref())
            .map(|((main, id), target)| (main, id, target));
        super::admit_graph_ledger(
            graph,
            limits,
            graph.root_xml(),
            None,
            generated_plan.as_ref().map(|plan| plan.bytes),
            false,
            owner,
        )?;
        if let Some(plan) = generated_plan {
            let xml = super::super::codec::write_plan_with_limits(
                &graph.catalog,
                graph.conformance,
                &plan,
                limits,
            )?;
            admit_xml(&xml, limits)?;
        }
        Ok(())
    }

    pub(in crate::glossary) fn admit_graph_with_owner(
        graph: &Graph,
        limits: &Limits,
        root_xml: Option<&[u8]>,
        generated_root_xml: Option<&[u8]>,
        generated_root_bytes: Option<usize>,
        root_xml_admitted: bool,
        owner_main: &str,
        owner_id: &str,
        owner_target: &str,
    ) -> Result<()> {
        super::admit_graph_ledger(
            graph,
            limits,
            root_xml,
            generated_root_xml,
            generated_root_bytes,
            root_xml_admitted,
            Some((owner_main, owner_id, owner_target)),
        )?;
        Ok(())
    }

    pub(in crate::glossary) fn admit_xml(xml: &[u8], limits: &Limits) -> Result<()> {
        if xml.len() > limits.xml_bytes() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::XmlBytes,
                actual: xml.len(),
                maximum: limits.xml_bytes(),
            });
        }
        if xml.len() > limits.dom_bytes() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::DomBytes,
                actual: xml.len(),
                maximum: limits.dom_bytes(),
            });
        }
        let mut reader = Reader::from_reader(xml);
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut nodes = 0usize;
        let mut attributes = 0usize;
        let mut attribute_bytes = 0usize;
        let mut tokens = 0usize;
        let mut contents = 0usize;
        loop {
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(|error| Error::Xml(error.to_string()))?;
            tokens = tokens.saturating_add(1);
            if tokens > limits.tokens() {
                return Err(Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::XmlToken,
                    actual: tokens,
                    maximum: limits.tokens(),
                });
            }
            match event {
                Event::Start(element) => {
                    depth = depth.saturating_add(1);
                    nodes = nodes.saturating_add(1);
                    check_xml_shape(depth, nodes, attributes, contents, limits)?;
                    check_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
                },
                Event::Empty(element) => {
                    nodes = nodes.saturating_add(1);
                    check_xml_shape(depth.saturating_add(1), nodes, attributes, contents, limits)?;
                    check_attributes(&element, &mut attributes, &mut attribute_bytes, limits)?;
                },
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::Invalid("glossary XML has an unexpected closing element".to_owned())
                    })?
                },
                Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_) => {
                    contents = contents.saturating_add(1);
                    check_xml_shape(depth, nodes, attributes, contents, limits)?;
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(Error::Invalid(
                        "glossary XML contains a forbidden declaration".to_owned(),
                    ));
                },
                Event::Decl(_) => {},
                Event::Eof => break,
            }
            buffer.clear();
        }
        if depth != 0 {
            return Err(Error::Invalid("glossary XML is unterminated".to_owned()));
        }
        Ok(())
    }

    fn check_xml_shape(
        depth: usize,
        nodes: usize,
        attributes: usize,
        contents: usize,
        limits: &Limits,
    ) -> Result<()> {
        if depth > limits.depth() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::XmlDepth,
                actual: depth,
                maximum: limits.depth(),
            });
        }
        if nodes > limits.nodes() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::XmlNode,
                actual: nodes,
                maximum: limits.nodes(),
            });
        }
        if attributes > limits.attributes() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::XmlAttribute,
                actual: attributes,
                maximum: limits.attributes(),
            });
        }
        if contents > limits.content() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::XmlContentToken,
                actual: contents,
                maximum: limits.content(),
            });
        }
        Ok(())
    }

    fn check_attributes(
        element: &quick_xml::events::BytesStart<'_>,
        attributes: &mut usize,
        attribute_bytes: &mut usize,
        limits: &Limits,
    ) -> Result<()> {
        for attribute in element.unchecked_attributes() {
            let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
            *attributes = attributes.saturating_add(1);
            *attribute_bytes = attribute_bytes
                .checked_add(
                    attribute
                        .key
                        .as_ref()
                        .len()
                        .saturating_add(attribute.value.len()),
                )
                .ok_or(Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::XmlAggregateAttribute,
                    actual: usize::MAX,
                    maximum: limits.dom_bytes(),
                })?;
            if *attributes > limits.attributes() {
                return Err(Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::XmlAttribute,
                    actual: *attributes,
                    maximum: limits.attributes(),
                });
            }
            if *attribute_bytes > limits.dom_bytes() {
                return Err(Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::XmlAggregateAttribute,
                    actual: *attribute_bytes,
                    maximum: limits.dom_bytes(),
                });
            }
        }
        Ok(())
    }
}

/// Admit the exact graph closure that a raw graph publication will own.
///
/// The source loader and the publisher must charge the same ledger: the
/// owner identity seed, the glossary root's own name/content type and bytes,
/// followed by every auxiliary part and relationship.  Keeping this pass
/// independent of package staging lets callers reject a selected policy
/// before cloning or retaining any destination parts.
fn admit_graph_ledger(
    graph: &raw::Graph,
    limits: &raw::Limits,
    root_xml: Option<&[u8]>,
    generated_root_xml: Option<&[u8]>,
    generated_root_bytes: Option<usize>,
    root_xml_admitted: bool,
    owner: Option<(&str, &str, &str)>,
) -> Result<ClosureLedger> {
    let part_limit = limits.parts().min(MAX_VALUES);
    if graph.parts.len() > part_limit {
        return Err(Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::AuxiliaryPart,
            actual: graph.parts.len(),
            maximum: part_limit,
        });
    }
    let root_xml_bytes = match (root_xml, generated_root_xml, generated_root_bytes) {
        (Some(xml), None, None) => {
            if !root_xml_admitted {
                raw::admit_xml(xml, limits)?;
            }
            xml.len()
        },
        (None, Some(xml), None) => {
            if !root_xml_admitted {
                raw::admit_xml(xml, limits)?;
            }
            xml.len()
        },
        (None, None, Some(bytes)) => {
            if bytes > limits.xml_bytes() {
                return Err(Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::XmlBytes,
                    actual: bytes,
                    maximum: limits.xml_bytes(),
                });
            }
            if bytes > limits.dom_bytes() {
                return Err(Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::DomBytes,
                    actual: bytes,
                    maximum: limits.dom_bytes(),
                });
            }
            bytes
        },
        (None, None, None) => 0,
        _ => {
            return Err(invalid(
                "glossary graph supplied both source and generated XML",
            ));
        },
    };

    let mut ledger = ClosureLedger::default();
    admit_graph_name(&graph.root_name, limits)?;
    let root_uri = PackURI::new(&graph.root_name).map_err(Error::Uri)?;
    if is_signature_part(&root_uri) || is_reserved_physical_part(&root_uri) {
        return Err(invalid(
            "glossary root cannot use reserved OPC package infrastructure",
        ));
    }
    if let Some((owner_main, owner_id, owner_target)) = owner {
        validate_owner_identity(owner_main, owner_id, owner_target, limits)?;
        ledger.metadata_bytes = owner_metadata_seed_values(
            owner_main,
            &graph.root_name,
            owner_id,
            owner_target,
            limits,
        )?;
    } else {
        add_selected_metadata(&mut ledger.metadata_bytes, &graph.root_name, limits)?;
        if let Some(owner_main) = graph.owner_main.as_deref() {
            if owner_main.len() > limits.name_bytes() {
                return Err(Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::StringBytes,
                    actual: owner_main.len(),
                    maximum: limits.name_bytes(),
                });
            }
            bounded(owner_main)?;
            PackURI::new(owner_main).map_err(Error::Uri)?;
            add_selected_metadata(&mut ledger.metadata_bytes, owner_main, limits)?;
        }
        if let Some(owner_id) = graph.owner_id.as_deref() {
            admit_graph_string(owner_id, limits)?;
            if !valid_ncname(owner_id) {
                return Err(invalid(
                    "glossary owner relationship ID is not an XML NCName",
                ));
            }
            add_selected_metadata(&mut ledger.metadata_bytes, owner_id, limits)?;
        }
        if let Some(owner_target) = graph.owner_target.as_deref() {
            admit_graph_string(owner_target, limits)?;
            if owner_target.is_empty() || owner_target.chars().any(char::is_control) {
                return Err(invalid("glossary owner relationship target is invalid"));
            }
            add_selected_metadata(&mut ledger.metadata_bytes, owner_target, limits)?;
        }
    }

    admit_graph_name(&graph.root_name, limits)?;
    add_selected_metadata(&mut ledger.metadata_bytes, &graph.root_name, limits)?;
    admit_graph_string(CT, limits)?;
    add_selected_metadata(&mut ledger.metadata_bytes, CT, limits)?;
    add_graph_bytes(&mut ledger.total_bytes, root_xml_bytes, limits)?;
    admit_graph_relationships(
        &graph.rels,
        graph.conformance,
        &mut ledger.relationship_count,
        &mut ledger.metadata_bytes,
        limits,
    )?;

    preflight_graph_parts(graph, &mut ledger, limits)?;

    let mut part_names = HashSet::new();
    part_names
        .try_reserve(graph.parts.len())
        .map_err(|source| Error::Allocation {
            resource: "glossary graph part names",
            source,
        })?;
    for part in &graph.parts {
        let part_uri = raw::admit_part(&part.name, &part.content_type, part.data().len(), limits)?;
        if part_uri == root_uri {
            return Err(invalid(format!(
                "glossary auxiliary part '{}' conflicts with the root",
                part_uri.as_str()
            )));
        }
        if !part_names.insert(part_uri.clone()) {
            return Err(invalid(format!(
                "duplicate glossary auxiliary part '{part_uri}'"
            )));
        }
        if is_xml_content_type(&part.content_type) {
            raw::admit_xml(part.data(), limits)?;
        }
    }
    ledger.part_count = graph.parts.len();
    Ok(ledger)
}

fn preflight_graph_parts(
    graph: &raw::Graph,
    ledger: &mut ClosureLedger,
    limits: &raw::Limits,
) -> Result<()> {
    for part in &graph.parts {
        admit_graph_name(&part.name, limits)?;
        admit_graph_string(&part.content_type, limits)?;
        add_selected_metadata(&mut ledger.metadata_bytes, &part.name, limits)?;
        add_selected_metadata(&mut ledger.metadata_bytes, &part.content_type, limits)?;
        add_graph_bytes(&mut ledger.total_bytes, part.data().len(), limits)?;
        admit_graph_relationships(
            &part.rels,
            graph.conformance,
            &mut ledger.relationship_count,
            &mut ledger.metadata_bytes,
            limits,
        )?;
    }
    Ok(())
}

fn admit_graph_name(value: &str, limits: &raw::Limits) -> Result<()> {
    if value.len() > limits.name_bytes() {
        return Err(Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::StringBytes,
            actual: value.len(),
            maximum: limits.name_bytes(),
        });
    }
    bounded(value)
}

fn is_xml_content_type(content_type: &str) -> bool {
    content_type.ends_with("+xml")
        || content_type.eq_ignore_ascii_case("application/xml")
        || content_type.eq_ignore_ascii_case("text/xml")
}

fn admit_graph_string(value: &str, limits: &raw::Limits) -> Result<()> {
    if value.len() > limits.string_bytes() {
        return Err(Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::StringBytes,
            actual: value.len(),
            maximum: limits.string_bytes(),
        });
    }
    bounded(value)
}

fn validate_owner_identity(
    owner_main: &str,
    owner_id: &str,
    owner_target: &str,
    limits: &raw::Limits,
) -> Result<()> {
    admit_graph_name(owner_main, limits)?;
    PackURI::new(owner_main).map_err(Error::Uri)?;
    admit_graph_string(owner_id, limits)?;
    if !valid_ncname(owner_id) {
        return Err(invalid(
            "glossary owner relationship ID is not an XML NCName",
        ));
    }
    admit_graph_string(owner_target, limits)?;
    if owner_target.is_empty() || owner_target.chars().any(char::is_control) {
        return Err(invalid("glossary owner relationship target is invalid"));
    }
    Ok(())
}

fn add_graph_bytes(total: &mut usize, bytes: usize, limits: &raw::Limits) -> Result<()> {
    *total = total.checked_add(bytes).ok_or(Error::GlossaryLimit {
        kind: crate::glossary::LimitKind::GraphBytes,
        actual: usize::MAX,
        maximum: limits.graph_bytes().min(MAX_GRAPH_BYTES),
    })?;
    if *total > limits.graph_bytes() || *total > MAX_GRAPH_BYTES {
        return Err(Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphBytes,
            actual: *total,
            maximum: limits.graph_bytes().min(MAX_GRAPH_BYTES),
        });
    }
    Ok(())
}

fn admit_graph_relationships(
    relationships: &[raw::Rel],
    conformance: Conformance,
    total_count: &mut usize,
    metadata_bytes: &mut usize,
    limits: &raw::Limits,
) -> Result<()> {
    let relationship_limit = limits.values().min(MAX_VALUES);
    if relationships.len() > relationship_limit {
        return Err(Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::Relationship,
            actual: relationships.len(),
            maximum: relationship_limit,
        });
    }
    *total_count = total_count
        .checked_add(relationships.len())
        .ok_or(Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphWideRelationship,
            actual: usize::MAX,
            maximum: relationship_limit,
        })?;
    if *total_count > relationship_limit {
        return Err(Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphWideRelationship,
            actual: *total_count,
            maximum: relationship_limit,
        });
    }
    for relationship in relationships {
        for value in [
            relationship.id.as_str(),
            relationship.kind.as_str(),
            relationship.target.as_str(),
        ] {
            admit_graph_string(value, limits)?;
            add_selected_metadata(metadata_bytes, value, limits)?;
        }
        validate_relationship_metadata(&relationship.id, &relationship.kind, &relationship.target)?;
        relationship_kind(conformance, &relationship.kind).ok_or_else(|| {
            invalid(format!(
                "unsupported glossary relationship type '{}' for {conformance:?} conformance",
                relationship.kind
            ))
        })?;
    }
    Ok(())
}

pub(in crate::glossary) fn load_graph_with_limits(
    package: &OpcPackage,
    limits: &raw::Limits,
) -> Result<Option<raw::Graph>> {
    let owner = locate(package, limits)?;
    load_graph_for_owner(package, limits, owner)
}

fn load_graph_for_owner(
    package: &OpcPackage,
    limits: &raw::Limits,
    owner: Option<Owner>,
) -> Result<Option<raw::Graph>> {
    let Some(owner) = owner else {
        return Ok(None);
    };
    let metadata = owner_metadata_seed(&owner, limits)?;
    let (owned, _) = glossary_owned_parts(
        package,
        &owner.root,
        owner.conformance,
        limits,
        metadata,
        false,
    )?;
    validate_exclusive_ownership(package, &owner, &owned)?;
    let root = package.get_part(&owner.root)?;
    let (catalog, conformance) = read_with_limits(root.blob(), limits)?;
    if conformance != owner.conformance {
        return Err(invalid(
            "glossary relationship and XML use different conformance families",
        ));
    }
    validate_relationship_integrity(&catalog, root, conformance)?;
    let rels = copy_relationships(root)?;
    let auxiliary_count = owned.len().saturating_sub(1);
    let part_limit = limits.parts().min(MAX_VALUES);
    if auxiliary_count > part_limit {
        return Err(Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::AuxiliaryPart,
            actual: auxiliary_count,
            maximum: part_limit,
        });
    }
    let mut names = Vec::new();
    names
        .try_reserve_exact(auxiliary_count)
        .map_err(|source| Error::Allocation {
            resource: "glossary auxiliary part names",
            source,
        })?;
    for uri in owned {
        if uri != owner.root {
            names.push(uri);
        }
    }
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
        if total_bytes > limits.graph_bytes() {
            return Err(Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::GraphBytes,
                actual: total_bytes,
                maximum: limits.graph_bytes(),
            });
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

struct OwnerPlan {
    id: String,
    insert: bool,
}

pub(in crate::glossary) struct PreparedGraph {
    root_uri: PackURI,
    main: PackURI,
    owner_target: String,
    owner_plan: OwnerPlan,
    preserve_root: bool,
    generated_xml: Option<Vec<u8>>,
    effective_catalog: Catalog,
}

fn plan_owner_relationship(
    main: &dyn Part,
    current_owner_id: Option<&str>,
    requested_id: Option<&str>,
    relationship_type: &str,
    target: &str,
) -> Result<OwnerPlan> {
    let occupied = |candidate: &str| {
        main.rels().iter().any(|relationship| {
            Some(relationship.r_id()) != current_owner_id && relationship.r_id() == candidate
        })
    };
    if let Some(requested_id) = requested_id.filter(|candidate| !occupied(candidate)) {
        return Ok(OwnerPlan {
            id: requested_id.to_owned(),
            insert: true,
        });
    }
    if let Some(existing) = main
        .rels()
        .iter()
        .filter(|relationship| {
            Some(relationship.r_id()) != current_owner_id
                && !relationship.is_external()
                && relationship.reltype() == relationship_type
                && relationship.target_ref() == target
        })
        .map(|relationship| relationship.r_id())
        .min()
    {
        return Ok(OwnerPlan {
            id: existing.to_owned(),
            insert: false,
        });
    }
    let mut next = 1u32;
    loop {
        let candidate = format!("rId{next}");
        if !occupied(&candidate) {
            return Ok(OwnerPlan {
                id: candidate,
                insert: true,
            });
        }
        next = next
            .checked_add(1)
            .ok_or_else(|| invalid("no glossary relationship ID is available"))?;
    }
}

pub(in crate::glossary) fn put_graph_checked(
    package: &mut OpcPackage,
    value: &raw::Graph,
    limits: &raw::Limits,
) -> Result<bool> {
    let prepared = prepare_graph_for_put(package, value, limits)?;
    put_prepared_graph(package, value, limits, prepared)
}

pub(in crate::glossary) fn prepare_graph_for_put(
    package: &OpcPackage,
    value: &raw::Graph,
    limits: &raw::Limits,
) -> Result<PreparedGraph> {
    validate_package_conformance(package, value.conformance)?;
    admit_graph_name(&value.root_name, limits)?;
    if let Some(owner_main) = value.owner_main.as_deref() {
        admit_graph_name(owner_main, limits)?;
    }
    if let Some(owner_id) = value.owner_id.as_deref() {
        admit_graph_string(owner_id, limits)?;
    }
    if let Some(owner_target) = value.owner_target.as_deref() {
        admit_graph_string(owner_target, limits)?;
    }
    validate_raw_graph_metadata(value)?;
    validate_catalog_limits(&value.catalog, limits)?;
    raw::admit_catalog(&value.catalog, limits)?;
    let root_uri = PackURI::new(&value.root_name).map_err(Error::Uri)?;
    if is_signature_part(&root_uri) || is_reserved_physical_part(&root_uri) {
        return Err(invalid(
            "glossary root cannot use reserved OPC package infrastructure",
        ));
    }
    let main = {
        let main_part = package.main_document_part()?;
        admit_graph_name(main_part.partname().as_str(), limits)?;
        main_part.partname().clone()
    };
    let generated_target = root_uri.relative_ref(main.base_uri());
    let owner_target = if value.owner_main.as_deref() == Some(main.as_str()) {
        value
            .owner_target
            .clone()
            .unwrap_or_else(|| generated_target.clone())
    } else {
        generated_target
    };
    admit_graph_string(&owner_target, limits)?;
    if owner_target.is_empty() || owner_target.chars().any(char::is_control) {
        return Err(invalid("glossary owner relationship target is invalid"));
    }
    if let Some(owner_id) = value.owner_id.as_deref() {
        admit_graph_string(owner_id, limits)?;
        if !valid_ncname(owner_id) {
            return Err(invalid(
                "glossary owner relationship ID is not an XML NCName",
            ));
        }
    }
    let owner_plan = {
        let main_part = package.main_document_part()?;
        let current_owner_id = main_part
            .rels()
            .iter()
            .find(|relationship| {
                !relationship.is_external()
                    && relationship.reltype() == value.conformance.glossary_relationship()
            })
            .map(|relationship| relationship.r_id());
        plan_owner_relationship(
            main_part,
            current_owner_id,
            value.owner_id.as_deref(),
            value.conformance.glossary_relationship(),
            &owner_target,
        )?
    };
    let preserved_catalog = if let Some(xml) = &value.root_xml {
        let (source, source_conformance) = read_with_limits(xml.as_slice(), limits)?;
        (source_conformance == value.conformance && source == value.catalog).then_some(source)
    } else {
        None
    };
    let preserve_root = preserved_catalog.is_some();
    let generated_plan = if preserve_root {
        None
    } else {
        Some(plan_write_with_limits(
            &value.catalog,
            value.conformance,
            limits,
        )?)
    };
    raw::admit_graph_with_owner(
        value,
        limits,
        preserve_root.then(|| {
            value
                .root_xml
                .as_ref()
                .expect("preserved glossary XML is present")
                .as_slice()
        }),
        None,
        generated_plan.as_ref().map(|plan| plan.bytes),
        true,
        main.as_str(),
        &owner_plan.id,
        &owner_target,
    )?;
    let generated_xml = if let Some(plan) = generated_plan {
        let xml = write_plan_with_limits(&value.catalog, value.conformance, &plan, limits)?;
        raw::admit_xml(&xml, limits)?;
        Some(xml)
    } else {
        None
    };
    let effective_catalog = if preserve_root {
        preserved_catalog.ok_or_else(|| invalid("preserved glossary catalog is missing"))?
    } else {
        let xml = generated_xml
            .as_deref()
            .ok_or_else(|| invalid("generated glossary XML is missing"))?;
        let (canonical, canonical_conformance) =
            super::codec::read_admitted_with_limits(xml, limits)?;
        if canonical_conformance != value.conformance {
            return Err(invalid("canonical glossary changed conformance"));
        }
        canonical
    };
    Ok(PreparedGraph {
        root_uri,
        main,
        owner_target,
        owner_plan,
        preserve_root,
        generated_xml,
        effective_catalog,
    })
}

pub(in crate::glossary) fn put_prepared_graph(
    package: &mut OpcPackage,
    value: &raw::Graph,
    limits: &raw::Limits,
    prepared: PreparedGraph,
) -> Result<bool> {
    let PreparedGraph {
        root_uri,
        main,
        owner_target,
        owner_plan,
        preserve_root,
        generated_xml,
        effective_catalog,
    } = prepared;
    let current_owner = locate(package, limits)?;
    let current = load_graph_for_owner(package, limits, current_owner)?;
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
        let xml = generated_xml.ok_or_else(|| invalid("generated glossary XML is missing"))?;
        BlobPart::new(root_uri.clone(), CT.to_owned(), xml)
    };
    add_relationships(&mut root, &value.rels, value.conformance)?;
    validate_relationship_integrity(&effective_catalog, &root, value.conformance)?;
    let staged_root_xml = root.blob_arc();
    if value.parts.len() > MAX_VALUES {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::AuxiliaryPart,
            actual: value.parts.len(),
            maximum: MAX_VALUES,
        });
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

    if let Some(current) = current.as_ref()
        && graph_matches_catalog(current, value, &effective_catalog, Some(&staged_root_xml))?
    {
        return Ok(false);
    }

    validate_all_internal_targets(package)?;
    let old_owned = if let Some(current) = current.as_ref() {
        let mut owned = HashSet::new();
        owned
            .try_reserve(current.parts.len().saturating_add(1))
            .map_err(|source| Error::Allocation {
                resource: "glossary previous ownership closure",
                source,
            })?;
        owned.insert(PackURI::new(&current.root_name).map_err(Error::Uri)?);
        for part in &current.parts {
            owned.insert(PackURI::new(&part.name).map_err(Error::Uri)?);
        }
        owned
    } else {
        HashSet::new()
    };

    crate::signature_policy::ensure_unsigned(package)?;
    let mut candidate = package.clone();
    if let Some(owner_id) = current.as_ref().and_then(|graph| graph.owner_id.as_deref()) {
        candidate.get_part_mut(&main)?.rels_mut().remove(owner_id);
    }
    for uri in &old_owned {
        candidate.remove_part(uri);
    }
    for (_, part) in staged {
        candidate.try_add_part(Box::new(part))?;
    }
    candidate.try_add_part(Box::new(root))?;
    if owner_plan.insert {
        candidate
            .get_part_mut(&main)?
            .rels_mut()
            .try_add_relationship(
                value.conformance.glossary_relationship().to_owned(),
                owner_target.clone(),
                owner_plan.id.clone(),
                litchi_opc::TargetMode::Internal,
            )?;
    }

    validate_all_internal_targets(&candidate)?;
    let round_trip = load_graph_with_limits(&candidate, limits)?
        .ok_or_else(|| invalid("staged glossary graph is missing"))?;
    if !graph_matches_catalog_with_owner(
        &round_trip,
        value,
        &effective_catalog,
        Some(&staged_root_xml),
        main.as_str(),
        &owner_plan.id,
        &owner_target,
    )? {
        return Err(invalid("staged glossary graph did not round-trip"));
    }
    *package = candidate;
    Ok(true)
}

pub(in crate::glossary) fn remove_graph_with_limits(
    package: &mut OpcPackage,
    limits: &raw::Limits,
) -> Result<Option<raw::Graph>> {
    let Some(graph) = load_graph_with_limits(package, limits)? else {
        return Ok(None);
    };
    let owner_main = PackURI::new(
        graph
            .owner_main
            .as_deref()
            .ok_or_else(|| invalid("glossary owner main part is missing"))?,
    )
    .map_err(Error::Uri)?;
    let owner_id = graph
        .owner_id
        .as_deref()
        .ok_or_else(|| invalid("glossary owner relationship ID is missing"))?;
    let mut owned = HashSet::new();
    owned
        .try_reserve(graph.parts.len().saturating_add(1))
        .map_err(|source| Error::Allocation {
            resource: "glossary ownership closure",
            source,
        })?;
    owned.insert(PackURI::new(&graph.root_name).map_err(Error::Uri)?);
    for part in &graph.parts {
        owned.insert(PackURI::new(&part.name).map_err(Error::Uri)?);
    }
    validate_all_internal_targets(package)?;

    crate::signature_policy::ensure_unsigned(package)?;
    let mut candidate = package.clone();
    candidate
        .get_part_mut(&owner_main)?
        .rels_mut()
        .remove(owner_id);
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
    graph_matches_catalog_inner(actual, expected, expected_catalog, expected_root_xml, None)
}

fn graph_matches_catalog_with_owner(
    actual: &raw::Graph,
    expected: &raw::Graph,
    expected_catalog: &Catalog,
    expected_root_xml: Option<&Arc<Vec<u8>>>,
    owner_main: &str,
    owner_id: &str,
    owner_target: &str,
) -> Result<bool> {
    graph_matches_catalog_inner(
        actual,
        expected,
        expected_catalog,
        expected_root_xml,
        Some((owner_main, owner_id, owner_target)),
    )
}

fn graph_matches_catalog_inner(
    actual: &raw::Graph,
    expected: &raw::Graph,
    expected_catalog: &Catalog,
    expected_root_xml: Option<&Arc<Vec<u8>>>,
    owner_identity: Option<(&str, &str, &str)>,
) -> Result<bool> {
    let owner_matches = match owner_identity {
        Some((owner_main, owner_id, owner_target)) => {
            actual.owner_main.as_deref() == Some(owner_main)
                && actual.owner_id.as_deref() == Some(owner_id)
                && actual.owner_target.as_deref() == Some(owner_target)
        },
        None => {
            expected
                .owner_main
                .as_ref()
                .is_none_or(|value| actual.owner_main.as_ref() == Some(value))
                && expected
                    .owner_id
                    .as_ref()
                    .is_none_or(|value| actual.owner_id.as_ref() == Some(value))
                && expected
                    .owner_target
                    .as_ref()
                    .is_none_or(|value| actual.owner_target.as_ref() == Some(value))
        },
    };
    Ok(actual.catalog == *expected_catalog
        && actual.conformance == expected.conformance
        && actual.root_name == expected.root_name
        && owner_matches
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

pub(in crate::glossary) fn locate(
    package: &OpcPackage,
    limits: &raw::Limits,
) -> Result<Option<Owner>> {
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
    let main = main_part.partname();
    if main.as_str().len() > limits.name_bytes() {
        return Err(Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::StringBytes,
            actual: main.as_str().len(),
            maximum: limits.name_bytes(),
        });
    }
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
            if relationship.r_id().len() > limits.string_bytes()
                || relationship.target_ref().len() > limits.string_bytes()
            {
                return Err(Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::StringBytes,
                    actual: relationship
                        .r_id()
                        .len()
                        .max(relationship.target_ref().len()),
                    maximum: limits.string_bytes(),
                });
            }
            if part.partname() != main {
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
            let target_part = package.get_part(&requested)?;
            let target = target_part.partname();
            owner_metadata_seed_values(
                main.as_str(),
                target.as_str(),
                relationship.r_id(),
                relationship.target_ref(),
                limits,
            )?;
            if target_part.content_type() != CT {
                return Err(Error::ContentType {
                    expected: CT.to_owned(),
                    actual: target_part.content_type().to_owned(),
                });
            }
            found = Some(Owner {
                main: main.clone(),
                root: target.clone(),
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
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::Relationship,
            actual: relationships.len(),
            maximum: MAX_VALUES,
        });
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
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::AuxiliaryPart,
            actual: graph.parts.len(),
            maximum: MAX_VALUES,
        });
    }
    let mut relationship_count = 0usize;
    let mut metadata_bytes = 0usize;
    add_graph_metadata(&mut metadata_bytes, &graph.root_name)?;
    add_graph_metadata(&mut metadata_bytes, CT)?;
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
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::Relationship,
            actual: relationships.len(),
            maximum: MAX_VALUES,
        });
    }
    *total_count = total_count
        .checked_add(relationships.len())
        .ok_or_else(|| invalid("glossary relationship count overflow"))?;
    if *total_count > MAX_VALUES {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphWideRelationship,
            actual: *total_count,
            maximum: MAX_VALUES,
        });
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

/// Add metadata while enforcing the caller-selected graph policy as well as
/// the format-wide ceiling.  The source closure uses this before retaining
/// any newly discovered part or relationship.
pub(in crate::glossary) fn add_selected_metadata(
    total: &mut usize,
    value: &str,
    limits: &raw::Limits,
) -> Result<()> {
    add_graph_metadata(total, value)?;
    if *total > limits.metadata_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphMetadataBytes,
            actual: *total,
            maximum: limits.metadata_bytes(),
        });
    }
    Ok(())
}

pub(in crate::glossary) fn owner_metadata_seed(
    owner: &Owner,
    limits: &raw::Limits,
) -> Result<usize> {
    owner_metadata_seed_values(
        owner.main.as_str(),
        owner.root.as_str(),
        &owner.relationship_id,
        &owner.relationship_target,
        limits,
    )
}

pub(in crate::glossary) fn owner_metadata_seed_values(
    main: &str,
    root: &str,
    relationship_id: &str,
    relationship_target: &str,
    limits: &raw::Limits,
) -> Result<usize> {
    let mut metadata = 0usize;
    for value in [main, root] {
        if value.len() > limits.name_bytes() {
            return Err(crate::Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::StringBytes,
                actual: value.len(),
                maximum: limits.name_bytes(),
            });
        }
        add_selected_metadata(&mut metadata, value, limits)?;
    }
    for value in [relationship_id, relationship_target] {
        if value.len() > limits.string_bytes() {
            return Err(crate::Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::StringBytes,
                actual: value.len(),
                maximum: limits.string_bytes(),
            });
        }
        add_selected_metadata(&mut metadata, value, limits)?;
    }
    Ok(metadata)
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

fn admit_owned_part(
    part: &dyn Part,
    limits: &raw::Limits,
    total_bytes: &mut usize,
    metadata_bytes: &mut usize,
    relationship_count: &mut usize,
    admit_payload: bool,
) -> Result<()> {
    let name = part.partname().as_str();
    let content_type = part.content_type();
    if name.len() > limits.name_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::StringBytes,
            actual: name.len(),
            maximum: limits.name_bytes(),
        });
    }
    if content_type.len() > limits.string_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::StringBytes,
            actual: content_type.len(),
            maximum: limits.string_bytes(),
        });
    }
    validate_physical_part(name, content_type, part.blob().len())?;
    add_selected_metadata(metadata_bytes, name, limits)?;
    add_selected_metadata(metadata_bytes, content_type, limits)?;
    *total_bytes =
        total_bytes
            .checked_add(part.blob().len())
            .ok_or(crate::Error::GlossaryLimit {
                kind: crate::glossary::LimitKind::GraphBytes,
                actual: usize::MAX,
                maximum: limits.graph_bytes().min(MAX_GRAPH_BYTES),
            })?;
    if *total_bytes > limits.graph_bytes() || *total_bytes > MAX_GRAPH_BYTES {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphBytes,
            actual: *total_bytes,
            maximum: limits.graph_bytes().min(MAX_GRAPH_BYTES),
        });
    }
    if admit_payload
        && (content_type == CT
            || content_type.ends_with("+xml")
            || content_type.eq_ignore_ascii_case("application/xml")
            || content_type.eq_ignore_ascii_case("text/xml"))
    {
        raw::admit_xml(part.blob(), limits)?;
    }
    let part_relationship_count = part.rels().iter().count();
    let relationship_limit = limits.values().min(MAX_VALUES);
    if part_relationship_count > relationship_limit {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::Relationship,
            actual: part_relationship_count,
            maximum: relationship_limit,
        });
    }
    *relationship_count = relationship_count
        .checked_add(part_relationship_count)
        .ok_or(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphWideRelationship,
            actual: usize::MAX,
            maximum: relationship_limit,
        })?;
    if *relationship_count > relationship_limit {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphWideRelationship,
            actual: *relationship_count,
            maximum: relationship_limit,
        });
    }
    for relationship in part.rels().iter() {
        for value in [
            relationship.r_id(),
            relationship.reltype(),
            relationship.target_ref(),
        ] {
            if value.len() > limits.string_bytes() {
                return Err(crate::Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::StringBytes,
                    actual: value.len(),
                    maximum: limits.string_bytes(),
                });
            }
            add_selected_metadata(metadata_bytes, value, limits)?;
        }
        validate_relationship_metadata(
            relationship.r_id(),
            relationship.reltype(),
            relationship.target_ref(),
        )?;
    }
    Ok(())
}

pub(in crate::glossary) fn glossary_owned_parts(
    package: &OpcPackage,
    root: &PackURI,
    conformance: Conformance,
    limits: &raw::Limits,
    metadata_seed: usize,
    admit_root_payload: bool,
) -> Result<(HashSet<PackURI>, ClosureLedger)> {
    if is_signature_part(root) || is_reserved_physical_part(root) {
        return Err(invalid(
            "glossary root cannot use reserved OPC package infrastructure",
        ));
    }
    if root.as_str().len() > limits.name_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::StringBytes,
            actual: root.as_str().len(),
            maximum: limits.name_bytes(),
        });
    }
    let root = package.get_part(root)?.partname().clone();
    let max_owned = limits
        .parts()
        .saturating_add(1)
        .min(MAX_VALUES.saturating_add(1));
    let mut owned = HashSet::new();
    let mut roles = HashMap::new();
    let mut queue = VecDeque::new();
    let mut relationship_count = 0usize;
    let mut total_bytes = 0usize;
    let mut metadata_bytes = metadata_seed;
    if metadata_bytes > limits.metadata_bytes() {
        return Err(crate::Error::GlossaryLimit {
            kind: crate::glossary::LimitKind::GraphMetadataBytes,
            actual: metadata_bytes,
            maximum: limits.metadata_bytes(),
        });
    }
    let root_part = package.get_part(&root)?;
    admit_owned_part(
        root_part,
        limits,
        &mut total_bytes,
        &mut metadata_bytes,
        &mut relationship_count,
        admit_root_payload,
    )?;
    owned.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "glossary owned-part closure",
        source,
    })?;
    roles.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "glossary owned-part roles",
        source,
    })?;
    queue.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "glossary owned-part queue",
        source,
    })?;
    owned.insert(root.clone());
    roles.insert(root.clone(), GraphRole::Glossary);
    queue.push_back((root, GraphRole::Glossary));
    while let Some((uri, role)) = queue.pop_front() {
        let part = package.get_part(&uri)?;
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
            let requested_part = package.get_part(&requested)?;
            if requested_part.partname().as_str().len() > limits.name_bytes() {
                return Err(crate::Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::StringBytes,
                    actual: requested_part.partname().as_str().len(),
                    maximum: limits.name_bytes(),
                });
            }
            if requested_part.content_type().len() > limits.string_bytes() {
                return Err(crate::Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::StringBytes,
                    actual: requested_part.content_type().len(),
                    maximum: limits.string_bytes(),
                });
            }
            let target = requested_part.partname().clone();
            if is_signature_part(&target) || is_reserved_physical_part(&target) {
                return Err(invalid(format!(
                    "glossary relationship '{}' targets reserved OPC package infrastructure",
                    relationship.r_id()
                )));
            }
            let target_part = requested_part;
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
            if owned.len() >= max_owned {
                return Err(crate::Error::GlossaryLimit {
                    kind: crate::glossary::LimitKind::OwnedPart,
                    actual: owned.len().saturating_add(1),
                    maximum: max_owned,
                });
            }
            admit_owned_part(
                target_part,
                limits,
                &mut total_bytes,
                &mut metadata_bytes,
                &mut relationship_count,
                true,
            )?;
            roles.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "glossary owned-part roles",
                source,
            })?;
            owned.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "glossary owned-part closure",
                source,
            })?;
            queue.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "glossary owned-part queue",
                source,
            })?;
            roles.insert(target.clone(), target_role);
            owned.insert(target.clone());
            queue.push_back((target, target_role));
        }
    }
    let part_count = owned.len().saturating_sub(1);
    Ok((
        owned,
        ClosureLedger {
            total_bytes,
            metadata_bytes,
            relationship_count,
            part_count,
        },
    ))
}

pub(in crate::glossary) fn admit_owner_closure(
    package: &OpcPackage,
    owner: Option<&Owner>,
    limits: &raw::Limits,
) -> Result<(HashSet<PackURI>, ClosureLedger)> {
    let Some(owner) = owner else {
        return Ok((HashSet::new(), ClosureLedger::default()));
    };
    let metadata = owner_metadata_seed(owner, limits)?;
    let result = glossary_owned_parts(
        package,
        &owner.root,
        owner.conformance,
        limits,
        metadata,
        true,
    )?;
    // `glossary_owned_parts` performs the complete source-closure admission,
    // including selected payload, metadata, relationship and XML limits,
    // before this caller retains any graph projection.
    Ok(result)
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
