//! Package-facing glossary load/store/remove orchestration.

use super::codec::*;
use super::graph::*;
use super::model::*;
use super::*;
pub(in crate::glossary) struct Owner {
    pub(in crate::glossary) main: PackURI,
    pub(in crate::glossary) root: PackURI,
    pub(in crate::glossary) relationship_id: String,
    pub(in crate::glossary) relationship_target: String,
    pub(in crate::glossary) conformance: Conformance,
}

/// Load the semantic catalog and its namespace dialect, without copying auxiliaries.
pub fn load(package: &OpcPackage) -> Result<Option<(Catalog, Conformance)>> {
    let Some(graph) = load_graph(package)? else {
        return Ok(None);
    };
    let binding = Binding::from_graph(&graph);
    let conformance = graph.conformance;
    let mut catalog = graph.catalog;
    for entry in &mut catalog.entries {
        if entry.has_relationship_references() {
            entry.lineage = Some(Arc::clone(&binding.lineage));
        }
    }
    if !catalog.background_refs.is_empty() {
        catalog.background_lineage = Some(Arc::clone(&binding.lineage));
    }
    catalog.binding = Some(Box::new(binding));
    Ok(Some((catalog, conformance)))
}

/// Move a semantic catalog into the package while preserving its auxiliary graph.
///
/// An unchanged package-loaded catalog is a byte- and signature-preserving no-op.
pub fn put(
    package: &mut OpcPackage,
    mut catalog: Catalog,
    conformance: Conformance,
) -> Result<bool> {
    validate_package_conformance(package, conformance)?;
    catalog.validate_bound_lineages()?;
    let existing = load_graph(package)?;
    let binding = catalog.binding.take();
    let bound_to_destination = binding
        .as_deref()
        .zip(existing.as_ref())
        .is_some_and(|(binding, graph)| binding.matches(graph));
    if bound_to_destination
        && existing
            .as_ref()
            .is_some_and(|graph| graph.catalog == catalog)
    {
        return Ok(false);
    }
    if !catalog_relationship_references(&catalog, conformance)?.is_empty() && !bound_to_destination
    {
        return Err(invalid(
            "glossary relationship references are bound to another physical graph; use glossary::raw for graph transfer",
        ));
    }
    let mut graph = if bound_to_destination {
        existing.ok_or_else(|| invalid("bound glossary graph disappeared"))?
    } else {
        let mut graph = raw::Graph::new(Catalog::new(), conformance);
        seed_semantic_graph(package, &mut graph)?;
        graph
    };
    graph.catalog = catalog;
    graph.conformance = conformance;
    put_graph(package, &graph)
}

pub(in crate::glossary) fn seed_semantic_graph(
    package: &OpcPackage,
    graph: &mut raw::Graph,
) -> Result<()> {
    if !graph.rels.is_empty() || !graph.parts.is_empty() {
        return Err(invalid("semantic glossary seed graph is not empty"));
    }
    let root_uri = free_part_name(
        package,
        "/word/glossary/document.xml",
        "/word/glossary/document%d.xml",
    )?;
    graph.root_name = root_uri.as_str().to_owned();
    let namespace = graph.conformance.word();
    for (index, kind, preferred, template, content_type, root) in [
        (
            1,
            "styles",
            "/word/glossary/styles.xml",
            "/word/glossary/styles%d.xml",
            ct::WML_STYLES,
            "styles",
        ),
        (
            2,
            "settings",
            "/word/glossary/settings.xml",
            "/word/glossary/settings%d.xml",
            ct::WML_SETTINGS,
            "settings",
        ),
        (
            3,
            "fontTable",
            "/word/glossary/fontTable.xml",
            "/word/glossary/fontTable%d.xml",
            ct::WML_FONT_TABLE,
            "fonts",
        ),
        (
            4,
            "webSettings",
            "/word/glossary/webSettings.xml",
            "/word/glossary/webSettings%d.xml",
            ct::WML_WEB_SETTINGS,
            "webSettings",
        ),
    ] {
        let part_uri = free_part_name(package, preferred, template)?;
        graph.rels.push(raw::Rel {
            id: format!("rId{index}"),
            kind: format!("{}/{kind}", graph.conformance.relationships()),
            target: part_uri.relative_ref(root_uri.base_uri()),
            external: false,
        });
        let source = package
            .main_document_part()?
            .rels()
            .iter()
            .find_map(|relationship| {
                (!relationship.is_external()
                    && relationship_kind(graph.conformance, relationship.reltype()) == Some(kind))
                .then(|| relationship.target_partname().ok())
                .flatten()
                .and_then(|target| package.get_part(&target).ok())
                .filter(|part| {
                    part.content_type() == content_type && part.rels().iter().next().is_none()
                })
                .map(Part::blob_arc)
            });
        let data = source.unwrap_or_else(|| {
            Arc::new(format!(r#"<w:{root} xmlns:w="{namespace}"/>"#).into_bytes())
        });
        graph.parts.push(raw::Part::from_shared(
            part_uri.as_str().to_owned(),
            content_type.to_owned(),
            data,
            Vec::new(),
        ));
    }
    Ok(())
}

pub(in crate::glossary) fn free_part_name(
    package: &OpcPackage,
    preferred: &str,
    template: &str,
) -> Result<PackURI> {
    let preferred = PackURI::new(preferred).map_err(Error::Uri)?;
    if package.validate_new_part_name(&preferred).is_ok() {
        return Ok(preferred);
    }
    let marker = template
        .find("%d")
        .ok_or_else(|| invalid("glossary part-name template is missing '%d'"))?;
    for index in 1..=10_000u32 {
        let mut candidate = String::new();
        candidate
            .try_reserve(template.len().saturating_add(10))
            .map_err(|source| Error::Allocation {
                resource: "glossary part name",
                source,
            })?;
        candidate.push_str(&template[..marker]);
        candidate.push_str(&index.to_string());
        candidate.push_str(&template[marker + 2..]);
        let candidate = PackURI::new(&candidate).map_err(Error::Uri)?;
        if package.validate_new_part_name(&candidate).is_ok() {
            return Ok(candidate);
        }
    }
    Err(invalid(
        "glossary part-name allocation exhausted 10,000 candidates",
    ))
}

/// Remove the glossary graph. Absence is a signature-preserving no-op.
///
/// Use [`raw::remove`] to move the complete physical graph elsewhere.
pub fn remove(package: &mut OpcPackage) -> Result<bool> {
    Ok(remove_graph(package)?.is_some())
}
