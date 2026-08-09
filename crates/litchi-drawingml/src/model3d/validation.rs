//! Semantic and package-graph validation for the model3d slice.

use quick_xml::{Reader, events::Event};
use thiserror::Error;

use super::package::{Relationship, Resolver};
use super::{
    Child, Inert, MAX_CHILDREN, MAX_DEPTH, MAX_FRAGMENT_BYTES, MAX_NAMESPACE_DECLARATIONS,
    MAX_NODES, MAX_RENDERER_TEXT_BYTES, Metadata, NAMESPACE, Raster, RasterChild, Reference,
};

/// Failure to validate a model3d semantic value or relationship graph.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// A bounded collection exceeded its safe limit.
    #[error("model3d {resource} exceeds the limit of {limit}")]
    Limit {
        /// Bounded resource.
        resource: &'static str,
        /// Maximum accepted count or byte length.
        limit: usize,
    },
    /// A required scene child is absent.
    #[error("model3d required child '{0}' is missing")]
    MissingChild(&'static str),
    /// A schema singleton child occurred more than once.
    #[error("model3d child '{0}' occurs more than once")]
    DuplicateChild(&'static str),
    /// A known child appears outside the normative `CT_Model3D` sequence.
    #[error("model3d child '{child}' is out of sequence")]
    ChildOrder {
        /// Local child name.
        child: &'static str,
    },
    /// The two mutually distinct viewport alternatives were both supplied.
    #[error("model3d object and window viewports cannot both be present")]
    MultipleViewports,
    /// A retained child fragment is not one complete XML element.
    #[error("invalid inert model3d child: {0}")]
    Inert(String),
    /// Both relationship attributes reuse one relationship occurrence.
    #[error("model3d {field} reuses relationship ID '{id}'")]
    DuplicateReference {
        /// Relationship-bearing field.
        field: &'static str,
        /// Reused ID.
        id: String,
    },
    /// A relationship ID is absent from the owning package graph.
    #[error("model3d {field} relationship '{id}' is missing")]
    MissingRelationship {
        /// Relationship-bearing field.
        field: &'static str,
        /// Missing ID.
        id: String,
    },
    /// An `r:embed` relationship resolved to an external target.
    #[error("model3d {field} relationship '{id}' must target an internal part")]
    EmbeddedTargetIsExternal {
        /// Relationship-bearing field.
        field: &'static str,
        /// Relationship ID.
        id: String,
    },
    /// An `r:link` relationship resolved to an internal target.
    #[error("model3d {field} relationship '{id}' must target an external resource")]
    LinkedTargetIsInternal {
        /// Relationship-bearing field.
        field: &'static str,
        /// Relationship ID.
        id: String,
    },
    /// A relationship target is empty or contains forbidden internal syntax.
    #[error("model3d {field} relationship target is invalid")]
    InvalidTarget {
        /// Relationship-bearing field.
        field: &'static str,
    },
    /// The package resolver returned a relationship without a type URI.
    #[error("model3d {field} relationship has an empty type URI")]
    EmptyRelationshipType {
        /// Relationship-bearing field.
        field: &'static str,
    },
}

/// Validate the typed model and its retained scene sequence.
pub fn validate(metadata: &Metadata) -> Result<(), Error> {
    validate_counts(metadata)?;
    validate_reference(&metadata.reference, "model")?;

    let mut stage = 0u8;
    let mut sp_pr = false;
    let mut camera = false;
    let mut transform = false;
    let mut viewport = None;
    let mut raster = false;
    let mut extension_list = false;
    let mut ambient = false;

    for child in &metadata.children {
        let (local_name, namespace) = match child {
            Child::Raster(value) => {
                validate_raster(value)?;
                if raster {
                    return Err(Error::DuplicateChild("raster"));
                }
                raster = true;
                ("raster", NAMESPACE)
            },
            Child::Opaque(value) => {
                validate_inert_fragment(value)?;
                (value.local_name(), value.namespace())
            },
        };

        let model = namespace == NAMESPACE;
        match (local_name, model) {
            ("spPr", true) => {
                if sp_pr || stage != 0 {
                    return Err(Error::ChildOrder { child: "spPr" });
                }
                sp_pr = true;
                stage = 1;
            },
            ("camera", true) => {
                if camera || stage > 1 {
                    return Err(Error::ChildOrder { child: "camera" });
                }
                camera = true;
                stage = 2;
            },
            ("trans", true) => {
                if transform || stage > 2 {
                    return Err(Error::ChildOrder { child: "trans" });
                }
                transform = true;
                stage = 3;
            },
            ("attrSrcUrl", true) => {
                if !(3..=4).contains(&stage) {
                    return Err(Error::ChildOrder {
                        child: "attrSrcUrl",
                    });
                }
                stage = 4;
            },
            ("raster", true) => {
                if !(3..=5).contains(&stage) {
                    return Err(Error::ChildOrder { child: "raster" });
                }
                stage = 5;
            },
            ("extLst", true) => {
                if extension_list || !(3..=6).contains(&stage) {
                    return Err(Error::ChildOrder { child: "extLst" });
                }
                extension_list = true;
                stage = 6;
            },
            ("objViewport" | "winViewport", true) => {
                if viewport.is_some() || !(3..=7).contains(&stage) {
                    return if viewport.is_some() {
                        Err(Error::MultipleViewports)
                    } else {
                        Err(Error::ChildOrder { child: "viewport" })
                    };
                }
                viewport = Some(local_name);
                stage = 7;
            },
            ("ambientLight", true) => {
                if ambient || stage < 7 {
                    return Err(Error::ChildOrder {
                        child: "ambientLight",
                    });
                }
                ambient = true;
                stage = 8;
            },
            ("ptLight" | "spotLight" | "dirLight" | "unkLight", true) => {
                if stage < 7 {
                    return Err(Error::ChildOrder { child: "light" });
                }
                stage = 8;
            },
            _ => {
                // Future namespaces and unmodeled child vocabulary remain inert.
            },
        }
    }

    if !sp_pr {
        return Err(Error::MissingChild("spPr"));
    }
    if !camera {
        return Err(Error::MissingChild("camera"));
    }
    if !transform {
        return Err(Error::MissingChild("trans"));
    }
    if viewport.is_none() {
        return Err(Error::MissingChild("objViewport or winViewport"));
    }
    Ok(())
}

/// Validate semantic metadata against a host-provided package relationship graph.
pub fn validate_relationships<R: Resolver + ?Sized>(
    metadata: &Metadata,
    resolver: &R,
) -> Result<(), Error> {
    validate(metadata)?;
    validate_reference_graph(&metadata.reference, "model", resolver)?;
    for child in &metadata.children {
        if let Child::Raster(raster) = child {
            for raster_child in &raster.children {
                if let RasterChild::Blip(blip) = raster_child {
                    validate_reference_graph(&blip.reference, "raster.blip", resolver)?;
                }
            }
        }
    }
    Ok(())
}

fn validate_counts(metadata: &Metadata) -> Result<(), Error> {
    if metadata.children.len() > MAX_CHILDREN {
        return Err(Error::Limit {
            resource: "children",
            limit: MAX_CHILDREN,
        });
    }
    if metadata.namespaces.len() > MAX_NAMESPACE_DECLARATIONS {
        return Err(Error::Limit {
            resource: "namespace declarations",
            limit: MAX_NAMESPACE_DECLARATIONS,
        });
    }
    Ok(())
}

fn validate_raster(raster: &Raster) -> Result<(), Error> {
    if raster.renderer_name.len() > MAX_RENDERER_TEXT_BYTES
        || raster.renderer_version.len() > MAX_RENDERER_TEXT_BYTES
    {
        return Err(Error::Limit {
            resource: "renderer text",
            limit: MAX_RENDERER_TEXT_BYTES,
        });
    }
    if raster.children.len() > super::MAX_RASTER_CHILDREN {
        return Err(Error::Limit {
            resource: "raster children",
            limit: super::MAX_RASTER_CHILDREN,
        });
    }
    let mut blip = false;
    for child in &raster.children {
        match child {
            RasterChild::Blip(value) => {
                if blip {
                    return Err(Error::DuplicateChild("raster/blip"));
                }
                blip = true;
                validate_reference(&value.reference, "raster.blip")?;
                for inert in &value.children {
                    validate_inert_fragment(inert)?;
                }
            },
            RasterChild::Opaque(value) => validate_inert_fragment(value)?,
        }
    }
    Ok(())
}

fn validate_reference(reference: &Reference, field: &'static str) -> Result<(), Error> {
    if let (Some(embedded), Some(linked)) = (&reference.embedded, &reference.linked)
        && embedded == linked
    {
        return Err(Error::DuplicateReference {
            field,
            id: embedded.to_string(),
        });
    }
    Ok(())
}

fn validate_reference_graph<R: Resolver + ?Sized>(
    reference: &Reference,
    field: &'static str,
    resolver: &R,
) -> Result<(), Error> {
    validate_reference(reference, field)?;
    if let Some(id) = &reference.embedded {
        validate_target(field, id, resolver.relationship(id), false)?;
    }
    if let Some(id) = &reference.linked {
        validate_target(field, id, resolver.relationship(id), true)?;
    }
    Ok(())
}

fn validate_target(
    field: &'static str,
    id: &super::Id,
    relationship: Option<Relationship<'_>>,
    linked: bool,
) -> Result<(), Error> {
    let Some(relationship) = relationship else {
        return Err(Error::MissingRelationship {
            field,
            id: id.to_string(),
        });
    };
    if relationship.relationship_type.is_empty() {
        return Err(Error::EmptyRelationshipType { field });
    }
    let target = relationship.target;
    if target.as_str().is_empty() || (!target.is_external() && target.as_str().contains(['?', '#']))
    {
        return Err(Error::InvalidTarget { field });
    }
    if linked && !target.is_external() {
        return Err(Error::LinkedTargetIsInternal {
            field,
            id: id.to_string(),
        });
    }
    if !linked && target.is_external() {
        return Err(Error::EmbeddedTargetIsExternal {
            field,
            id: id.to_string(),
        });
    }
    Ok(())
}

pub(crate) fn validate_inert_fragment(value: &Inert) -> Result<(), Error> {
    if value.as_bytes().is_empty() || value.as_bytes().len() > MAX_FRAGMENT_BYTES {
        return Err(Error::Limit {
            resource: "inert fragment bytes",
            limit: MAX_FRAGMENT_BYTES,
        });
    }
    let mut reader = Reader::from_reader(value.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::Inert(error.to_string()))?;
        match event {
            Event::Start(_) => {
                if depth == 0 && (root_seen || root_closed) {
                    return Err(Error::Inert("multiple roots".into()));
                }
                nodes = nodes.saturating_add(1);
                if nodes > MAX_NODES {
                    return Err(Error::Limit {
                        resource: "inert fragment nodes",
                        limit: MAX_NODES,
                    });
                }
                depth = depth.saturating_add(1);
                if depth > MAX_DEPTH {
                    return Err(Error::Limit {
                        resource: "inert fragment depth",
                        limit: MAX_DEPTH,
                    });
                }
                root_seen = true;
            },
            Event::Empty(_) => {
                if depth == 0 && (root_seen || root_closed) {
                    return Err(Error::Inert("multiple roots".into()));
                }
                nodes = nodes.saturating_add(1);
                if nodes > MAX_NODES {
                    return Err(Error::Limit {
                        resource: "inert fragment nodes",
                        limit: MAX_NODES,
                    });
                }
                if depth == 0 {
                    root_seen = true;
                    root_closed = true;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Inert("unexpected closing element".into()))?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(text)
                if depth == 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(Error::Inert("text outside root".into()));
            },
            Event::CData(_) if depth == 0 => {
                return Err(Error::Inert("CDATA outside root".into()));
            },
            Event::DocType(_) | Event::Decl(_) => {
                return Err(Error::Inert("document-level markup is forbidden".into()));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen || depth != 0 || !root_closed {
        return Err(Error::Inert("fragment is not one complete element".into()));
    }
    if value.local_name().is_empty() {
        return Err(Error::Inert("fragment has no local name".into()));
    }
    Ok(())
}
