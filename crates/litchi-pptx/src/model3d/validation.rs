//! PPTX relationship, content-type, and bounded-resource validation.

use litchi_drawingml::model3d as drawing;
use litchi_drawingml::{Error as DrawingError, LimitKind};
use litchi_opc::constants::relationship_type::{IMAGE, STRICT_IMAGE};
use litchi_opc::{OpcPackage, Part};
use quick_xml::{Reader, events::Event};

use super::codec::Limits;
use super::model::{Asset, Data, Model, Preview, Relation};
use super::{MODEL_CONTENT_TYPE, MODEL_RELATIONSHIP};
use crate::{Error, Result};

/// Validate the detached semantic owner before package publication.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn validate(model: &Model, limits: Limits) -> Result<()> {
    validate_selected_structure(&model.scene.wire, limits)?;
    drawing::validate(&model.scene.wire).map_err(super::model::drawing_model3d_error)?;
    validate_asset(&model.asset, limits)?;
    if let Some(preview) = &model.preview {
        validate_preview(preview, limits)?;
        if !model.scene.has_raster() {
            return Err(invalid("preview requires a raster child"));
        }
    }
    let total = model
        .asset
        .embedded_data()
        .map_or(0, Data::len)
        .checked_add(
            model
                .preview
                .as_ref()
                .and_then(Preview::data)
                .map_or(0, Data::len),
        )
        .ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::PayloadBytes,
                usize::MAX,
                limits.max_total_payload_bytes(),
            )
        })?;
    if total > limits.max_total_payload_bytes() {
        return Err(limit(
            crate::error::CodecLimitKind::PayloadBytes,
            total,
            limits.max_total_payload_bytes(),
        ));
    }
    Ok(())
}

pub(crate) fn validate_loaded_graph(
    package: &OpcPackage,
    source: &dyn Part,
    model: &Model,
    limits: Limits,
) -> Result<()> {
    validate(model, limits)?;
    let resolver = Resolver { source };
    drawing::validate_relationships(&model.scene.wire, &resolver)
        .map_err(super::model::drawing_model3d_error)?;

    let mut referenced = Vec::new();
    validate_reference(
        package,
        source,
        model.scene.wire.reference.embedded.as_ref(),
        false,
        "model",
        MODEL_RELATIONSHIP,
        &mut referenced,
        limits,
    )?;
    validate_reference(
        package,
        source,
        model.scene.wire.reference.linked.as_ref(),
        true,
        "model",
        MODEL_RELATIONSHIP,
        &mut referenced,
        limits,
    )?;
    if let Some(raster) = model.scene.wire.raster() {
        for child in &raster.children {
            let drawing::RasterChild::Blip(blip) = child else {
                continue;
            };
            validate_reference(
                package,
                source,
                blip.reference.embedded.as_ref(),
                false,
                "raster.blip",
                image_relationship(source, false),
                &mut referenced,
                limits,
            )?;
            validate_reference(
                package,
                source,
                blip.reference.linked.as_ref(),
                true,
                "raster.blip",
                image_relationship(source, true),
                &mut referenced,
                limits,
            )?;
        }
    }

    Ok(())
}

/// Validate the source-wide model relationship inventory after all model
/// anchors have been discovered. A slide can own more than one graphic frame,
/// so this check intentionally runs once for the complete inventory.
pub(crate) fn validate_model_relationship_inventory(
    source: &dyn Part,
    models: &[Model],
    limits: Limits,
) -> Result<()> {
    let mut referenced = Vec::new();
    for model in models {
        validate_selected_structure(&model.scene.wire, limits)?;
        collect_reference(&model.scene.wire.reference, &mut referenced);
        if let Some(raster) = model.scene.wire.raster()
            && let Some(blip) = raster.children.iter().find_map(|child| match child {
                drawing::RasterChild::Blip(blip) => Some(blip),
                drawing::RasterChild::Opaque(_) => None,
                _ => None,
            })
        {
            collect_reference(&blip.reference, &mut referenced);
        }
    }
    for relationship in source.rels().iter() {
        if relationship.reltype() == MODEL_RELATIONSHIP
            && !referenced.iter().any(|id| id == relationship.r_id())
            && !source
                .blob()
                .windows(relationship.r_id().len())
                .any(|window| window == relationship.r_id().as_bytes())
        {
            return Err(invalid(format!(
                "unreferenced model3d relationship '{}'",
                relationship.r_id()
            )));
        }
    }
    Ok(())
}

/// Apply the PPTX adapter's selected structural policy before the shared
/// DrawingML validator runs.  The shared semantic validator intentionally
/// uses its crate defaults; a package owner must not lose tighter limits when
/// it crosses this adapter boundary.
fn validate_selected_structure(metadata: &drawing::Metadata, limits: Limits) -> Result<()> {
    if metadata.children.len() > limits.max_children() {
        return Err(drawing_limit(
            LimitKind::ModelChildren,
            metadata.children.len(),
            limits.max_children(),
        ));
    }
    if metadata.namespaces.len() > limits.max_nodes() {
        return Err(drawing_limit(
            LimitKind::NamespaceDeclarations,
            metadata.namespaces.len(),
            limits.max_nodes(),
        ));
    }
    for child in &metadata.children {
        match child {
            drawing::Child::Raster(raster) => {
                if raster.children.len() > limits.max_raster_children() {
                    return Err(drawing_limit(
                        LimitKind::RasterChildren,
                        raster.children.len(),
                        limits.max_raster_children(),
                    ));
                }
                if let Some(blip) = raster.blip() {
                    for inert in &blip.children {
                        validate_fragment_limits(inert, limits)?;
                    }
                }
                for child in &raster.children {
                    if let drawing::RasterChild::Opaque(inert) = child {
                        validate_fragment_limits(inert, limits)?;
                    }
                }
            },
            drawing::Child::Opaque(inert) => validate_fragment_limits(inert, limits)?,
            _ => return Err(invalid("unsupported model3d child variant")),
        }
    }
    Ok(())
}

fn validate_fragment_limits(inert: &drawing::Inert, limits: Limits) -> Result<()> {
    let bytes = inert.as_bytes();
    if bytes.len() > limits.max_fragment_bytes() {
        return Err(drawing_limit(
            LimitKind::InertFragmentBytes,
            bytes.len(),
            limits.max_fragment_bytes(),
        ));
    }
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_)) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    drawing_limit(
                        LimitKind::InertFragmentNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                if nodes > limits.max_nodes() {
                    return Err(drawing_limit(
                        LimitKind::InertFragmentNodes,
                        nodes,
                        limits.max_nodes(),
                    ));
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    drawing_limit(
                        LimitKind::InertFragmentDepth,
                        usize::MAX,
                        limits.max_depth(),
                    )
                })?;
                if depth > limits.max_depth() {
                    return Err(drawing_limit(
                        LimitKind::InertFragmentDepth,
                        depth,
                        limits.max_depth(),
                    ));
                }
            },
            Ok(Event::Empty(_)) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    drawing_limit(
                        LimitKind::InertFragmentNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                if nodes > limits.max_nodes() {
                    return Err(drawing_limit(
                        LimitKind::InertFragmentNodes,
                        nodes,
                        limits.max_nodes(),
                    ));
                }
            },
            Ok(Event::End(_)) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Drawing(DrawingError::Xml(
                        "unexpected model3d fragment closing element".into(),
                    ))
                })?;
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(Error::Drawing(DrawingError::Xml(error.to_string()))),
        }
        buffer.clear();
    }
    if depth != 0 || nodes == 0 {
        return Err(Error::Drawing(DrawingError::Xml(
            "incomplete model3d fragment".into(),
        )));
    }
    Ok(())
}

fn drawing_limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::Drawing(DrawingError::Limit {
        kind,
        observed,
        maximum,
    })
}

fn collect_reference(reference: &drawing::Reference, referenced: &mut Vec<String>) {
    if let Some(id) = &reference.embedded {
        referenced.push(id.as_str().to_owned());
    }
    if let Some(id) = &reference.linked {
        referenced.push(id.as_str().to_owned());
    }
}

fn validate_asset(asset: &Asset, limits: Limits) -> Result<()> {
    if let Some(data) = asset.embedded_data()
        && data.len() > limits.max_model_bytes()
    {
        return Err(limit(
            crate::error::CodecLimitKind::PayloadBytes,
            data.len(),
            limits.max_model_bytes(),
        ));
    }
    if let Some(link) = asset.linked_target()
        && (link.as_str().is_empty() || link.as_str().len() > limits.max_link_bytes())
    {
        return Err(invalid("external model3d target exceeds its bound"));
    }
    Ok(())
}

fn validate_preview(preview: &Preview, limits: Limits) -> Result<()> {
    if let Some(data) = preview.data() {
        if data.len() > limits.max_preview_bytes() {
            return Err(limit(
                crate::error::CodecLimitKind::PayloadBytes,
                data.len(),
                limits.max_preview_bytes(),
            ));
        }
        let content_type = preview
            .content_type()
            .ok_or_else(|| invalid("embedded model3d preview has no content type"))?;
        if !is_image_content_type(content_type) {
            return Err(invalid(format!(
                "model3d preview has non-image content type '{content_type}'"
            )));
        }
    }
    if let Some(link) = preview.linked_target()
        && (link.as_str().is_empty() || link.as_str().len() > limits.max_link_bytes())
    {
        return Err(invalid("external model3d preview target exceeds its bound"));
    }
    Ok(())
}

fn validate_reference(
    package: &OpcPackage,
    source: &dyn Part,
    id: Option<&drawing::Id>,
    external: bool,
    field: &'static str,
    expected_type: &str,
    referenced: &mut Vec<String>,
    limits: Limits,
) -> Result<()> {
    let Some(id) = id else {
        return Ok(());
    };
    referenced.push(id.as_str().to_owned());
    let relationship = source
        .rels()
        .get(id.as_str())
        .ok_or_else(|| invalid(format!("{field} relationship '{id}' is missing")))?;
    if relationship.reltype() != expected_type
        && !(is_image_relationship(expected_type) && is_image_relationship(relationship.reltype()))
    {
        return Err(invalid(format!(
            "{field} relationship '{}' has type '{}'",
            id,
            relationship.reltype()
        )));
    }
    if relationship.is_external() != external {
        return Err(invalid(format!(
            "{field} relationship '{id}' has an invalid target mode"
        )));
    }
    if external {
        if relationship.target_ref().is_empty()
            || relationship.target_ref().len() > limits.max_link_bytes()
        {
            return Err(invalid(format!(
                "{field} relationship '{id}' has an invalid target"
            )));
        }
        return Ok(());
    }
    let target = relationship.target_partname()?;
    let part = package.get_part(&target)?;
    if field == "model" {
        if part.content_type() != MODEL_CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: MODEL_CONTENT_TYPE.into(),
                actual: part.content_type().into(),
            });
        }
    } else if !is_image_content_type(part.content_type()) {
        return Err(invalid(format!(
            "{field} relationship '{}' targets non-image content type '{}'",
            id,
            part.content_type()
        )));
    }
    if part.blob().len()
        > if field == "model" {
            limits.max_model_bytes()
        } else {
            limits.max_preview_bytes()
        }
    {
        return Err(limit(
            crate::error::CodecLimitKind::PayloadBytes,
            part.blob().len(),
            if field == "model" {
                limits.max_model_bytes()
            } else {
                limits.max_preview_bytes()
            },
        ));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid(format!(
            "{field} part '{target}' has outbound relationships"
        )));
    }
    Ok(())
}

/// Build the shared relationship view without allocating an absolute target;
/// shared validation only needs target mode and lexical safety.
pub(crate) struct Resolver<'a> {
    pub(crate) source: &'a dyn Part,
}

impl drawing::Resolver for Resolver<'_> {
    fn relationship<'a>(&'a self, id: &drawing::Id) -> Option<drawing::Relationship<'a>> {
        let relationship = self.source.rels().get(id.as_str())?;
        Some(drawing::Relationship {
            relationship_type: relationship.reltype(),
            target: if relationship.is_external() {
                drawing::Target::External(relationship.target_ref())
            } else {
                drawing::Target::Internal(relationship.target_ref())
            },
        })
    }
}

pub(crate) fn relation_from_part(source: &dyn Part, id: &drawing::Id) -> Result<Relation> {
    let relationship = source
        .rels()
        .get(id.as_str())
        .ok_or_else(|| invalid(format!("relationship '{id}' is missing")))?;
    Ok(Relation {
        id: id.as_str().to_owned(),
        target: relationship.target_ref().to_owned(),
        external: relationship.is_external(),
    })
}

pub(crate) fn is_image_content_type(value: &str) -> bool {
    value.starts_with("image/") || matches!(value, "application/x-emf" | "application/x-wmf")
}

pub(crate) fn image_relationship(_source: &dyn Part, _external: bool) -> &'static str {
    IMAGE
}

fn is_image_relationship(value: &str) -> bool {
    matches!(value, IMAGE | STRICT_IMAGE)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(format!("PPTX model3d {}", message.into()))
}

fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}
