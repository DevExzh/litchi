//! Structural checks shared by semantic authoring and package emission.

use super::snapshot::Snapshot;
use litchi_core::{Error, Result};

pub(super) fn validate(snapshot: Snapshot<'_>) -> Result<()> {
    if snapshot.slides.len() > 65_536 {
        return Err(Error::InvalidFormat(
            "ODP document exceeds 65536 slides".to_string(),
        ));
    }
    if snapshot.media_files.len() > 65_536 {
        return Err(Error::InvalidFormat(
            "ODP package exceeds 65536 embedded media files".to_string(),
        ));
    }
    snapshot.page_layouts.validate()?;
    if let Some(settings) = snapshot.settings {
        settings.validate()?;
    }
    if let Some(declarations) = snapshot.declarations {
        declarations.validate()?;
    }
    if let Some(metadata) = snapshot.page_metadata {
        metadata.validate_for_slides(snapshot.slides.len())?;
    }
    Ok(())
}

pub(super) fn push_drawing_attributes(
    output: &mut String,
    attributes: &[crate::DrawingAttribute],
) -> Result<()> {
    use std::collections::BTreeSet;

    let mut names = BTreeSet::new();
    for attribute in attributes {
        let qualified_name = format!("{}:{}", attribute.namespace.prefix(), attribute.local_name);
        if !names.insert(qualified_name.clone()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate drawing attribute '{qualified_name}'"
            )));
        }
        output.push(' ');
        output.push_str(&qualified_name);
        output.push_str("=\"");
        output.push_str(&litchi_core::xml::escape_xml(&attribute.value));
        output.push('"');
    }
    Ok(())
}

pub(super) fn validate_drawing_shape_parent(
    kind: crate::DrawingShapeKind,
    parent: Option<crate::DrawingShapeKind>,
) -> Result<()> {
    use crate::DrawingShapeKind;

    match parent {
        None if kind.is_three_dimensional() && kind != DrawingShapeKind::ThreeDimensionalScene => {
            Err(Error::InvalidFormat(
                "3D drawing objects require a dr3d:scene parent".to_string(),
            ))
        },
        None => Ok(()),
        Some(DrawingShapeKind::Group) => {
            if kind.is_three_dimensional() && kind != DrawingShapeKind::ThreeDimensionalScene {
                Err(Error::InvalidFormat(
                    "3D drawing objects require a dr3d:scene parent".to_string(),
                ))
            } else {
                Ok(())
            }
        },
        Some(DrawingShapeKind::ThreeDimensionalScene) if kind.is_three_dimensional() => Ok(()),
        Some(DrawingShapeKind::ThreeDimensionalScene) => Err(Error::InvalidFormat(
            "dr3d:scene can only contain 3D lights and objects".to_string(),
        )),
        Some(_) => Err(Error::InvalidFormat(
            "nested drawing shapes require a draw:g or dr3d:scene parent".to_string(),
        )),
    }
}

pub(super) fn validate_three_dimensional_child_order(children: &[crate::Shape]) -> Result<()> {
    use crate::DrawingShapeKind;

    let mut object_seen = false;
    for child in children {
        let kind = child.drawing_kind().ok_or_else(|| {
            Error::InvalidFormat(
                "dr3d:scene child is missing its exact 3D element kind".to_string(),
            )
        })?;
        if kind == DrawingShapeKind::ThreeDimensionalLight {
            if object_seen {
                return Err(Error::InvalidFormat(
                    "dr3d:light elements must precede 3D objects".to_string(),
                ));
            }
        } else {
            object_seen = true;
        }
    }
    Ok(())
}

pub(super) fn validate_required_three_dimensional_attributes(
    kind: crate::DrawingShapeKind,
    attributes: &[crate::DrawingAttribute],
) -> Result<()> {
    use crate::{DrawingAttributeNamespace, DrawingShapeKind};

    let has = |namespace, local_name| {
        attributes.iter().any(|attribute| {
            attribute.namespace() == namespace && attribute.local_name() == local_name
        })
    };
    if kind == DrawingShapeKind::ThreeDimensionalLight
        && !has(DrawingAttributeNamespace::Dr3d, "direction")
    {
        return Err(Error::InvalidFormat(
            "dr3d:light requires dr3d:direction".to_string(),
        ));
    }
    if matches!(
        kind,
        DrawingShapeKind::ThreeDimensionalExtrude | DrawingShapeKind::ThreeDimensionalRotate
    ) {
        for (namespace, local_name) in [
            (DrawingAttributeNamespace::Svg, "viewBox"),
            (DrawingAttributeNamespace::Svg, "d"),
        ] {
            if !has(namespace, local_name) {
                return Err(Error::InvalidFormat(format!(
                    "{} requires svg:{local_name}",
                    kind.element_name()
                )));
            }
        }
    }
    Ok(())
}
