//! Structural and value validation for parsed ODP content.

use super::*;

impl Parser {
    pub(super) fn validate_shape_parent(
        parent: &ShapeBuilder,
        child: DrawingShapeKind,
    ) -> Result<()> {
        match parent.drawing_kind {
            Some(DrawingShapeKind::Group) => {
                if child.is_three_dimensional() && child != DrawingShapeKind::ThreeDimensionalScene
                {
                    return Err(Error::InvalidFormat(
                        "3D drawing objects require a dr3d:scene parent".to_string(),
                    ));
                }
            },
            Some(DrawingShapeKind::ThreeDimensionalScene) => {
                if !child.is_three_dimensional() {
                    return Err(Error::InvalidFormat(
                        "dr3d:scene can only contain 3D lights and objects".to_string(),
                    ));
                }
                if child == DrawingShapeKind::ThreeDimensionalLight
                    && parent.children.iter().any(|existing| {
                        existing.drawing_kind() != Some(DrawingShapeKind::ThreeDimensionalLight)
                    })
                {
                    return Err(Error::InvalidFormat(
                        "dr3d:light elements must precede 3D objects".to_string(),
                    ));
                }
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "nested drawing shapes require a draw:g or dr3d:scene parent".to_string(),
                ));
            },
        }
        Ok(())
    }

    pub(super) fn validate_three_dimensional_child_element(
        parent: Option<&ShapeBuilder>,
        child: Element,
    ) -> Result<()> {
        let Some(parent_kind) = parent.and_then(|builder| builder.drawing_kind) else {
            return Ok(());
        };
        if !parent_kind.is_three_dimensional() {
            return Ok(());
        }
        if parent_kind != DrawingShapeKind::ThreeDimensionalScene {
            return Err(Error::InvalidFormat(
                "3D light and object elements cannot contain child elements".to_string(),
            ));
        }
        match child {
            Element::Shape(shape) if Self::drawing_kind(shape).is_three_dimensional() => Ok(()),
            // `svg:title`, `svg:desc`, `draw:glue-point`, and foreign
            // extension elements are intentionally handled as opaque content.
            Element::Other => Ok(()),
            _ => Err(Error::InvalidFormat(
                "dr3d:scene can only contain 3D content".to_string(),
            )),
        }
    }

    pub(super) fn validate_required_three_dimensional_attributes(
        kind: DrawingShapeKind,
        attributes: &[DrawingAttribute],
    ) -> Result<()> {
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
            for local_name in ["viewBox", "d"] {
                if !has(DrawingAttributeNamespace::Svg, local_name) {
                    return Err(Error::InvalidFormat(format!(
                        "{} requires svg:{local_name}",
                        kind.element_name()
                    )));
                }
            }
        }
        Ok(())
    }

    pub(super) fn required_attr(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        namespace: &[u8],
        local_name: &[u8],
        qualified_name: &str,
    ) -> Result<String> {
        Self::get_attr(reader, element, namespace, local_name)?.ok_or_else(|| {
            Error::InvalidFormat(format!("element is missing required {qualified_name}"))
        })
    }

    pub(super) fn require_simple_xlink(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        description: &str,
    ) -> Result<()> {
        let link_type =
            Self::required_attr(reader, element, XLINK_NAMESPACE, b"type", "xlink:type")?;
        if link_type != "simple" {
            return Err(Error::InvalidFormat(format!(
                "{description} xlink:type must be 'simple', found '{link_type}'"
            )));
        }
        Ok(())
    }

    pub(super) fn parse_optional_bool(
        value: Option<String>,
        attribute: &str,
    ) -> Result<Option<bool>> {
        value
            .map(|value| match value.as_str() {
                "true" | "1" => Ok(true),
                "false" | "0" => Ok(false),
                _ => Err(Error::InvalidFormat(format!(
                    "invalid {attribute} value '{value}'"
                ))),
            })
            .transpose()
    }
}
