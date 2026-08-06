//! Namespace, attribute, and structural validation for ODP XML.

use super::super::*;

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
    pub(super) fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
        matches!(namespace, ResolveResult::Bound(XmlNamespace(value)) if *value == expected)
    }

    /// Rewinds the running element depth for a subtree consumed in place.
    ///
    /// Nested parsers such as [`Self::parse_enhanced_geometry`] read through
    /// their own element's `Event::End`, so the main event loop never observes
    /// it and would otherwise keep counting that element as open. Spreadsheet
    /// `table:shapes` container tracking compares depths exactly, so the
    /// increment taken on `Event::Start` has to be given back here.
    pub(super) const fn rewind_consumed_subtree(element_depth: usize) -> usize {
        element_depth.saturating_sub(1)
    }

    pub(super) fn animation_attributes(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Vec<Attribute>> {
        if element.attributes().count() > 256 {
            return Err(Error::InvalidFormat(
                "ODP animation node exceeds 256 attributes".to_string(),
            ));
        }
        let mut attributes = Vec::with_capacity(element.attributes().count());
        let mut expanded_names = HashSet::new();
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let qualified_name = attribute.key.as_ref();
            if qualified_name == b"xmlns" || qualified_name.starts_with(b"xmlns:") {
                continue;
            }
            let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
            let namespace_uri = match namespace {
                ResolveResult::Unbound => None,
                ResolveResult::Bound(XmlNamespace(uri)) => {
                    Some(std::str::from_utf8(uri).map_err(|_| {
                        Error::InvalidFormat("non-UTF-8 animation namespace URI".to_string())
                    })?)
                },
                ResolveResult::Unknown(prefix) => {
                    return Err(Error::InvalidFormat(format!(
                        "unknown animation attribute namespace prefix '{}'",
                        String::from_utf8_lossy(&prefix)
                    )));
                },
            };
            let local_name = std::str::from_utf8(local_name.as_ref())
                .map_err(|_| {
                    Error::InvalidFormat("non-UTF-8 animation attribute name".to_string())
                })?
                .to_string();
            let namespace = Namespace::from_uri(namespace_uri);
            if !expanded_names.insert((namespace.clone(), local_name.clone())) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate animation attribute '{local_name}'"
                )));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid XML attribute value: {error}"))
                })?
                .into_owned();
            if value.len() > 1_048_576 {
                return Err(Error::InvalidFormat(
                    "ODP animation attribute exceeds 1 MiB".to_string(),
                ));
            }
            attributes.push(Attribute::from_parsed(namespace, local_name, value)?);
        }
        Ok(attributes)
    }

    pub(super) fn exact_geometry_attributes(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Vec<DrawingAttribute>> {
        let mut attributes = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| {
                Error::InvalidFormat(format!("invalid enhanced-geometry attribute: {error}"))
            })?;
            let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
            let namespace = if Self::is_namespace(&namespace, DRAW_NAMESPACE) {
                DrawingAttributeNamespace::Drawing
            } else if Self::is_namespace(&namespace, SVG_NAMESPACE) {
                DrawingAttributeNamespace::Svg
            } else if Self::is_namespace(&namespace, DR3D_NAMESPACE) {
                DrawingAttributeNamespace::Dr3d
            } else {
                continue;
            };
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!(
                        "invalid enhanced-geometry attribute value: {error}"
                    ))
                })?
                .into_owned();
            attributes.push(DrawingAttribute::new(
                namespace,
                String::from_utf8(local_name.as_ref().to_vec()).map_err(|_| {
                    Error::InvalidFormat("non-UTF-8 enhanced-geometry attribute name".to_string())
                })?,
                value,
            )?);
        }
        Ok(attributes)
    }

    /// Helper to extract attribute values
    pub(super) fn get_attr(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        namespace_uri: &[u8],
        local_name: &[u8],
    ) -> Result<Option<String>> {
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            if Self::is_namespace(&namespace, namespace_uri) && local.as_ref() == local_name {
                return attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map(|value| Some(value.into_owned()))
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid XML attribute value: {error}"))
                    });
            }
        }
        Ok(None)
    }
}
