//! Inert custom-shape enhanced geometry declarations.

use litchi_core::{Error, Result};

const MAX_ATTRIBUTES: usize = 65_536;
const MAX_VALUE_BYTES: usize = 16 * 1024 * 1024;
const MAX_GEOMETRY_BYTES: usize = 8 * 1024 * 1024;

/// Namespace of a recognized enhanced-geometry attribute.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum DrawingAttributeNamespace {
    /// `draw:*`.
    Drawing,
    /// `svg:*`.
    Svg,
    /// `dr3d:*`.
    Dr3d,
}

/// A recognized drawing attribute retained as a bounded lexical value.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DrawingAttribute {
    namespace: DrawingAttributeNamespace,
    local_name: String,
    value: String,
}

impl DrawingAttribute {
    pub(crate) fn parsed(
        namespace: DrawingAttributeNamespace,
        local_name: String,
        value: String,
    ) -> Result<Self> {
        if local_name.is_empty()
            || local_name.len() > 256
            || !local_name.bytes().enumerate().all(|(index, byte)| {
                byte.is_ascii_alphanumeric() || byte == b'_' || (index > 0 && byte == b'-')
            })
        {
            return Err(Error::InvalidFormat(
                "invalid ODG enhanced-geometry attribute local name".into(),
            ));
        }
        if value.len() > MAX_VALUE_BYTES || value.contains('\0') {
            return Err(Error::InvalidFormat(
                "ODG enhanced-geometry attribute exceeds the limit".into(),
            ));
        }
        Ok(Self {
            namespace,
            local_name,
            value,
        })
    }

    /// Attribute namespace.
    #[must_use]
    pub const fn namespace(&self) -> DrawingAttributeNamespace {
        self.namespace
    }

    /// Attribute local name.
    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    /// Decoded attribute value.
    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Child owner kind in `draw:enhanced-geometry`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EnhancedGeometryChildKind {
    /// `draw:equation` formula declaration.
    Equation,
    /// `draw:handle` adjustment handle.
    Handle,
}

/// One inert equation or handle declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EnhancedGeometryChild {
    kind: EnhancedGeometryChildKind,
    attributes: Vec<DrawingAttribute>,
}

impl EnhancedGeometryChild {
    pub(crate) fn parsed(
        kind: EnhancedGeometryChildKind,
        attributes: Vec<DrawingAttribute>,
    ) -> Self {
        Self { kind, attributes }
    }

    /// Child owner kind.
    #[must_use]
    pub const fn kind(&self) -> EnhancedGeometryChildKind {
        self.kind
    }

    /// Recognized attributes in source order.
    #[must_use]
    pub fn attributes(&self) -> &[DrawingAttribute] {
        &self.attributes
    }
}

/// Inert enhanced geometry attached to a custom shape.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct EnhancedGeometry {
    attributes: Vec<DrawingAttribute>,
    children: Vec<EnhancedGeometryChild>,
}

impl EnhancedGeometry {
    pub(crate) fn parsed(
        attributes: Vec<DrawingAttribute>,
        children: Vec<EnhancedGeometryChild>,
    ) -> Result<Self> {
        if attributes.len().saturating_add(children.len()) > MAX_ATTRIBUTES {
            return Err(Error::InvalidFormat(
                "ODG enhanced-geometry aggregate exceeds the limit".into(),
            ));
        }
        let bytes = attributes
            .iter()
            .map(|attribute| {
                attribute
                    .local_name
                    .len()
                    .saturating_add(attribute.value.len())
            })
            .chain(children.iter().flat_map(|child| {
                child.attributes.iter().map(|attribute| {
                    attribute
                        .local_name
                        .len()
                        .saturating_add(attribute.value.len())
                })
            }))
            .try_fold(0usize, |total, value| total.checked_add(value))
            .ok_or_else(|| Error::InvalidFormat("ODG enhanced-geometry size overflow".into()))?;
        if bytes > MAX_GEOMETRY_BYTES {
            return Err(Error::InvalidFormat(
                "ODG enhanced-geometry aggregate exceeds the byte limit".into(),
            ));
        }
        let mut handles_seen = false;
        for child in &children {
            match child.kind {
                EnhancedGeometryChildKind::Equation => {
                    if handles_seen {
                        return Err(Error::InvalidFormat(
                            "ODG enhanced-geometry equations must precede handles".into(),
                        ));
                    }
                },
                EnhancedGeometryChildKind::Handle => {
                    handles_seen = true;
                    validate_handle_attributes(&child.attributes)?;
                },
            }
        }
        Ok(Self {
            attributes,
            children,
        })
    }

    /// Recognized geometry attributes in source order.
    #[must_use]
    pub fn attributes(&self) -> &[DrawingAttribute] {
        &self.attributes
    }

    /// Equations and handles in source order.
    #[must_use]
    pub fn children(&self) -> &[EnhancedGeometryChild] {
        &self.children
    }
}

fn validate_handle_attributes(attributes: &[DrawingAttribute]) -> Result<()> {
    let has = |name: &str| {
        attributes.iter().any(|attribute| {
            attribute.namespace == DrawingAttributeNamespace::Drawing
                && attribute.local_name == name
        })
    };
    for name in [
        "handle-mirror-horizontal",
        "handle-mirror-vertical",
        "handle-switched",
    ] {
        if has(name)
            && !attributes
                .iter()
                .filter(|attribute| {
                    attribute.namespace == DrawingAttributeNamespace::Drawing
                        && attribute.local_name == name
                })
                .all(|attribute| matches!(attribute.value.as_str(), "true" | "false"))
        {
            return Err(Error::InvalidFormat(
                "ODG enhanced-geometry handle Boolean attribute is invalid".into(),
            ));
        }
    }

    let has_xy_x = has("handle-position-x");
    let has_xy_y = has("handle-position-y");
    let has_polar_x = has("handle-polar-pole-x");
    let has_polar_y = has("handle-polar-pole-y");
    let xy_complete = has_xy_x && has_xy_y;
    let polar_complete = has_polar_x && has_polar_y;
    let xy_partial = has_xy_x != has_xy_y;
    let polar_partial = has_polar_x != has_polar_y;
    if xy_partial || polar_partial || (xy_complete && polar_complete) {
        return Err(Error::InvalidFormat(
            "ODG enhanced-geometry handle position pairs are invalid".into(),
        ));
    }

    // ODF 1.4 retains the deprecated legacy attributes for compatibility with
    // native LibreOffice documents. If no modern pair is present, one of those
    // legacy forms is the only accepted fallback.
    if !xy_complete && !polar_complete && !has("handle-position") && !has("handle-polar") {
        return Err(Error::InvalidFormat(
            "ODG enhanced-geometry handle has no position pair".into(),
        ));
    }
    Ok(())
}
