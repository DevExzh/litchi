//! Read-only projection of ODF drawing frames and text boxes.

/// The inert payload kind carried by a drawing frame.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    TextBox,
    Image,
    Object,
    OleObject,
    Plugin,
    FloatingFrame,
    Generic,
}

/// A projected frame or text box. Geometry remains lexical and inert.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Frame {
    anchor_type: Option<String>,
    height: Option<String>,
    href: Option<String>,
    kind: Kind,
    name: Option<String>,
    style_name: Option<String>,
    text: String,
    width: Option<String>,
    x: Option<String>,
    y: Option<String>,
}

impl Frame {
    #[allow(
        clippy::too_many_arguments,
        reason = "projection fields mirror ODF frame attributes"
    )]
    pub(crate) const fn projected(
        kind: Kind,
        name: Option<String>,
        style_name: Option<String>,
        anchor_type: Option<String>,
        x: Option<String>,
        y: Option<String>,
        width: Option<String>,
        height: Option<String>,
        href: Option<String>,
        text: String,
    ) -> Self {
        Self {
            anchor_type,
            height,
            href,
            kind,
            name,
            style_name,
            text,
            width,
            x,
            y,
        }
    }

    /// Frame payload kind.
    #[must_use]
    pub const fn kind(&self) -> &Kind {
        &self.kind
    }

    /// Producer-visible frame name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Frame style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Inert ODF anchoring value.
    #[must_use]
    pub fn anchor_type(&self) -> Option<&str> {
        self.anchor_type.as_deref()
    }

    /// Lexical SVG x coordinate.
    #[must_use]
    pub fn x(&self) -> Option<&str> {
        self.x.as_deref()
    }

    /// Lexical SVG y coordinate.
    #[must_use]
    pub fn y(&self) -> Option<&str> {
        self.y.as_deref()
    }

    /// Lexical SVG width.
    #[must_use]
    pub fn width(&self) -> Option<&str> {
        self.width.as_deref()
    }

    /// Lexical SVG height.
    #[must_use]
    pub fn height(&self) -> Option<&str> {
        self.height.as_deref()
    }

    /// Inert xlink target, if present.
    #[must_use]
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }

    /// Visible text nested in a text box or caption.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}
