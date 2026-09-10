//! Inert drawing child owners: image maps, image contours, and glue points.

use super::enhanced::DrawingAttribute;
use litchi_core::{Error, Result};

const MAX_VALUE_BYTES: usize = 1 << 20;
const MAX_SOURCE_BYTES: usize = 8 * 1024 * 1024;

/// Geometry of one `draw:image-map` area.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ImageMapAreaShape {
    /// `draw:area-rectangle`.
    Rectangle {
        /// `svg:x`.
        x: String,
        /// `svg:y`.
        y: String,
        /// `svg:width`.
        width: String,
        /// `svg:height`.
        height: String,
    },
    /// `draw:area-circle`.
    Circle {
        /// `svg:cx`.
        cx: String,
        /// `svg:cy`.
        cy: String,
        /// `svg:r`.
        r: String,
    },
    /// `draw:area-polygon`.
    Polygon {
        /// `svg:x`.
        x: String,
        /// `svg:y`.
        y: String,
        /// `svg:width`.
        width: String,
        /// `svg:height`.
        height: String,
        /// `svg:viewBox`.
        view_box: String,
        /// `draw:points`.
        points: String,
    },
}

/// One inert image-map area and its link metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ImageMapArea {
    shape: ImageMapAreaShape,
    href: Option<String>,
    target_frame_name: Option<String>,
    show: Option<String>,
    no_href: bool,
    name: Option<String>,
    source_xml: String,
}

impl ImageMapArea {
    pub(crate) fn parsed(
        shape: ImageMapAreaShape,
        href: Option<String>,
        target_frame_name: Option<String>,
        show: Option<String>,
        no_href: bool,
        name: Option<String>,
        source_xml: String,
    ) -> Result<Self> {
        match &shape {
            ImageMapAreaShape::Rectangle {
                x,
                y,
                width,
                height,
            } => {
                for value in [x, y, width, height] {
                    bounded(value, "ODG image-map rectangle value")?;
                }
            },
            ImageMapAreaShape::Circle { cx, cy, r } => {
                for value in [cx, cy, r] {
                    bounded(value, "ODG image-map circle value")?;
                }
            },
            ImageMapAreaShape::Polygon {
                x,
                y,
                width,
                height,
                view_box,
                points,
            } => {
                for value in [x, y, width, height, view_box, points] {
                    bounded(value, "ODG image-map polygon value")?;
                }
            },
        }
        for value in [
            href.as_deref(),
            target_frame_name.as_deref(),
            show.as_deref(),
            name.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            bounded_optional(value, "ODG image-map area value")?;
        }
        bounded_source(&source_xml, "ODG image-map area source")?;
        Ok(Self {
            shape,
            href,
            target_frame_name,
            show,
            no_href,
            name,
            source_xml,
        })
    }

    /// Area geometry.
    #[must_use]
    pub fn shape(&self) -> &ImageMapAreaShape {
        &self.shape
    }

    /// Optional inert `xlink:href` target.
    #[must_use]
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }

    /// Optional `office:target-frame-name`.
    #[must_use]
    pub fn target_frame_name(&self) -> Option<&str> {
        self.target_frame_name.as_deref()
    }

    /// Optional `xlink:show` value.
    #[must_use]
    pub fn show(&self) -> Option<&str> {
        self.show.as_deref()
    }

    /// Whether `draw:nohref` is present.
    #[must_use]
    pub const fn no_href(&self) -> bool {
        self.no_href
    }

    /// Optional `office:name`.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Exact source fragment for this area.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        &self.source_xml
    }
}

/// One inert `draw:image-map` owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ImageMap {
    areas: Vec<ImageMapArea>,
    source_xml: String,
}

impl ImageMap {
    pub(crate) fn parsed(areas: Vec<ImageMapArea>, source_xml: String) -> Result<Self> {
        if areas.len() > 65_536 {
            return Err(Error::InvalidFormat(
                "ODG image-map area count exceeds the limit".into(),
            ));
        }
        bounded_source(&source_xml, "ODG image-map source")?;
        let bytes = areas
            .iter()
            .map(|area| area.source_xml.len())
            .try_fold(source_xml.len(), |total, value| total.checked_add(value))
            .ok_or_else(|| Error::InvalidFormat("ODG image-map size overflow".into()))?;
        if bytes > MAX_SOURCE_BYTES {
            return Err(Error::InvalidFormat(
                "ODG image-map aggregate exceeds the byte limit".into(),
            ));
        }
        Ok(Self { areas, source_xml })
    }

    /// Areas in document order.
    #[must_use]
    pub fn areas(&self) -> &[ImageMapArea] {
        &self.areas
    }

    /// Exact source fragment for this image map.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        &self.source_xml
    }
}

/// Inert contour owner kind.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ContourKind {
    /// `draw:contour-polygon`.
    Polygon,
    /// `draw:contour-path`.
    Path,
}

/// One inert image contour with recognized drawing/SVG attributes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Contour {
    kind: ContourKind,
    attributes: Vec<DrawingAttribute>,
    source_xml: String,
}

impl Contour {
    pub(crate) fn parsed(
        kind: ContourKind,
        attributes: Vec<DrawingAttribute>,
        source_xml: String,
    ) -> Result<Self> {
        bounded_source(&source_xml, "ODG contour source")?;
        if attributes.len() > 65_536 {
            return Err(Error::InvalidFormat(
                "ODG contour attribute count exceeds the limit".into(),
            ));
        }
        let bytes = attributes
            .iter()
            .map(|attribute| {
                attribute
                    .local_name()
                    .len()
                    .saturating_add(attribute.value().len())
            })
            .try_fold(source_xml.len(), |total, value| total.checked_add(value))
            .ok_or_else(|| Error::InvalidFormat("ODG contour size overflow".into()))?;
        if bytes > MAX_SOURCE_BYTES {
            return Err(Error::InvalidFormat(
                "ODG contour aggregate exceeds the byte limit".into(),
            ));
        }
        Ok(Self {
            kind,
            attributes,
            source_xml,
        })
    }

    /// Contour kind.
    #[must_use]
    pub const fn kind(&self) -> ContourKind {
        self.kind
    }

    /// Recognized contour attributes in source order.
    #[must_use]
    pub fn attributes(&self) -> &[DrawingAttribute] {
        &self.attributes
    }

    /// Exact source fragment for this contour.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        &self.source_xml
    }
}

/// One inert `draw:glue-point` owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GluePoint {
    id: String,
    x: String,
    y: String,
    align: Option<String>,
    escape_direction: String,
    source_xml: String,
}

impl GluePoint {
    pub(crate) fn parsed(
        id: String,
        x: String,
        y: String,
        align: Option<String>,
        escape_direction: String,
        source_xml: String,
    ) -> Result<Self> {
        for value in [
            id.as_str(),
            x.as_str(),
            y.as_str(),
            escape_direction.as_str(),
        ] {
            bounded(value, "ODG glue-point value")?;
        }
        if let Some(align) = align.as_deref() {
            bounded(align, "ODG glue-point alignment")?;
        }
        bounded_source(&source_xml, "ODG glue-point source")?;
        Ok(Self {
            id,
            x,
            y,
            align,
            escape_direction,
            source_xml,
        })
    }

    /// `draw:id` lexical value.
    #[must_use]
    pub fn id(&self) -> &str {
        &self.id
    }

    /// `svg:x` lexical value.
    #[must_use]
    pub fn x(&self) -> &str {
        &self.x
    }

    /// `svg:y` lexical value.
    #[must_use]
    pub fn y(&self) -> &str {
        &self.y
    }

    /// Optional `draw:align` value.
    #[must_use]
    pub fn align(&self) -> Option<&str> {
        self.align.as_deref()
    }

    /// `draw:escape-direction` value.
    #[must_use]
    pub fn escape_direction(&self) -> &str {
        &self.escape_direction
    }

    /// Exact source fragment for this glue point.
    #[must_use]
    pub fn source_xml(&self) -> &str {
        &self.source_xml
    }
}

fn bounded(value: &str, owner: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_VALUE_BYTES || value.contains('\0') {
        return Err(Error::InvalidFormat(format!("{owner} exceeds the limit")));
    }
    Ok(())
}

fn bounded_optional(value: &str, owner: &str) -> Result<()> {
    if value.len() > MAX_VALUE_BYTES || value.contains('\0') {
        return Err(Error::InvalidFormat(format!("{owner} exceeds the limit")));
    }
    Ok(())
}

fn bounded_source(value: &str, owner: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_SOURCE_BYTES {
        return Err(Error::InvalidFormat(format!("{owner} exceeds the limit")));
    }
    Ok(())
}
