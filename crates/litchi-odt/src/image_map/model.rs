/// The geometry of one clickable image-map area.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ImageMapAreaShape {
    /// `draw:area-rectangle` with `svg:x`/`svg:y`/`svg:width`/`svg:height`.
    Rectangle {
        /// Left edge coordinate.
        x: String,
        /// Top edge coordinate.
        y: String,
        /// Area width.
        width: String,
        /// Area height.
        height: String,
    },
    /// `draw:area-circle` with `svg:cx`/`svg:cy`/`svg:r`.
    Circle {
        /// Center x coordinate.
        cx: String,
        /// Center y coordinate.
        cy: String,
        /// Radius.
        r: String,
    },
    /// `draw:area-polygon` with extents, view box, and point list.
    Polygon {
        /// Left edge coordinate of the extent.
        x: String,
        /// Top edge coordinate of the extent.
        y: String,
        /// Extent width.
        width: String,
        /// Extent height.
        height: String,
        /// `svg:viewBox` of the polygon coordinate space.
        view_box: String,
        /// `svg:points` vertex list.
        points: String,
    },
}

/// One clickable area of an image map, with inert link metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageMapArea {
    /// The area geometry.
    pub shape: ImageMapAreaShape,
    /// `xlink:href` link target, stored verbatim and never resolved.
    pub href: Option<String>,
    /// `office:target-frame-name` link frame target.
    pub target_frame_name: Option<String>,
    /// `xlink:show` presentation hint (`new` or `replace`).
    pub show: Option<String>,
    /// `draw:nohref`: the area has no link target.
    pub no_href: bool,
    /// `office:name` of the area.
    pub name: Option<String>,
    /// Exact `svg:title` child XML, when present.
    pub title_xml: Option<String>,
    /// Exact `svg:desc` child XML, when present.
    pub description_xml: Option<String>,
    /// Exact `office:event-listeners` child XML, preserved without
    /// interpretation.
    pub event_listeners_xml: Option<String>,
    /// Exact area element XML.
    pub xml: String,
}

/// A `draw:image-map` element and its clickable areas.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageMap {
    /// The clickable areas, in document order.
    pub areas: Vec<ImageMapArea>,
    /// Exact `draw:image-map` element XML.
    pub xml: String,
}
