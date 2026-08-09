//! Inert ODF client-side image-map semantics.
/// One client-side image map attached directly to the document frame.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ImageMap {
    areas: Vec<Area>,
}

impl ImageMap {
    /// Creates a map from areas in document order.
    #[must_use]
    pub fn new(areas: Vec<Area>) -> Self {
        Self { areas }
    }

    /// Returns map areas in document order.
    #[must_use]
    pub fn areas(&self) -> &[Area] {
        &self.areas
    }
}

/// One inert link region. URI values are never resolved or followed.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Area {
    kind: AreaKind,
    href: Option<String>,
    target_frame_name: Option<String>,
    name: Option<String>,
    no_href: bool,
    link_type: Option<String>,
    show: Option<String>,
    title: Option<String>,
    description: Option<String>,
}

impl Area {
    pub(crate) fn new(kind: AreaKind, properties: AreaProperties) -> Self {
        Self {
            kind,
            href: properties.href,
            target_frame_name: properties.target_frame_name,
            name: properties.name,
            no_href: properties.no_href,
            link_type: properties.link_type,
            show: properties.show,
            title: properties.title,
            description: properties.description,
        }
    }

    /// Creates a rectangular area without a link target.
    #[must_use]
    pub fn rectangle(
        x: impl Into<String>,
        y: impl Into<String>,
        width: impl Into<String>,
        height: impl Into<String>,
    ) -> Self {
        Self::new(
            AreaKind::Rectangle {
                x: x.into(),
                y: y.into(),
                width: width.into(),
                height: height.into(),
            },
            AreaProperties::default(),
        )
    }

    /// Creates a circular area without a link target.
    #[must_use]
    pub fn circle(
        center_x: impl Into<String>,
        center_y: impl Into<String>,
        radius: impl Into<String>,
    ) -> Self {
        Self::new(
            AreaKind::Circle {
                center_x: center_x.into(),
                center_y: center_y.into(),
                radius: radius.into(),
            },
            AreaProperties::default(),
        )
    }

    /// Creates a polygon area without a link target.
    #[must_use]
    pub fn polygon(
        geometry: [impl Into<String>; 4],
        view_box: impl Into<String>,
        points: impl Into<String>,
    ) -> Self {
        let [x, y, width, height] = geometry.map(Into::into);
        Self::new(
            AreaKind::Polygon {
                x,
                y,
                width,
                height,
                view_box: view_box.into(),
                points: points.into(),
            },
            AreaProperties::default(),
        )
    }

    /// Sets an inert link target.
    #[must_use]
    pub fn with_href(mut self, href: impl Into<String>) -> Self {
        self.href = Some(href.into());
        self.no_href = false;
        self.link_type = Some("simple".to_owned());
        self
    }

    /// Marks this region as explicitly unlinked.
    #[must_use]
    pub fn with_no_href(mut self) -> Self {
        self.href = None;
        self.no_href = true;
        self.link_type = None;
        self.show = None;
        self
    }

    /// Sets the human-readable area name.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Sets the short accessible title.
    #[must_use]
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Sets the accessible description.
    #[must_use]
    pub fn with_description(mut self, description: impl Into<String>) -> Self {
        self.description = Some(description.into());
        self
    }

    /// Returns the geometric region.
    #[must_use]
    pub const fn kind(&self) -> &AreaKind {
        &self.kind
    }

    /// Returns the inert target URI, if present.
    #[must_use]
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }

    /// Returns the target frame name, if present.
    #[must_use]
    pub fn target_frame_name(&self) -> Option<&str> {
        self.target_frame_name.as_deref()
    }

    /// Returns the area name, if present.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns whether the area explicitly has no link target.
    #[must_use]
    pub const fn has_no_href(&self) -> bool {
        self.no_href
    }

    /// Returns the `XLink` type, if present.
    #[must_use]
    pub fn link_type(&self) -> Option<&str> {
        self.link_type.as_deref()
    }

    /// Returns the `XLink` presentation behavior, if present.
    #[must_use]
    pub fn show(&self) -> Option<&str> {
        self.show.as_deref()
    }

    /// Returns the short accessible title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Returns the accessible description.
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }
}

/// ODF 1.4 image-map geometry retained lexically.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum AreaKind {
    /// A rectangular region.
    Rectangle {
        x: String,
        y: String,
        width: String,
        height: String,
    },
    /// A circular region.
    Circle {
        center_x: String,
        center_y: String,
        radius: String,
    },
    /// A polygon region inside a declared viewport.
    Polygon {
        x: String,
        y: String,
        width: String,
        height: String,
        view_box: String,
        points: String,
    },
}

#[derive(Clone, Debug, Default)]
pub(crate) struct AreaProperties {
    pub(crate) href: Option<String>,
    pub(crate) target_frame_name: Option<String>,
    pub(crate) name: Option<String>,
    pub(crate) no_href: bool,
    pub(crate) link_type: Option<String>,
    pub(crate) show: Option<String>,
    pub(crate) title: Option<String>,
    pub(crate) description: Option<String>,
}
