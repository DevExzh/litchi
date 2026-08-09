//! Built-in `PresentationML` producer templates.
//!
//! These assets are deliberately owned by the format crate. Package
//! orchestration can therefore construct a valid document without depending
//! on the former OOXML umbrella or duplicating large XML literals in writer
//! code.

/// The default presentation part used for a newly authored package.
pub(crate) const PRESENTATION: &str = include_str!("generated/presentation.xml");

/// The default slide-master part used for a newly authored package.
pub(crate) const SLIDE_MASTER: &str = include_str!("generated/slideMasters/slideMaster1.xml");

/// The default slide-layout parts used for a newly authored package.
///
/// The built-in slide master carries the complete eleven-entry
/// `p:sldLayoutIdLst`, so package construction must materialize the same
/// relationship graph rather than retaining only the first layout.
pub(crate) const SLIDE_LAYOUTS: [&str; 11] = [
    include_str!("generated/slideLayouts/slideLayout1.xml"),
    include_str!("generated/slideLayouts/slideLayout2.xml"),
    include_str!("generated/slideLayouts/slideLayout3.xml"),
    include_str!("generated/slideLayouts/slideLayout4.xml"),
    include_str!("generated/slideLayouts/slideLayout5.xml"),
    include_str!("generated/slideLayouts/slideLayout6.xml"),
    include_str!("generated/slideLayouts/slideLayout7.xml"),
    include_str!("generated/slideLayouts/slideLayout8.xml"),
    include_str!("generated/slideLayouts/slideLayout9.xml"),
    include_str!("generated/slideLayouts/slideLayout10.xml"),
    include_str!("generated/slideLayouts/slideLayout11.xml"),
];

/// The default theme part used for a newly authored package.
pub(crate) const THEME: &str = include_str!("generated/theme/theme1.xml");

/// The default presentation-view properties part.
pub(crate) const VIEW_PROPERTIES: &str = include_str!("generated/viewProps.xml");

/// The default presentation properties part.
pub(crate) const PRESENTATION_PROPERTIES: &str = include_str!("generated/presProps.xml");

/// The default core-properties part.
pub(crate) const CORE_PROPERTIES: &str = include_str!("generated/docProps/core.xml");

/// The default extended-properties part.
pub(crate) const EXTENDED_PROPERTIES: &str = include_str!("generated/docProps/app.xml");
