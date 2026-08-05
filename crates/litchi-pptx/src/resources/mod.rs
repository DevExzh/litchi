//! Built-in PresentationML producer templates.
//!
//! These assets are deliberately owned by the format crate. Package
//! orchestration can therefore construct a valid document without depending
//! on the former OOXML umbrella or duplicating large XML literals in writer
//! code.

/// The default presentation part used for a newly authored package.
pub(crate) const PRESENTATION: &str = include_str!("generated/presentation.xml");

/// The default slide-master part used for a newly authored package.
pub(crate) const SLIDE_MASTER: &str = include_str!("generated/slideMasters/slideMaster1.xml");

/// The first default slide-layout part used for a newly authored package.
pub(crate) const SLIDE_LAYOUT: &str = include_str!("generated/slideLayouts/slideLayout1.xml");

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
