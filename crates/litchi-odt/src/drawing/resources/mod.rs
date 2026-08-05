//! Named drawing resources stored in ODF common styles.
//!
//! Each resource family owns its typed model, bounded XML codec, and focused
//! regression coverage. Package and flat-document accessors live in `style`.

pub mod fill_image;
pub mod gradient;
pub mod hatch;
pub mod marker;
pub mod opacity;
pub mod stroke_dash;
pub mod style;

// Resource types stay under their family modules so contextual names such as
// `Length`, `Style`, and `Collection` remain unambiguous.
