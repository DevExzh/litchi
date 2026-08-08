//! Named drawing resources stored in ODF common styles.
//!
//! Each resource family owns its typed model, bounded XML codec, and focused
//! regression coverage. Package and flat-document accessors live in `style`.

pub mod fill_image {
    pub use litchi_odf_common::drawing::resources::fill_image::*;
}
pub mod gradient {
    pub use litchi_odf_common::drawing::resources::gradient::*;
}
pub mod hatch {
    pub use litchi_odf_common::drawing::resources::hatch::*;
}
pub mod marker {
    pub use litchi_odf_common::drawing::resources::marker::*;
}
pub mod opacity {
    pub use litchi_odf_common::drawing::resources::opacity::*;
}
pub mod stroke_dash {
    pub use litchi_odf_common::drawing::resources::stroke_dash::*;
}
pub mod style;

// Resource types stay under their family modules so contextual names such as
// `Length`, `Style`, and `Collection` remain unambiguous.
