//! Compatibility re-exports for PPTX slide backgrounds.
//!
//! The package-independent semantic model and XML codec live in
//! the canonical litchi_pptx::backgrounds module. Package relationship
//! resolution remains in the OOXML host adapters.

pub use litchi_pptx::backgrounds::{
    GradientStop, GradientType, PatternType, PictureStyle, SlideBackground,
};
