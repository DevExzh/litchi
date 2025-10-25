/// EMF to SVG Conversion
///
/// High-performance SVG generation with SVGO-like optimization
mod buffer;
pub mod converter;
pub mod path;
pub mod state;

pub use converter::EmfSvgConverter;
pub use path::PathBuilder;
pub use state::{DeviceContext, RenderState};
