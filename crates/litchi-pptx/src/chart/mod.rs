//! Layered `DrawingML` chart semantics.
//!
//! `model` contains contextual authoring and discovery values, `codec` owns
//! bounded chart XML parsing and writing, and `package` owns chart-part
//! relationships. The extension and style submodules own `ChartEx` companion
//! graphs without leaking their physical package prefixes into the API.

pub mod codec;
pub mod extension;
pub mod model;
pub mod package;
pub mod style;

pub use codec::{encode, write_graphic_frame};
pub use model::{Chart, Info, Series, Type};
pub use package::{Part, add, load, related};
