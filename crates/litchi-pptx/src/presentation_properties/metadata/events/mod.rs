//! Slide-show event values and bounded extension storage.

mod codec;
mod model;
mod package;

pub use codec::EXTENSION_URI;
pub use model::{Draft, Event, Kind, Trigger};
pub use package::store;
