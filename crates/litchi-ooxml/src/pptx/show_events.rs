//! Temporary host boundary for the canonical inert slide-event owner.

use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, PackURI};

pub use litchi_pptx::presentation_properties::metadata::events::{
    Draft as EventDraft, Event, Kind as EventKind, Trigger,
};

pub const SHOW_EVENT_EXTENSION_URI: &str =
    litchi_pptx::presentation_properties::metadata::events::EXTENSION_URI;

pub fn store_slide_show_events(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    events: &[EventDraft],
) -> Result<()> {
    litchi_pptx::presentation_properties::metadata::events::store(package, slide_name, events)
        .map_err(OoxmlError::from)
}
