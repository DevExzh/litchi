//! Shared markup-compatibility preprocessing for OOXML parts.

pub mod alternative;
pub mod stream;

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::{
    active_offsets, process_markup_compatibility, process_ooxml, process_part, process_part_arc,
    process_str,
};
pub use model::{Capabilities, Error, Limits, NAMESPACE, Name, OffsetLimits, Output, Report};
pub use stream::{
    EventLimitExceeded, InputLimitExceeded, RawAttribute, RawElement, RawElementKind,
    SemanticAttribute, SemanticDecl, SemanticElement, SemanticEnd, SemanticEvent,
    SemanticGeneralRef, SemanticText, StreamError, StreamLimits, StreamReport, XMLNS_NAMESPACE,
    process_markup_compatibility_stream, process_markup_compatibility_stream_with_active_observer,
    process_markup_compatibility_stream_with_observers,
};
