//! Shared markup-compatibility preprocessing for OOXML parts.

pub mod alternative;
// The stream reads attributes with quick-xml's duplicate check off, under its
// per-event attribute limit, and checks expanded names in a keyed set. Record
// 0771 is rewriting that file, so record 0770 allows it here.
#[allow(clippy::disallowed_methods)]
pub mod stream;

mod codec;
mod fragment;
mod model;
mod patterns;
mod scope;

#[cfg(test)]
mod bounds_tests;
#[cfg(test)]
mod expanded_duplicates_tests;
#[cfg(test)]
mod shared_names_tests;
#[cfg(test)]
mod tests;

pub use codec::{
    active_offsets, process_markup_compatibility, process_ooxml, process_part, process_part_arc,
    process_str,
};
pub use fragment::{InScopeNamespaces, self_contained_fragment};
pub use model::{
    ATTRIBUTES_PER_ELEMENT_CEILING, Capabilities, DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT, Error,
    ExpandedName, Limits, NAMESPACE, Name, NamespaceUri, OffsetLimits, Output, Report,
};
pub use stream::{
    ActiveFlow, EventLimitExceeded, InputLimitExceeded, RawAttribute, RawElement, RawElementKind,
    SemanticAttribute, SemanticDecl, SemanticElement, SemanticEnd, SemanticEvent,
    SemanticGeneralRef, SemanticText, StreamError, StreamLimits, StreamReport, XMLNS_NAMESPACE,
    process_markup_compatibility_stream, process_markup_compatibility_stream_with_active_observer,
    process_markup_compatibility_stream_with_observers,
    process_markup_compatibility_stream_with_stoppable_observers,
};
