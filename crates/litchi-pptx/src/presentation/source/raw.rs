//! Explicit access to the immutable presentation root captured at open.
//!
//! These accessors expose the retained snapshot, without reading the source.
//! Semantic reads on the source owner continue to check source and execution
//! state. Native slide identities are available in [`super::related::raw`].

use litchi_opc::PackURI;

use super::SourceBackedPresentation;

/// The native URI of the retained main presentation part.
pub fn part_uri(presentation: &SourceBackedPresentation) -> &PackURI {
    presentation.inner._presentation.source_partname()
}

/// Original presentation XML captured when the source catalog was opened.
pub fn data(presentation: &SourceBackedPresentation) -> &[u8] {
    presentation.inner._presentation.source_data().as_bytes()
}
