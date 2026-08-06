//! Contextual BIFF writers for MS-XLS PivotTables and pivot caches.
//!
//! The facade keeps the existing `writer::biff` call surface while separating
//! typed configurations, wire emission, validation, and focused tests.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub(crate) use codec::{
    generate_pivot_cache_stream, write_dconref, write_mso_drawing_group,
    write_pivot_modern_extensions, write_sx_stream_id, write_sxdi, write_sxex, write_sxivd,
    write_sxli, write_sxpi, write_sxvd, write_sxvdex, write_sxvi, write_sxview, write_sxvs,
};
pub(crate) use model::{
    PivotCacheFieldInfo, PivotCacheSourceRow, PivotCacheStreamInfo, SxDiConfig, SxExConfig,
    SxVdConfig, SxViConfig, SxViewConfig,
};

// The drawing owner is a sibling of this owner under `biff`; retain the
// existing restricted internal entry point without widening its visibility.
pub(super) use codec::write_pivot_page_obj;
