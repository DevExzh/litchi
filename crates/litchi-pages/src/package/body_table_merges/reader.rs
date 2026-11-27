//! Lazy, metadata-only Pages body-table merge reader.

use std::path::Path;
use std::sync::Arc;

use litchi_iwa_archive::{ComponentCatalog, Limits, SourceCatalog};
use litchi_numbers_wire::table_merges;

use super::super::table_lock;
use super::{
    Error, map_lock_error, map_merge_error, map_package_error, model_message_from_components,
    read_limits, select_reader_target,
};
use crate::OpenOptions;
use crate::selector::BodyTableSelector;
use crate::table::merge::Region;

/// A lazy, metadata-only reader for rooted Pages body-table merges.
///
/// `MergeReader` retains the parsed physical component catalog and performs
/// the rooted body/table ownership proof only when a selector is queried. It
/// does not construct [`crate::Document`] or materialize section values. The
/// complete body storage tree is validated, then the bounded placeholder text
/// needed to validate table attachment offsets is read as borrowed wire
/// fragments. No semantic body `String` or run vector is materialized. Native
/// object identifiers and generated protobuf values remain private. Cloning a
/// reader shares the immutable parsed archives and the compact reader state.
///
/// Use [`super::Package::body_table_merges`] when a fully parsed
/// [`super::Package`] is already available. `MergeReader` is useful when an
/// application needs only merged-cell geometry and wants to avoid the
/// package's full semantic document projection.
///
/// ```no_run
/// use litchi_pages::{BodyTableSelector, MergeReader};
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let reader = MergeReader::open("budget.pages")?;
/// let regions = reader.body_table_merges(BodyTableSelector::name("Revenue"))?;
/// for region in regions {
///     println!("{region:?}");
/// }
/// # Ok(())
/// # }
/// ```
pub struct MergeReader {
    state: Arc<MergeReaderState>,
}

struct MergeReaderState {
    components: Arc<ComponentCatalog>,
    limits: Limits,
}

impl std::fmt::Debug for MergeReader {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("MergeReader")
            .field("limits", &self.state.limits)
            .finish_non_exhaustive()
    }
}

impl Clone for MergeReader {
    fn clone(&self) -> Self {
        Self {
            state: Arc::clone(&self.state),
        }
    }
}

impl MergeReader {
    /// Open a Pages package for lazy merged-cell queries with default limits.
    ///
    /// Filesystem ingress is bounded and descriptor-first, then the parsed
    /// component catalog is retained without constructing a semantic
    /// document.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the path cannot be read, physical ingress
    /// exceeds its limits, or the source is not a valid Pages package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self, Error> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open a Pages package for lazy merged-cell queries under the physical
    /// source limits of `options`.
    ///
    /// # Errors
    ///
    /// Returns the same typed failures as [`Self::from_bytes_with_options`],
    /// plus bounded filesystem ingress failures.
    pub fn open_with_options(path: impl AsRef<Path>, options: OpenOptions) -> Result<Self, Error> {
        let limits = archive_limits(options)?;
        let bytes = super::super::read_path(path.as_ref(), limits).map_err(map_package_error)?;
        Self::from_shared_bytes_with_archive_limits(bytes, limits)
    }

    /// Parse a Pages package for lazy merged-cell queries with default limits.
    ///
    /// Unlike [`super::Package::from_bytes`], this constructor stops after
    /// physical component validation and Pages root qualification; it does
    /// not build a semantic document or materialize section values.
    ///
    /// # Errors
    ///
    /// Returns a typed error when physical ingress, Pages root qualification,
    /// or the component object inventory is malformed.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, Error> {
        Self::from_bytes_with_options(bytes, OpenOptions::default())
    }

    /// Parse a Pages package for lazy merged-cell queries under the physical
    /// source limits of `options`.
    ///
    /// # Errors
    ///
    /// Returns a typed error when physical ingress, Pages root qualification,
    /// or the component object inventory is malformed.
    pub fn from_bytes_with_options(bytes: &[u8], options: OpenOptions) -> Result<Self, Error> {
        let limits = archive_limits(options)?;
        let source = SourceCatalog::from_bytes(bytes, limits)
            .map_err(|error| map_lock_error(table_lock::map_archive_error(error)))?;
        Self::from_components_with_limits(source.into_components(), limits)
    }

    /// Parse an already-owned immutable Pages package without copying its
    /// source allocation.
    ///
    /// The package bytes are consumed by physical ingress and released after
    /// the component catalog is built; subsequent queries retain only parsed
    /// IWA archives and reader metadata.
    ///
    /// # Errors
    ///
    /// Returns the same typed failures as [`Self::from_bytes_with_options`].
    pub fn from_shared_bytes(bytes: Arc<[u8]>) -> Result<Self, Error> {
        Self::from_shared_bytes_with_options(bytes, OpenOptions::default())
    }

    /// Parse an already-owned immutable Pages package under the physical
    /// source limits of `options`.
    ///
    /// # Errors
    ///
    /// Returns the same typed failures as [`Self::from_bytes_with_options`].
    pub fn from_shared_bytes_with_options(
        bytes: Arc<[u8]>,
        options: OpenOptions,
    ) -> Result<Self, Error> {
        Self::from_shared_bytes_with_archive_limits(bytes, archive_limits(options)?)
    }

    fn from_shared_bytes_with_archive_limits(
        bytes: Arc<[u8]>,
        limits: Limits,
    ) -> Result<Self, Error> {
        let source = SourceCatalog::from_shared_bytes(bytes, limits)
            .map_err(|error| map_lock_error(table_lock::map_archive_error(error)))?;
        Self::from_components_with_limits(source.into_components(), limits)
    }

    fn from_components_with_limits(
        components: ComponentCatalog,
        limits: Limits,
    ) -> Result<Self, Error> {
        Self::from_shared_components_with_limits(Arc::new(components), limits)
    }

    fn from_shared_components_with_limits(
        components: Arc<ComponentCatalog>,
        limits: Limits,
    ) -> Result<Self, Error> {
        limits
            .effective_archive_limits()
            .map_err(|error| map_lock_error(table_lock::map_archive_error(error)))?;
        super::super::validate_components(components.as_ref()).map_err(map_package_error)?;
        // Qualify the Pages root at construction without projecting the body,
        // section, or table semantics. A body-less Pages root remains a valid
        // reader and reports TableNotFound when queried.
        super::super::root_references_with_limits(components.as_ref(), limits)
            .map_err(map_package_error)?;
        Ok(Self {
            state: Arc::new(MergeReaderState { components, limits }),
        })
    }

    /// Read all validated merged-cell regions for one selected body table.
    ///
    /// The selector is resolved by exact visible name or checked zero-based
    /// body order. Only rooted body/table metadata, the selected ownership
    /// graph, and the selected merge payload are decoded for the query.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, or finite-resource error. Malformed
    /// or overlapping merge regions are rejected before any partial result is
    /// published.
    pub fn body_table_merges<'table, S>(&self, selector: S) -> Result<Vec<Region>, Error>
    where
        S: Into<BodyTableSelector<'table>>,
    {
        let mut budget = table_lock::WireBudget::new(self.state.limits).map_err(map_lock_error)?;
        budget
            .charge_component_catalog(self.state.components.as_ref())
            .map_err(map_lock_error)?;
        let targets = table_lock::body_table_targets_from_components_with_budget(
            self.state.components.as_ref(),
            self.state.limits,
            &mut budget,
        )
        .map_err(map_lock_error)?;
        let target =
            select_reader_target(targets, selector.into(), &mut budget).map_err(map_lock_error)?;
        let message = model_message_from_components(self.state.components.as_ref(), &target)?;
        let limits = read_limits(&budget)?;
        let read =
            table_merges::read_table_merges(&message.data, limits).map_err(map_merge_error)?;
        budget
            .charge_payload_work(read.report.input_bytes())
            .map_err(map_lock_error)?;
        budget
            .charge_codec_report(read.report.fields(), read.report.work(), 0, 0)
            .map_err(map_lock_error)?;
        for region in &read.regions {
            if region.end_row() >= target.table_rows || region.end_column() >= target.table_columns
            {
                return Err(Error::InvalidSource);
            }
        }
        Ok(read.regions)
    }
}

/// The archive profile `options` selects for physical merge-reader ingress.
fn archive_limits(options: OpenOptions) -> Result<Limits, Error> {
    options.archive().map_err(map_package_error)
}

#[cfg(test)]
mod tests {
    use litchi_iwa_archive::SourceCatalog;

    use super::*;

    const NATIVE: &[u8] =
        include_bytes!("../../../../../test-data/iwork/pages/body-table-merges-native.pages");

    /// The extent of `component`'s parsed archive stream.
    fn stream_extent(component: &litchi_iwa_archive::Component) -> u64 {
        component.archive().objects.last().map_or(0, |object| {
            object
                .data_offset
                .checked_add(object.data_length)
                .expect("native archive stream extent fits u64")
        })
    }

    fn component_payload_limits(components: &ComponentCatalog, total: u64) -> Limits {
        let largest = components
            .iter()
            .map(stream_extent)
            .max()
            .expect("native Pages package has components");
        assert!(largest > 0);
        Limits::new(
            Limits::MAX_INPUT_BYTES,
            Limits::MAX_ENTRIES,
            Limits::MAX_ENTRY_BYTES,
            total,
            usize::try_from(largest).expect("native archive stream extent fits usize"),
        )
        .expect("native payload budget is within the physical hard ceilings")
    }

    fn native_components() -> Arc<ComponentCatalog> {
        let source = SourceCatalog::from_bytes(NATIVE, Limits::default()).expect("native source");
        Arc::new(source.into_components())
    }

    #[test]
    fn reader_accepts_the_exact_component_payload_budget() {
        let components = native_components();
        let total = components.iter().map(stream_extent).sum();
        let limits = component_payload_limits(&components, total);
        let reader = MergeReader::from_shared_components_with_limits(components, limits)
            .expect("the exact budget admits the reader");
        assert!(!reader.body_table_merges(0usize).expect("merges").is_empty());
    }

    #[test]
    fn reader_rejects_one_below_component_payload_budget() {
        let components = native_components();
        let total: u64 = components.iter().map(stream_extent).sum();
        assert!(total > 0);
        let limits = component_payload_limits(&components, total - 1);
        let reader = MergeReader::from_shared_components_with_limits(components, limits)
            .expect("the root graph stays within the reduced budget");
        assert!(matches!(
            reader.body_table_merges(0usize),
            Err(Error::Limit {
                kind: crate::table::merge::transaction::LimitKind::TotalPayloadBytes,
                observed,
                maximum,
            }) if observed == total && maximum == total - 1
        ));
    }
}
