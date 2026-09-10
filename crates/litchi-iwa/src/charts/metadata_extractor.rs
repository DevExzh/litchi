//! Chart metadata extraction for the legacy document facade.
//!
//! The object index still supplies the historical global ordering, including
//! chart objects that are not reachable from a document root. The payload
//! projection itself is owned by the shared Buffa chart metadata codec; this
//! adapter only resolves the optional title object and transfers the bounded
//! borrowed labels into the format-neutral semantic value.

use std::mem::size_of;

use litchi_iwa_common::chart::metadata::ChartMetadata;
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::chart_metadata_codec::{self, DecodeOptions};

use crate::Result;
use crate::bundle::Bundle;
use crate::charts::Kind;
use crate::charts::options::read_chart_non_style_title;
use crate::object_index::{ObjectIndex, ResolvedObjectRef};

const LEGACY_CHART_MESSAGE_TYPE: u32 = 5_000;
const CHART_DRAWABLE_MESSAGE_TYPE: u32 = 5_021;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const MODERN_CHART_EXTENSION: &str = "TSCH.ChartDrawableArchive.unity";
const MAX_CHART_METADATA_LABELS: usize = WireLimits::MAX_FIELDS;
const MAX_CHART_METADATA_TEXT_BYTES: usize = WireLimits::MAX_INPUT_BYTES;
const MAX_CHART_METADATA_RETAINED_BYTES: usize = WireLimits::MAX_INPUT_BYTES;
const TITLE_OUTER_SCAN_WORK_MULTIPLIER: usize = 2;
const TITLE_CODEC_WORK_MULTIPLIER: usize = 16;
const TITLE_CODEC_WORK_OVERHEAD: usize = 32;

#[derive(Debug, Clone, Copy, Default)]
struct ChartMetadataBudget {
    fields: usize,
    work: usize,
    output: usize,
    retained: usize,
    labels: usize,
    text: usize,
    nesting: usize,
}

impl ChartMetadataBudget {
    fn charge(
        &mut self,
        report: chart_metadata_codec::DecodeReport,
        output: usize,
        retained_extra: usize,
    ) -> Result<()> {
        self.charge_with_work(report, report.work_bytes(), output, retained_extra)
    }

    fn charge_failure(&mut self, report: chart_metadata_codec::DecodeReport) -> Result<()> {
        let work = report
            .work_bytes()
            .checked_add(report.failure_work_bytes())
            .ok_or_else(|| {
                chart_metadata_budget_error(
                    "chart metadata aggregate work",
                    usize::MAX,
                    WireLimits::MAX_REWRITE_WORK,
                    Some(LimitKind::RewriteWork),
                )
            })?;
        self.charge_with_work(report, work, 0, 0)
    }

    fn charge_with_work(
        &mut self,
        report: chart_metadata_codec::DecodeReport,
        work: usize,
        output: usize,
        retained_extra: usize,
    ) -> Result<()> {
        let mut next = *self;
        next.fields = charge_one(
            "chart metadata aggregate fields",
            report.fields(),
            next.fields,
            WireLimits::MAX_FIELDS,
            Some(LimitKind::Fields),
        )?;
        next.labels = charge_one(
            "chart metadata aggregate labels",
            report.label_count(),
            next.labels,
            MAX_CHART_METADATA_LABELS,
            Some(LimitKind::Fields),
        )?;
        next.text = charge_one(
            "chart metadata aggregate text bytes",
            report.text_bytes(),
            next.text,
            MAX_CHART_METADATA_TEXT_BYTES,
            Some(LimitKind::InputBytes),
        )?;
        next.work = charge_one(
            "chart metadata aggregate work",
            work,
            next.work,
            WireLimits::MAX_REWRITE_WORK,
            Some(LimitKind::RewriteWork),
        )?;
        next.output = charge_one(
            "chart metadata aggregate output",
            output,
            next.output,
            WireLimits::MAX_OUTPUT_BYTES,
            Some(LimitKind::OutputBytes),
        )?;
        next.retained = charge_one(
            "chart metadata aggregate retained bytes",
            report
                .retained_bytes()
                .checked_add(retained_extra)
                .ok_or_else(|| {
                    chart_metadata_budget_error(
                        "chart metadata aggregate retained bytes",
                        usize::MAX,
                        MAX_CHART_METADATA_RETAINED_BYTES,
                        None,
                    )
                })?,
            next.retained,
            MAX_CHART_METADATA_RETAINED_BYTES,
            None,
        )?;
        let depth = usize::try_from(report.max_depth()).unwrap_or(usize::MAX);
        if depth > WireLimits::MAX_NESTING {
            return Err(chart_metadata_budget_error(
                "chart metadata aggregate nesting",
                depth,
                WireLimits::MAX_NESTING,
                Some(LimitKind::Nesting),
            ));
        }
        next.nesting = next.nesting.max(depth);
        *self = next;
        Ok(())
    }

    /// Build codec limits from the aggregate budget's residual capacity.
    ///
    /// `DecodeOptions::for_source` is intentionally not used here: it gives
    /// every chart a fresh source-sized work allowance, which would let a
    /// document with many chart objects exceed the extractor's aggregate
    /// ceilings before the report is charged. Refuse an exhausted residual
    /// budget before entering the codec, then pass the remaining capacity to
    /// its strict preflight.
    fn decode_options(&self, source: &[u8]) -> Result<DecodeOptions> {
        let max_input = residual_limit(
            "chart metadata aggregate retained bytes",
            self.retained,
            MAX_CHART_METADATA_RETAINED_BYTES,
            None,
        )?;
        let source_bytes = source.len().max(1);
        if source_bytes > max_input {
            return Err(chart_metadata_budget_error(
                "chart metadata aggregate retained bytes",
                source_bytes,
                max_input,
                None,
            ));
        }
        let max_fields = residual_limit(
            "chart metadata aggregate fields",
            self.fields,
            WireLimits::MAX_FIELDS,
            Some(LimitKind::Fields),
        )?;
        let max_work = residual_limit(
            "chart metadata aggregate work",
            self.work,
            WireLimits::MAX_REWRITE_WORK,
            Some(LimitKind::RewriteWork),
        )?;
        let max_labels = residual_limit(
            "chart metadata aggregate labels",
            self.labels,
            MAX_CHART_METADATA_LABELS,
            Some(LimitKind::Fields),
        )?;
        let max_text = residual_limit(
            "chart metadata aggregate text bytes",
            self.text,
            MAX_CHART_METADATA_TEXT_BYTES,
            Some(LimitKind::InputBytes),
        )?;
        let max_depth = residual_limit(
            "chart metadata aggregate nesting",
            self.nesting,
            WireLimits::MAX_NESTING,
            Some(LimitKind::Nesting),
        )?;
        let max_output = residual_limit(
            "chart metadata aggregate output",
            self.output,
            WireLimits::MAX_OUTPUT_BYTES,
            Some(LimitKind::OutputBytes),
        )?;
        if max_output < size_of::<ChartMetadata>() {
            return Err(chart_metadata_budget_error(
                "chart metadata aggregate output",
                size_of::<ChartMetadata>(),
                max_output,
                Some(LimitKind::OutputBytes),
            ));
        }
        Ok(DecodeOptions::new(
            source_bytes,
            max_fields,
            max_work,
            u32::try_from(max_depth).unwrap_or(u32::MAX),
            max_labels,
            max_text,
        ))
    }

    /// Reserve the bounded work performed by the legacy title helper.
    ///
    /// The helper owns an older title codec and therefore cannot return the
    /// chart metadata codec's report. Its options permit up to sixteen times
    /// the generated extension length, while the outer extension scan also
    /// traverses the enclosing payload. Reserve that conservative amount
    /// before calling it so malformed or oversized titles cannot bypass the
    /// aggregate work ceiling on an error path.
    fn reserve_title_work(&mut self, source_bytes: usize) -> Result<()> {
        let outer_work = source_bytes
            .checked_mul(TITLE_OUTER_SCAN_WORK_MULTIPLIER)
            .ok_or_else(|| {
                chart_metadata_budget_error(
                    "chart metadata title work",
                    usize::MAX,
                    WireLimits::MAX_REWRITE_WORK,
                    Some(LimitKind::RewriteWork),
                )
            })?;
        let codec_work = source_bytes
            .checked_add(TITLE_CODEC_WORK_OVERHEAD)
            .and_then(|bytes| bytes.checked_mul(TITLE_CODEC_WORK_MULTIPLIER))
            .ok_or_else(|| {
                chart_metadata_budget_error(
                    "chart metadata title work",
                    usize::MAX,
                    WireLimits::MAX_REWRITE_WORK,
                    Some(LimitKind::RewriteWork),
                )
            })?;
        let work = outer_work.checked_add(codec_work).ok_or_else(|| {
            chart_metadata_budget_error(
                "chart metadata title work",
                usize::MAX,
                WireLimits::MAX_REWRITE_WORK,
                Some(LimitKind::RewriteWork),
            )
        })?;
        self.work = charge_one(
            "chart metadata title work",
            work,
            self.work,
            WireLimits::MAX_REWRITE_WORK,
            Some(LimitKind::RewriteWork),
        )?;
        Ok(())
    }
}

fn residual_limit(
    resource: &'static str,
    current: usize,
    maximum: usize,
    kind: Option<LimitKind>,
) -> Result<usize> {
    maximum
        .checked_sub(current)
        .ok_or_else(|| chart_metadata_budget_error(resource, usize::MAX, maximum, kind))
}

fn charge_one(
    resource: &'static str,
    amount: usize,
    current: usize,
    maximum: usize,
    kind: Option<LimitKind>,
) -> Result<usize> {
    let observed = current
        .checked_add(amount)
        .ok_or_else(|| chart_metadata_budget_error(resource, usize::MAX, maximum, kind))?;
    if observed > maximum {
        return Err(chart_metadata_budget_error(
            resource, observed, maximum, kind,
        ));
    }
    Ok(observed)
}

fn chart_metadata_budget_error(
    resource: &'static str,
    observed: usize,
    maximum: usize,
    kind: Option<LimitKind>,
) -> crate::Error {
    if let Some(kind) = kind {
        return crate::Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit: maximum,
        });
    }
    crate::Error::InvalidFormat(format!(
        "{resource} limit exceeded: observed {observed}, limit {maximum}"
    ))
}

/// Extractor for chart metadata in the legacy document facade.
pub(crate) struct ChartMetadataExtractor<'a> {
    bundle: &'a Bundle,
    object_index: &'a ObjectIndex,
}

impl<'a> ChartMetadataExtractor<'a> {
    /// Create a chart metadata extractor over the package's existing object
    /// index. The index is deliberately not rebuilt or restricted to rooted
    /// objects so legacy global and orphan ordering remains unchanged.
    pub(crate) fn new(bundle: &'a Bundle, object_index: &'a ObjectIndex) -> Self {
        Self {
            bundle,
            object_index,
        }
    }

    /// Extract metadata from all charts in the document.
    pub(crate) fn extract_all_charts(&self) -> Result<Vec<ChartMetadata>> {
        let mut charts = Vec::new();
        let mut budget = ChartMetadataBudget::default();

        // Keep the legacy type-group ordering: all 5000 entries precede all
        // 5021 entries, and each group retains ObjectIndex ordering.
        for chart_type in [LEGACY_CHART_MESSAGE_TYPE, CHART_DRAWABLE_MESSAGE_TYPE] {
            let chart_entries = self.object_index.iter_entries_by_type(chart_type);

            for entry in chart_entries {
                if let Some(resolved) = self.object_index.resolve_ref(self.bundle, entry.id())?
                    && let Some(metadata) = self.extract_chart_metadata(&resolved, &mut budget)?
                {
                    charts.try_reserve(1).map_err(|_| {
                        crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                            resource: "chart metadata results",
                            amount: charts.len().saturating_add(1),
                        })
                    })?;
                    charts.push(metadata);
                }
            }
        }

        Ok(charts)
    }

    /// Extract metadata from a single chart object.
    fn extract_chart_metadata(
        &self,
        object: &ResolvedObjectRef<'_>,
        budget: &mut ChartMetadataBudget,
    ) -> Result<Option<ChartMetadata>> {
        for message in object.messages {
            match message.type_ {
                CHART_DRAWABLE_MESSAGE_TYPE => {
                    let options = budget.decode_options(&message.data)?;
                    let (snapshot, report) = match chart_metadata_codec::decode_modern_with_report(
                        &message.data,
                        &options,
                    ) {
                        Ok(decoded) => decoded,
                        // A 5021 object without its chart extension was
                        // ignored by the old extension-aware decoder. Keep
                        // scanning this object for a later supported message.
                        Err(error)
                            if error.missing_required_field() == Some(MODERN_CHART_EXTENSION) =>
                        {
                            budget.charge_failure(error.report())?;
                            continue;
                        },
                        Err(error) => {
                            budget.charge_failure(error.report())?;
                            return Err(codec_error(error));
                        },
                    };

                    let title_data = snapshot
                        .non_style_ref()
                        .map(|reference| self.extract_chart_title_data(reference.get()))
                        .transpose()?
                        .flatten();
                    let title_source_bytes = title_data.map_or(0, |data| data.len());
                    let output_bytes = metadata_output_bytes(
                        report.label_count(),
                        report.text_bytes(),
                        title_source_bytes,
                    )?;
                    budget.charge(report, output_bytes, title_source_bytes)?;

                    let title = title_data
                        .map(|data| {
                            budget.reserve_title_work(data.len())?;
                            read_chart_non_style_title(data)
                        })
                        .transpose()?
                        .flatten();
                    let row_labels = snapshot.row_labels();
                    let column_labels = snapshot.column_labels();
                    let row_names = own_labels(row_labels.iter(), row_labels.len())?;
                    let column_names = own_labels(column_labels.iter(), column_labels.len())?;

                    return Ok(Some(ChartMetadata::from_owned(
                        Kind::from_native(snapshot.chart_type()),
                        title,
                        row_names,
                        column_names,
                        snapshot.series_count(),
                        snapshot.contains_default_data().unwrap_or(false),
                    )));
                },
                LEGACY_CHART_MESSAGE_TYPE => {
                    let options = budget.decode_options(&message.data)?;
                    let (snapshot, report) = match chart_metadata_codec::decode_legacy_with_report(
                        &message.data,
                        &options,
                    ) {
                        Ok(decoded) => decoded,
                        Err(error) => {
                            budget.charge_failure(error.report())?;
                            return Err(codec_error(error));
                        },
                    };
                    let output_bytes =
                        metadata_output_bytes(report.label_count(), report.text_bytes(), 0)?;
                    budget.charge(report, output_bytes, 0)?;
                    let row_labels = snapshot.row_labels();
                    let column_labels = snapshot.column_labels();
                    let row_names = own_labels(row_labels.iter(), row_labels.len())?;
                    let column_names = own_labels(column_labels.iter(), column_labels.len())?;

                    return Ok(Some(ChartMetadata::from_owned(
                        Kind::from_native(snapshot.chart_type()),
                        None,
                        row_names,
                        column_names,
                        snapshot.series_count(),
                        false,
                    )));
                },
                _ => {},
            }
        }

        Ok(None)
    }

    /// Locate the title payload referenced by a modern chart payload without
    /// allocating its decoded text. The caller charges the source span before
    /// invoking the title codec, which preserves aggregate output accounting.
    fn extract_chart_title_data(&self, reference: u64) -> Result<Option<&[u8]>> {
        let Some(resolved) = self.object_index.resolve_ref_id(self.bundle, reference)? else {
            return Ok(None);
        };
        let mut messages = resolved
            .messages
            .iter()
            .filter(|message| message.type_ == CHART_NON_STYLE_MESSAGE_TYPE);
        let Some(message) = messages.next() else {
            return Err(crate::Error::InvalidFormat(format!(
                "chart non-style {reference} must have exactly one payload"
            )));
        };
        if messages.next().is_some() {
            return Err(crate::Error::InvalidFormat(format!(
                "chart non-style {reference} must have exactly one payload"
            )));
        }
        Ok(Some(message.data.as_slice()))
    }
}

fn codec_error(error: chart_metadata_codec::DecodeError) -> crate::Error {
    crate::Error::InvalidFormat(format!("chart metadata decode failed: {error}"))
}

fn metadata_output_bytes(
    label_count: usize,
    text_bytes: usize,
    title_source_bytes: usize,
) -> Result<usize> {
    let label_slots = label_count
        .checked_mul(size_of::<String>())
        .ok_or_else(|| {
            crate::Error::InvalidFormat("chart metadata output size overflow".to_owned())
        })?;
    size_of::<ChartMetadata>()
        .checked_add(label_slots)
        .and_then(|size| size.checked_add(text_bytes))
        .and_then(|size| size.checked_add(title_source_bytes))
        .ok_or_else(|| {
            crate::Error::InvalidFormat("chart metadata output size overflow".to_owned())
        })
}

fn own_labels<'source>(
    labels: impl Iterator<Item = &'source str>,
    label_count: usize,
) -> Result<Vec<String>> {
    let mut owned = Vec::new();
    owned.try_reserve_exact(label_count).map_err(|_| {
        crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "chart metadata label slots",
            amount: label_count,
        })
    })?;
    for label in labels {
        owned.try_reserve(1).map_err(|_| {
            crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "chart metadata label slots",
                amount: owned.len().saturating_add(1),
            })
        })?;

        let mut value = String::new();
        value.try_reserve_exact(label.len()).map_err(|_| {
            crate::Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "chart metadata label text",
                amount: label.len(),
            })
        })?;
        value.push_str(label);
        owned.push(value);
    }
    Ok(owned)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::IWorkPackage;
    use crate::archive::{Archive, ArchiveObject, RawMessage};
    use crate::bundle::Bundle;
    use crate::object_index::ObjectIndex;
    use crate::wire::{append_length_delimited_field, append_varint_field};

    #[test]
    fn chart_metadata_creation_uses_shared_semantic_projection() {
        let metadata =
            ChartMetadata::from_owned(Kind::Undefined, None, Vec::new(), Vec::new(), 0, false);
        assert_eq!(metadata.title(), None);
        assert_eq!(metadata.series_count(), 0);
        assert!(!metadata.has_content());
    }

    #[test]
    fn chart_metadata_with_content_preserves_text_order() {
        let metadata = ChartMetadata::from_owned(
            Kind::Column2d,
            Some("Sales Chart".to_owned()),
            vec!["Q1".to_owned(), "Q2".to_owned()],
            vec!["Revenue".to_owned()],
            2,
            false,
        );

        assert!(metadata.has_content());
        let all_text = metadata.all_text().collect::<Vec<_>>();
        assert_eq!(all_text.len(), 4);
        assert!(all_text.contains(&"Sales Chart"));
    }

    #[test]
    fn chart_metadata_type_remains_lossless() {
        let metadata =
            ChartMetadata::from_owned(Kind::Bar2d, None, Vec::new(), Vec::new(), 0, false);

        assert_eq!(metadata.kind(), Kind::Bar2d);
    }

    #[test]
    fn extraction_preserves_global_type_groups_and_skips_missing_modern_extension() {
        let archive = Archive {
            objects: vec![
                chart_object(30, LEGACY_CHART_MESSAGE_TYPE, legacy_chart("legacy-30")),
                chart_object(10, LEGACY_CHART_MESSAGE_TYPE, legacy_chart("legacy-10")),
                chart_object(40, CHART_DRAWABLE_MESSAGE_TYPE, modern_chart("modern-40")),
                chart_object(20, CHART_DRAWABLE_MESSAGE_TYPE, modern_chart("modern-20")),
                chart_object(50, CHART_DRAWABLE_MESSAGE_TYPE, Vec::new()),
            ],
        };
        let mut package = IWorkPackage::new();
        package
            .replace_archive("Index/SyntheticCharts.iwa", &archive)
            .expect("synthetic archive");
        let bytes = package.to_bytes().expect("synthetic package");
        let bundle = Bundle::from_bytes(&bytes).expect("synthetic bundle");
        let object_index = ObjectIndex::from_bundle(&bundle).expect("synthetic object index");

        let charts = ChartMetadataExtractor::new(&bundle, &object_index)
            .extract_all_charts()
            .expect("chart metadata");
        let rows = charts
            .iter()
            .map(|chart| chart.row_names()[0].as_str())
            .collect::<Vec<_>>();
        assert_eq!(rows, ["legacy-10", "legacy-30", "modern-20", "modern-40"]);
    }

    #[test]
    fn extraction_rejects_malformed_legacy_payload() {
        let archive = Archive {
            objects: vec![chart_object(10, LEGACY_CHART_MESSAGE_TYPE, Vec::new())],
        };
        let mut package = IWorkPackage::new();
        package
            .replace_archive("Index/MalformedCharts.iwa", &archive)
            .expect("synthetic archive");
        let bytes = package.to_bytes().expect("synthetic package");
        let bundle = Bundle::from_bytes(&bytes).expect("synthetic bundle");
        let object_index = ObjectIndex::from_bundle(&bundle).expect("synthetic object index");

        assert!(
            ChartMetadataExtractor::new(&bundle, &object_index)
                .extract_all_charts()
                .is_err()
        );
    }

    fn chart_object(identifier: u64, type_: u32, data: Vec<u8>) -> ArchiveObject {
        ArchiveObject::new(identifier, vec![RawMessage { type_, data }])
            .expect("synthetic chart object")
    }

    fn legacy_chart(row_name: &str) -> Vec<u8> {
        let mut grid = Vec::new();
        append_length_delimited_field(&mut grid, 2, row_name.as_bytes()).unwrap();
        append_length_delimited_field(&mut grid, 3, b"series").unwrap();
        append_length_delimited_field(&mut grid, 4, &[]).unwrap();
        append_length_delimited_field(&mut grid, 4, &[]).unwrap();

        let mut model = Vec::new();
        append_length_delimited_field(&mut model, 5, &grid).unwrap();

        let mut chart = Vec::new();
        append_length_delimited_field(&mut chart, 2, &model).unwrap();
        append_varint_field(&mut chart, 4, 0).unwrap();
        chart
    }

    fn modern_chart(row_name: &str) -> Vec<u8> {
        let mut grid = Vec::new();
        append_length_delimited_field(&mut grid, 1, row_name.as_bytes()).unwrap();
        append_length_delimited_field(&mut grid, 2, b"series").unwrap();
        append_length_delimited_field(&mut grid, 3, &[]).unwrap();
        append_length_delimited_field(&mut grid, 3, &[]).unwrap();

        let mut chart = Vec::new();
        append_varint_field(&mut chart, 1, 0).unwrap();
        append_length_delimited_field(&mut chart, 7, &grid).unwrap();

        let mut drawable = Vec::new();
        append_length_delimited_field(&mut drawable, 10_000, &chart).unwrap();
        drawable
    }
}
