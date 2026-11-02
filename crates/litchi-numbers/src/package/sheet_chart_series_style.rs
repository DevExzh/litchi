//! The per-series properties stored in a chart's series styles (strokes,
//! fills, data-symbol fills and outlines) for [`super::sheet_chart_series`]
//! (record 0378): the host editor's `set_sheet_chart_series_{strokes,
//! fills, symbol_fills, symbol_outlines}` and `reset_*` routes on W1-E's
//! series-style write rules (`chart_style_codec::property`, record 0375a).
//!
//! The value of a series is read through its effective style (the private
//! override, else the theme style) and its `TSS.StyleArchive.parent` chain;
//! a write rewrites the effective style in place, which only this chart may
//! use (the host's `ensure_exclusive`). Nothing is allocated and the package
//! metadata does not change.

use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::chart_style_codec::{self as codec, Archive, Limits};
use litchi_iwa_style::series::style as engine;
use litchi_iwa_style::series::{Refusal as EngineRefusal, SymbolError};

use super::Package;
use super::sheet_drawable_graph::{self as graph, CHART_MESSAGE_TYPE};
use super::structure::{self, Fault, FaultResult, Published};

/// One style object's in-place rewrite.
type StyleRewrite<'a> =
    dyn FnMut(&mut litchi_iwa_core::ArchiveObject, litchi_iwa_core::Limits) -> FaultResult<()> + 'a;

/// Replace message `index` (of `message_type`) with `data`, keeping the
/// archive header.
fn rewriter(
    index: usize,
    message_type: u32,
    data: Vec<u8>,
) -> impl FnMut(&mut litchi_iwa_core::ArchiveObject, litchi_iwa_core::Limits) -> FaultResult<()> {
    let mut data = Some(data);
    move |object, limits| {
        object
            .replace_message_preserving_header(
                index,
                RawMessage {
                    type_: message_type,
                    data: data.take().ok_or(Fault::Invalid)?,
                },
                limits,
            )
            .map(drop)
            .map_err(|error| structure::map_core(&error))
    }
}

pub(super) use litchi_iwa_style::series::style::{Kind, Value};

/// Why a style property is refused.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Refused {
    /// The chart kind has no such property.
    Kind,
    /// Another chart uses the series style.
    Shared,
    /// An image fill needs the data-reference graph.
    ImageFill,
}

#[derive(Debug)]
pub(super) enum Failure {
    Refused(Refused),
    Fault(Fault),
    Budget(super::sheet_chart_series::Error),
}

impl From<Fault> for Failure {
    fn from(fault: Fault) -> Self {
        Self::Fault(fault)
    }
}

impl From<super::sheet_chart_series::Error> for Failure {
    fn from(error: super::sheet_chart_series::Error) -> Self {
        Self::Budget(error)
    }
}

impl From<SymbolError> for Failure {
    fn from(error: SymbolError) -> Self {
        match error {
            SymbolError::Refused(EngineRefusal::UnsupportedKind) => Self::Refused(Refused::Kind),
            SymbolError::Refused(EngineRefusal::ImageFill) => Self::Refused(Refused::ImageFill),
            SymbolError::Refused(EngineRefusal::SharedStyle) => Self::Refused(Refused::Shared),
            SymbolError::Graph(litchi_iwa_style::graph::Error::Codec(error)) => error.into(),
            _ => Self::Fault(Fault::Invalid),
        }
    }
}

impl From<codec::Error> for Failure {
    fn from(error: codec::Error) -> Self {
        match error {
            codec::Error::Allocation { bytes, .. } => Self::Fault(structure::allocation(bytes)),
            _ => Self::Fault(Fault::Invalid),
        }
    }
}

type Result<T> = std::result::Result<T, Failure>;

/// Transaction-owned admission for the style publisher.  The chart-series
/// facade implements this once so staging, member encoding, physical
/// reassembly, clone, and candidate execution all consume one ledger.  The
/// graph module adapts these operations to its generic publisher seam.
pub(super) trait PublicationAdmission {
    fn encode_member(
        &mut self,
        name: &str,
        archive: &litchi_iwa_core::Archive,
        archive_limits: litchi_iwa_core::Limits,
        snappy_limits: litchi_iwa_core::SnappyLimits,
    ) -> Result<Vec<u8>>;

    fn admit_reassembly(
        &mut self,
        charge: litchi_iwa_archive::unstable::CandidateIngressCharge,
    ) -> Result<()>;

    fn admit_execution(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<()>;

    fn admit_candidate(
        &mut self,
        candidate: &[u8],
        limits: litchi_iwa_archive::Limits,
    ) -> Result<()>;

    fn reopen_candidate(&mut self, source: &Package, candidate: Vec<u8>) -> Result<Package>;

    fn admit_clone(&mut self, chunk: litchi_iwa_core::ArchiveCloneInventoryChunk) -> Result<()>;

    fn admit_clone_arithmetic(&mut self, resource: &'static str) -> Result<()>;

    fn admit_staging(&mut self, charge: graph::PublicationStagingCharge) -> Result<()>;
}

impl<T: PublicationAdmission> graph::PublicationAdmission for T {
    type Error = Failure;

    fn encode_member(
        &mut self,
        name: &str,
        archive: &litchi_iwa_core::Archive,
        archive_limits: litchi_iwa_core::Limits,
        snappy_limits: litchi_iwa_core::SnappyLimits,
    ) -> Result<Vec<u8>> {
        PublicationAdmission::encode_member(self, name, archive, archive_limits, snappy_limits)
    }

    fn admit_reassembly(
        &mut self,
        charge: litchi_iwa_archive::unstable::CandidateIngressCharge,
    ) -> Result<()> {
        PublicationAdmission::admit_reassembly(self, charge)
    }

    fn admit_execution(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<()> {
        PublicationAdmission::admit_execution(self, requirements)
    }

    fn admit_candidate(
        &mut self,
        candidate: &[u8],
        limits: litchi_iwa_archive::Limits,
    ) -> Result<()> {
        PublicationAdmission::admit_candidate(self, candidate, limits)
    }

    fn reopen_candidate(&mut self, source: &Package, candidate: Vec<u8>) -> Result<Package> {
        PublicationAdmission::reopen_candidate(self, source, candidate)
    }

    fn admit_clone(&mut self, chunk: litchi_iwa_core::ArchiveCloneInventoryChunk) -> Result<()> {
        PublicationAdmission::admit_clone(self, chunk)
    }

    fn admit_clone_arithmetic(&mut self, resource: &'static str) -> Result<()> {
        PublicationAdmission::admit_clone_arithmetic(self, resource)
    }

    fn admit_staging(&mut self, charge: graph::PublicationStagingCharge) -> Result<()> {
        PublicationAdmission::admit_staging(self, charge)
    }
}

/// The effective series styles of the chart's first `count` series.
fn styles(package: &Package, chart: u64, count: usize) -> Result<Vec<u64>> {
    let object = graph::object(package, chart)?;
    let index = structure::unique_message(object, &[CHART_MESSAGE_TYPE])?;
    let references = codec::references(&object.messages[index].data, Limits::DEFAULT)?;
    (0..count)
        .map(|series| {
            references
                .series_style(series)
                .ok_or(Failure::Fault(Fault::Invalid))
        })
        .collect()
}

fn style_payload(package: &Package, style: u64) -> FaultResult<(usize, &[u8])> {
    let object = graph::object(package, style)?;
    let index = structure::unique_message(object, &[Archive::SeriesStyle.message_type()])?;
    Ok((index, &object.messages[index].data))
}

/// The effective value of `kind` in `style` (and whether `style` defines
/// it itself), on the shared series-style engine.
fn effective(package: &Package, style: u64, chart: ChartKind, kind: Kind) -> Result<(Value, bool)> {
    engine::effective(style, chart, kind, |identifier| {
        style_payload(package, identifier)
            .map(|(_, payload)| payload)
            .map_err(|_fault| SymbolError::Invalid("a series style is missing"))
    })
    .map_err(Failure::from)
}

/// The effective value of `kind` for each of the chart's `count` series.
pub(super) fn read(
    package: &Package,
    chart: u64,
    chart_kind: ChartKind,
    count: usize,
    kind: Kind,
) -> Result<Vec<Value>> {
    styles(package, chart, count)?
        .into_iter()
        .map(|style| effective(package, style, chart_kind, kind).map(|(value, _)| value))
        .collect()
}

/// Whether only one chart's series use `style` (the host's
/// `ensure_exclusive`).
fn exclusive(package: &Package, style: u64) -> Result<bool> {
    let mut owners = 0_usize;
    for object in package.state.components.iter_objects() {
        for message in &object.messages {
            if message.type_ != CHART_MESSAGE_TYPE {
                continue;
            }
            let references = codec::references(&message.data, Limits::DEFAULT)?;
            let private = references
                .series_private_styles()
                .is_some_and(|sparse| sparse.entries().iter().any(|(_, id)| *id == style));
            if private || references.series_theme_styles().contains(&style) {
                owners += 1;
            }
        }
    }
    Ok(owners == 1)
}

/// The rewritten payload of `style` storing `value` (`None`: reset).
fn rewrite(payload: &[u8], chart: ChartKind, kind: Kind, value: Option<&Value>) -> Result<Vec<u8>> {
    engine::rewrite(payload, chart, kind, value).map_err(Failure::from)
}

/// Store `values` (`Some`: set; `None`: reset) in the effective styles of
/// the chart's series whose value changes; `None` without a change.
pub(super) fn write(
    package: &Package,
    chart: u64,
    chart_kind: ChartKind,
    kind: Kind,
    values: &[Option<Value>],
    admission: &mut impl PublicationAdmission,
) -> Result<Option<Published>> {
    let slots = styles(package, chart, values.len())?;
    let mut replacements: Vec<(u64, usize, Vec<u8>)> = Vec::new();
    for (&style, value) in slots.iter().zip(values) {
        let (current, defined) = effective(package, style, chart_kind, kind)?;
        let unchanged = match value {
            Some(value) => current == *value,
            None => !defined,
        };
        if unchanged {
            continue;
        }
        if !exclusive(package, style)? {
            return Err(Failure::Refused(Refused::Shared));
        }
        if replacements.iter().any(|(existing, ..)| *existing == style) {
            return Err(Failure::Refused(Refused::Shared));
        }
        let (index, payload) = style_payload(package, style)?;
        let data = rewrite(payload, chart_kind, kind, value.as_ref())?;
        if data != payload {
            let charge = graph::PublicationStagingCharge::vector::<(u64, usize, Vec<u8>)>(1)
                .ok_or(Failure::Fault(Fault::Invalid))?;
            admission.admit_staging(charge)?;
            structure::reserve(&mut replacements, 1)?;
            replacements.push((style, index, data));
        }
    }
    if replacements.is_empty() {
        return Ok(None);
    }
    let message_type = Archive::SeriesStyle.message_type();
    let mut closures = Vec::new();
    let closure_charge =
        graph::PublicationStagingCharge::vector::<(u64, Vec<u8>)>(replacements.len())
            .ok_or(Failure::Fault(Fault::Invalid))?;
    admission.admit_staging(closure_charge)?;
    structure::reserve(&mut closures, replacements.len())?;
    for (style, index, data) in replacements {
        closures.push((style, rewriter(index, message_type, data)));
    }
    let mut targets: Vec<(u64, &mut StyleRewrite<'_>)> = Vec::new();
    let target_charge =
        graph::PublicationStagingCharge::vector::<(u64, &mut StyleRewrite<'_>)>(closures.len())
            .ok_or(Failure::Fault(Fault::Invalid))?;
    admission.admit_staging(target_charge)?;
    structure::reserve(&mut targets, closures.len())?;
    for (style, closure) in &mut closures {
        targets.push((*style, closure));
    }
    let published = graph::publish_in_place_with_admission(package, &mut targets, admission)?;
    Ok(Some(published))
}
