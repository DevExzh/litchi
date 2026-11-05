//! The private references of a chart drawable and their remap (record
//! 0426; the host's `IWorkChartArchive::{typed_reference_identifiers,
//! remap_references}` and `remap_chart_preset_wire`).
//!
//! A chart's private graph is reached from its `TSCH.ChartDrawableArchive`:
//! the drawable's title and caption, the chart's preset and its chart,
//! legend, axis, series and reference-line styles. [`references`] lists
//! them in the host's discovery order (`body_chart_graph`), so a copy that
//! allocates identifiers in that order takes the host's identifiers.
//! [`remap`] rewrites every reference a copied message holds at its known
//! paths ([`reference_paths`]) in place, keeping every other byte, and
//! reports the references it saw, so the owner can refuse a private
//! reference outside them as the host does.

use std::collections::{BTreeSet, HashMap};

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::wire::{
    WireView, patch_varint_field, transform_length_delimited_fields_at_path,
};

use super::Error;
use super::author::{CHART_MESSAGE_TYPE, MEDIATOR_MESSAGE_TYPE, PRESET_MESSAGE_TYPE};

/// `TSCH.ChartStyleArchive`.
pub const CHART_STYLE_MESSAGE_TYPE: u32 = 5_022;
/// `TSCH.ChartNonStyleArchive`.
pub const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
/// `TSCH.LegendStyleArchive`.
pub const LEGEND_STYLE_MESSAGE_TYPE: u32 = 5_024;
/// `TSCH.LegendNonStyleArchive`.
pub const LEGEND_NON_STYLE_MESSAGE_TYPE: u32 = 5_025;
/// `TSCH.ChartAxisStyleArchive`.
pub const AXIS_STYLE_MESSAGE_TYPE: u32 = 5_026;
/// `TSCH.ChartAxisNonStyleArchive`.
pub const AXIS_NON_STYLE_MESSAGE_TYPE: u32 = 5_027;
/// `TSCH.ChartSeriesStyleArchive`.
pub const SERIES_STYLE_MESSAGE_TYPE: u32 = 5_028;
/// `TSCH.ChartSeriesNonStyleArchive`.
pub const SERIES_NON_STYLE_MESSAGE_TYPE: u32 = 5_029;
/// `TSCH.ReferenceLineStyleArchive`.
pub const REFERENCE_LINE_STYLE_MESSAGE_TYPE: u32 = 5_030;
/// `TSCH.ReferenceLineNonStyleArchive`.
pub const REFERENCE_LINE_NON_STYLE_MESSAGE_TYPE: u32 = 5_031;

/// `TSCH.ChartArchive` (the drawable's `unity` extension) and its private
/// reference-line extension.
const UNITY: u32 = 10_000;
const REFERENCE_LINES: u32 = 10_005;

/// `TSCH.ChartDrawableArchive` reference paths: the drawable's parent,
/// comment, pencil annotations, title and caption; the chart's preset,
/// mediator, styles (sparse arrays through their entries), paragraph
/// styles and owned preset; the reference lines' non-styles, styles and
/// theme preset style.
const CHART_PATHS: &[&[u32]] = &[
    &[1, 2],
    &[1, 6],
    &[1, 9],
    &[1, 10],
    &[1, 11],
    &[UNITY, 4],
    &[UNITY, 8],
    &[UNITY, 9],
    &[UNITY, 10],
    &[UNITY, 11],
    &[UNITY, 12],
    &[UNITY, 13],
    &[UNITY, 14],
    &[UNITY, 15],
    &[UNITY, 16],
    &[UNITY, 17],
    &[UNITY, 18, 2, 2],
    &[UNITY, 19, 2, 2],
    &[UNITY, 20],
    &[UNITY, 23],
    &[UNITY, REFERENCE_LINES, 1, 2, 1],
    &[UNITY, REFERENCE_LINES, 2, 2, 2, 2],
    &[UNITY, REFERENCE_LINES, 3],
];
/// `TSCH.ChartStylePreset`: chart, legend, value-axis, category-axis,
/// series and paragraph styles.
const PRESET_PATHS: &[&[u32]] = &[&[1], &[2], &[3], &[4], &[5], &[6]];
/// `TN.ChartMediatorArchive`: its chart.
const MEDIATOR_PATHS: &[&[u32]] = &[&[1, 1]];

/// The reference paths a chart-graph message type holds; `None` for a
/// message whose references a copy cannot follow.
#[must_use]
pub fn reference_paths(message_type: u32) -> Option<&'static [&'static [u32]]> {
    match message_type {
        CHART_MESSAGE_TYPE => Some(CHART_PATHS),
        PRESET_MESSAGE_TYPE => Some(PRESET_PATHS),
        MEDIATOR_MESSAGE_TYPE => Some(MEDIATOR_PATHS),
        _ => None,
    }
}

/// One private style reference of a chart and the message type its object
/// must hold.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Style {
    /// The referenced object.
    pub identifier: u64,
    /// The message type the object holds.
    pub message_type: u32,
}

/// The references a chart drawable payload holds.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct References {
    /// The drawable's parent.
    pub parent: Option<u64>,
    /// The drawable's comment.
    pub comment: Option<u64>,
    /// The drawable's title.
    pub title: Option<u64>,
    /// The drawable's caption.
    pub caption: Option<u64>,
    /// The chart's active preset.
    pub preset: Option<u64>,
    /// The chart's own preset (zero references are absent).
    pub owned_preset: Option<u64>,
    /// The chart's mediator (zero references are absent).
    pub mediator: Option<u64>,
    /// The chart, legend, axis, series and reference-line styles in the
    /// host's discovery order (`chart_style` … `series_non_styles`, then
    /// the reference-line non-styles and styles), repeats kept.
    pub styles: Vec<Style>,
    /// Every typed reference of the payload.
    pub typed: BTreeSet<u64>,
}

fn reference_identifier(reference: &[u8], limits: WireLimits) -> Result<u64, Error> {
    let view = WireView::parse(reference, limits)?;
    let mut identifier = None;
    for field in view.fields() {
        if field.number() == 1 {
            if field.wire_type() != 0 || identifier.is_some() {
                return Err(Error::Invalid);
            }
            let (value, length) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                .map_err(|_error| Error::Invalid)?;
            if length != field.payload().len() {
                return Err(Error::Invalid);
            }
            identifier = Some(value);
        }
    }
    identifier.ok_or(Error::Invalid)
}

/// Every reference identifier at `path`, in payload order.
fn at_path(payload: &[u8], path: &[u32], limits: WireLimits) -> Result<Vec<u64>, Error> {
    let mut found = Vec::new();
    transform_length_delimited_fields_at_path(payload, path, |reference| {
        found.push(reference_identifier(reference, limits)?);
        Ok::<_, Error>(reference.to_vec())
    })?;
    Ok(found)
}

fn single(payload: &[u8], path: &[u32], limits: WireLimits) -> Result<Option<u64>, Error> {
    match at_path(payload, path, limits)?.as_slice() {
        [] => Ok(None),
        [identifier] => Ok(Some(*identifier)),
        _ => Err(Error::Invalid),
    }
}

/// The references of a chart drawable payload.
///
/// # Errors
///
/// [`Error`] for a malformed payload or a repeated singular reference.
pub fn references(payload: &[u8], limits: WireLimits) -> Result<References, Error> {
    let nonzero = |identifier: Option<u64>| identifier.filter(|identifier| *identifier != 0);
    let mut styles = Vec::new();
    let style_paths: [(&[u32], u32); 13] = [
        (&[UNITY, 9], CHART_STYLE_MESSAGE_TYPE),
        (&[UNITY, 10], CHART_NON_STYLE_MESSAGE_TYPE),
        (&[UNITY, 11], LEGEND_STYLE_MESSAGE_TYPE),
        (&[UNITY, 12], LEGEND_NON_STYLE_MESSAGE_TYPE),
        (&[UNITY, 13], AXIS_STYLE_MESSAGE_TYPE),
        (&[UNITY, 14], AXIS_NON_STYLE_MESSAGE_TYPE),
        (&[UNITY, 15], AXIS_STYLE_MESSAGE_TYPE),
        (&[UNITY, 16], AXIS_NON_STYLE_MESSAGE_TYPE),
        (&[UNITY, 17], SERIES_STYLE_MESSAGE_TYPE),
        (&[UNITY, 18, 2, 2], SERIES_STYLE_MESSAGE_TYPE),
        (&[UNITY, 19, 2, 2], SERIES_NON_STYLE_MESSAGE_TYPE),
        (
            &[UNITY, REFERENCE_LINES, 1, 2, 1],
            REFERENCE_LINE_NON_STYLE_MESSAGE_TYPE,
        ),
        (
            &[UNITY, REFERENCE_LINES, 2, 2, 2, 2],
            REFERENCE_LINE_STYLE_MESSAGE_TYPE,
        ),
    ];
    for (path, message_type) in style_paths {
        for identifier in at_path(payload, path, limits)? {
            styles.push(Style {
                identifier,
                message_type,
            });
        }
    }
    let mut typed = BTreeSet::new();
    for path in CHART_PATHS {
        typed.extend(at_path(payload, path, limits)?);
    }
    Ok(References {
        parent: single(payload, &[1, 2], limits)?,
        comment: single(payload, &[1, 6], limits)?,
        title: single(payload, &[1, 10], limits)?,
        caption: single(payload, &[1, 11], limits)?,
        preset: single(payload, &[UNITY, 4], limits)?,
        owned_preset: nonzero(single(payload, &[UNITY, 23], limits)?),
        mediator: nonzero(single(payload, &[UNITY, 8], limits)?),
        styles,
        typed,
    })
}

/// Whether a chart style object is one the document stylesheet registers:
/// every chart, legend, axis and series style kind, and a series non-style
/// whose `TSS.StyleArchive` names a stylesheet (the host's
/// `is_registered_style`).
///
/// # Errors
///
/// [`Error`] for a malformed series non-style.
pub fn registered_style(
    message_type: u32,
    payload: &[u8],
    limits: WireLimits,
) -> Result<bool, Error> {
    Ok(match message_type {
        SERIES_NON_STYLE_MESSAGE_TYPE => !at_path(payload, &[1, 5], limits)?.is_empty(),
        CHART_STYLE_MESSAGE_TYPE
        | CHART_NON_STYLE_MESSAGE_TYPE
        | LEGEND_STYLE_MESSAGE_TYPE
        | LEGEND_NON_STYLE_MESSAGE_TYPE
        | AXIS_STYLE_MESSAGE_TYPE
        | AXIS_NON_STYLE_MESSAGE_TYPE
        | SERIES_STYLE_MESSAGE_TYPE => true,
        _ => false,
    })
}

/// `payload` (of a message with reference `paths`) with every reference
/// `remap` names rewritten in place, and the references seen at `paths`.
///
/// # Errors
///
/// [`Error`] for a malformed payload.
pub fn remap(
    payload: &[u8],
    paths: &[&[u32]],
    remap: &HashMap<u64, u64>,
    limits: WireLimits,
) -> Result<(Vec<u8>, BTreeSet<u64>), Error> {
    let mut output = payload.to_vec();
    let mut seen = BTreeSet::new();
    for path in paths {
        output = transform_length_delimited_fields_at_path(&output, path, |reference| {
            let identifier = reference_identifier(reference, limits)?;
            seen.insert(identifier);
            match remap.get(&identifier) {
                Some(new) => Ok::<_, Error>(patch_varint_field(reference, 1, true, Some(*new))?),
                None => Ok(reference.to_vec()),
            }
        })?;
    }
    Ok((output, seen))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn reference(identifier: u64) -> Vec<u8> {
        let mut out = vec![0x08];
        litchi_iwa_common::varint::encode_varint_into(&mut out, identifier);
        out
    }

    fn field(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut out,
            number,
            payload,
            WireLimits::default(),
        )
        .expect("field");
        out
    }

    #[test]
    fn styles_follow_the_host_order_and_remap_in_place() {
        let mut drawable = field(10, &reference(7));
        drawable.extend(field(11, &reference(8)));
        let mut unity = field(4, &reference(20));
        unity.extend(field(13, &reference(31)));
        unity.extend(field(9, &reference(30)));
        let mut entry = vec![0x08, 0x00];
        entry.extend(field(2, &reference(32)));
        let mut sparse = vec![0x08, 0x01];
        sparse.extend(field(2, &entry));
        unity.extend(field(18, &sparse));
        unity.extend(field(23, &reference(20)));
        let mut payload = field(1, &drawable);
        payload.extend(field(UNITY, &unity));

        let found = references(&payload, WireLimits::default()).expect("references");
        assert_eq!(found.title, Some(7));
        assert_eq!(found.caption, Some(8));
        assert_eq!(found.preset, Some(20));
        assert_eq!(found.owned_preset, Some(20));
        assert_eq!(
            found
                .styles
                .iter()
                .map(|style| style.identifier)
                .collect::<Vec<_>>(),
            [30, 31, 32]
        );
        assert_eq!(found.typed, BTreeSet::from([7, 8, 20, 30, 31, 32]));

        let map = HashMap::from([(7, 107), (32, 132), (20, 120)]);
        let (copy, seen) = remap(
            &payload,
            reference_paths(CHART_MESSAGE_TYPE).expect("paths"),
            &map,
            WireLimits::default(),
        )
        .expect("remap");
        assert_eq!(seen, found.typed);
        let copied = references(&copy, WireLimits::default()).expect("copy");
        assert_eq!(copied.title, Some(107));
        assert_eq!(copied.owned_preset, Some(120));
        assert_eq!(copied.styles[2].identifier, 132);
    }
}
