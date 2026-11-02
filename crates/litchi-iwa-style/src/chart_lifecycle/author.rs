//! The source-free chart graph every format authors when it creates a chart
//! (the host's `charts::source` builder, lifted from the Keynote leaf in
//! record 0418).
//!
//! A new chart is a fixed private graph: the chart drawable, its title and
//! caption stand-ins, a Numbers chart mediator, the chart's theme style
//! preset and sixteen chart, legend, axis and series styles. This module
//! allocates the graph's identifiers in the host's order and writes every
//! object's payload field by field in tag-number order (how Prost encodes
//! them, so the bytes equal the host's). The owner inserts the objects,
//! joins the styles to its stylesheet and the preset to its theme, attaches
//! the drawable and registers the cross-component references and UUIDs,
//! because those steps differ by format.

use litchi_iwa_common::chart::data::ChartData;
use litchi_iwa_common::chart::gaps::Spacing;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_core::{ArchiveObject, Limits, RawMessage};
use litchi_iwa_protos::chart_grid_creation_codec as grid_codec;

use super::Error;

/// `TSCH.ChartDrawableArchive`.
pub const CHART_MESSAGE_TYPE: u32 = 5_021;
/// `TSCH.ChartStylePreset`.
pub const PRESET_MESSAGE_TYPE: u32 = 5_020;
/// `TSD.StandinCaptionArchive`.
pub const STANDIN_MESSAGE_TYPE: u32 = 3_097;
/// `TN.ChartMediatorArchive`.
pub const MEDIATOR_MESSAGE_TYPE: u32 = 12_006;
const CHART_STYLE: u32 = 5_022;
const CHART_NON_STYLE: u32 = 5_023;
const LEGEND_STYLE: u32 = 5_024;
const LEGEND_NON_STYLE: u32 = 5_025;
const AXIS_STYLE: u32 = 5_026;
const AXIS_NON_STYLE: u32 = 5_027;
const SERIES_STYLE: u32 = 5_028;
const STANDARD_VERSION: &[u32] = &[1, 0, 5];
const STANDIN_VERSION: &[u32] = &[10, 1, 0];
const EXTENSION: u32 = 10_000;
const SUPPORTS_PRIMARY: u32 = 10_001;
const SUPPORTS_SECONDARY: u32 = 10_002;
const SCATTER_FORMAT_SHARED_X: u64 = 2;
const SERIES_DIRECTION_BY_ROW: u64 = 1;
const NUMBER_FORMAT_TYPE: u64 = 256;
const AUTOMATIC_DECIMAL_PLACES: u64 = 253;
const RADAR_STROKE_WIDTH: f32 = 4.0;
const MITER_LIMIT: f32 = 4.0;
const STROKE_PATTERN_CAPACITY: usize = 6;
const RADAR_FILL_STROKE_ALPHA: f32 = 0.15;
const MEDIATOR_LOCAL_SERIES_SENTINEL: u64 = u32::MAX as u64;
const SERIES_COLORS: [(f32, f32, f32); 6] = [
    (0.16, 0.55, 0.88),
    (0.29, 0.70, 0.39),
    (0.57, 0.57, 0.60),
    (0.95, 0.65, 0.16),
    (0.72, 0.25, 0.23),
    (0.62, 0.25, 0.55),
];
/// `TSCH.ChartArchive` boolean extensions every new chart records.
const CHART_BOOL_EXTENSIONS: [u32; 7] = [10_002, 10_023, 10_024, 10_026, 10_027, 10_028, 10_029];
const CHART_CACHED_FORMATTERS: u32 = 10_030;

/// The per-application differences of a new chart graph (the host's
/// `ChartApplicationProfile`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Profile {
    /// Numbers: thirty paragraph styles and a chart mediator.
    Numbers,
    /// Pages: thirty paragraph styles and an owned preset reference.
    Pages,
    /// Keynote: twenty-five paragraph styles.
    Keynote,
}

impl Profile {
    const fn paragraph_style_count(self) -> usize {
        match self {
            Self::Numbers | Self::Pages => 30,
            Self::Keynote => 25,
        }
    }
}

/// Identifiers of a new chart graph, in the host's allocation order.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Ids {
    /// The chart drawable.
    pub drawable: u64,
    /// The caption stand-in.
    pub caption: u64,
    /// The title stand-in.
    pub title: u64,
    /// The Numbers chart mediator.
    pub mediator: Option<u64>,
    /// The chart's theme style preset.
    pub preset: u64,
    chart_style: u64,
    chart_non_style: u64,
    legend_style: u64,
    legend_non_style: u64,
    value_axis_styles: [u64; 2],
    value_axis_non_styles: [u64; 2],
    category_axis_style: u64,
    category_axis_non_style: u64,
    series_styles: [u64; 6],
}

impl Ids {
    /// Take the graph's identifiers from `take`, in the host's order.
    ///
    /// # Errors
    ///
    /// The first error `take` returns.
    pub fn allocate<E>(
        profile: Profile,
        mut take: impl FnMut() -> Result<u64, E>,
    ) -> Result<Self, E> {
        let drawable = take()?;
        let caption = take()?;
        let title = take()?;
        let mediator = match profile {
            Profile::Numbers => Some(take()?),
            Profile::Pages | Profile::Keynote => None,
        };
        Ok(Self {
            drawable,
            caption,
            title,
            mediator,
            preset: take()?,
            chart_style: take()?,
            chart_non_style: take()?,
            legend_style: take()?,
            legend_non_style: take()?,
            value_axis_styles: [take()?, take()?],
            value_axis_non_styles: [take()?, take()?],
            category_axis_style: take()?,
            category_axis_non_style: take()?,
            series_styles: [take()?, take()?, take()?, take()?, take()?, take()?],
        })
    }

    /// The sixteen styles, in the order they join the stylesheet.
    #[must_use]
    pub fn styles(&self) -> Vec<u64> {
        let mut styles = vec![
            self.chart_style,
            self.chart_non_style,
            self.legend_style,
            self.legend_non_style,
        ];
        styles.extend(self.value_axis_styles);
        styles.extend(self.value_axis_non_styles);
        styles.extend([self.category_axis_style, self.category_axis_non_style]);
        styles.extend(self.series_styles);
        styles
    }

    /// Every identifier of the graph, in allocation order.
    #[must_use]
    pub fn all(&self) -> Vec<u64> {
        let mut all = vec![self.drawable, self.caption, self.title];
        all.extend(self.mediator);
        all.push(self.preset);
        all.extend(self.styles());
        all
    }

    /// The last identifier of the graph.
    #[must_use]
    pub const fn last(&self) -> u64 {
        self.series_styles[5]
    }
}

/// One chart to author.
#[derive(Debug, Clone, Copy)]
pub struct Request<'a> {
    /// The application profile.
    pub profile: Profile,
    /// The chart kind.
    pub kind: Kind,
    /// The inline data grid.
    pub data: &'a ChartData,
    /// The chart's `TSD.DrawableArchive` (frame, parent, text wrap, title
    /// and caption references), which each owner writes.
    pub drawable: &'a [u8],
    /// The document stylesheet the styles join.
    pub stylesheet: u64,
    /// The theme's first paragraph style preset.
    pub paragraph_style: u64,
}

/// The number of visible series of `data` by rows (host
/// `chart_series_count`); `None` when the data cannot form the kind.
#[must_use]
pub fn series_count(kind: Kind, data: &ChartData) -> Option<usize> {
    let rows = data.row_names().len();
    let count = match kind {
        Kind::Scatter2d => (rows == 2).then_some(1)?,
        Kind::Bubble2d => (rows == 3).then_some(1)?,
        Kind::MultiDataScatter2d | Kind::MultiDataBubble2d => 1,
        _ => rows,
    };
    (count > 0).then_some(count)
}

/// The objects of the new graph, in the host's insertion order.
///
/// # Errors
///
/// [`Error::InvalidData`] when the data cannot form the kind or its grid
/// cannot be encoded; [`Error::Object`] when an object exceeds `limits`.
pub fn objects(
    ids: &Ids,
    request: &Request<'_>,
    limits: Limits,
) -> Result<Vec<ArchiveObject>, Error> {
    let series = series_count(request.kind, request.data).ok_or(Error::InvalidData)?;
    let grid_request = grid_codec::ChartGridCreationRequest::new(
        request.data.row_names(),
        request.data.column_names(),
        request.data.values(),
        ids.drawable,
    );
    let options = grid_codec::EncodeOptions::for_request(&grid_request);
    let grid = grid_codec::encode_chart_grid(grid_request, options)
        .map_err(|_error| Error::InvalidData)?;
    let paragraph = request.paragraph_style;
    let paragraph_count = request.profile.paragraph_style_count();

    #[allow(
        clippy::cast_sign_loss,
        reason = "chart kinds are positive enum values"
    )]
    let mut unity = Writer::new()
        .varint(1, request.kind.native_value() as u64)
        .varint(2, SCATTER_FORMAT_SHARED_X)
        .reference(4, ids.preset)
        .varint(5, SERIES_DIRECTION_BY_ROW)
        .bytes(7, grid.bytes());
    if let Some(mediator) = ids.mediator {
        unity = unity.reference(8, mediator);
    }
    unity = unity
        .reference(9, ids.chart_style)
        .reference(10, ids.chart_non_style)
        .reference(11, ids.legend_style)
        .reference(12, ids.legend_non_style);
    for identifier in ids.value_axis_styles {
        unity = unity.reference(13, identifier);
    }
    for identifier in ids.value_axis_non_styles {
        unity = unity.reference(14, identifier);
    }
    unity = unity
        .reference(15, ids.category_axis_style)
        .reference(16, ids.category_axis_non_style);
    for identifier in ids.series_styles {
        unity = unity.reference(17, identifier);
    }
    unity = unity
        .message(18, &Writer::new().varint(1, 0))
        .message(19, &Writer::new().varint(1, 0));
    for _ in 0..paragraph_count {
        unity = unity.reference(20, paragraph);
    }
    unity = unity.varint(21, 0).boolean(22, false);
    if request.profile == Profile::Pages {
        unity = unity.reference(23, ids.preset);
    }
    unity = unity.boolean(24, false);
    for number in CHART_BOOL_EXTENSIONS {
        unity = unity.boolean(number, true);
    }
    unity = unity.message(CHART_CACHED_FORMATTERS, &cached_formatters(series));
    let drawable = Writer::new()
        .bytes(1, request.drawable)
        .message(EXTENSION, &unity)
        .finish();

    let mut chart_references = vec![ids.caption, ids.title];
    chart_references.extend(ids.mediator);
    chart_references.push(ids.preset);
    chart_references.extend(ids.styles());
    chart_references.push(paragraph);
    let mut preset = Writer::new()
        .reference(1, ids.chart_style)
        .reference(2, ids.legend_style);
    for identifier in ids.value_axis_styles {
        preset = preset.reference(3, identifier);
    }
    preset = preset.reference(4, ids.category_axis_style);
    for identifier in ids.series_styles {
        preset = preset.reference(5, identifier);
    }
    for _ in 0..paragraph_count {
        preset = preset.reference(6, paragraph);
    }
    let mut preset_references = vec![ids.chart_style, ids.legend_style];
    preset_references.extend(ids.value_axis_styles);
    preset_references.push(ids.category_axis_style);
    preset_references.extend(ids.series_styles);
    preset_references.push(paragraph);

    let mut objects = Vec::with_capacity(ids.all().len());
    objects.push(object(
        ids.drawable,
        CHART_MESSAGE_TYPE,
        drawable,
        STANDARD_VERSION,
        &chart_references,
        limits,
    )?);
    objects.push(object(
        ids.caption,
        STANDIN_MESSAGE_TYPE,
        Vec::new(),
        STANDIN_VERSION,
        &[],
        limits,
    )?);
    objects.push(object(
        ids.title,
        STANDIN_MESSAGE_TYPE,
        Vec::new(),
        STANDIN_VERSION,
        &[],
        limits,
    )?);
    if let Some(mediator) = ids.mediator {
        objects.push(object(
            mediator,
            MEDIATOR_MESSAGE_TYPE,
            mediator_payload(mediator),
            STANDARD_VERSION,
            &[],
            limits,
        )?);
    }
    objects.push(object(
        ids.preset,
        PRESET_MESSAGE_TYPE,
        preset.finish(),
        STANDARD_VERSION,
        &preset_references,
        limits,
    )?);
    styles(ids, request.stylesheet, limits, &mut objects)?;
    Ok(objects)
}

fn styles(
    ids: &Ids,
    stylesheet: u64,
    limits: Limits,
    objects: &mut Vec<ArchiveObject>,
) -> Result<(), Error> {
    let gaps = Spacing::DEFAULT;
    let axis_style = Writer::new()
        .boolean(24, true)
        .boolean(25, true)
        .boolean(26, true)
        .boolean(27, false)
        .boolean(28, true)
        .boolean(31, true);
    let axis_non_style = Writer::new()
        .varint(5, 5)
        .varint(6, 1)
        .boolean(9, true)
        .boolean(10, true)
        .boolean(11, true);
    let both = [SUPPORTS_PRIMARY, SUPPORTS_SECONDARY];
    let style = |identifier: u64, message_type: u32, extension: &Writer, features: &[u32]| {
        object(
            identifier,
            message_type,
            style_payload(stylesheet, extension, features),
            STANDARD_VERSION,
            &[],
            limits,
        )
    };
    objects.push(style(
        ids.chart_style,
        CHART_STYLE,
        &Writer::new()
            .float(15, 1.0)
            .float(16, gaps.between_items().percent())
            .float(17, gaps.between_sets().percent())
            .boolean(18, false),
        &[SUPPORTS_PRIMARY],
    )?);
    objects.push(style(
        ids.chart_non_style,
        CHART_NON_STYLE,
        &Writer::new()
            .boolean(20, true)
            .boolean(21, false)
            .boolean(22, false),
        &[],
    )?);
    objects.push(style(
        ids.legend_style,
        LEGEND_STYLE,
        &Writer::new().float(3, 1.0),
        &[],
    )?);
    objects.push(style(
        ids.legend_non_style,
        LEGEND_NON_STYLE,
        &Writer::new(),
        &[],
    )?);
    for identifier in ids.value_axis_styles {
        objects.push(style(
            identifier,
            AXIS_STYLE,
            &axis_style,
            &[SUPPORTS_PRIMARY],
        )?);
    }
    for identifier in ids.value_axis_non_styles {
        objects.push(style(identifier, AXIS_NON_STYLE, &axis_non_style, &both)?);
    }
    objects.push(style(
        ids.category_axis_style,
        AXIS_STYLE,
        &axis_style,
        &[SUPPORTS_PRIMARY],
    )?);
    objects.push(style(
        ids.category_axis_non_style,
        AXIS_NON_STYLE,
        &axis_non_style,
        &both,
    )?);
    for (index, identifier) in ids.series_styles.into_iter().enumerate() {
        objects.push(style(
            identifier,
            SERIES_STYLE,
            &series_style_extension(index),
            &both,
        )?);
    }
    Ok(())
}

fn object(
    identifier: u64,
    message_type: u32,
    payload: Vec<u8>,
    versions: &[u32],
    references: &[u64],
    limits: Limits,
) -> Result<ArchiveObject, Error> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: payload,
        }],
        limits,
    )
    .map_err(|_error| Error::Object)?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::Object)?;
    info.versions = versions.to_vec().into();
    info.object_references = references.to_vec();
    Ok(object)
}

/// The host's deterministic mediator entity identifier.
fn mediator_payload(identifier: u64) -> Vec<u8> {
    let suffix = identifier & 0x0000_ffff_ffff_ffff;
    let entity = format!("00000000-0000-4000-8000-{suffix:012X}");
    // TN.ChartMediatorArchive: super (TSCH.ChartMediatorArchive with one
    // local sentinel and one remote index, unpacked), entity_id, formulas
    // (direction and scheme).
    Writer::new()
        .message(
            1,
            &Writer::new()
                .varint(2, MEDIATOR_LOCAL_SERIES_SENTINEL)
                .varint(3, 0),
        )
        .bytes(2, entity.as_bytes())
        .message(3, &Writer::new().varint(5, 0).varint(10, 0))
        .finish()
}

fn color(red: f32, green: f32, blue: f32) -> Writer {
    // TSP.Color in field-number order: model, r, g, b, a, rgbspace.
    Writer::new()
        .varint(1, 1)
        .float(3, red)
        .float(4, green)
        .float(5, blue)
        .float(6, 1.0)
        .varint(12, 1)
}

fn series_style_extension(index: usize) -> Writer {
    let (red, green, blue) = SERIES_COLORS[index];
    let fill = Writer::new().message(1, &color(red, green, blue));
    let mut pattern = Writer::new().varint(1, 1).float(2, 0.0).varint(3, 0);
    for _ in 0..STROKE_PATTERN_CAPACITY {
        pattern = pattern.float(4, 0.0);
    }
    let stroke = Writer::new()
        .message(1, &color(red, green, blue))
        .float(2, RADAR_STROKE_WIDTH)
        .varint(3, 0)
        .varint(4, 0)
        .float(5, MITER_LIMIT)
        .message(6, &pattern);
    // Field-number order: area, bar, column, default, pie and radar-area
    // fills, the radar stroke, its stroke flag and alpha multiplier.
    Writer::new()
        .message(11, &fill)
        .message(12, &fill)
        .message(13, &fill)
        .message(14, &fill)
        .message(17, &fill)
        .message(165, &fill)
        .message(172, &stroke)
        .boolean(188, true)
        .float(189, RADAR_FILL_STROKE_ALPHA)
}

fn style_payload(stylesheet: u64, extension: &Writer, features: &[u32]) -> Vec<u8> {
    let mut payload = Writer::new()
        .message(1, &Writer::new().reference(5, stylesheet))
        .message(EXTENSION, extension);
    for feature in features {
        payload = payload.varint(*feature, 1);
    }
    payload.finish()
}

fn number_format() -> Writer {
    Writer::new()
        .varint(1, NUMBER_FORMAT_TYPE)
        .varint(2, AUTOMATIC_DECIMAL_PLACES)
        .varint(4, 0)
        .boolean(5, true)
}

fn cached_formatters(series: usize) -> Writer {
    let mut formatters = Writer::new();
    for axis in [1, 2] {
        formatters = formatters.message(
            1,
            &Writer::new()
                .message(1, &Writer::new().varint(1, axis).varint(2, 0))
                .message(2, &number_format()),
        );
    }
    for index in 0..series {
        formatters = formatters.message(
            2,
            &Writer::new()
                .varint(1, index as u64)
                .message(2, &number_format()),
        );
    }
    formatters.message(3, &number_format())
}

/// A minimal canonical protobuf writer (fields in call order).
struct Writer(Vec<u8>);

impl Writer {
    const fn new() -> Self {
        Self(Vec::new())
    }

    fn key(&mut self, number: u32, wire_type: u8) {
        self.raw(u64::from(number) << 3 | u64::from(wire_type));
    }

    fn raw(&mut self, mut value: u64) {
        loop {
            #[allow(
                clippy::cast_possible_truncation,
                reason = "the value is masked to seven bits"
            )]
            let byte = (value & 0x7f) as u8;
            value >>= 7;
            if value == 0 {
                self.0.push(byte);
                return;
            }
            self.0.push(byte | 0x80);
        }
    }

    fn varint(mut self, number: u32, value: u64) -> Self {
        self.key(number, 0);
        self.raw(value);
        self
    }

    fn boolean(self, number: u32, value: bool) -> Self {
        self.varint(number, u64::from(value))
    }

    fn float(mut self, number: u32, value: f32) -> Self {
        self.key(number, 5);
        self.0.extend_from_slice(&value.to_le_bytes());
        self
    }

    fn bytes(mut self, number: u32, value: &[u8]) -> Self {
        self.key(number, 2);
        self.raw(value.len() as u64);
        self.0.extend_from_slice(value);
        self
    }

    fn message(self, number: u32, inner: &Self) -> Self {
        self.bytes(number, &inner.0)
    }

    fn reference(self, number: u32, identifier: u64) -> Self {
        self.message(number, &Self::new().varint(1, identifier))
    }

    fn finish(self) -> Vec<u8> {
        self.0
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn data() -> ChartData {
        ChartData::new(
            vec!["a".to_owned(), "b".to_owned()],
            vec!["x".to_owned()],
            vec![vec![Some(1.0)], vec![Some(2.0)]],
        )
        .expect("data")
    }

    #[test]
    fn series_follow_the_host_topology() {
        let data = data();
        assert_eq!(series_count(Kind::Bar2d, &data), Some(2));
        assert_eq!(series_count(Kind::Scatter2d, &data), Some(1));
        assert_eq!(series_count(Kind::Bubble2d, &data), None);
    }

    #[test]
    fn identifiers_follow_the_host_profile_order() {
        let mut next = 100_u64;
        let numbers = Ids::allocate(Profile::Numbers, || -> Result<u64, ()> {
            next += 1;
            Ok(next)
        })
        .expect("ids");
        assert_eq!(numbers.mediator, Some(104));
        assert_eq!(numbers.preset, 105);
        assert_eq!(numbers.all().len(), 21);
        assert_eq!(numbers.last(), 121);
        let mut next = 0_u64;
        let pages = Ids::allocate(Profile::Pages, || -> Result<u64, ()> {
            next += 1;
            Ok(next)
        })
        .expect("ids");
        assert_eq!(pages.mediator, None);
        assert_eq!(pages.all(), (1..=20).collect::<Vec<_>>());
        assert_eq!(pages.styles().len(), 16);
    }

    #[test]
    fn every_profile_authors_the_whole_graph() {
        let data = data();
        for profile in [Profile::Numbers, Profile::Pages, Profile::Keynote] {
            let mut next = 10_u64;
            let ids = Ids::allocate(profile, || -> Result<u64, ()> {
                next += 1;
                Ok(next)
            })
            .expect("ids");
            let request = Request {
                profile,
                kind: Kind::Bar2d,
                data: &data,
                drawable: &[],
                stylesheet: 1,
                paragraph_style: 2,
            };
            let objects = objects(&ids, &request, Limits::default()).expect("objects");
            let identifiers: Vec<u64> = objects
                .iter()
                .filter_map(|object| object.archive_info.identifier)
                .collect();
            assert_eq!(identifiers, ids.all());
        }
        let ids = Ids::allocate(Profile::Keynote, || -> Result<u64, ()> { Ok(1) }).expect("ids");
        let request = Request {
            profile: Profile::Keynote,
            kind: Kind::Bubble2d,
            data: &data,
            drawable: &[],
            stylesheet: 1,
            paragraph_style: 2,
        };
        assert_eq!(
            objects(&ids, &request, Limits::default()).err(),
            Some(Error::InvalidData)
        );
    }
}
