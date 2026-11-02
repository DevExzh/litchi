//! Shared chart style engine (wave W1-E, record 0375): chart-level inspector
//! properties through the host editors and through the leaf engine — the
//! `litchi_iwa_common::{chart, chart3d}` values and the
//! `litchi_iwa_protos::chart_style_codec` wire codec — on the same input.
//!
//! The leaf route here is the engine the format families (W2-A/B/C) wrap: it
//! locates the chart's style object through the codec's reference
//! projection, maps the common value to codec writes (the table in
//! [`writes`]), rewrites the payload and republishes the one changed IWA
//! component through `litchi_iwa_archive`'s exact-source reassembly. The
//! chart is selected by the host's identifier of a chart the host added to a
//! native-saved fixture; selector resolution belongs to the families.

use litchi_iwa::charts::{
    Chart3dDepth, Chart3dRotation, Chart3dSeriesGap, ChartCornerRadius, ChartDonutInnerRadius,
    ChartPieLabelDistance, ChartPieStartAngle, ChartPieWedgeExplosion, ChartRadarStartAngle,
    ChartRoundedCorners,
};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_archive::Limits as ArchiveLimits;
use litchi_iwa_archive::iwa::{
    Archive as IwaArchive, Limits as CoreLimits, RawMessage, SnappyLimits, SnappyStream,
};
use litchi_iwa_archive::package::Catalog as PackageCatalog;
use litchi_iwa_archive::unstable::EntryEdit;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::{ChartData, corner, donut, legend, pie, radar};
use litchi_iwa_common::chart3d::{self, Scale};
use litchi_iwa_protos::chart_style_codec::{
    self as codec, Archive, EmptyExtension, Limits, Properties, Write, field, property,
};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const POSITION: DrawablePoint = DrawablePoint { x: 120.0, y: 140.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 360.0,
    height: 240.0,
};

/// Where the host added the chart.
#[derive(Clone, Copy, Debug)]
pub(crate) enum Target {
    Keynote { slide: usize },
    Numbers { sheet: u64 },
    Pages,
}

/// A native-saved fixture with one chart of a given kind added by the host.
pub(crate) struct Input {
    pub(crate) fixture: Fixture,
    pub(crate) format: Format,
    pub(crate) target: Target,
    pub(crate) chart: u64,
}

fn pie_data() -> TestResult<ChartData> {
    Ok(ChartData::new(
        vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
        vec!["Revenue".to_owned()],
        vec![vec![Some(12.0)], vec![Some(18.0)], vec![Some(24.0)]],
    )?)
}

fn series_data() -> TestResult<ChartData> {
    Ok(ChartData::new(
        vec!["Region 1".to_owned(), "Region 2".to_owned()],
        vec!["April".to_owned(), "May".to_owned(), "June".to_owned()],
        vec![
            vec![Some(17.0), Some(26.0), Some(53.0)],
            vec![Some(55.0), Some(43.0), Some(70.0)],
        ],
    )?)
}

fn data_for(kind: Kind) -> TestResult<ChartData> {
    if kind.supports_pie_start_angle() {
        pie_data()
    } else {
        series_data()
    }
}

pub(crate) fn input(format: Format, kind: Kind) -> TestResult<Input> {
    let (path, extension) = match format {
        Format::Keynote => ("iwork/keynote/chart-data-native.key", "key"),
        Format::Numbers => ("iwork/numbers/chart-data-native.numbers", "numbers"),
        Format::Pages => ("iwork/pages/chart-data-native.pages", "pages"),
    };
    let native = fixture(path)?;
    let data = data_for(kind)?;
    let (bytes, target, chart) = match format {
        Format::Keynote => {
            let mut editor = keynote::Draft::from_bytes(&native.bytes)?;
            let chart = editor.add_slide_chart(0, kind, data, POSITION, SIZE)?;
            (
                editor.to_bytes()?,
                Target::Keynote { slide: 0 },
                chart.drawable_object_id,
            )
        },
        Format::Numbers => {
            let mut editor = numbers::Draft::from_bytes(&native.bytes)?;
            let sheet = editor.sheets()?.first().ok_or("no sheet")?.id();
            let chart = editor.add_sheet_chart(sheet, kind, data, POSITION, SIZE)?;
            (
                editor.to_bytes()?,
                Target::Numbers { sheet },
                chart.drawable_object_id,
            )
        },
        Format::Pages => {
            let mut editor = pages::Draft::from_bytes(&native.bytes)?;
            let chart = editor.add_body_chart(0, kind, data, POSITION, SIZE)?;
            (editor.to_bytes()?, Target::Pages, chart.drawable_object_id)
        },
    };
    let _ = extension;
    Ok(Input {
        fixture: Fixture { path, bytes },
        format,
        target,
        chart,
    })
}

/// One chart-level property value, as the common vocabulary spells it.
#[derive(Clone, Copy, Debug)]
enum Property {
    PieStartAngle(pie::StartAngle),
    DonutInnerRadius(donut::InnerRadius),
    RadarStartAngle(radar::StartAngle),
    SeriesGap(chart3d::SeriesGap),
    Rotation(chart3d::Rotation),
    Depth(chart3d::Depth, Scale),
    Rounding(corner::Rounding),
}

impl Property {
    /// The native default of the same property (what a reset writes).
    fn native_default(self) -> Self {
        match self {
            Self::PieStartAngle(_) => Self::PieStartAngle(pie::StartAngle::ZERO),
            Self::DonutInnerRadius(_) => Self::DonutInnerRadius(donut::InnerRadius::DEFAULT),
            Self::RadarStartAngle(_) => Self::RadarStartAngle(radar::StartAngle::ZERO),
            Self::SeriesGap(_) => Self::SeriesGap(chart3d::SeriesGap::NATIVE_DEFAULT),
            Self::Rotation(_) => Self::Rotation(chart3d::Rotation::DEFAULT),
            Self::Depth(_, scale) => Self::Depth(chart3d::Depth::DEFAULT, scale),
            Self::Rounding(_) => Self::Rounding(corner::Rounding::NONE),
        }
    }

    const fn archive(self) -> Archive {
        match self {
            Self::Rounding(_) => Archive::ChartStyle,
            _ => Archive::ChartNonStyle,
        }
    }
}

const fn scale_field(scale: Scale) -> u32 {
    match scale {
        Scale::Area => field::chart_non_style::AREA_SCALE_3D,
        Scale::Bar => field::chart_non_style::BAR_SCALE_3D,
        Scale::Column => field::chart_non_style::COLUMN_SCALE_3D,
        Scale::Line => field::chart_non_style::LINE_SCALE_3D,
        _ => field::chart_non_style::PIE_SCALE_3D,
    }
}

/// The engine's read of a property from decoded style properties.
fn read(properties: &Properties<'_>, property: Property) -> TestResult<Property> {
    use field::{chart_non_style as non_style, chart_style as style};
    Ok(match property {
        Property::PieStartAngle(_) => Property::PieStartAngle(
            properties
                .float(non_style::PIE_START_ANGLE)?
                .map(pie::StartAngle::from_degrees)
                .transpose()?
                .unwrap_or_default(),
        ),
        Property::DonutInnerRadius(_) => Property::DonutInnerRadius(
            properties
                .float(non_style::DONUT_INNER_RADIUS)?
                .map(donut::InnerRadius::from_native)
                .transpose()?
                .unwrap_or_default(),
        ),
        Property::RadarStartAngle(_) => Property::RadarStartAngle(
            properties
                .float(non_style::RADAR_START_ANGLE)?
                .map(radar::StartAngle::from_degrees)
                .transpose()?
                .unwrap_or_default(),
        ),
        Property::SeriesGap(_) => Property::SeriesGap(
            properties
                .float(non_style::SERIES_GAP_3D)?
                .map(chart3d::SeriesGap::from_native_ratio)
                .transpose()?
                .unwrap_or_default(),
        ),
        Property::Rotation(_) => Property::Rotation(
            properties
                .vector(non_style::ROTATION_3D)?
                .map(|v| chart3d::Rotation::from_native([v.x, v.y, v.z, v.w]))
                .transpose()?
                .unwrap_or_default(),
        ),
        Property::Depth(_, scale) => Property::Depth(
            properties
                .vector(scale_field(scale))?
                .map(|v| chart3d::Depth::from_native(scale, [v.x, v.y, v.z, v.w]))
                .transpose()?
                .unwrap_or_default(),
            scale,
        ),
        Property::Rounding(_) => Property::Rounding(corner::Rounding::new(
            properties
                .float(style::ROUNDED_CORNER_RADIUS)?
                .map(corner::Radius::from_native)
                .transpose()?
                .unwrap_or(corner::Radius::ZERO),
            properties
                .bool(style::ROUNDED_CORNER_OUTER_END_ONLY)?
                .unwrap_or(false),
        )),
    })
}

fn same(left: Property, right: Property) -> bool {
    match (left, right) {
        (Property::PieStartAngle(a), Property::PieStartAngle(b)) => a == b,
        (Property::DonutInnerRadius(a), Property::DonutInnerRadius(b)) => a == b,
        (Property::RadarStartAngle(a), Property::RadarStartAngle(b)) => a == b,
        (Property::SeriesGap(a), Property::SeriesGap(b)) => a == b,
        (Property::Rotation(a), Property::Rotation(b)) => a == b,
        (Property::Depth(a, _), Property::Depth(b, _)) => a == b,
        (Property::Rounding(a), Property::Rounding(b)) => a == b,
        _ => false,
    }
}

const fn scale3d(scale: Scale) -> property::Scale3d {
    match scale {
        Scale::Area => property::Scale3d::Area,
        Scale::Bar => property::Scale3d::Bar,
        Scale::Column => property::Scale3d::Column,
        Scale::Line => property::Scale3d::Line,
        _ => property::Scale3d::Pie,
    }
}

/// The library write rules (`chart_style_codec::property`) for `value`:
/// what a format family calls with its common value.
fn property_write(value: Property) -> property::PropertyWrite<'static> {
    match value {
        Property::PieStartAngle(angle) => property::pie_start_angle(angle.degrees()),
        Property::DonutInnerRadius(radius) => {
            property::donut_inner_radius(radius.native_fraction())
        },
        Property::RadarStartAngle(angle) => property::radar_start_angle(angle.degrees()),
        Property::SeriesGap(gap) => property::series_gap_3d(gap.native_ratio()),
        Property::Rotation(rotation) => {
            property::rotation_3d(rotation.x_degrees(), rotation.y_degrees())
        },
        Property::Depth(depth, scale) => property::depth_3d(
            scale3d(scale),
            depth.native_ratio(scale),
            (depth != chart3d::Depth::DEFAULT).then(|| depth.native_vector(scale)),
        ),
        Property::Rounding(rounding) => property::rounded_corners(
            rounding.radius().native_fraction(),
            rounding.outside_corners_only(),
        ),
    }
}

/// Every IWA component of a package, decoded once.
pub(crate) struct Components {
    pub(crate) catalog: PackageCatalog,
    pub(crate) archives: Vec<(String, IwaArchive)>,
}

impl Components {
    pub(crate) fn read(source: &[u8]) -> TestResult<Self> {
        // The ZIP member table only: each IWA component is parsed once below.
        let catalog = PackageCatalog::from_bytes(source, ArchiveLimits::default())?;
        let mut archives = Vec::new();
        for entry in catalog.iter() {
            if !entry.name().ends_with(".iwa") {
                continue;
            }
            let stream = SnappyStream::decompress(entry.data(), SnappyLimits::default())?;
            let archive = IwaArchive::parse(stream.as_bytes(), CoreLimits::default())?;
            archives.push((entry.name().to_owned(), archive));
        }
        Ok(Self { catalog, archives })
    }

    /// Every IWA object of every component.
    pub(crate) fn objects(
        &self,
    ) -> impl Iterator<Item = &litchi_iwa_archive::iwa::ArchiveObject> + Clone {
        self.archives
            .iter()
            .flat_map(|(_, archive)| archive.objects.iter())
    }

    /// The length of a package member, if present.
    pub(crate) fn member_size(&self, name: &str) -> Option<u64> {
        self.catalog
            .iter()
            .find(|entry| entry.name() == name)
            .map(|entry| entry.data().len() as u64)
    }

    /// `(component index, message index)` of the unique `message_type`
    /// message of object `identifier`.
    pub(crate) fn locate(&self, identifier: u64, message_type: u32) -> TestResult<(usize, usize)> {
        let mut found = None;
        for (component, (_, archive)) in self.archives.iter().enumerate() {
            let Some(object) = archive.object(identifier) else {
                continue;
            };
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == message_type && found.replace((component, index)).is_some() {
                    return Err(
                        format!("object {identifier} has several {message_type} messages").into(),
                    );
                }
            }
        }
        found.ok_or_else(|| format!("object {identifier} has no {message_type} message").into())
    }

    pub(crate) fn payload(
        &self,
        identifier: u64,
        (component, index): (usize, usize),
    ) -> TestResult<&[u8]> {
        let object = self.archives[component]
            .1
            .object(identifier)
            .ok_or("object vanished")?;
        Ok(object.messages[index].data.as_slice())
    }

    /// The style object of `archive` for chart drawable `chart`.
    fn style_object(&self, chart: u64, archive: Archive) -> TestResult<u64> {
        let drawable = self.locate(chart, CHART_MESSAGE_TYPE)?;
        let references = codec::references(self.payload(chart, drawable)?, Limits::DEFAULT)?;
        Ok(references
            .chart_level(archive)
            .ok_or("the chart has no reference to the style archive")?)
    }
}

/// Replace the payload of message `message_type` of each object and
/// republish every changed component once through exact-source reassembly.
pub(crate) fn publish(
    source: &[u8],
    components: Components,
    replacements: Vec<(u64, u32, Vec<u8>)>,
) -> TestResult<Vec<u8>> {
    publish_referencing(
        source,
        components,
        replacements
            .into_iter()
            .map(|(object, message_type, data)| (object, message_type, data, Vec::new()))
            .collect(),
    )
}

/// [`publish`], also adding object references to each replaced message's
/// metadata (a fixture that makes another object consume a style).
pub(crate) fn publish_referencing(
    source: &[u8],
    mut components: Components,
    replacements: Vec<(u64, u32, Vec<u8>, Vec<u64>)>,
) -> TestResult<Vec<u8>> {
    let mut changed = std::collections::BTreeSet::new();
    for (object, message_type, data, references) in replacements {
        let (component, index) = components.locate(object, message_type)?;
        if components.payload(object, (component, index))? == data.as_slice()
            && references.is_empty()
        {
            continue;
        }
        let target = components.archives[component]
            .1
            .object_mut(object)
            .ok_or("object vanished")?;
        target.replace_message_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
            CoreLimits::default(),
        )?;
        let info = target
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or("message metadata vanished")?;
        for reference in references {
            if !info.object_references.contains(&reference) {
                info.object_references.push(reference);
            }
        }
        changed.insert(component);
    }
    if changed.is_empty() {
        return Ok(source.to_vec());
    }
    let mut compressed = Vec::new();
    for component in changed {
        let (name, archive) = &components.archives[component];
        let bytes = archive.to_bytes(CoreLimits::default())?;
        compressed.push((
            name.clone(),
            SnappyStream::compress(&bytes, SnappyLimits::default())?,
        ));
    }
    let edits: Vec<EntryEdit<'_>> = compressed
        .iter()
        .map(|(name, data)| EntryEdit::new(name.as_str(), data.as_slice()))
        .collect();
    let prepared = components
        .catalog
        .prepare_reassembly(&edits, ArchiveLimits::default())?;
    let requirements = prepared.execution_requirements();
    Ok(prepared.execute(requirements.exact_limits())?)
}

/// Engine read of one chart property.
fn engine_read(bytes: &[u8], chart: u64, property: Property) -> TestResult<Property> {
    let components = Components::read(bytes)?;
    let archive = property.archive();
    let style = components.style_object(chart, archive)?;
    let location = components.locate(style, archive.message_type())?;
    let properties = codec::decode(
        components.payload(style, location)?,
        archive,
        Limits::DEFAULT,
    )?;
    read(&properties, property)
}

/// Rewrite style objects and republish every changed component once.
/// Each edit names the object, its archive, the writes and the
/// empty-extension policy; unchanged payloads leave the source as is.
fn engine_rewrite(
    source: &[u8],
    components: Components,
    edits: &[(u64, Archive, Vec<Write>, EmptyExtension)],
) -> TestResult<Vec<u8>> {
    let mut replacements = Vec::new();
    for (object, archive, writes, empty) in edits {
        let location = components.locate(*object, archive.message_type())?;
        let payload = components.payload(*object, location)?;
        let rewritten = codec::rewrite(payload, *archive, writes, *empty, Limits::DEFAULT)?;
        if rewritten.changed() {
            replacements.push((
                *object,
                archive.message_type(),
                rewritten.into_bytes().into_owned(),
            ));
        }
    }
    publish(source, components, replacements)
}

/// Engine write of one chart property; an equal value is an exact no-op.
fn engine_write(source: &[u8], chart: u64, property: Property) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let archive = property.archive();
    let style = components.style_object(chart, archive)?;
    let location = components.locate(style, archive.message_type())?;
    let properties = codec::decode(
        components.payload(style, location)?,
        archive,
        Limits::DEFAULT,
    )?;
    if same(read(&properties, property)?, property) {
        return Ok(source.to_vec());
    }
    let write = property_write(property);
    assert_eq!(write.archive(), archive);
    let edits = [(style, archive, write.writes().to_vec(), write.empty())];
    engine_rewrite(source, components, &edits)
}

/// The host value of a property.
enum HostValue {
    PieStartAngle(ChartPieStartAngle),
    DonutInnerRadius(ChartDonutInnerRadius),
    RadarStartAngle(ChartRadarStartAngle),
    SeriesGap(Chart3dSeriesGap),
    Rotation(Chart3dRotation),
    Depth(Chart3dDepth),
    Rounding(ChartRoundedCorners),
}

fn host_value(property: Property) -> TestResult<HostValue> {
    Ok(match property {
        Property::PieStartAngle(angle) => {
            HostValue::PieStartAngle(ChartPieStartAngle::from_degrees(angle.degrees())?)
        },
        Property::DonutInnerRadius(radius) => {
            HostValue::DonutInnerRadius(ChartDonutInnerRadius::from_percent(radius.percent())?)
        },
        Property::RadarStartAngle(angle) => {
            HostValue::RadarStartAngle(ChartRadarStartAngle::from_degrees(angle.degrees())?)
        },
        Property::SeriesGap(gap) => HostValue::SeriesGap(Chart3dSeriesGap::new(gap.percent())?),
        Property::Rotation(rotation) => HostValue::Rotation(Chart3dRotation::from_degrees(
            rotation.x_degrees(),
            rotation.y_degrees(),
        )?),
        Property::Depth(depth, _) => HostValue::Depth(Chart3dDepth::from_percent(depth.percent())?),
        Property::Rounding(rounding) => HostValue::Rounding(ChartRoundedCorners::new(
            ChartCornerRadius::new(rounding.radius().percent())?,
            rounding.outside_corners_only(),
        )),
    })
}

/// Host write of one chart property through the format's `Draft`.
fn host_write(
    source: &[u8],
    target: Target,
    chart: u64,
    property: Property,
) -> TestResult<Vec<u8>> {
    let value = host_value(property)?;
    macro_rules! apply {
        ($editor:expr, $($scope:expr,)? ; $pie:ident, $donut:ident, $radar:ident, $gap:ident, $rotation:ident, $depth:ident, $rounding:ident) => {{
            let editor = &mut $editor;
            match value {
                HostValue::PieStartAngle(v) => editor.$pie($($scope,)? chart, v)?,
                HostValue::DonutInnerRadius(v) => editor.$donut($($scope,)? chart, v)?,
                HostValue::RadarStartAngle(v) => editor.$radar($($scope,)? chart, v)?,
                HostValue::SeriesGap(v) => editor.$gap($($scope,)? chart, v)?,
                HostValue::Rotation(v) => editor.$rotation($($scope,)? chart, v)?,
                HostValue::Depth(v) => editor.$depth($($scope,)? chart, v)?,
                HostValue::Rounding(v) => editor.$rounding($($scope,)? chart, v)?,
            }
            editor.to_bytes()?
        }};
    }
    Ok(match target {
        Target::Keynote { slide } => apply!(
            keynote::Draft::from_bytes(source)?, slide, ;
            set_slide_chart_pie_start_angle, set_slide_chart_donut_inner_radius,
            set_slide_chart_radar_start_angle, set_slide_chart_3d_series_gap,
            set_slide_chart_3d_rotation, set_slide_chart_3d_depth, set_slide_chart_rounded_corners
        ),
        Target::Numbers { sheet } => apply!(
            numbers::Draft::from_bytes(source)?, sheet, ;
            set_sheet_chart_pie_start_angle, set_sheet_chart_donut_inner_radius,
            set_sheet_chart_radar_start_angle, set_sheet_chart_3d_series_gap,
            set_sheet_chart_3d_rotation, set_sheet_chart_3d_depth, set_sheet_chart_rounded_corners
        ),
        Target::Pages => apply!(
            pages::Draft::from_bytes(source)?, ;
            set_body_chart_pie_start_angle, set_body_chart_donut_inner_radius,
            set_body_chart_radar_start_angle, set_body_chart_3d_series_gap,
            set_body_chart_3d_rotation, set_body_chart_3d_depth, set_body_chart_rounded_corners
        ),
    })
}

/// The host's read of the property, rendered for comparison.
fn host_read(bytes: &[u8], target: Target, chart: u64, property: Property) -> TestResult<String> {
    macro_rules! read_with {
        ($editor:expr, $($scope:expr,)? ; $pie:ident, $donut:ident, $radar:ident, $gap:ident, $rotation:ident, $depth:ident, $rounding:ident) => {{
            let editor = $editor;
            match property {
                Property::PieStartAngle(_) => format!("{:?}", editor.$pie($($scope,)? chart)?),
                Property::DonutInnerRadius(_) => format!("{:?}", editor.$donut($($scope,)? chart)?),
                Property::RadarStartAngle(_) => format!("{:?}", editor.$radar($($scope,)? chart)?),
                Property::SeriesGap(_) => format!("{:?}", editor.$gap($($scope,)? chart)?),
                Property::Rotation(_) => format!("{:?}", editor.$rotation($($scope,)? chart)?),
                Property::Depth(..) => format!("{:?}", editor.$depth($($scope,)? chart)?),
                Property::Rounding(_) => format!("{:?}", editor.$rounding($($scope,)? chart)?),
            }
        }};
    }
    Ok(match target {
        Target::Keynote { slide } => read_with!(
            keynote::Draft::from_bytes(bytes)?, slide, ;
            slide_chart_pie_start_angle, slide_chart_donut_inner_radius,
            slide_chart_radar_start_angle, slide_chart_3d_series_gap,
            slide_chart_3d_rotation, slide_chart_3d_depth, slide_chart_rounded_corners
        ),
        Target::Numbers { sheet } => read_with!(
            numbers::Draft::from_bytes(bytes)?, sheet, ;
            sheet_chart_pie_start_angle, sheet_chart_donut_inner_radius,
            sheet_chart_radar_start_angle, sheet_chart_3d_series_gap,
            sheet_chart_3d_rotation, sheet_chart_3d_depth, sheet_chart_rounded_corners
        ),
        Target::Pages => read_with!(
            pages::Draft::from_bytes(bytes)?, ;
            body_chart_pie_start_angle, body_chart_donut_inner_radius,
            body_chart_radar_start_angle, body_chart_3d_series_gap,
            body_chart_3d_rotation, body_chart_3d_depth, body_chart_rounded_corners
        ),
    })
}

/// Whether a case sets the property on the added chart or resets a value the
/// host set first.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Step {
    Set,
    Reset,
}

/// Set `property` on a chart of `kind` (or, for [`Step::Reset`], reset it to
/// the native default after the host set it) through both routes.
fn style_case(
    name: &'static str,
    format: Format,
    kind: Kind,
    tier: Tier,
    evidence: Evidence,
    property: Property,
    step: Step,
) -> TestResult {
    let mut input = input(format, kind)?;
    let (target, chart) = (input.target, input.chart);
    let property = if step == Step::Reset {
        input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, property)?;
        property.native_default()
    } else {
        property
    };
    Case::new(name, input.format, tier, evidence, &input.fixture)
        .host(|source| host_write(source, target, chart, property))
        .leaf(|source| Ok(LeafRun::output(engine_write(source, chart, property)?)))
        .read("host property read", move |bytes| {
            host_read(bytes, target, chart, property)
        })
        .read("engine property read", move |bytes| {
            Ok(format!("{:?}", engine_read(bytes, chart, property)?))
        })
        .run()?;
    Ok(())
}

fn pie_start_angle() -> TestResult<Property> {
    Ok(Property::PieStartAngle(pie::StartAngle::from_degrees(
        123.5,
    )?))
}

fn donut_inner_radius() -> TestResult<Property> {
    Ok(Property::DonutInnerRadius(
        donut::InnerRadius::from_percent(40.0)?,
    ))
}

fn radar_start_angle() -> TestResult<Property> {
    Ok(Property::RadarStartAngle(radar::StartAngle::from_degrees(
        45.0,
    )?))
}

fn series_gap() -> TestResult<Property> {
    Ok(Property::SeriesGap(chart3d::SeriesGap::new(150)?))
}

fn rotation() -> TestResult<Property> {
    Ok(Property::Rotation(chart3d::Rotation::from_degrees(
        12.0, -30.0,
    )?))
}

fn depth(scale: Scale) -> TestResult<Property> {
    Ok(Property::Depth(chart3d::Depth::from_percent(62.5)?, scale))
}

fn rounding() -> TestResult<Property> {
    Ok(Property::Rounding(corner::Rounding::new(
        corner::Radius::new(35.0)?,
        true,
    )))
}

#[test]
fn keynote_chart_pie_start_angle_parity() -> TestResult {
    style_case(
        "keynote_chart_pie_start_angle_parity",
        Format::Keynote,
        Kind::Pie2d,
        Tier::A,
        Evidence::Queued,
        pie_start_angle()?,
        Step::Set,
    )
}

#[test]
fn numbers_chart_pie_start_angle_parity() -> TestResult {
    style_case(
        "numbers_chart_pie_start_angle_parity",
        Format::Numbers,
        Kind::Pie2d,
        Tier::A,
        Evidence::Queued,
        pie_start_angle()?,
        Step::Set,
    )
}

#[test]
fn pages_chart_pie_start_angle_parity() -> TestResult {
    style_case(
        "pages_chart_pie_start_angle_parity",
        Format::Pages,
        Kind::Pie2d,
        Tier::A,
        Evidence::Queued,
        pie_start_angle()?,
        Step::Set,
    )
}

#[test]
fn keynote_chart_donut_inner_radius_parity() -> TestResult {
    style_case(
        "keynote_chart_donut_inner_radius_parity",
        Format::Keynote,
        Kind::Donut2d,
        Tier::A,
        Evidence::Queued,
        donut_inner_radius()?,
        Step::Set,
    )
}

#[test]
fn numbers_chart_donut_inner_radius_parity() -> TestResult {
    style_case(
        "numbers_chart_donut_inner_radius_parity",
        Format::Numbers,
        Kind::Donut2d,
        Tier::A,
        Evidence::Queued,
        donut_inner_radius()?,
        Step::Set,
    )
}

#[test]
fn pages_chart_donut_inner_radius_parity() -> TestResult {
    style_case(
        "pages_chart_donut_inner_radius_parity",
        Format::Pages,
        Kind::Donut2d,
        Tier::A,
        Evidence::Queued,
        donut_inner_radius()?,
        Step::Set,
    )
}

#[test]
fn keynote_chart_radar_start_angle_parity() -> TestResult {
    style_case(
        "keynote_chart_radar_start_angle_parity",
        Format::Keynote,
        Kind::Radar2d,
        Tier::A,
        Evidence::Queued,
        radar_start_angle()?,
        Step::Set,
    )
}

#[test]
fn numbers_chart_radar_start_angle_parity() -> TestResult {
    style_case(
        "numbers_chart_radar_start_angle_parity",
        Format::Numbers,
        Kind::Radar2d,
        Tier::A,
        Evidence::Queued,
        radar_start_angle()?,
        Step::Set,
    )
}

#[test]
fn pages_chart_radar_start_angle_parity() -> TestResult {
    style_case(
        "pages_chart_radar_start_angle_parity",
        Format::Pages,
        Kind::Radar2d,
        Tier::A,
        Evidence::Queued,
        radar_start_angle()?,
        Step::Set,
    )
}

#[test]
fn keynote_chart_3d_series_gap_parity() -> TestResult {
    style_case(
        "keynote_chart_3d_series_gap_parity",
        Format::Keynote,
        Kind::Line3d,
        Tier::A,
        Evidence::Queued,
        series_gap()?,
        Step::Set,
    )
}

#[test]
fn numbers_chart_3d_series_gap_parity() -> TestResult {
    style_case(
        "numbers_chart_3d_series_gap_parity",
        Format::Numbers,
        Kind::Area3d,
        Tier::A,
        Evidence::Queued,
        series_gap()?,
        Step::Set,
    )
}

#[test]
fn pages_chart_3d_series_gap_parity() -> TestResult {
    style_case(
        "pages_chart_3d_series_gap_parity",
        Format::Pages,
        Kind::Line3d,
        Tier::A,
        Evidence::Queued,
        series_gap()?,
        Step::Set,
    )
}

#[test]
fn keynote_chart_3d_rotation_parity() -> TestResult {
    style_case(
        "keynote_chart_3d_rotation_parity",
        Format::Keynote,
        Kind::Column3d,
        Tier::A,
        Evidence::Queued,
        rotation()?,
        Step::Set,
    )
}

#[test]
fn numbers_chart_3d_rotation_parity() -> TestResult {
    style_case(
        "numbers_chart_3d_rotation_parity",
        Format::Numbers,
        Kind::Pie3d,
        Tier::A,
        Evidence::Queued,
        rotation()?,
        Step::Set,
    )
}

#[test]
fn pages_chart_3d_rotation_parity() -> TestResult {
    style_case(
        "pages_chart_3d_rotation_parity",
        Format::Pages,
        Kind::Bar3d,
        Tier::A,
        Evidence::Queued,
        rotation()?,
        Step::Set,
    )
}

#[test]
fn keynote_chart_3d_depth_parity() -> TestResult {
    style_case(
        "keynote_chart_3d_depth_parity",
        Format::Keynote,
        Kind::Bar3d,
        Tier::A,
        Evidence::Queued,
        depth(Scale::Bar)?,
        Step::Set,
    )
}

#[test]
fn numbers_chart_3d_depth_parity() -> TestResult {
    style_case(
        "numbers_chart_3d_depth_parity",
        Format::Numbers,
        Kind::Column3d,
        Tier::A,
        Evidence::Queued,
        depth(Scale::Column)?,
        Step::Set,
    )
}

#[test]
fn pages_chart_3d_depth_parity() -> TestResult {
    style_case(
        "pages_chart_3d_depth_parity",
        Format::Pages,
        Kind::Pie3d,
        Tier::A,
        Evidence::Queued,
        depth(Scale::Pie)?,
        Step::Set,
    )
}

#[test]
fn keynote_chart_rounded_corners_parity() -> TestResult {
    style_case(
        "keynote_chart_rounded_corners_parity",
        Format::Keynote,
        Kind::Column2d,
        Tier::A,
        Evidence::Queued,
        rounding()?,
        Step::Set,
    )
}

#[test]
fn numbers_chart_rounded_corners_parity() -> TestResult {
    style_case(
        "numbers_chart_rounded_corners_parity",
        Format::Numbers,
        Kind::Bar2d,
        Tier::A,
        Evidence::Queued,
        rounding()?,
        Step::Set,
    )
}

#[test]
fn pages_chart_rounded_corners_parity() -> TestResult {
    style_case(
        "pages_chart_rounded_corners_parity",
        Format::Pages,
        Kind::Column2d,
        Tier::A,
        Evidence::Queued,
        rounding()?,
        Step::Set,
    )
}

#[test]
fn keynote_chart_pie_start_angle_reset_parity() -> TestResult {
    style_case(
        "keynote_chart_pie_start_angle_reset_parity",
        Format::Keynote,
        Kind::Pie2d,
        Tier::A,
        Evidence::Queued,
        pie_start_angle()?,
        Step::Reset,
    )
}

#[test]
fn numbers_chart_pie_start_angle_reset_parity() -> TestResult {
    style_case(
        "numbers_chart_pie_start_angle_reset_parity",
        Format::Numbers,
        Kind::Pie2d,
        Tier::A,
        Evidence::Queued,
        pie_start_angle()?,
        Step::Reset,
    )
}

#[test]
fn pages_chart_pie_start_angle_reset_parity() -> TestResult {
    style_case(
        "pages_chart_pie_start_angle_reset_parity",
        Format::Pages,
        Kind::Pie2d,
        Tier::A,
        Evidence::Queued,
        pie_start_angle()?,
        Step::Reset,
    )
}

#[test]
fn keynote_chart_donut_inner_radius_reset_parity() -> TestResult {
    style_case(
        "keynote_chart_donut_inner_radius_reset_parity",
        Format::Keynote,
        Kind::Donut2d,
        Tier::A,
        Evidence::Queued,
        donut_inner_radius()?,
        Step::Reset,
    )
}

#[test]
fn numbers_chart_donut_inner_radius_reset_parity() -> TestResult {
    style_case(
        "numbers_chart_donut_inner_radius_reset_parity",
        Format::Numbers,
        Kind::Donut2d,
        Tier::A,
        Evidence::Queued,
        donut_inner_radius()?,
        Step::Reset,
    )
}

#[test]
fn pages_chart_donut_inner_radius_reset_parity() -> TestResult {
    style_case(
        "pages_chart_donut_inner_radius_reset_parity",
        Format::Pages,
        Kind::Donut2d,
        Tier::A,
        Evidence::Queued,
        donut_inner_radius()?,
        Step::Reset,
    )
}

#[test]
fn keynote_chart_radar_start_angle_reset_parity() -> TestResult {
    style_case(
        "keynote_chart_radar_start_angle_reset_parity",
        Format::Keynote,
        Kind::Radar2d,
        Tier::A,
        Evidence::Queued,
        radar_start_angle()?,
        Step::Reset,
    )
}

#[test]
fn numbers_chart_radar_start_angle_reset_parity() -> TestResult {
    style_case(
        "numbers_chart_radar_start_angle_reset_parity",
        Format::Numbers,
        Kind::Radar2d,
        Tier::A,
        Evidence::Queued,
        radar_start_angle()?,
        Step::Reset,
    )
}

#[test]
fn pages_chart_radar_start_angle_reset_parity() -> TestResult {
    style_case(
        "pages_chart_radar_start_angle_reset_parity",
        Format::Pages,
        Kind::Radar2d,
        Tier::A,
        Evidence::Queued,
        radar_start_angle()?,
        Step::Reset,
    )
}

#[test]
fn keynote_chart_3d_series_gap_reset_parity() -> TestResult {
    style_case(
        "keynote_chart_3d_series_gap_reset_parity",
        Format::Keynote,
        Kind::Line3d,
        Tier::A,
        Evidence::Queued,
        series_gap()?,
        Step::Reset,
    )
}

#[test]
fn numbers_chart_3d_series_gap_reset_parity() -> TestResult {
    style_case(
        "numbers_chart_3d_series_gap_reset_parity",
        Format::Numbers,
        Kind::Area3d,
        Tier::A,
        Evidence::Queued,
        series_gap()?,
        Step::Reset,
    )
}

#[test]
fn pages_chart_3d_series_gap_reset_parity() -> TestResult {
    style_case(
        "pages_chart_3d_series_gap_reset_parity",
        Format::Pages,
        Kind::Line3d,
        Tier::A,
        Evidence::Queued,
        series_gap()?,
        Step::Reset,
    )
}

#[test]
fn keynote_chart_3d_rotation_reset_parity() -> TestResult {
    style_case(
        "keynote_chart_3d_rotation_reset_parity",
        Format::Keynote,
        Kind::Column3d,
        Tier::A,
        Evidence::Queued,
        rotation()?,
        Step::Reset,
    )
}

#[test]
fn numbers_chart_3d_rotation_reset_parity() -> TestResult {
    style_case(
        "numbers_chart_3d_rotation_reset_parity",
        Format::Numbers,
        Kind::Pie3d,
        Tier::A,
        Evidence::Queued,
        rotation()?,
        Step::Reset,
    )
}

#[test]
fn pages_chart_3d_rotation_reset_parity() -> TestResult {
    style_case(
        "pages_chart_3d_rotation_reset_parity",
        Format::Pages,
        Kind::Bar3d,
        Tier::A,
        Evidence::Queued,
        rotation()?,
        Step::Reset,
    )
}

#[test]
fn keynote_chart_3d_depth_reset_parity() -> TestResult {
    style_case(
        "keynote_chart_3d_depth_reset_parity",
        Format::Keynote,
        Kind::Bar3d,
        Tier::A,
        Evidence::Queued,
        depth(Scale::Bar)?,
        Step::Reset,
    )
}

#[test]
fn numbers_chart_3d_depth_reset_parity() -> TestResult {
    style_case(
        "numbers_chart_3d_depth_reset_parity",
        Format::Numbers,
        Kind::Column3d,
        Tier::A,
        Evidence::Queued,
        depth(Scale::Column)?,
        Step::Reset,
    )
}

#[test]
fn pages_chart_3d_depth_reset_parity() -> TestResult {
    style_case(
        "pages_chart_3d_depth_reset_parity",
        Format::Pages,
        Kind::Pie3d,
        Tier::A,
        Evidence::Queued,
        depth(Scale::Pie)?,
        Step::Reset,
    )
}

#[test]
fn keynote_chart_rounded_corners_reset_parity() -> TestResult {
    style_case(
        "keynote_chart_rounded_corners_reset_parity",
        Format::Keynote,
        Kind::Column2d,
        Tier::A,
        Evidence::Queued,
        rounding()?,
        Step::Reset,
    )
}

#[test]
fn numbers_chart_rounded_corners_reset_parity() -> TestResult {
    style_case(
        "numbers_chart_rounded_corners_reset_parity",
        Format::Numbers,
        Kind::Bar2d,
        Tier::A,
        Evidence::Queued,
        rounding()?,
        Step::Reset,
    )
}

#[test]
fn pages_chart_rounded_corners_reset_parity() -> TestResult {
    style_case(
        "pages_chart_rounded_corners_reset_parity",
        Format::Pages,
        Kind::Column2d,
        Tier::A,
        Evidence::Queued,
        rounding()?,
        Step::Reset,
    )
}

/// A per-wedge (series non-style) property for every wedge in series order.
#[derive(Clone, Debug)]
enum SeriesProperty {
    WedgeExplosions(Vec<pie::WedgeExplosion>),
    LabelDistances(Vec<pie::LabelDistance>),
}

impl SeriesProperty {
    fn len(&self) -> usize {
        match self {
            Self::WedgeExplosions(values) => values.len(),
            Self::LabelDistances(values) => values.len(),
        }
    }

    /// The library write rules for wedge `index`.
    fn write(&self, index: usize) -> property::PropertyWrite<'static> {
        match self {
            Self::WedgeExplosions(values) => {
                property::pie_wedge_explosion(values[index].native_fraction())
            },
            Self::LabelDistances(values) => property::pie_label_distance(values[index].percent()),
        }
    }
}

/// The engine's read of every wedge's value (absent objects read as the
/// default).
fn engine_series_read(bytes: &[u8], chart: u64, template: &SeriesProperty) -> TestResult<String> {
    use field::series_non_style as series;
    let components = Components::read(bytes)?;
    let drawable = components.locate(chart, CHART_MESSAGE_TYPE)?;
    let references = codec::references(components.payload(chart, drawable)?, Limits::DEFAULT)?;
    let non_styles = references.series_non_styles().cloned().unwrap_or_default();
    let mut explosions = Vec::new();
    let mut distances = Vec::new();
    for index in 0..template.len() {
        let index = u32::try_from(index)?;
        let properties = match non_styles.get(index) {
            Some(object) => {
                let location = components.locate(object, Archive::SeriesNonStyle.message_type())?;
                Some(components.payload(object, location)?)
            },
            None => None,
        };
        let decoded = properties
            .map(|payload| codec::decode(payload, Archive::SeriesNonStyle, Limits::DEFAULT))
            .transpose()?;
        let float = |number| -> TestResult<Option<f32>> {
            Ok(match &decoded {
                Some(properties) => properties.float(number)?,
                None => None,
            })
        };
        explosions.push(
            float(series::PIE_WEDGE_EXPLOSION)?
                .map(pie::WedgeExplosion::from_native)
                .transpose()?
                .unwrap_or_default(),
        );
        let legacy = float(series::PIE_LEGACY_LABEL_EXPLOSION)?;
        let modern = float(series::PIE_LABEL_EXPLOSION)?;
        distances.push(
            legacy
                .or(modern)
                .map(pie::LabelDistance::from_percent)
                .transpose()?
                .unwrap_or_default(),
        );
    }
    Ok(match template {
        SeriesProperty::WedgeExplosions(_) => format!("{explosions:?}"),
        SeriesProperty::LabelDistances(_) => format!("{distances:?}"),
    })
}

/// Engine write: rewrite every wedge's existing series non-style object.
fn engine_series_write(
    source: &[u8],
    chart: u64,
    property: &SeriesProperty,
) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let drawable = components.locate(chart, CHART_MESSAGE_TYPE)?;
    let references = codec::references(components.payload(chart, drawable)?, Limits::DEFAULT)?;
    let non_styles = references
        .series_non_styles()
        .ok_or("the chart has no series non-styles")?;
    let mut edits = Vec::new();
    for index in 0..property.len() {
        let object = non_styles.get(u32::try_from(index)?).ok_or(
            "a wedge has no series non-style object; allocation is a format-family graph operation",
        )?;
        let write = property.write(index);
        edits.push((
            object,
            write.archive(),
            write.writes().to_vec(),
            write.empty(),
        ));
    }
    engine_rewrite(source, components, &edits)
}

fn host_series_write(
    source: &[u8],
    target: Target,
    chart: u64,
    property: &SeriesProperty,
) -> TestResult<Vec<u8>> {
    let explosions = |values: &[pie::WedgeExplosion]| -> TestResult<Vec<ChartPieWedgeExplosion>> {
        values
            .iter()
            .map(|value| Ok(ChartPieWedgeExplosion::from_percent(value.percent())?))
            .collect()
    };
    let distances = |values: &[pie::LabelDistance]| -> TestResult<Vec<ChartPieLabelDistance>> {
        values
            .iter()
            .map(|value| Ok(ChartPieLabelDistance::from_percent(value.percent())?))
            .collect()
    };
    Ok(match (target, property) {
        (Target::Keynote { slide }, SeriesProperty::WedgeExplosions(values)) => {
            let mut editor = keynote::Draft::from_bytes(source)?;
            editor.set_slide_chart_pie_wedge_explosions(slide, chart, &explosions(values)?)?;
            editor.to_bytes()?
        },
        (Target::Keynote { slide }, SeriesProperty::LabelDistances(values)) => {
            let mut editor = keynote::Draft::from_bytes(source)?;
            editor.set_slide_chart_pie_label_distances(slide, chart, &distances(values)?)?;
            editor.to_bytes()?
        },
        (Target::Numbers { sheet }, SeriesProperty::WedgeExplosions(values)) => {
            let mut editor = numbers::Draft::from_bytes(source)?;
            editor.set_sheet_chart_pie_wedge_explosions(sheet, chart, &explosions(values)?)?;
            editor.to_bytes()?
        },
        (Target::Numbers { sheet }, SeriesProperty::LabelDistances(values)) => {
            let mut editor = numbers::Draft::from_bytes(source)?;
            editor.set_sheet_chart_pie_label_distances(sheet, chart, &distances(values)?)?;
            editor.to_bytes()?
        },
        (Target::Pages, SeriesProperty::WedgeExplosions(values)) => {
            let mut editor = pages::Draft::from_bytes(source)?;
            editor.set_body_chart_pie_wedge_explosions(chart, &explosions(values)?)?;
            editor.to_bytes()?
        },
        (Target::Pages, SeriesProperty::LabelDistances(values)) => {
            let mut editor = pages::Draft::from_bytes(source)?;
            editor.set_body_chart_pie_label_distances(chart, &distances(values)?)?;
            editor.to_bytes()?
        },
    })
}

fn host_series_read(
    bytes: &[u8],
    target: Target,
    chart: u64,
    property: &SeriesProperty,
) -> TestResult<String> {
    let percents = |values: Vec<f32>| format!("{values:?}");
    Ok(match (target, property) {
        (Target::Keynote { slide }, SeriesProperty::WedgeExplosions(_)) => percents(
            keynote::Draft::from_bytes(bytes)?
                .slide_chart_pie_wedge_explosions(slide, chart)?
                .into_iter()
                .map(ChartPieWedgeExplosion::percent)
                .collect(),
        ),
        (Target::Keynote { slide }, SeriesProperty::LabelDistances(_)) => percents(
            keynote::Draft::from_bytes(bytes)?
                .slide_chart_pie_label_distances(slide, chart)?
                .into_iter()
                .map(ChartPieLabelDistance::percent)
                .collect(),
        ),
        (Target::Numbers { sheet }, SeriesProperty::WedgeExplosions(_)) => percents(
            numbers::Draft::from_bytes(bytes)?
                .sheet_chart_pie_wedge_explosions(sheet, chart)?
                .into_iter()
                .map(ChartPieWedgeExplosion::percent)
                .collect(),
        ),
        (Target::Numbers { sheet }, SeriesProperty::LabelDistances(_)) => percents(
            numbers::Draft::from_bytes(bytes)?
                .sheet_chart_pie_label_distances(sheet, chart)?
                .into_iter()
                .map(ChartPieLabelDistance::percent)
                .collect(),
        ),
        (Target::Pages, SeriesProperty::WedgeExplosions(_)) => percents(
            pages::Draft::from_bytes(bytes)?
                .body_chart_pie_wedge_explosions(chart)?
                .into_iter()
                .map(ChartPieWedgeExplosion::percent)
                .collect(),
        ),
        (Target::Pages, SeriesProperty::LabelDistances(_)) => percents(
            pages::Draft::from_bytes(bytes)?
                .body_chart_pie_label_distances(chart)?
                .into_iter()
                .map(ChartPieLabelDistance::percent)
                .collect(),
        ),
    })
}

/// Change per-wedge values on a pie chart whose wedges the host gave series
/// non-style objects first (`before`), so both routes rewrite existing
/// objects.
fn series_case(
    name: &'static str,
    format: Format,
    tier: Tier,
    evidence: Evidence,
    before: SeriesProperty,
    after: SeriesProperty,
) -> TestResult {
    let mut input = input(format, Kind::Pie2d)?;
    let (target, chart) = (input.target, input.chart);
    input.fixture.bytes = host_series_write(&input.fixture.bytes, target, chart, &before)?;
    let after_ref = &after;
    Case::new(name, input.format, tier, evidence, &input.fixture)
        .host(|source| host_series_write(source, target, chart, after_ref))
        .leaf(|source| {
            Ok(LeafRun::output(engine_series_write(
                source, chart, after_ref,
            )?))
        })
        .read("host per-wedge read", move |bytes| {
            host_series_read(bytes, target, chart, after_ref)
        })
        .read("engine per-wedge read", move |bytes| {
            engine_series_read(bytes, chart, after_ref)
        })
        .run()?;
    Ok(())
}

fn explosions(percents: [f32; 3]) -> TestResult<SeriesProperty> {
    Ok(SeriesProperty::WedgeExplosions(
        percents
            .into_iter()
            .map(pie::WedgeExplosion::from_percent)
            .collect::<Result<_, _>>()?,
    ))
}

fn distances(percents: [f32; 3]) -> TestResult<SeriesProperty> {
    Ok(SeriesProperty::LabelDistances(
        percents
            .into_iter()
            .map(pie::LabelDistance::from_percent)
            .collect::<Result<_, _>>()?,
    ))
}

#[test]
fn keynote_chart_pie_wedge_explosion_parity() -> TestResult {
    series_case(
        "keynote_chart_pie_wedge_explosion_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        explosions([10.0, 20.0, 30.0])?,
        explosions([50.0, 25.0, 75.0])?,
    )
}

#[test]
fn numbers_chart_pie_wedge_explosion_parity() -> TestResult {
    series_case(
        "numbers_chart_pie_wedge_explosion_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        explosions([10.0, 20.0, 30.0])?,
        explosions([50.0, 25.0, 75.0])?,
    )
}

#[test]
fn pages_chart_pie_wedge_explosion_parity() -> TestResult {
    series_case(
        "pages_chart_pie_wedge_explosion_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        explosions([10.0, 20.0, 30.0])?,
        explosions([50.0, 25.0, 75.0])?,
    )
}

#[test]
fn keynote_chart_pie_label_distance_parity() -> TestResult {
    series_case(
        "keynote_chart_pie_label_distance_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        distances([40.0, 90.0, 150.0])?,
        distances([120.0, 35.0, 200.0])?,
    )
}

#[test]
fn numbers_chart_pie_label_distance_parity() -> TestResult {
    series_case(
        "numbers_chart_pie_label_distance_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        distances([40.0, 90.0, 150.0])?,
        distances([120.0, 35.0, 200.0])?,
    )
}

#[test]
fn pages_chart_pie_label_distance_parity() -> TestResult {
    series_case(
        "pages_chart_pie_label_distance_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        distances([40.0, 90.0, 150.0])?,
        distances([120.0, 35.0, 200.0])?,
    )
}

/// Callgrind toggle target: the host route of one chart property set
/// (open, set, the host's own reopen check, serialize).
#[inline(never)]
fn host_style_route(
    source: &[u8],
    target: Target,
    chart: u64,
    property: Property,
) -> TestResult<Vec<u8>> {
    host_write(source, target, chart, property)
}

/// Callgrind toggle target: the engine route of the same set (catalog,
/// component decode, reference projection, codec rewrite, reassembly).
#[inline(never)]
fn leaf_style_route(source: &[u8], chart: u64, property: Property) -> TestResult<Vec<u8>> {
    engine_write(source, chart, property)
}

/// ABBA cost probe of the chart property set route (record 0375 §5):
///
/// ```sh
/// CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --no-run
/// taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*chart_style_cases::host_style_route*' \
///   target/release/deps/parity-<hash> --ignored --exact \
///   chart_style_cases::chart_style_set_cost_abba --test-threads=1
/// # likewise with '--toggle-collect=*chart_style_cases::leaf_style_route*'
/// ```
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn chart_style_set_cost_abba() -> TestResult {
    let cases = [
        (Kind::Pie2d, pie_start_angle()?),
        (Kind::Column3d, depth(Scale::Column)?),
        (Kind::Column2d, rounding()?),
    ];
    for format in [Format::Keynote, Format::Numbers, Format::Pages] {
        for (kind, property) in cases {
            let input = input(format, kind)?;
            let source = input.fixture.bytes.as_slice();
            let (target, chart) = (input.target, input.chart);
            std::hint::black_box(host_style_route(source, target, chart, property)?);
            std::hint::black_box(leaf_style_route(source, chart, property)?);
            std::hint::black_box(leaf_style_route(source, chart, property)?);
            std::hint::black_box(host_style_route(source, target, chart, property)?);
        }
    }
    Ok(())
}

// ---- legend frame ------------------------------------------------------

fn host_legend_frame(frame: legend::Frame) -> TestResult<litchi_iwa::charts::ChartLegendFrame> {
    use litchi_iwa::charts::{ChartLegendFrame, ChartLegendRect};
    Ok(match frame {
        legend::Frame::Automatic => ChartLegendFrame::Automatic,
        legend::Frame::Frame(rect) => {
            let [x, y, width, height] = rect.native();
            ChartLegendFrame::Frame(ChartLegendRect::from_points(x, y, width, height)?)
        },
    })
}

fn host_legend_write(
    source: &[u8],
    target: Target,
    chart: u64,
    frame: legend::Frame,
) -> TestResult<Vec<u8>> {
    let frame = host_legend_frame(frame)?;
    Ok(match target {
        Target::Keynote { slide } => {
            let mut editor = keynote::Draft::from_bytes(source)?;
            editor.set_slide_chart_legend_frame(slide, chart, frame)?;
            editor.to_bytes()?
        },
        Target::Numbers { sheet } => {
            let mut editor = numbers::Draft::from_bytes(source)?;
            editor.set_sheet_chart_legend_frame(sheet, chart, frame)?;
            editor.to_bytes()?
        },
        Target::Pages => {
            let mut editor = pages::Draft::from_bytes(source)?;
            editor.set_body_chart_legend_frame(chart, frame)?;
            editor.to_bytes()?
        },
    })
}

fn host_legend_read(bytes: &[u8], target: Target, chart: u64) -> TestResult<String> {
    use litchi_iwa::charts::ChartLegendFrame;
    let frame = match target {
        Target::Keynote { slide } => {
            keynote::Draft::from_bytes(bytes)?.slide_chart_legend_frame(slide, chart)?
        },
        Target::Numbers { sheet } => {
            numbers::Draft::from_bytes(bytes)?.sheet_chart_legend_frame(sheet, chart)?
        },
        Target::Pages => pages::Draft::from_bytes(bytes)?.body_chart_legend_frame(chart)?,
    };
    let common = match frame {
        ChartLegendFrame::Automatic => legend::Frame::Automatic,
        ChartLegendFrame::Frame(rect) => legend::Frame::Frame(legend::Rect::from_points(
            rect.x().points(),
            rect.y().points(),
            rect.width().points(),
            rect.height().points(),
        )?),
    };
    Ok(format!("{common:?}"))
}

fn engine_legend_read(bytes: &[u8], chart: u64) -> TestResult<String> {
    let components = Components::read(bytes)?;
    let location = components.locate(chart, CHART_MESSAGE_TYPE)?;
    let stored = codec::legend_frame(components.payload(chart, location)?, Limits::DEFAULT)?;
    Ok(format!("{:?}", legend::Frame::from_native(stored)?))
}

fn engine_legend_write(source: &[u8], chart: u64, frame: legend::Frame) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let location = components.locate(chart, CHART_MESSAGE_TYPE)?;
    let payload = components.payload(chart, location)?;
    if legend::Frame::from_native(codec::legend_frame(payload, Limits::DEFAULT)?)? == frame {
        return Ok(source.to_vec());
    }
    let rewritten = codec::rewrite_legend_frame(payload, frame.native(), Limits::DEFAULT)?;
    let data = rewritten.into_bytes().into_owned();
    publish(source, components, vec![(chart, CHART_MESSAGE_TYPE, data)])
}

fn legend_case(
    name: &'static str,
    format: Format,
    tier: Tier,
    evidence: Evidence,
    step: Step,
) -> TestResult {
    let mut input = input(format, Kind::Column2d)?;
    let (target, chart) = (input.target, input.chart);
    let explicit = legend::Frame::Frame(legend::Rect::from_points(24.0, 36.0, 0.0, 48.0)?);
    let frame = if step == Step::Reset {
        input.fixture.bytes = host_legend_write(&input.fixture.bytes, target, chart, explicit)?;
        legend::Frame::Automatic
    } else {
        explicit
    };
    Case::new(name, input.format, tier, evidence, &input.fixture)
        .host(|source| host_legend_write(source, target, chart, frame))
        .leaf(|source| Ok(LeafRun::output(engine_legend_write(source, chart, frame)?)))
        .read("host legend frame", move |bytes| {
            host_legend_read(bytes, target, chart)
        })
        .read("engine legend frame", move |bytes| {
            engine_legend_read(bytes, chart)
        })
        .run()?;
    Ok(())
}

#[test]
fn keynote_chart_legend_frame_parity() -> TestResult {
    legend_case(
        "keynote_chart_legend_frame_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        Step::Set,
    )
}

#[test]
fn numbers_chart_legend_frame_parity() -> TestResult {
    legend_case(
        "numbers_chart_legend_frame_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        Step::Set,
    )
}

#[test]
fn pages_chart_legend_frame_parity() -> TestResult {
    legend_case(
        "pages_chart_legend_frame_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        Step::Set,
    )
}

#[test]
fn keynote_chart_legend_frame_reset_parity() -> TestResult {
    legend_case(
        "keynote_chart_legend_frame_reset_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        Step::Reset,
    )
}

#[test]
fn numbers_chart_legend_frame_reset_parity() -> TestResult {
    legend_case(
        "numbers_chart_legend_frame_reset_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        Step::Reset,
    )
}

#[test]
fn pages_chart_legend_frame_reset_parity() -> TestResult {
    legend_case(
        "pages_chart_legend_frame_reset_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        Step::Reset,
    )
}
