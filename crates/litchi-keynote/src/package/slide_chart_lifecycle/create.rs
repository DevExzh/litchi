//! Source-free chart creation (a port of the migration host's
//! `add_slide_chart` and its `charts::source` graph builder).
//!
//! A new chart is a fixed private graph in the slide's component: the
//! chart drawable, its title and caption stand-ins, the theme style preset
//! and sixteen chart, legend, axis and series styles. The styles join the
//! document stylesheet, the preset joins the theme's chart presets, and the
//! cross-component edges are registered as external references.

use litchi_iwa_common::chart::data::ChartData;
use litchi_iwa_common::chart::gaps::Spacing;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_core::{ArchiveObject, RawMessage};

use super::super::structure::metadata::ExternalReference;
use super::super::structure::wire::Writer;
use super::{
    Error, Fields, Geometry, Package, Refusal, STYLESHEET_MESSAGE_TYPE, STYLESHEET_STYLES,
    THEME_CHART_PRESETS, THEME_MESSAGE_TYPE, THEME_SUPER, Workspace, append_references, lift,
    theme_and_stylesheet,
};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const PRESET_MESSAGE_TYPE: u32 = 5_020;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CHART_STYLE: u32 = 5_022;
const CHART_NON_STYLE: u32 = 5_023;
const LEGEND_STYLE: u32 = 5_024;
const LEGEND_NON_STYLE: u32 = 5_025;
const AXIS_STYLE: u32 = 5_026;
const AXIS_NON_STYLE: u32 = 5_027;
const SERIES_STYLE: u32 = 5_028;
const STANDARD_VERSION: &[u32] = &[1, 0, 5];
const STANDIN_VERSION: &[u32] = &[10, 1, 0];
/// `TSWP.ThemePresetsArchive` extension of `TSS.ThemeArchive`.
const THEME_TEXT_PRESETS: u32 = 110;
const THEME_PARAGRAPH_STYLE_PRESETS: u32 = 7;
const EXTENSION: u32 = 10_000;
const SUPPORTS_PRIMARY: u32 = 10_001;
const SUPPORTS_SECONDARY: u32 = 10_002;
const KEYNOTE_PARAGRAPH_STYLE_COUNT: usize = 25;
const SCATTER_FORMAT_SHARED_X: u64 = 2;
const SERIES_DIRECTION_BY_ROW: u64 = 1;
const NUMBER_FORMAT_TYPE: u64 = 256;
const AUTOMATIC_DECIMAL_PLACES: u64 = 253;
const RADAR_STROKE_WIDTH: f32 = 4.0;
const MITER_LIMIT: f32 = 4.0;
const STROKE_PATTERN_CAPACITY: usize = 6;
const RADAR_FILL_STROKE_ALPHA: f32 = 0.15;
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

/// One staged chart creation.
#[derive(Clone)]
pub(super) struct NewChart {
    pub(super) kind: Kind,
    pub(super) data: ChartData,
    pub(super) geometry: Geometry,
}

/// Identifiers of the new graph, allocated in the host's order.
struct Ids {
    drawable: u64,
    caption: u64,
    title: u64,
    preset: u64,
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
    fn allocate(workspace: &mut Workspace<'_>) -> Result<Self, Error> {
        let mut take = || workspace.allocate().map_err(lift);
        Ok(Self {
            drawable: take()?,
            caption: take()?,
            title: take()?,
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

    fn styles(&self) -> Vec<u64> {
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

    fn all(&self) -> Vec<u64> {
        let mut all = vec![self.drawable, self.caption, self.title, self.preset];
        all.extend(self.styles());
        all
    }
}

/// The number of visible series of `data` by rows (host `chart_series_count`).
pub(super) fn series_count(kind: Kind, data: &ChartData) -> Option<usize> {
    let rows = data.row_names().len();
    let count = match kind {
        Kind::Scatter2d => (rows == 2).then_some(1)?,
        Kind::Bubble2d => (rows == 3).then_some(1)?,
        Kind::MultiDataScatter2d | Kind::MultiDataBubble2d => 1,
        _ => rows,
    };
    (count > 0).then_some(count)
}

fn object(
    workspace: &Workspace<'_>,
    identifier: u64,
    message_type: u32,
    payload: Vec<u8>,
    versions: &[u32],
    references: &[u64],
) -> Result<ArchiveObject, Error> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: payload,
        }],
        workspace.archive_limits(),
    )
    .map_err(|_error| Error::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource)?;
    info.versions = versions.to_vec().into();
    info.object_references = references.to_vec();
    Ok(object)
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

/// The theme's first paragraph style preset.
fn paragraph_style(workspace: &Workspace<'_>, theme: u64) -> Result<u64, Error> {
    let wire = workspace.wire;
    let theme = Fields::parse(
        workspace.payload(theme, THEME_MESSAGE_TYPE).map_err(lift)?,
        wire,
    )
    .map_err(lift)?;
    let base = theme
        .unique(THEME_SUPER)
        .map_err(lift)?
        .ok_or(Error::InvalidSource)?;
    let base = Fields::parse(base.message().map_err(lift)?, wire).map_err(lift)?;
    let presets = base
        .unique(THEME_TEXT_PRESETS)
        .map_err(lift)?
        .ok_or(Error::InvalidSource)?;
    let presets = Fields::parse(presets.message().map_err(lift)?, wire).map_err(lift)?;
    for field in presets.all(THEME_PARAGRAPH_STYLE_PRESETS) {
        let identifier = super::super::structure::wire::reference_identifier(
            field.message().map_err(lift)?,
            wire,
        )
        .map_err(lift)?;
        if identifier != 0 {
            return Ok(identifier);
        }
    }
    Err(Error::InvalidSource)
}

fn component_of(workspace: &Workspace<'_>, object: u64) -> Result<u64, Error> {
    let member = workspace.component_of(object).ok_or(Error::InvalidSource)?;
    Ok(workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .identifier)
}

/// Build and register one new chart on `slide`; returns its drawable.
pub(super) fn add_chart(
    workspace: &mut Workspace<'_>,
    source: &Package,
    slide: u64,
    chart: &NewChart,
) -> Result<u64, Error> {
    let wire = workspace.wire;
    let series =
        series_count(chart.kind, &chart.data).ok_or(Error::Refused(Refusal::InvalidData))?;
    let (theme, _presets, stylesheet, _styles) = theme_and_stylesheet(workspace, source)?;
    let paragraph = paragraph_style(workspace, theme)?;
    let member = workspace.component_of(slide).ok_or(Error::InvalidSource)?;
    let ids = Ids::allocate(workspace)?;

    let request = litchi_iwa_protos::chart_grid_creation_codec::ChartGridCreationRequest::new(
        chart.data.row_names(),
        chart.data.column_names(),
        chart.data.values(),
        ids.drawable,
    );
    let options =
        litchi_iwa_protos::chart_grid_creation_codec::EncodeOptions::for_request(&request);
    let grid = litchi_iwa_protos::chart_grid_creation_codec::encode_chart_grid(request, options)
        .map_err(|_error| Error::Refused(Refusal::InvalidData))?;

    #[allow(
        clippy::cast_sign_loss,
        reason = "chart kinds are positive enum values"
    )]
    let mut unity = Writer::new()
        .varint(1, chart.kind.native_value() as u64)
        .varint(2, SCATTER_FORMAT_SHARED_X)
        .reference(4, ids.preset)
        .varint(5, SERIES_DIRECTION_BY_ROW)
        .bytes(7, grid.bytes())
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
    for _ in 0..KEYNOTE_PARAGRAPH_STYLE_COUNT {
        unity = unity.reference(20, paragraph);
    }
    unity = unity.varint(21, 0).boolean(22, false).boolean(24, false);
    for number in CHART_BOOL_EXTENSIONS {
        unity = unity.boolean(number, true);
    }
    unity = unity.message(CHART_CACHED_FORMATTERS, &cached_formatters(series));
    let geometry = Geometry {
        rotation: Some(chart.geometry.rotation.unwrap_or(0.0)),
        ..chart.geometry
    };
    let drawable = Writer::new()
        .bytes(
            1,
            &super::super::slide_drawable::new_drawable_payload(
                geometry,
                slide,
                ids.title,
                ids.caption,
            ),
        )
        .message(EXTENSION, &unity)
        .finish();

    let mut chart_references = vec![ids.caption, ids.title, ids.preset];
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
    for _ in 0..KEYNOTE_PARAGRAPH_STYLE_COUNT {
        preset = preset.reference(6, paragraph);
    }
    let mut preset_references = vec![ids.chart_style, ids.legend_style];
    preset_references.extend(ids.value_axis_styles);
    preset_references.push(ids.category_axis_style);
    preset_references.extend(ids.series_styles);
    preset_references.push(paragraph);

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
    let mut objects = vec![
        object(
            workspace,
            ids.drawable,
            CHART_MESSAGE_TYPE,
            drawable,
            STANDARD_VERSION,
            &chart_references,
        )?,
        object(
            workspace,
            ids.caption,
            STANDIN_MESSAGE_TYPE,
            Vec::new(),
            STANDIN_VERSION,
            &[],
        )?,
        object(
            workspace,
            ids.title,
            STANDIN_MESSAGE_TYPE,
            Vec::new(),
            STANDIN_VERSION,
            &[],
        )?,
        object(
            workspace,
            ids.preset,
            PRESET_MESSAGE_TYPE,
            preset.finish(),
            STANDARD_VERSION,
            &preset_references,
        )?,
    ];
    let style = |identifier: u64, message_type: u32, extension: &Writer, features: &[u32]| {
        object(
            workspace,
            identifier,
            message_type,
            style_payload(stylesheet, extension, features),
            STANDARD_VERSION,
            &[],
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
    for object in objects {
        workspace.insert(&member, object).map_err(lift)?;
    }

    // Styles join the document stylesheet; the preset joins the theme.
    let styles = ids.styles();
    append_references(
        workspace,
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        &[],
        STYLESHEET_STYLES,
        &styles,
    )?;
    super::super::slide_drawable::attach_drawable(workspace, slide, ids.drawable).map_err(lift)?;
    let has_presets = {
        let theme_fields = Fields::parse(
            workspace.payload(theme, THEME_MESSAGE_TYPE).map_err(lift)?,
            wire,
        )
        .map_err(lift)?;
        let base = theme_fields
            .unique(THEME_SUPER)
            .map_err(lift)?
            .ok_or(Error::InvalidSource)?;
        Fields::parse(base.message().map_err(lift)?, wire)
            .map_err(lift)?
            .unique(THEME_CHART_PRESETS)
            .map_err(lift)?
            .is_some()
    };
    if !has_presets {
        let payload = workspace
            .payload(theme, THEME_MESSAGE_TYPE)
            .map_err(lift)?
            .to_vec();
        let data = super::edit_nested(&payload, &[THEME_SUPER], wire, &mut |fields| {
            fields.set_message(THEME_CHART_PRESETS, &[]);
            Ok(())
        })
        .map_err(lift)?;
        super::super::slide_drawable::replace_message_payload(
            workspace,
            theme,
            THEME_MESSAGE_TYPE,
            data,
        )
        .map_err(lift)?;
    }
    append_references(
        workspace,
        theme,
        THEME_MESSAGE_TYPE,
        &[THEME_SUPER, THEME_CHART_PRESETS],
        1,
        &[ids.preset],
    )?;

    // Registry: cross-component edges in the host's order, then UUIDs.
    let owner = component_of(workspace, slide)?;
    let stylesheet_component = component_of(workspace, stylesheet)?;
    let theme_component = component_of(workspace, theme)?;
    let paragraph_component = component_of(workspace, paragraph)?;
    let mut edges = Vec::new();
    if stylesheet_component != owner {
        edges.push((owner, stylesheet_component, stylesheet));
        edges.extend(
            styles
                .iter()
                .map(|style| (stylesheet_component, owner, *style)),
        );
    }
    if theme_component != owner {
        edges.push((theme_component, owner, ids.preset));
    }
    if paragraph_component != owner {
        edges.push((owner, paragraph_component, paragraph));
    }
    for (from, to, object) in edges {
        workspace
            .registry
            .component_mut(from)
            .ok_or(Error::InvalidSource)?
            .add_external_reference(ExternalReference {
                component: to,
                object: Some(object),
                is_weak: None,
            });
    }
    let mut uuids = Vec::new();
    for identifier in ids.all() {
        uuids.push((
            identifier,
            workspace.clone_uuid(None, identifier).map_err(lift)?,
        ));
    }
    let entry = workspace
        .registry
        .component_mut(owner)
        .ok_or(Error::InvalidSource)?;
    for (identifier, uuid) in uuids {
        entry.add_object_uuid(identifier, uuid);
    }
    Ok(ids.drawable)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn series_follow_the_host_topology() {
        let data = ChartData::new(
            vec!["a".to_owned(), "b".to_owned()],
            vec!["x".to_owned()],
            vec![vec![Some(1.0)], vec![Some(2.0)]],
        )
        .expect("data");
        assert_eq!(series_count(Kind::Bar2d, &data), Some(2));
        assert_eq!(series_count(Kind::Scatter2d, &data), Some(1));
        assert_eq!(series_count(Kind::Bubble2d, &data), None);
    }
}
