//! Source-free chart creation (a port of the migration host's
//! `add_slide_chart` and its `charts::source` graph builder).
//!
//! A new chart is a fixed private graph in the slide's component: the
//! chart drawable, its title and caption stand-ins, the theme style preset
//! and sixteen chart, legend, axis and series styles. The styles join the
//! document stylesheet, the preset joins the theme's chart presets, and the
//! cross-component edges are registered as external references.

use litchi_iwa_common::chart::data::ChartData;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_style::chart_lifecycle::author::{self, Ids, Profile, Request};

use super::super::structure::metadata::ExternalReference;
use super::{
    Error, Fields, Geometry, Package, Refusal, STYLESHEET_MESSAGE_TYPE, STYLESHEET_STYLES,
    THEME_CHART_PRESETS, THEME_MESSAGE_TYPE, THEME_SUPER, Workspace, append_references, lift,
    theme_and_stylesheet,
};

/// `TSWP.ThemePresetsArchive` extension of `TSS.ThemeArchive`.
const THEME_TEXT_PRESETS: u32 = 110;
const THEME_PARAGRAPH_STYLE_PRESETS: u32 = 7;

/// One staged chart creation.
#[derive(Clone)]
pub(super) struct NewChart {
    pub(super) kind: Kind,
    pub(super) data: ChartData,
    pub(super) geometry: Geometry,
}

/// The number of visible series of `data` by rows (host `chart_series_count`).
pub(super) fn series_count(kind: Kind, data: &ChartData) -> Option<usize> {
    author::series_count(kind, data)
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
    if series_count(chart.kind, &chart.data).is_none() {
        return Err(Error::Refused(Refusal::InvalidData));
    }
    let (theme, _presets, stylesheet, _styles) = theme_and_stylesheet(workspace, source)?;
    let paragraph = paragraph_style(workspace, theme)?;
    let member = workspace.component_of(slide).ok_or(Error::InvalidSource)?;
    let ids = Ids::allocate(Profile::Keynote, || workspace.allocate().map_err(lift))?;
    let geometry = Geometry {
        rotation: Some(chart.geometry.rotation.unwrap_or(0.0)),
        ..chart.geometry
    };
    let drawable =
        super::super::slide_drawable::new_drawable_payload(geometry, slide, ids.title, ids.caption);
    let request = Request {
        profile: Profile::Keynote,
        kind: chart.kind,
        data: &chart.data,
        drawable: &drawable,
        stylesheet,
        paragraph_style: paragraph,
    };
    let objects =
        author::objects(&ids, &request, workspace.archive_limits()).map_err(
            |error| match error {
                litchi_iwa_style::chart_lifecycle::Error::InvalidData => {
                    Error::Refused(Refusal::InvalidData)
                },
                _ => Error::InvalidSource,
            },
        )?;
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
