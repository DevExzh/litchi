//! Strict discovery and validation of Keynote theme layouts.

use super::*;

const DOCUMENT_OBJECT_ID: u64 = 1;
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHOW_MESSAGE_TYPE: u32 = 2;
const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const THEME_MESSAGE_TYPE: u32 = 10;

pub(super) struct LayoutGraph {
    pub(super) show_id: u64,
    pub(super) show_archive_name: String,
    pub(super) theme: kn::ThemeArchive,
}

pub(super) struct ResolvedLayout {
    pub(super) node_id: u64,
    pub(super) slide_id: u64,
    pub(super) archive_name: String,
    pub(super) slide: kn::SlideArchive,
}

pub(super) fn read_layout_graph(graph: &ObjectGraph) -> Result<LayoutGraph> {
    let document: kn::DocumentArchive = graph.decode_type(
        DOCUMENT_OBJECT_ID,
        DOCUMENT_MESSAGE_TYPE,
        "KN.DocumentArchive",
    )?;
    let show_id = document.show.identifier;
    let show: kn::ShowArchive = graph.decode_type(show_id, SHOW_MESSAGE_TYPE, "KN.ShowArchive")?;
    let theme = graph.decode_type(show.theme.identifier, THEME_MESSAGE_TYPE, "KN.ThemeArchive")?;
    Ok(LayoutGraph {
        show_id,
        show_archive_name: graph.archive_name(show_id)?.to_owned(),
        theme,
    })
}

pub(super) fn default_layout_node_id(theme: &kn::ThemeArchive) -> Result<u64> {
    theme
        .default_template_slide_node_reference
        .as_ref()
        .or(theme.default_template_slide_node.as_ref())
        .map(|reference| reference.identifier)
        .ok_or_else(|| Error::InvalidFormat("Keynote theme has no default slide layout".to_owned()))
}

pub(super) fn resolve_layout(graph: &ObjectGraph, node_id: u64) -> Result<ResolvedLayout> {
    let node: kn::SlideNodeArchive =
        graph.decode_type(node_id, SLIDE_NODE_MESSAGE_TYPE, "KN.SlideNodeArchive")?;
    let slide_id = node
        .slide
        .ok_or_else(|| Error::InvalidFormat(format!("Keynote layout node {node_id} has no slide")))?
        .identifier;
    let archive_name = graph.archive_name(slide_id)?.to_owned();
    let expected = format!("Index/TemplateSlide-{slide_id}.iwa");
    if archive_name != expected {
        return Err(Error::InvalidFormat(format!(
            "Keynote layout slide {slide_id} is not stored in {expected}"
        )));
    }
    Ok(ResolvedLayout {
        node_id,
        slide_id,
        archive_name,
        slide: graph.decode_type(slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?,
    })
}
