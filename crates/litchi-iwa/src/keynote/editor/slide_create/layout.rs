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

pub(in crate::keynote::editor) struct LayoutCatalog {
    infos: Vec<KeynoteSlideLayoutInfo>,
    by_template_slide: HashMap<u64, usize>,
}

impl LayoutCatalog {
    pub(in crate::keynote::editor) fn read(
        graph: &ObjectGraph,
        theme: &kn::ThemeArchive,
    ) -> Result<Self> {
        let default = default_layout_node_id(theme)?;
        let mut seen_nodes = HashSet::with_capacity(theme.templates.len());
        let mut infos = Vec::with_capacity(theme.templates.len());
        let mut by_template_slide = HashMap::with_capacity(theme.templates.len());
        let mut found_default = false;
        for reference in &theme.templates {
            if !seen_nodes.insert(reference.identifier) {
                return Err(Error::InvalidFormat(format!(
                    "Keynote theme duplicates layout node {}",
                    reference.identifier
                )));
            }
            let layout = resolve_layout(graph, reference.identifier)?;
            let name = layout
                .slide
                .name
                .filter(|name| !name.is_empty())
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote layout slide {} has no name",
                        layout.slide_id
                    ))
                })?;
            let is_default = reference.identifier == default;
            found_default |= is_default;
            let index = infos.len();
            infos.push(KeynoteSlideLayoutInfo {
                id: KeynoteSlideLayoutId(reference.identifier),
                name,
                is_default,
            });
            if by_template_slide.insert(layout.slide_id, index).is_some() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote theme maps template slide {} to multiple layout nodes",
                    layout.slide_id
                )));
            }
        }
        if !found_default {
            return Err(Error::InvalidFormat(format!(
                "Keynote default layout node {default} is not in the theme layout list"
            )));
        }
        Ok(Self {
            infos,
            by_template_slide,
        })
    }

    pub(in crate::keynote::editor) fn current(
        &self,
        template_slide_id: u64,
    ) -> Result<KeynoteSlideLayoutInfo> {
        self.by_template_slide
            .get(&template_slide_id)
            .map(|index| self.infos[*index].clone())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote theme has no layout for template slide {template_slide_id}"
                ))
            })
    }

    pub(super) fn into_infos(self) -> Vec<KeynoteSlideLayoutInfo> {
        self.infos
    }
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
