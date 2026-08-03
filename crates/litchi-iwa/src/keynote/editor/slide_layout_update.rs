//! Transactional reassignment of a live slide to a theme layout.

use super::*;
use slide_create::layout::{read_layout_graph, resolve_layout};

mod wire;

use wire::*;

const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
const SLIDE_STYLE_FIELD: u32 = 1;
const SLIDE_TITLE_PLACEHOLDER_FIELD: u32 = 5;
const SLIDE_BODY_PLACEHOLDER_FIELD: u32 = 6;
const SLIDE_TEMPLATE_FIELD: u32 = 17;
const SLIDE_NODE_TEMPLATE_UUID_FIELD: u32 = 29;
const PLACEHOLDER_GEOMETRY_PATH: &[u32] = &[1, 1, 1, 1];
const PLACEHOLDER_STYLE_PATH: &[u32] = &[1, 1, 2];
const PLACEHOLDER_PATH_SOURCE_PATH: &[u32] = &[1, 1, 3];

impl KeynoteEditor {
    /// Reassign an existing slide to a theme layout without replacing user content.
    ///
    /// Keynote retains the slide's text, notes, builds, transition, and ordinary
    /// drawables during this operation. The selected layout supplies the slide
    /// style, title/body placeholder presentation and visibility, and cloned
    /// layout-owned images or live-video objects. Replaced layout media graphs
    /// are removed safely.
    pub fn set_slide_layout(
        &mut self,
        slide_index: usize,
        layout: KeynoteSlideLayoutId,
    ) -> Result<()> {
        let slides = self.slides()?;
        let before = slides.get(slide_index).cloned().ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        if before.layout.as_ref().map(|current| current.id) == Some(layout) {
            return Ok(());
        }

        let graph = ObjectGraph::read(self.package())?;
        let layout_graph = read_layout_graph(&graph)?;
        let layout_id = layout.as_u64();
        if !layout_graph
            .theme
            .templates
            .iter()
            .any(|reference| reference.identifier == layout_id)
        {
            return Err(Error::ParseError(format!(
                "Keynote theme has no slide layout {}",
                layout_id
            )));
        }
        let target = resolve_layout(&graph, layout_id)?;
        let target_node: kn::SlideNodeArchive = graph.decode_type(
            target.node_id,
            SLIDE_NODE_MESSAGE_TYPE,
            "KN.SlideNodeArchive",
        )?;
        let current_slide: kn::SlideArchive =
            graph.decode_type(before.slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
        let current_node: kn::SlideNodeArchive = graph.decode_type(
            before.node_id,
            SLIDE_NODE_MESSAGE_TYPE,
            "KN.SlideNodeArchive",
        )?;

        let title = placeholder_plan(
            slide_index,
            &current_slide,
            &target.slide,
            KeynoteSlideTextPlaceholder::Title,
        )?;
        let body = placeholder_plan(
            slide_index,
            &current_slide,
            &target.slide,
            KeynoteSlideTextPlaceholder::Body,
        )?;
        let preserved_text_boxes = self
            .slide_text_storages(slide_index)?
            .into_iter()
            .filter(|storage| storage.role == KeynoteSlideTextRole::TextBox)
            .collect::<Vec<_>>();

        let mut staged = self.package().clone();
        for plan in [&title, &body].into_iter().flatten() {
            if let Some(target_id) = plan.target_id {
                patch_placeholder_presentation(
                    &mut staged,
                    &graph,
                    plan.current_id,
                    target_id,
                    plan.label,
                )?;
            }
        }
        patch_slide_relationship(
            &mut staged,
            &graph,
            before.slide_id,
            &current_slide,
            &target,
        )?;
        for plan in [&title, &body].into_iter().flatten() {
            if plan.current_visible != plan.target_visible {
                placeholder_ownership::patch(
                    &mut staged,
                    graph.archive_name(before.slide_id)?,
                    before.slide_id,
                    plan.reference_field,
                    plan.current_id,
                    plan.target_visible,
                    plan.label,
                )?;
            }
        }
        slide_layout_media::materialize(
            &mut staged,
            &graph,
            before.slide_id,
            &current_slide,
            &target,
        )?;
        patch_node_template_uuid(
            &mut staged,
            &graph,
            before.node_id,
            &current_node,
            &target_node,
        )?;
        slide_preview::invalidate(
            &mut staged,
            graph.archive_name(before.node_id)?,
            before.node_id,
        )?;

        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        let after = verified
            .slides()?
            .get(slide_index)
            .cloned()
            .ok_or_else(|| {
                Error::InvalidFormat("Keynote layout update lost its slide".to_owned())
            })?;
        verify_slide_semantics(&before, &after, layout, &title, &body)?;
        let after_text_boxes = verified
            .slide_text_storages(slide_index)?
            .into_iter()
            .filter(|storage| storage.role == KeynoteSlideTextRole::TextBox)
            .collect::<Vec<_>>();
        if after_text_boxes != preserved_text_boxes {
            return Err(Error::InvalidFormat(
                "Keynote layout update changed ordinary slide text".to_owned(),
            ));
        }

        *self = verified;
        Ok(())
    }
}

struct PlaceholderPlan {
    current_id: u64,
    target_id: Option<u64>,
    current_visible: bool,
    target_visible: bool,
    reference_field: u32,
    label: &'static str,
}

fn placeholder_plan(
    slide_index: usize,
    current: &kn::SlideArchive,
    target: &kn::SlideArchive,
    placeholder: KeynoteSlideTextPlaceholder,
) -> Result<Option<PlaceholderPlan>> {
    let (current_reference, target_reference, reference_field, label) = match placeholder {
        KeynoteSlideTextPlaceholder::Title => (
            current.title_placeholder.as_ref(),
            target.title_placeholder.as_ref(),
            SLIDE_TITLE_PLACEHOLDER_FIELD,
            "title",
        ),
        KeynoteSlideTextPlaceholder::Body => (
            current.body_placeholder.as_ref(),
            target.body_placeholder.as_ref(),
            SLIDE_BODY_PLACEHOLDER_FIELD,
            "body",
        ),
    };
    let Some(current_reference) = current_reference else {
        if target_reference.is_none() {
            return Ok(None);
        }
        return Err(Error::InvalidFormat(format!(
            "Keynote slide {slide_index} cannot adopt a layout with a {label} placeholder because its retained placeholder is missing"
        )));
    };
    let current_visible =
        placeholder_ownership::validate(slide_index, current, current_reference.identifier, label)?;
    let target_visible = target_reference
        .map(|target_reference| {
            placeholder_ownership::validate(
                slide_index,
                target,
                target_reference.identifier,
                &format!("layout {label}"),
            )
        })
        .transpose()?
        .unwrap_or(false);
    Ok(Some(PlaceholderPlan {
        current_id: current_reference.identifier,
        target_id: target_reference.map(|reference| reference.identifier),
        current_visible,
        target_visible,
        reference_field,
        label,
    }))
}

fn verify_slide_semantics(
    before: &KeynoteSlideInfo,
    after: &KeynoteSlideInfo,
    layout: KeynoteSlideLayoutId,
    title: &Option<PlaceholderPlan>,
    body: &Option<PlaceholderPlan>,
) -> Result<()> {
    if after.layout.as_ref().map(|current| current.id) != Some(layout)
        || after.node_id != before.node_id
        || after.slide_id != before.slide_id
        || after.name != before.name
        || after.is_skipped != before.is_skipped
        || after.is_slide_number_visible != before.is_slide_number_visible
        || after.transition != before.transition
        || after.title_storage_id != before.title_storage_id
        || after.title != before.title
        || after.body_storage_id != before.body_storage_id
        || after.body != before.body
        || after.notes_storage_id != before.notes_storage_id
        || after.notes != before.notes
        || title
            .as_ref()
            .is_some_and(|plan| after.is_title_visible != Some(plan.target_visible))
        || body
            .as_ref()
            .is_some_and(|plan| after.is_body_visible != Some(plan.target_visible))
    {
        return Err(Error::InvalidFormat(
            "Keynote slide layout update failed semantic validation".to_owned(),
        ));
    }
    Ok(())
}
