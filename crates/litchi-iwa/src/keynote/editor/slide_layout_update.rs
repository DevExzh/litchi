//! Transactional reassignment of a live slide to a theme layout.

use super::*;
use crate::archive::ArchiveLimits;
use litchi_keynote::slide::placeholder::Kind as PlaceholderKind;
use slide_create::layout::{read_layout_graph, resolve_layout};

mod dependencies;
mod wire;

use dependencies::{capture_slide_layout_dependencies, reconcile_slide_layout_dependencies};
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

fn slide_preview_wire_limits(package: &IWorkPackage) -> Result<WireLimits> {
    slide_preview_wire_limits_for(
        package.limits().effective_archive_limits()?,
        package.limits().max_iwa_stream_bytes(),
    )
}

fn slide_preview_wire_limits_for(
    archive_limits: ArchiveLimits,
    max_iwa_stream_bytes: usize,
) -> Result<WireLimits> {
    let source_bytes = archive_limits
        .max_message_bytes()
        .min(archive_limits.max_archive_bytes())
        .min(max_iwa_stream_bytes)
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    // Archive header budgets do not describe the semantic protobuf payload.
    // Keep the payload profile source-sized and give its traversal an explicit
    // positive depth budget independent of header-only ceilings.
    let fields = source_bytes
        .saturating_mul(4)
        .clamp(1, WireLimits::MAX_FIELDS);
    let work = source_bytes
        .saturating_mul(8)
        .clamp(1, WireLimits::MAX_REWRITE_WORK);
    WireLimits::default()
        .with_input_bytes(source_bytes)
        .and_then(|limits| limits.with_fields(fields))
        .and_then(|limits| limits.with_output_bytes(source_bytes))
        .and_then(|limits| limits.with_nesting(WireLimits::MAX_NESTING))
        .and_then(|limits| limits.with_rewrite_work(work))
        .map_err(|error| {
            Error::InvalidFormat(format!("invalid Keynote slide preview limits: {error}"))
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn slide_preview_payload_limits_do_not_use_header_only_budgets() {
        let archive_limits = ArchiveLimits::default()
            .with_header_fields(1)
            .unwrap()
            .with_header_nesting(1)
            .unwrap();
        let wire_limits =
            slide_preview_wire_limits_for(archive_limits, archive_limits.max_archive_bytes())
                .unwrap();

        assert!(wire_limits.max_fields() > archive_limits.max_header_fields());
        assert_eq!(wire_limits.max_nesting(), WireLimits::MAX_NESTING);
        assert!(wire_limits.max_rewrite_work() > 1);
    }
}

fn invalidate_slide_preview(
    package: &mut IWorkPackage,
    archive_name: &str,
    node_id: u64,
) -> Result<()> {
    let archive_limits = package.limits().effective_archive_limits()?;
    let wire_limits = slide_preview_wire_limits(package)?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(node_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide node {node_id} is missing"))
        })?;
        litchi_keynote::__invalidate_slide_preview(object, archive_limits, wire_limits).map_err(
            |error| {
                Error::InvalidFormat(format!(
                    "Keynote slide preview invalidation failed: {error}"
                ))
            },
        )
    })
}

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
        let before_ids = before.native_ids()?;
        let before_slide_id = before_ids.slide.get();
        let before_node_id = before_ids.node.get();
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
            graph.decode_type(before_slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
        let current_node: kn::SlideNodeArchive = graph.decode_type(
            before_node_id,
            SLIDE_NODE_MESSAGE_TYPE,
            "KN.SlideNodeArchive",
        )?;
        let dependencies = capture_slide_layout_dependencies(
            self.package(),
            &graph,
            before_slide_id,
            &current_slide,
        )?;

        let title = placeholder_plan(
            slide_index,
            &current_slide,
            &target.slide,
            PlaceholderKind::Title,
        )?;
        let body = placeholder_plan(
            slide_index,
            &current_slide,
            &target.slide,
            PlaceholderKind::Body,
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
            before_slide_id,
            &current_slide,
            &target,
        )?;
        for plan in [&title, &body].into_iter().flatten() {
            if plan.current_visible != plan.target_visible {
                placeholder_ownership::patch(
                    &mut staged,
                    graph.archive_name(before_slide_id)?,
                    before_slide_id,
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
            before_slide_id,
            &current_slide,
            &target,
        )?;
        patch_node_template_uuid(
            &mut staged,
            &graph,
            before_node_id,
            &current_node,
            &target_node,
        )?;
        invalidate_slide_preview(
            &mut staged,
            graph.archive_name(before_node_id)?,
            before_node_id,
        )?;
        reconcile_slide_layout_dependencies(
            &mut staged,
            graph.archive_name(before_slide_id)?,
            target.archive_name.as_str(),
            target.slide_id,
            &dependencies,
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
    placeholder: PlaceholderKind,
) -> Result<Option<PlaceholderPlan>> {
    let (current_reference, target_reference, reference_field, label) = match placeholder {
        PlaceholderKind::Title => (
            current.title_placeholder.as_ref(),
            target.title_placeholder.as_ref(),
            SLIDE_TITLE_PLACEHOLDER_FIELD,
            "title",
        ),
        PlaceholderKind::Body => (
            current.body_placeholder.as_ref(),
            target.body_placeholder.as_ref(),
            SLIDE_BODY_PLACEHOLDER_FIELD,
            "body",
        ),
        _ => {
            return Err(Error::InvalidFormat(
                "unsupported Keynote slide placeholder kind".to_owned(),
            ));
        },
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
    let before_ids = before.native_ids()?;
    let after_ids = after.native_ids()?;
    if after.layout.as_ref().map(|current| current.id) != Some(layout)
        || after_ids.node != before_ids.node
        || after_ids.slide != before_ids.slide
        || after.name != before.name
        || after.is_skipped != before.is_skipped
        || after.is_slide_number_visible != before.is_slide_number_visible
        || after.transition != before.transition
        || after_ids.title_storage != before_ids.title_storage
        || after.title != before.title
        || after_ids.body_storage != before_ids.body_storage
        || after.body != before.body
        || after_ids.notes_storage != before_ids.notes_storage
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
