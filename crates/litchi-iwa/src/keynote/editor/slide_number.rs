//! Per-slide number visibility and its drawable-ownership invariant.

use super::*;

const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
const NODE_SLIDE_NUMBER_VISIBLE_FIELD: u32 = 18;
const SLIDE_NUMBER_PLACEHOLDER_FIELD: u32 = 20;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_DRAWABLES_Z_ORDER_FIELD: u32 = 42;

impl KeynoteEditor {
    /// Show or hide the layout-provided slide number on one slide.
    ///
    /// Keynote keeps the slide-number placeholder alive while hidden and
    /// changes only the node visibility bit plus drawable ownership/z-order.
    /// Slides whose layout has no slide-number placeholder are rejected.
    pub fn set_slide_number_visible(&mut self, slide_index: usize, visible: bool) -> Result<()> {
        let slides = self.slides()?;
        let info = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        let graph = ObjectGraph::read(self.package())?;
        let node: kn::SlideNodeArchive =
            graph.decode_type(info.node_id, SLIDE_NODE_MESSAGE_TYPE, "KN.SlideNodeArchive")?;
        let slide: kn::SlideArchive =
            graph.decode_type(info.slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
        let placeholder_id = slide
            .slide_number_placeholder
            .as_ref()
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Keynote slide {slide_index} has no layout-provided slide-number placeholder"
                ))
            })?
            .identifier;
        graph.decode_type::<kn::PlaceholderArchive>(
            placeholder_id,
            PLACEHOLDER_MESSAGE_TYPE,
            "KN.PlaceholderArchive",
        )?;
        let slide_archive_name = graph.archive_name(info.slide_id)?.to_owned();
        if graph.archive_name(placeholder_id)? != slide_archive_name {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide-number placeholder {placeholder_id} is outside slide component {}",
                info.slide_id
            )));
        }
        validate_visibility_invariant(slide_index, &node, &slide, placeholder_id)?;
        if node.is_slide_number_visible.unwrap_or(false) == visible {
            return Ok(());
        }

        let node_archive_name = graph.archive_name(info.node_id)?.to_owned();
        let mut staged = self.package().clone();
        patch_node_visibility(&mut staged, &node_archive_name, info.node_id, visible)?;
        patch_slide_number_ownership(
            &mut staged,
            &slide_archive_name,
            info.slide_id,
            placeholder_id,
            visible,
        )?;

        let verified = Self::from_package(staged)?;
        let verified_slides = verified.slides()?;
        let verified_info = verified_slides.get(slide_index).ok_or_else(|| {
            Error::InvalidFormat("Keynote slide disappeared during number update".to_owned())
        })?;
        let verified_graph = ObjectGraph::read(verified.package())?;
        let verified_node: kn::SlideNodeArchive = verified_graph.decode_type(
            verified_info.node_id,
            SLIDE_NODE_MESSAGE_TYPE,
            "KN.SlideNodeArchive",
        )?;
        let verified_slide: kn::SlideArchive = verified_graph.decode_type(
            verified_info.slide_id,
            SLIDE_MESSAGE_TYPE,
            "KN.SlideArchive",
        )?;
        validate_visibility_invariant(
            slide_index,
            &verified_node,
            &verified_slide,
            placeholder_id,
        )?;
        if verified_info.is_slide_number_visible != Some(visible) {
            return Err(Error::InvalidFormat(
                "Keynote slide-number visibility failed round-trip validation".to_owned(),
            ));
        }
        *self = verified;
        Ok(())
    }
}

fn validate_visibility_invariant(
    slide_index: usize,
    node: &kn::SlideNodeArchive,
    slide: &kn::SlideArchive,
    placeholder_id: u64,
) -> Result<()> {
    let owned_count = slide
        .owned_drawables
        .iter()
        .filter(|reference| reference.identifier == placeholder_id)
        .count();
    let z_order_count = slide
        .drawables_z_order
        .iter()
        .filter(|reference| reference.identifier == placeholder_id)
        .count();
    if owned_count > 1 || z_order_count > 1 || owned_count != z_order_count {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide {slide_index} has inconsistent slide-number drawable ownership"
        )));
    }
    let visible = node.is_slide_number_visible.unwrap_or(false);
    if visible != (owned_count == 1) {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide {slide_index} number visibility disagrees with drawable ownership"
        )));
    }
    Ok(())
}

fn patch_node_visibility(
    package: &mut IWorkPackage,
    archive_name: &str,
    node_id: u64,
    visible: bool,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(node_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide node {node_id} is missing"))
        })?;
        let message_index =
            unique_message_index(object, SLIDE_NODE_MESSAGE_TYPE, "Keynote slide node")?;
        let original = object.messages[message_index].data.as_slice();
        let decoded = kn::SlideNodeArchive::decode(original)?;
        let data = patch_varint_field(
            original,
            NODE_SLIDE_NUMBER_VISIBLE_FIELD,
            decoded.is_slide_number_visible.is_some(),
            Some(u64::from(visible)),
        )?;
        if kn::SlideNodeArchive::decode(data.as_slice())?.is_slide_number_visible != Some(visible) {
            return Err(Error::InvalidFormat(
                "Keynote slide-number node patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: SLIDE_NODE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn patch_slide_number_ownership(
    package: &mut IWorkPackage,
    archive_name: &str,
    slide_id: u64,
    placeholder_id: u64,
    visible: bool,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(slide_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote slide {slide_id} is missing"))
        })?;
        let message_index = unique_message_index(
            object,
            SLIDE_MESSAGE_TYPE,
            "Keynote slide",
        )?;
        let original = object.messages[message_index].data.as_slice();
        let raw_placeholder = repeated_length_delimited_payloads(
            original,
            SLIDE_NUMBER_PLACEHOLDER_FIELD,
        )?;
        let [raw_placeholder] = raw_placeholder.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_id} must contain exactly one raw slide-number reference"
            )));
        };
        if tsp::Reference::decode(*raw_placeholder)?.identifier != placeholder_id {
            return Err(Error::InvalidFormat(
                "Keynote slide-number reference changed during update".to_owned(),
            ));
        }
        let mut data = original.to_vec();
        for field in [SLIDE_OWNED_DRAWABLES_FIELD, SLIDE_DRAWABLES_Z_ORDER_FIELD] {
            let raw = repeated_length_delimited_payloads(&data, field)?;
            let matches = raw.iter().try_fold(0usize, |count, payload| {
                let identifier = tsp::Reference::decode(*payload)?.identifier;
                Ok::<_, Error>(count + usize::from(identifier == placeholder_id))
            })?;
            if visible {
                if matches != 0 {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {slide_id} already owns its hidden slide-number placeholder"
                    )));
                }
                data = append_repeated_length_delimited_field(&data, field, raw_placeholder)?;
            } else {
                if matches != 1 {
                    return Err(Error::InvalidFormat(format!(
                        "Keynote slide {slide_id} must own its visible slide-number placeholder once"
                    )));
                }
                data = remove_repeated_length_delimited_field_where(
                    &data,
                    field,
                    |payload| Ok(tsp::Reference::decode(payload)?.identifier == placeholder_id),
                )?;
            }
        }
        let verified = kn::SlideArchive::decode(data.as_slice())?;
        let expected_count = usize::from(visible);
        for references in [&verified.owned_drawables, &verified.drawables_z_order] {
            if references
                .iter()
                .filter(|reference| reference.identifier == placeholder_id)
                .count()
                != expected_count
            {
                return Err(Error::InvalidFormat(
                    "Keynote slide-number ownership patch failed validation".to_owned(),
                ));
            }
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn unique_message_index(object: &ArchiveObject, message_type: u32, context: &str) -> Result<usize> {
    let mut indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == message_type)
        .map(|(index, _)| index);
    let Some(index) = indexes.next() else {
        return Err(Error::InvalidFormat(format!(
            "{context} has no message type {message_type} payload"
        )));
    };
    if indexes.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{context} repeats message type {message_type} payload"
        )));
    }
    Ok(index)
}
