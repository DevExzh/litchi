//! Lossless direct-slide transition splicing for deferred presentations.

use std::ops::Range;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::transition::{Kind, Transition};
use crate::{Error, Result};

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 1_000_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum OwnerKind {
    Direct,
    AlternateContent,
}

#[derive(Debug)]
struct Layout {
    transition: Option<Range<usize>>,
    owner: Option<OwnerKind>,
    insertion: usize,
    prefix: Box<str>,
    namespace: Box<str>,
}

/// Validate a source-backed slide while allowing its direct transition owner
/// to use markup-compatibility. Shape edits still require every other owner to
/// remain byte-stable after MCE preprocessing.
pub(super) fn validate_source(xml: &[u8], operation: &'static str) -> Result<()> {
    let layout = locate(xml, operation)?;
    let _ = validate_owner_for_read(xml, &layout)?;
    let scene_is_rewritten = if let Some(range) = layout.transition.as_ref() {
        let scene_xml = remove_owner(xml, range)?;
        crate::shape::Scene::read(&scene_xml)?.is_rewritten() && has_mce_markup(&scene_xml)?
    } else {
        crate::shape::Scene::read(xml)?.is_rewritten() && has_mce_markup(xml)?
    };
    if scene_is_rewritten {
        return Err(Error::UnsafeEdit {
            operation,
            reason: "source-backed slide edits do not support markup-compatibility branch selection outside the direct transition owner",
        });
    }
    Ok(())
}

// A namespace declaration alone can cause the MCE processor to allocate and
// normalize XML. It cannot select a branch or alter scene semantics. Keep the
// source scene byte-exact while still refusing actual MCE owners/directives.
fn has_mce_markup(xml: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                if is_markup_element(&namespace) {
                    return Ok(true);
                }
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
                    if is_markup_element(&namespace) {
                        return Ok(true);
                    }
                }
            },
            Event::Eof => return Ok(false),
            _ => {},
        }
    }
}

pub(super) fn read_direct(xml: &[u8]) -> Result<Option<Transition>> {
    let layout = locate(xml, "read_transition")?;
    if layout.transition.is_none() {
        return Ok(None);
    }
    validate_owner_for_read(xml, &layout)
}

pub(super) fn stage(
    xml: &[u8],
    presentation_xml: &[u8],
    target: Option<&Transition>,
    max_output_bytes: usize,
    operation: &'static str,
) -> Result<(Option<Vec<u8>>, bool)> {
    let layout = locate(xml, operation)?;
    let current = if layout.transition.is_some() {
        validate_owner_for_read(xml, &layout)?
    } else {
        None
    };

    // A value read from this exact source is an authoring no-op, including
    // retained opaque children or an unknown effect.  Equality includes the
    // retained source bytes, so this path cannot silently discard data.  Do
    // this before edit-admissibility checks: those checks are for destructive
    // rewrites and must not reject an exact source-sharing operation.
    if current.as_ref() == target {
        return Ok((None, false));
    }
    if let Some(current) = current.as_ref() {
        // Source owners are edited only when their active subtree is known.
        // Semantic equality below remains deliberately conservative: an
        // unknown retained child must not make a target that drops it look
        // like a no-op.
        validate_owner_for_edit(xml, &layout, current, operation)?;
    }
    if let Some(target) = target {
        validate_target(target, operation)?;
    }
    if same_semantics(current.as_ref(), target) {
        return Ok((None, false));
    }
    if presentation_is_protected(presentation_xml)? {
        return Err(Error::UnsafeEdit {
            operation,
            reason: "source-backed transition edits refuse modification-protected presentations",
        });
    }

    let replaced_len = layout.transition.as_ref().map_or(0, Range::len);
    let replacement_len = preflight_replacement_len(xml, &layout, target, operation)?;
    let output_len = xml
        .len()
        .checked_sub(replaced_len)
        .and_then(|len| len.checked_add(replacement_len))
        .ok_or_else(|| invalid("direct transition output size overflow"))?;
    if output_len > max_output_bytes {
        return Err(Error::Limit {
            resource: "source-backed slide transition output bytes",
            limit: max_output_bytes,
        });
    }

    let replacement = match (layout.owner, target) {
        (Some(OwnerKind::AlternateContent), Some(value)) => {
            Some(alternate_fragment(xml, &layout, value, operation)?)
        },
        (Some(OwnerKind::AlternateContent), None) => {
            Some(clear_alternate_fragment(xml, &layout, operation)?)
        },
        (_, Some(value)) => Some(direct_fragment(value, &layout.prefix, operation)?),
        (_, None) => None,
    };

    let at = layout
        .transition
        .as_ref()
        .map_or(layout.insertion, |range| range.start);
    let end = layout.transition.as_ref().map_or(at, |range| range.end);
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "source-backed slide transition XML",
            source,
        })?;
    output.extend_from_slice(&xml[..at]);
    if let Some(replacement) = replacement {
        output.extend_from_slice(&replacement);
    }
    output.extend_from_slice(&xml[end..]);
    if output.len() != output_len {
        return Err(invalid(
            "direct transition output length changed during emission",
        ));
    }

    let published_layout = locate(&output, operation)?;
    let published = if published_layout.transition.is_some() {
        validate_owner_for_read(&output, &published_layout)?
    } else {
        None
    };
    if !same_semantics(published.as_ref(), target) {
        return Err(invalid(
            "staged direct transition did not round-trip semantically",
        ));
    }
    // The source slide was validated before this operation.  Emission only
    // replaces the located transition owner, while `locate` and the semantic
    // readback above validate the complete new owner and its direct-child
    // placement; re-running shape MCE preprocessing would copy the whole
    // slide without adding a new safety guarantee.
    Ok((Some(output), true))
}

fn preflight_replacement_len(
    xml: &[u8],
    layout: &Layout,
    target: Option<&Transition>,
    operation: &'static str,
) -> Result<usize> {
    match (layout.owner, target) {
        (Some(OwnerKind::AlternateContent), Some(value)) => {
            alternate_fragment_len(xml, layout, value, operation)
        },
        (Some(OwnerKind::AlternateContent), None) => clear_alternate_fragment_len(xml, layout),
        (_, Some(value)) => {
            let canonical = crate::transition::write(value)?;
            rewrite_pml_qnames_len(canonical.as_bytes(), &layout.prefix)
        },
        (_, None) => Ok(0),
    }
}

fn same_semantics(current: Option<&Transition>, target: Option<&Transition>) -> bool {
    match (current, target) {
        (None, None) => true,
        (Some(current), Some(target)) => current.same_semantics(target),
        _ => false,
    }
}

fn validate_owner_for_read(xml: &[u8], layout: &Layout) -> Result<Option<Transition>> {
    let Some(range) = layout.transition.as_ref() else {
        return Ok(None);
    };
    if xml.get(range.clone()).is_none() {
        return Err(invalid("direct transition range is outside its slide"));
    }
    match layout.owner {
        Some(OwnerKind::Direct) => crate::transition::read(xml)?
            .map(Some)
            .ok_or_else(|| invalid("direct transition disappeared during semantic read")),
        Some(OwnerKind::AlternateContent) => {
            // Clearing the active branch leaves the AlternateContent owner in
            // place so inactive choices and fallback bytes remain intact.
            crate::transition::read(xml)
        },
        None => Err(invalid("direct transition owner kind is unavailable")),
    }
}

fn validate_owner_for_edit(
    xml: &[u8],
    layout: &Layout,
    current: &Transition,
    operation: &'static str,
) -> Result<()> {
    let Some(range) = layout.transition.as_ref() else {
        return Ok(());
    };
    let bytes = xml
        .get(range.clone())
        .ok_or_else(|| invalid("direct transition range is outside its slide"))?;
    match layout.owner {
        Some(OwnerKind::Direct) => {
            validate_direct_subtree(bytes, &layout.prefix, &layout.namespace)
        },
        Some(OwnerKind::AlternateContent) => {
            if matches!(current.kind(), Kind::Raw(_)) {
                return Err(Error::UnsafeEdit {
                    operation,
                    reason: "source-backed transition edits refuse an unknown active extension effect",
                });
            }
            validate_active_transition(xml, operation)
        },
        None => Err(invalid("direct transition owner kind is unavailable")),
    }
}

fn validate_active_transition(xml: &[u8], operation: &'static str) -> Result<()> {
    let mut capabilities = litchi_ooxml_common::mce::Capabilities::ooxml_baseline();
    capabilities
        .understand_namespace("http://schemas.microsoft.com/office/powerpoint/2010/main")
        .understand_namespace("http://schemas.microsoft.com/office/powerpoint/2012/main")
        .understand_namespace("http://schemas.microsoft.com/office/powerpoint/2015/09/main");
    let mce_limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: 64 * 1024 * 1024,
        max_output_bytes: 64 * 1024 * 1024,
        max_depth: MAX_DEPTH,
        ..litchi_ooxml_common::mce::Limits::default()
    };
    let processed =
        litchi_ooxml_common::mce::process_markup_compatibility(xml, &capabilities, &mce_limits)?
            .xml;
    let active_layout = locate(processed.as_ref(), operation)?;
    let Some(range) = active_layout.transition.as_ref() else {
        return Err(invalid(
            "active transition owner has no selected transition",
        ));
    };
    if active_layout.owner != Some(OwnerKind::Direct) {
        return Err(invalid(
            "active transition owner did not reduce to one direct transition",
        ));
    }
    // The preprocessor declares each namespace once, where XML requires it
    // (change 0653), so the sliced owner is re-declared with every binding it
    // inherits before it is validated standalone.
    let active = litchi_ooxml_common::mce::self_contained_fragment(
        processed.as_ref(),
        range.start,
        range.len(),
        &mce_limits,
    )?;
    validate_direct_subtree_with_extensions(
        active.as_ref(),
        &active_layout.prefix,
        &active_layout.namespace,
        true,
    )?;
    // The structural validator above deliberately knows only source-edit
    // placement and attribute names. Run the bounded typed reader as well so
    // token/boolean domains, duplicate attributes, and exact p14 timing are
    // checked before a destructive branch splice.
    let _ = crate::transition::read(processed.as_ref())?
        .ok_or_else(|| invalid("active transition disappeared during typed validation"))?;
    Ok(())
}

fn remove_owner(xml: &[u8], range: &Range<usize>) -> Result<Vec<u8>> {
    let (start, end) = (range.start, range.end);
    if start > end || end > xml.len() {
        return Err(invalid("direct transition range is outside its slide"));
    }
    let output_len = xml
        .len()
        .checked_sub(range.len())
        .ok_or_else(|| invalid("source slide owner output size underflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "source-backed slide transition validation XML",
            source,
        })?;
    output.extend_from_slice(&xml[..start]);
    output.extend_from_slice(&xml[end..]);
    Ok(output)
}

fn direct_fragment(value: &Transition, prefix: &str, operation: &'static str) -> Result<Vec<u8>> {
    validate_target(value, operation)?;
    let canonical = crate::transition::write(value)?;
    Ok(rewrite_pml_qnames(canonical.as_bytes(), prefix)?.into_bytes())
}

/// Splice a replacement into the active branch of an existing
/// `mc:AlternateContent` owner.  The wrapper and every inactive branch remain
/// byte-for-byte source material; only the selected transition subtree is
/// replaced.  This is what lets a typed edit preserve vendor choices,
/// fallback attributes, comments, and processing instructions it does not
/// own.
fn alternate_fragment(
    xml: &[u8],
    layout: &Layout,
    target: &Transition,
    operation: &'static str,
) -> Result<Vec<u8>> {
    let owner = layout
        .transition
        .as_ref()
        .ok_or_else(|| invalid("alternate transition owner has no source range"))?;
    let active = active_transition_range(xml, owner, operation)?;
    let Some(active) = active else {
        return Err(Error::UnsafeEdit {
            operation,
            reason: "source-backed transition edit cannot insert into an AlternateContent owner without an active transition",
        });
    };
    let canonical = crate::transition::write(target)?;
    let canonical_range = first_transition_range(&canonical)?;
    let fragment = canonical
        .as_bytes()
        .get(canonical_range)
        .ok_or_else(|| invalid("canonical transition range is outside its output"))?;
    let fragment = rewrite_pml_qnames(fragment, &layout.prefix)?;
    let owner_start = owner.start;
    let relative_start = active
        .start
        .checked_sub(owner_start)
        .ok_or_else(|| invalid("active transition precedes its AlternateContent owner"))?;
    let relative_end = active
        .end
        .checked_sub(owner_start)
        .ok_or_else(|| invalid("active transition end precedes its owner"))?;
    let owner_bytes = xml
        .get(owner.clone())
        .ok_or_else(|| invalid("AlternateContent owner range is outside its slide"))?;
    if relative_start > relative_end || relative_end > owner_bytes.len() {
        return Err(invalid("active transition range is outside its owner"));
    }
    let mut output = replace_range_bytes(
        owner_bytes,
        relative_start..relative_end,
        fragment.as_bytes(),
        "source-backed AlternateContent transition XML",
    )?;
    if let Some(requires) = transition_requirements(target) {
        let choice = active_choice(xml, owner, &active, requires)?.ok_or(Error::UnsafeEdit {
            operation,
            reason: "an extension transition requires an active Choice branch",
        })?;
        let opening = &owner_bytes[choice.opening.clone()];
        let mut replacement = replace_range_bytes(
            opening,
            choice.requires.start - choice.opening.start
                ..choice.requires.end - choice.opening.start,
            requires.as_bytes(),
            "source-backed AlternateContent Requires attribute",
        )?;
        // Scope new declarations to the selected Choice. Declaring them on
        // AlternateContent would change inherited bindings in inactive branches.
        ensure_extension_namespaces(&mut replacement, Some(requires))?;
        output = replace_range_bytes(
            &output,
            choice.opening,
            &replacement,
            "source-backed AlternateContent Choice opening tag",
        )?;
    }
    Ok(output)
}

fn alternate_fragment_len(
    xml: &[u8],
    layout: &Layout,
    target: &Transition,
    operation: &'static str,
) -> Result<usize> {
    let owner = layout
        .transition
        .as_ref()
        .ok_or_else(|| invalid("alternate transition owner has no source range"))?;
    let active = active_transition_range(xml, owner, operation)?.ok_or(Error::UnsafeEdit {
        operation,
        reason: "source-backed transition edit cannot insert into an AlternateContent owner without an active transition",
    })?;
    let canonical = crate::transition::write(target)?;
    let canonical_range = first_transition_range(&canonical)?;
    let fragment = canonical
        .as_bytes()
        .get(canonical_range)
        .ok_or_else(|| invalid("canonical transition range is outside its output"))?;
    let fragment_len = rewrite_pml_qnames_len(fragment, &layout.prefix)?;
    let owner_start = owner.start;
    let relative_start = active
        .start
        .checked_sub(owner_start)
        .ok_or_else(|| invalid("active transition precedes its AlternateContent owner"))?;
    let relative_end = active
        .end
        .checked_sub(owner_start)
        .ok_or_else(|| invalid("active transition end precedes its owner"))?;
    let owner_bytes = xml
        .get(owner.clone())
        .ok_or_else(|| invalid("AlternateContent owner range is outside its slide"))?;
    if relative_start > relative_end || relative_end > owner_bytes.len() {
        return Err(invalid("active transition range is outside its owner"));
    }
    let mut output_len = owner_bytes
        .len()
        .checked_sub(relative_end - relative_start)
        .and_then(|length| length.checked_add(fragment_len))
        .ok_or_else(|| invalid("AlternateContent replacement size overflow"))?;
    if let Some(requires) = transition_requirements(target) {
        let choice = active_choice(xml, owner, &active, requires)?.ok_or(Error::UnsafeEdit {
            operation,
            reason: "an extension transition requires an active Choice branch",
        })?;
        let opening = owner_bytes
            .get(choice.opening.clone())
            .ok_or_else(|| invalid("Choice opening range is outside its owner"))?;
        let additions = extension_namespace_additions_len(opening, Some(requires))?;
        let replacement_opening_len = opening
            .len()
            .checked_sub(choice.requires.end - choice.requires.start)
            .and_then(|length| length.checked_add(requires.len()))
            .and_then(|length| length.checked_add(additions))
            .ok_or_else(|| invalid("Choice opening replacement size overflow"))?;
        output_len = output_len
            .checked_sub(choice.opening.end - choice.opening.start)
            .and_then(|length| length.checked_add(replacement_opening_len))
            .ok_or_else(|| invalid("AlternateContent replacement size overflow"))?;
    }
    Ok(output_len)
}

fn transition_requirements(value: &Transition) -> Option<&'static str> {
    if matches!(
        value.kind(),
        Kind::Ripple(_)
            | Kind::Conveyor(_)
            | Kind::Doors(_)
            | Kind::Ferris(_)
            | Kind::Flash
            | Kind::Flip(_)
            | Kind::FlyThrough(_)
            | Kind::Gallery(_)
            | Kind::Glitter(_)
            | Kind::Honeycomb
            | Kind::Pan(_)
            | Kind::Prism(_)
            | Kind::Reveal(_)
            | Kind::Shred(_)
            | Kind::Switch(_)
            | Kind::Vortex(_)
            | Kind::Warp(_)
            | Kind::WheelReverse(_)
            | Kind::Window(_)
    ) {
        Some("p14")
    } else if matches!(value.kind(), Kind::Morph(_)) {
        if value.duration_offset().is_some() {
            Some("p14 p159")
        } else {
            Some("p159")
        }
    } else if matches!(value.kind(), Kind::Preset(_)) {
        if value.duration_offset().is_some() {
            Some("p14 p15")
        } else {
            Some("p15")
        }
    } else if value.duration_offset().is_some() {
        Some("p14")
    } else {
        None
    }
}

fn replace_range_bytes(
    source: &[u8],
    range: Range<usize>,
    replacement: &[u8],
    resource: &'static str,
) -> Result<Vec<u8>> {
    let start = range.start;
    let end = range.end;
    if start > end || end > source.len() {
        return Err(invalid("replacement range is outside its source"));
    }
    let output_len = source
        .len()
        .checked_sub(end - start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| invalid("replacement size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation { resource, source })?;
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[end..]);
    if output.len() != output_len {
        return Err(invalid("replacement length changed during emission"));
    }
    Ok(output)
}

fn ensure_extension_namespaces(xml: &mut Vec<u8>, requirements: Option<&str>) -> Result<()> {
    let Some(requirements) = requirements else {
        return Ok(());
    };
    let root_end = find_tag_end(xml, 0)?;
    let root = xml
        .get(..root_end)
        .ok_or_else(|| invalid("AlternateContent root range is outside its output"))?;
    let mut additions = String::new();
    for token in requirements.split_ascii_whitespace() {
        let uri = match token {
            "p14" => "http://schemas.microsoft.com/office/powerpoint/2010/main",
            "p15" => "http://schemas.microsoft.com/office/powerpoint/2012/main",
            "p159" => "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
            _ => continue,
        };
        let declaration = format!("xmlns:{token}=\"");
        if !root_has_namespace(root, token, uri)? {
            additions.push(' ');
            additions.push_str(&declaration);
            additions.push_str(uri);
            additions.push('"');
        }
    }
    if additions.is_empty() {
        return Ok(());
    }
    let insertion = root_end
        .checked_sub(1)
        .ok_or_else(|| invalid("AlternateContent root has no closing delimiter"))?;
    let output_len = xml
        .len()
        .checked_add(additions.len())
        .ok_or_else(|| invalid("AlternateContent namespace output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "source-backed AlternateContent namespace XML",
            source,
        })?;
    output.extend_from_slice(&xml[..insertion]);
    output.extend_from_slice(additions.as_bytes());
    output.extend_from_slice(&xml[insertion..]);
    if output.len() != output_len {
        return Err(invalid(
            "AlternateContent namespace output length changed during emission",
        ));
    }
    *xml = output;
    Ok(())
}

fn extension_namespace_additions_len(root: &[u8], requirements: Option<&str>) -> Result<usize> {
    let Some(requirements) = requirements else {
        return Ok(0);
    };
    let root_end = find_tag_end(root, 0)?;
    let root = root
        .get(..root_end)
        .ok_or_else(|| invalid("namespace root range is outside its opening tag"))?;
    let mut additions = 0usize;
    for token in requirements.split_ascii_whitespace() {
        let uri = match token {
            "p14" => "http://schemas.microsoft.com/office/powerpoint/2010/main",
            "p15" => "http://schemas.microsoft.com/office/powerpoint/2012/main",
            "p159" => "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
            _ => continue,
        };
        if !root_has_namespace(root, token, uri)? {
            let declaration_len = 1usize
                .checked_add("xmlns:".len())
                .and_then(|length| length.checked_add(token.len()))
                .and_then(|length| length.checked_add(2))
                .and_then(|length| length.checked_add(uri.len()))
                .and_then(|length| length.checked_add(1))
                .ok_or_else(|| invalid("namespace declaration size overflow"))?;
            additions = additions
                .checked_add(declaration_len)
                .ok_or_else(|| invalid("namespace declaration size overflow"))?;
        }
    }
    Ok(additions)
}

fn root_has_namespace(root: &[u8], prefix: &str, expected: &str) -> Result<bool> {
    let mut reader = NsReader::from_reader(root);
    let event = reader
        .read_event()
        .map_err(|error| Error::Xml(error.to_string()))?;
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element,
        _ => return Err(invalid("AlternateContent root opening tag is unavailable")),
    };
    let expected_name = format!("xmlns:{prefix}");
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() != expected_name.as_bytes() {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value != expected {
            return Err(invalid(
                "AlternateContent extension prefix has a conflicting namespace binding",
            ));
        }
        return Ok(true);
    }
    Ok(false)
}

fn clear_alternate_fragment(
    xml: &[u8],
    layout: &Layout,
    operation: &'static str,
) -> Result<Vec<u8>> {
    let owner = layout
        .transition
        .as_ref()
        .ok_or_else(|| invalid("alternate transition owner has no source range"))?;
    let Some(active) = active_transition_range(xml, owner, operation)? else {
        return Ok(xml
            .get(owner.clone())
            .ok_or_else(|| invalid("AlternateContent owner range is outside its slide"))?
            .to_vec());
    };
    let owner_bytes = xml
        .get(owner.clone())
        .ok_or_else(|| invalid("AlternateContent owner range is outside its slide"))?;
    let relative_start = active
        .start
        .checked_sub(owner.start)
        .ok_or_else(|| invalid("active transition precedes its AlternateContent owner"))?;
    let relative_end = active
        .end
        .checked_sub(owner.start)
        .ok_or_else(|| invalid("active transition end precedes its owner"))?;
    if relative_start > relative_end || relative_end > owner_bytes.len() {
        return Err(invalid("active transition range is outside its owner"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(owner_bytes.len() - (relative_end - relative_start))
        .map_err(|source| Error::Allocation {
            resource: "source-backed AlternateContent transition XML",
            source,
        })?;
    output.extend_from_slice(&owner_bytes[..relative_start]);
    output.extend_from_slice(&owner_bytes[relative_end..]);
    Ok(output)
}

fn clear_alternate_fragment_len(xml: &[u8], layout: &Layout) -> Result<usize> {
    let owner = layout
        .transition
        .as_ref()
        .ok_or_else(|| invalid("alternate transition owner has no source range"))?;
    let Some(active) = active_transition_range(xml, owner, "clear_transition")? else {
        return Ok(xml
            .get(owner.clone())
            .ok_or_else(|| invalid("AlternateContent owner range is outside its slide"))?
            .len());
    };
    let relative_start = active
        .start
        .checked_sub(owner.start)
        .ok_or_else(|| invalid("active transition precedes its AlternateContent owner"))?;
    let relative_end = active
        .end
        .checked_sub(owner.start)
        .ok_or_else(|| invalid("active transition end precedes its owner"))?;
    let owner_bytes = xml
        .get(owner.clone())
        .ok_or_else(|| invalid("AlternateContent owner range is outside its slide"))?;
    if relative_start > relative_end || relative_end > owner_bytes.len() {
        return Err(invalid("active transition range is outside its owner"));
    }
    owner_bytes
        .len()
        .checked_sub(relative_end - relative_start)
        .ok_or_else(|| invalid("AlternateContent replacement size underflow"))
}

#[derive(Clone)]
struct ActiveChoice {
    opening: Range<usize>,
    requires: Range<usize>,
    bindings_match: bool,
}

fn active_choice(
    xml: &[u8],
    owner: &Range<usize>,
    active: &Range<usize>,
    requirements: &str,
) -> Result<Option<ActiveChoice>> {
    // Read from the slide root so inherited PML and MCE bindings remain in
    // scope. A detached AlternateContent fragment cannot resolve those names.
    let mut reader = NsReader::from_reader(xml);
    let mut stack: Vec<Option<ActiveChoice>> = Vec::new();
    loop {
        let start = position(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = position(&reader)?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        match event {
            Event::Start(element) => {
                let choice = if start >= owner.start
                    && start < owner.end
                    && is_markup_element(&namespace)
                    && element.name().local_name().as_ref() == b"Choice"
                {
                    let requires = requires_attribute_range(xml, start, end)?;
                    Some(ActiveChoice {
                        opening: start - owner.start..end - owner.start,
                        requires: requires.start - owner.start..requires.end - owner.start,
                        bindings_match: validate_extension_bindings(&reader, requirements).is_ok(),
                    })
                } else {
                    None
                };
                if is_pml_element(&namespace, element.name(), b"transition")
                    && start == active.start
                {
                    validate_extension_bindings(&reader, requirements)?;
                    return checked_active_choice(&stack);
                }
                stack.push(choice);
            },
            Event::Empty(element) => {
                if is_pml_element(&namespace, element.name(), b"transition")
                    && start == active.start
                {
                    validate_extension_bindings(&reader, requirements)?;
                    return checked_active_choice(&stack);
                }
            },
            Event::End(_) => {
                stack
                    .pop()
                    .ok_or_else(|| invalid("AlternateContent branch stack underflow"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(None)
}

fn checked_active_choice(stack: &[Option<ActiveChoice>]) -> Result<Option<ActiveChoice>> {
    let choice = stack.iter().rev().find_map(|value| value.clone());
    if choice.as_ref().is_some_and(|choice| !choice.bindings_match) {
        return Err(Error::UnsafeEdit {
            operation: "set_transition",
            reason: "transition extension prefix conflicts with the active Choice namespace binding",
        });
    }
    Ok(choice)
}

fn validate_extension_bindings(reader: &NsReader<&[u8]>, requirements: &str) -> Result<()> {
    for token in requirements.split_ascii_whitespace() {
        let expected = match token {
            "p14" => "http://schemas.microsoft.com/office/powerpoint/2010/main",
            "p15" => "http://schemas.microsoft.com/office/powerpoint/2012/main",
            "p159" => "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
            _ => return Err(invalid("unknown transition namespace requirement")),
        };
        let name = format!("{token}:transition");
        let (namespace, _) = reader.resolver().resolve_element(QName(name.as_bytes()));
        if let ResolveResult::Bound(namespace) = namespace
            && namespace.as_ref() != expected.as_bytes()
        {
            return Err(Error::UnsafeEdit {
                operation: "set_transition",
                reason: "transition extension prefix conflicts with an inherited namespace binding",
            });
        }
    }
    Ok(())
}

fn requires_attribute_range(xml: &[u8], start: usize, end: usize) -> Result<Range<usize>> {
    let tag = xml
        .get(start..end)
        .ok_or_else(|| invalid("Choice opening tag range is outside its owner"))?;
    let mut cursor = 1usize;
    if tag.get(cursor) == Some(&b'/') {
        cursor += 1;
    }
    while cursor < tag.len()
        && !tag[cursor].is_ascii_whitespace()
        && !matches!(tag[cursor], b'/' | b'>')
    {
        cursor += 1;
    }
    loop {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || tag[cursor] == b'>' || tag[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name = tag
            .get(name_start..cursor)
            .ok_or_else(|| invalid("Choice attribute name range is invalid"))?;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(invalid("Choice attribute is missing its equals sign"));
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .ok_or_else(|| invalid("Choice attribute value is truncated"))?;
        if quote != b'"' && quote != b'\'' {
            return Err(invalid("Choice attribute value is not quoted"));
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= tag.len() {
            return Err(invalid("Choice attribute value is unterminated"));
        }
        if name == b"Requires" {
            return Ok(start + value_start..start + value_end);
        }
        cursor += 1;
    }
    Err(invalid("Choice lacks its Requires attribute"))
}

/// Find the transition selected by the same bounded MCE processor used for
/// typed reads.  `active_offsets` preserves source coordinates while it
/// selects a branch, avoiding a lossy reserialization of the owner.
fn active_transition_range(
    xml: &[u8],
    owner: &Range<usize>,
    operation: &'static str,
) -> Result<Option<Range<usize>>> {
    let transition_offsets = transition_start_offsets(xml, owner, operation)?;
    if transition_offsets.is_empty() {
        return Ok(None);
    }
    let mut capabilities = litchi_ooxml_common::mce::Capabilities::ooxml_baseline();
    capabilities
        .understand_namespace("http://schemas.microsoft.com/office/powerpoint/2010/main")
        .understand_namespace("http://schemas.microsoft.com/office/powerpoint/2012/main")
        .understand_namespace("http://schemas.microsoft.com/office/powerpoint/2015/09/main");
    let limits = litchi_ooxml_common::mce::OffsetLimits {
        max_source_bytes: 64 * 1024 * 1024,
        max_offsets: MAX_NODES,
        max_marked_bytes: 64 * 1024 * 1024,
        processing: litchi_ooxml_common::mce::Limits {
            max_input_bytes: 64 * 1024 * 1024,
            max_output_bytes: 64 * 1024 * 1024,
            max_depth: MAX_DEPTH,
            ..litchi_ooxml_common::mce::Limits::default()
        },
    };
    let offsets = transition_offsets
        .iter()
        .map(|offset| {
            u32::try_from(*offset).map_err(|_| Error::Limit {
                resource: "source-backed transition offset",
                limit: u32::MAX as usize,
            })
        })
        .collect::<Result<Vec<_>>>()?;
    let selected = litchi_ooxml_common::mce::active_offsets(xml, &offsets, &capabilities, &limits)
        .map_err(|error| {
            Error::Invalid(format!(
                "source-backed transition MCE selection failed: {error}"
            ))
        })?;
    if selected.len() > 1 {
        return Err(invalid(
            "active AlternateContent branch contains multiple transitions",
        ));
    }
    let Some(offset) = selected.first().copied() else {
        return Ok(None);
    };
    transition_element_range(
        xml,
        usize::try_from(offset).map_err(|_| Error::Limit {
            resource: "source-backed transition offset",
            limit: usize::MAX,
        })?,
    )
}

fn transition_start_offsets(
    xml: &[u8],
    owner: &Range<usize>,
    _operation: &'static str,
) -> Result<Vec<usize>> {
    let owner_bytes = xml
        .get(owner.clone())
        .ok_or_else(|| invalid("AlternateContent owner range is outside its slide"))?;
    let mut reader = NsReader::from_reader(owner_bytes);
    let mut offsets = Vec::new();
    loop {
        let local_start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_pml_element(&namespace, element.name(), b"transition") =>
            {
                let absolute = owner
                    .start
                    .checked_add(local_start)
                    .ok_or_else(|| invalid("transition offset overflow"))?;
                offsets.push(absolute);
                if offsets.len() > MAX_NODES {
                    return Err(Error::Limit {
                        resource: "source-backed transition XML nodes",
                        limit: MAX_NODES,
                    });
                }
            },
            Event::Eof => break,
            Event::DocType(_) => return Err(invalid("DOCTYPE is forbidden in transition XML")),
            Event::PI(_) => {},
            _ => {},
        }
    }
    Ok(offsets)
}

fn transition_element_range(xml: &[u8], wanted_start: usize) -> Result<Option<Range<usize>>> {
    if wanted_start >= xml.len() {
        return Err(invalid("transition offset is outside its slide"));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut wanted = None;
    loop {
        let start = position(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                let name = element.name().as_ref().to_vec();
                stack.push(name);
                if start == wanted_start {
                    wanted = Some((stack.len(), start));
                }
            },
            Event::Empty(_) => {
                if start == wanted_start {
                    return Ok(Some(start..end));
                }
            },
            Event::End(element) => {
                let name = stack
                    .pop()
                    .ok_or_else(|| invalid("transition range has unmatched end"))?;
                if name.as_slice() != element.name().as_ref() {
                    return Err(invalid("transition range has mismatched end"));
                }
                if let Some((depth, start)) = wanted
                    && stack.len() + 1 == depth
                {
                    return Ok(Some(start..end));
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(None)
}

fn first_transition_range(xml: &str) -> Result<Range<usize>> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut stack: Vec<(Vec<u8>, usize)> = Vec::new();
    loop {
        let start = position(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = position(&reader)?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        match event {
            Event::Start(element) => {
                let is_transition = is_pml_element(&namespace, element.name(), b"transition");
                stack.push((element.name().as_ref().to_vec(), start));
                if is_transition {
                    loop {
                        let inner_end = position(&reader)?;
                        let event = reader
                            .read_event()
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        let final_end = position(&reader)?;
                        match event {
                            Event::Start(element) => {
                                stack.push((element.name().as_ref().to_vec(), inner_end));
                            },
                            Event::Empty(_) => {},
                            Event::End(element) => {
                                let name = stack.pop().ok_or_else(|| {
                                    invalid("canonical transition has unmatched end")
                                })?;
                                if name.0.as_slice() != element.name().as_ref() {
                                    return Err(invalid("canonical transition has mismatched end"));
                                }
                                if stack.last().is_none_or(|(_, parent)| *parent < start) {
                                    return Ok(start..final_end);
                                }
                            },
                            Event::Eof => break,
                            _ => {},
                        }
                    }
                }
            },
            Event::Empty(element) if is_pml_element(&namespace, element.name(), b"transition") => {
                return Ok(start..end);
            },
            Event::End(element) => {
                let (name, _) = stack
                    .pop()
                    .ok_or_else(|| invalid("canonical transition has unmatched end"))?;
                if name.as_slice() != element.name().as_ref() {
                    return Err(invalid("canonical transition has mismatched end"));
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Err(invalid("canonical transition has no transition element"))
}

/// Rewrite only XML qualified names in markup.  Attribute values, text, raw
/// retained effect XML, and URI strings are copied without interpretation.
fn rewrite_pml_qnames(xml: &[u8], prefix: &str) -> Result<String> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len().saturating_add(prefix.len().saturating_mul(8)))
        .map_err(|source| Error::Allocation {
            resource: "source-backed transition QName rewrite",
            source,
        })?;
    let mut cursor = 0usize;
    while cursor < xml.len() {
        let Some(relative) = xml[cursor..].iter().position(|byte| *byte == b'<') else {
            output.extend_from_slice(&xml[cursor..]);
            break;
        };
        let tag_start = cursor + relative;
        output.extend_from_slice(&xml[cursor..tag_start]);
        if let Some(end) = markup_passthrough_end(xml, tag_start)? {
            output.extend_from_slice(
                xml.get(tag_start..end)
                    .ok_or_else(|| invalid("transition passthrough range is invalid"))?,
            );
            cursor = end;
            continue;
        }
        let tag_end = find_tag_end(xml, tag_start)?;
        rewrite_tag(&xml[tag_start..tag_end], prefix, &mut output)?;
        cursor = tag_end;
    }
    String::from_utf8(output)
        .map_err(|error| Error::Xml(format!("transition QName rewrite is not UTF-8: {error}")))
}

fn rewrite_pml_qnames_len(xml: &[u8], prefix: &str) -> Result<usize> {
    let mut output_len = 0usize;
    let mut cursor = 0usize;
    while cursor < xml.len() {
        let Some(relative) = xml[cursor..].iter().position(|byte| *byte == b'<') else {
            output_len = output_len
                .checked_add(xml.len() - cursor)
                .ok_or_else(|| invalid("transition QName rewrite size overflow"))?;
            break;
        };
        let tag_start = cursor + relative;
        output_len = output_len
            .checked_add(tag_start - cursor)
            .ok_or_else(|| invalid("transition QName rewrite size overflow"))?;
        if let Some(end) = markup_passthrough_end(xml, tag_start)? {
            output_len = output_len
                .checked_add(end - tag_start)
                .ok_or_else(|| invalid("transition QName rewrite size overflow"))?;
            cursor = end;
            continue;
        }
        let tag_end = find_tag_end(xml, tag_start)?;
        output_len = output_len
            .checked_add(rewrite_tag_len(&xml[tag_start..tag_end], prefix)?)
            .ok_or_else(|| invalid("transition QName rewrite size overflow"))?;
        cursor = tag_end;
    }
    Ok(output_len)
}

fn rewrite_tag_len(tag: &[u8], prefix: &str) -> Result<usize> {
    if tag.len() < 2 || tag[0] != b'<' {
        return Err(invalid("transition QName rewrite saw an invalid tag"));
    }
    let mut output_len = tag.len();
    let mut cursor = 1usize;
    if tag.get(cursor) == Some(&b'/') {
        cursor += 1;
    }
    let name_end = tag[cursor..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
        .map_or(tag.len(), |offset| cursor + offset);
    output_len = rewritten_qname_len(output_len, &tag[cursor..name_end], prefix)?;
    cursor = name_end;
    let mut quote = None;
    while cursor < tag.len() {
        let byte = tag[cursor];
        if quote.is_none() {
            if byte == b'>' {
                cursor += 1;
                break;
            }
            if byte == b'"' || byte == b'\'' {
                quote = Some(byte);
                cursor += 1;
                continue;
            }
            if byte.is_ascii_whitespace() || byte == b'/' || byte == b'=' {
                cursor += 1;
                continue;
            }
            let attr_end = tag[cursor..]
                .iter()
                .position(|value| {
                    value.is_ascii_whitespace() || matches!(value, b'=' | b'/' | b'>')
                })
                .map_or(tag.len(), |offset| cursor + offset);
            output_len = rewritten_qname_len(output_len, &tag[cursor..attr_end], prefix)?;
            cursor = attr_end;
        } else {
            if byte == quote.unwrap() {
                quote = None;
            }
            cursor += 1;
        }
    }
    if cursor != tag.len() {
        return Err(invalid("transition QName rewrite stopped before tag end"));
    }
    Ok(output_len)
}

fn rewritten_qname_len(current: usize, name: &[u8], prefix: &str) -> Result<usize> {
    if !name.starts_with(b"p:") {
        return Ok(current);
    }
    let replacement = if prefix.is_empty() {
        name.len()
            .checked_sub(2)
            .ok_or_else(|| invalid("transition QName is shorter than its prefix"))?
    } else {
        prefix
            .len()
            .checked_add(1)
            .and_then(|length| length.checked_add(name.len().checked_sub(2)?))
            .ok_or_else(|| invalid("transition QName replacement size overflow"))?
    };
    current
        .checked_sub(name.len())
        .and_then(|length| length.checked_add(replacement))
        .ok_or_else(|| invalid("transition QName rewrite size underflow"))
}

fn markup_passthrough_end(xml: &[u8], start: usize) -> Result<Option<usize>> {
    let tail = xml
        .get(start..)
        .ok_or_else(|| invalid("transition passthrough starts outside its output"))?;
    let marker = if tail.starts_with(b"<!--") {
        Some(b"-->".as_slice())
    } else if tail.starts_with(b"<![CDATA[") {
        Some(b"]]>".as_slice())
    } else if tail.starts_with(b"<?") {
        Some(b"?>".as_slice())
    } else {
        None
    };
    let Some(marker) = marker else {
        return Ok(None);
    };
    let relative = tail
        .windows(marker.len())
        .position(|window| window == marker)
        .ok_or_else(|| invalid("transition passthrough markup is unterminated"))?;
    start
        .checked_add(relative + marker.len())
        .map(Some)
        .ok_or_else(|| invalid("transition passthrough range overflow"))
}

fn find_tag_end(xml: &[u8], start: usize) -> Result<usize> {
    let mut quote = None;
    for (offset, byte) in xml
        .get(start..)
        .ok_or_else(|| invalid("transition tag starts outside its output"))?
        .iter()
        .enumerate()
    {
        match (quote, *byte) {
            (None, b'"' | b'\'') => quote = Some(*byte),
            (Some(value), byte) if value == byte => quote = None,
            (None, b'>') => {
                return start
                    .checked_add(offset + 1)
                    .ok_or_else(|| invalid("transition tag end offset overflow"));
            },
            _ => {},
        }
    }
    Err(invalid("transition tag is not closed"))
}

fn rewrite_tag(tag: &[u8], prefix: &str, output: &mut Vec<u8>) -> Result<()> {
    if tag.len() < 2 || tag[0] != b'<' {
        return Err(invalid("transition QName rewrite saw an invalid tag"));
    }
    output.push(b'<');
    let mut cursor = 1usize;
    if tag.get(cursor) == Some(&b'/') {
        output.push(b'/');
        cursor += 1;
    }
    let name_end = tag[cursor..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'/' | b'>'))
        .map_or(tag.len(), |offset| cursor + offset);
    let name = tag
        .get(cursor..name_end)
        .ok_or_else(|| invalid("transition element name range is invalid"))?;
    write_rewritten_qname(name, prefix, output)?;
    cursor = name_end;
    let mut quote = None;
    while cursor < tag.len() {
        let byte = tag[cursor];
        if quote.is_none() {
            if byte == b'>' {
                output.push(byte);
                cursor += 1;
                break;
            }
            if byte == b'"' || byte == b'\'' {
                quote = Some(byte);
                output.push(byte);
                cursor += 1;
                continue;
            }
            if byte.is_ascii_whitespace() || byte == b'/' || byte == b'=' {
                output.push(byte);
                cursor += 1;
                continue;
            }
            let attr_end = tag[cursor..]
                .iter()
                .position(|value| {
                    value.is_ascii_whitespace() || matches!(value, b'=' | b'/' | b'>')
                })
                .map_or(tag.len(), |offset| cursor + offset);
            let attr = tag
                .get(cursor..attr_end)
                .ok_or_else(|| invalid("transition attribute name range is invalid"))?;
            write_rewritten_qname(attr, prefix, output)?;
            cursor = attr_end;
        } else {
            output.push(byte);
            if byte == quote.unwrap() {
                quote = None;
            }
            cursor += 1;
        }
    }
    if cursor != tag.len() {
        return Err(invalid("transition QName rewrite stopped before tag end"));
    }
    Ok(())
}

fn write_rewritten_qname(name: &[u8], prefix: &str, output: &mut Vec<u8>) -> Result<()> {
    if name.starts_with(b"p:") {
        if !prefix.is_empty() {
            output.extend_from_slice(prefix.as_bytes());
            output.push(b':');
        }
        output.extend_from_slice(&name[2..]);
    } else {
        output.extend_from_slice(name);
    }
    Ok(())
}

fn validate_target(value: &Transition, operation: &'static str) -> Result<()> {
    if matches!(value.kind(), Kind::Raw(_)) {
        return Err(Error::UnsafeEdit {
            operation,
            reason: "source-backed transition edits refuse untyped transition effects",
        });
    }
    // The public model can only retain raw children produced by the bounded
    // reader. Ask the writer to validate portability before any output
    // allocation; preserved effect/auxiliary XML is intentionally retained
    // when a caller edits timing around a typed effect.
    crate::transition::write(value)?;
    Ok(())
}

fn locate(xml: &[u8], operation: &'static str) -> Result<Layout> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_prefix = None;
    let mut root_namespace = None;
    let mut root_close = None;
    let mut transition = None;
    let mut owner = None;
    let mut open_transition = None;
    let mut open_owner = None;
    let mut previous_rank = 0u8;
    let mut insertion = None;
    let mut child_counts = [0u8; 5];

    loop {
        let start = position(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = position(&reader)?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        match event {
            Event::Start(element) => {
                nodes = add_node(nodes)?;
                depth = enter_depth(depth)?;
                if depth == 1 {
                    if root_seen || !is_pml(&namespace, element.name(), b"sld") {
                        return Err(invalid("slide XML must contain exactly one p:sld root"));
                    }
                    root_seen = true;
                    root_prefix = Some(qname_prefix(element.name())?);
                    root_namespace = Some(pml_namespace(&namespace)?);
                } else if depth == 2 {
                    let rank = child_rank(&namespace, element.name()).ok_or(Error::UnsafeEdit {
                        operation,
                        reason: "source-backed transition edits refuse unknown direct slide children",
                    })?;
                    record_child(
                        rank,
                        start,
                        &mut previous_rank,
                        &mut insertion,
                        &mut child_counts,
                    )?;
                    if rank == 3 {
                        if transition.is_some() || open_transition.is_some() {
                            return Err(invalid("slide contains duplicate direct transitions"));
                        }
                        open_transition = Some(start);
                        open_owner = Some(if is_pml(&namespace, element.name(), b"transition") {
                            OwnerKind::Direct
                        } else {
                            OwnerKind::AlternateContent
                        });
                    }
                } else if is_markup_element(&namespace)
                    && !matches!(open_owner, Some(OwnerKind::AlternateContent))
                {
                    return Err(Error::UnsafeEdit {
                        operation,
                        reason: "source-backed transition edits refuse markup-compatibility outside the direct transition owner",
                    });
                } else if is_pml(&namespace, element.name(), b"transition")
                    && !matches!(open_owner, Some(OwnerKind::AlternateContent))
                {
                    return Err(Error::UnsafeEdit {
                        operation,
                        reason: "source-backed transition edits refuse nested or markup-compatibility transitions",
                    });
                }
            },
            Event::Empty(element) => {
                nodes = add_node(nodes)?;
                let event_depth = enter_depth(depth)?;
                if event_depth == 1 {
                    return Err(invalid("slide root cannot be empty"));
                }
                if event_depth == 2 {
                    let rank = child_rank(&namespace, element.name()).ok_or(Error::UnsafeEdit {
                        operation,
                        reason: "source-backed transition edits refuse unknown direct slide children",
                    })?;
                    record_child(
                        rank,
                        start,
                        &mut previous_rank,
                        &mut insertion,
                        &mut child_counts,
                    )?;
                    if rank == 3 {
                        if transition.replace(start..end).is_some() || open_transition.is_some() {
                            return Err(invalid("slide contains duplicate direct transitions"));
                        }
                        owner = Some(if is_pml(&namespace, element.name(), b"transition") {
                            OwnerKind::Direct
                        } else {
                            OwnerKind::AlternateContent
                        });
                    }
                } else if is_markup_element(&namespace)
                    && !matches!(open_owner, Some(OwnerKind::AlternateContent))
                {
                    return Err(Error::UnsafeEdit {
                        operation,
                        reason: "source-backed transition edits refuse markup-compatibility outside the direct transition owner",
                    });
                } else if is_pml(&namespace, element.name(), b"transition")
                    && !matches!(open_owner, Some(OwnerKind::AlternateContent))
                {
                    return Err(Error::UnsafeEdit {
                        operation,
                        reason: "source-backed transition edits refuse nested or markup-compatibility transitions",
                    });
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("slide XML contains an unmatched end element"));
                }
                if depth == 2 {
                    let is_direct_close = is_pml(&namespace, element.name(), b"transition");
                    let is_alternate_close = is_markup_element(&namespace)
                        && element.name().local_name().as_ref() == b"AlternateContent";
                    if is_direct_close || is_alternate_close {
                        let open = open_transition.take().ok_or_else(|| {
                            invalid("direct transition close has no matching open element")
                        })?;
                        let expected = open_owner.take().ok_or_else(|| {
                            invalid("direct transition owner kind is unavailable")
                        })?;
                        let actual = if is_direct_close {
                            OwnerKind::Direct
                        } else {
                            OwnerKind::AlternateContent
                        };
                        if expected != actual {
                            return Err(invalid(
                                "direct transition owner close does not match its open element",
                            ));
                        }
                        transition = Some(open..end);
                        owner = Some(actual);
                    }
                }
                if depth == 1 {
                    if !is_pml(&namespace, element.name(), b"sld") {
                        return Err(invalid("slide root close does not match p:sld"));
                    }
                    root_close = Some(start);
                }
                depth -= 1;
            },
            Event::DocType(_) => {
                return Err(invalid("DOCTYPE is forbidden in source-backed slide XML"));
            },
            Event::Text(text) => {
                if depth == 0 && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("slide XML has text outside its root"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(invalid("slide XML has data outside its root"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0
        || !root_seen
        || open_transition.is_some()
        || open_owner.is_some()
        || child_counts[0] != 1
    {
        return Err(invalid(
            "slide XML has an incomplete direct-child structure",
        ));
    }
    let root_close = root_close.ok_or_else(|| invalid("slide root is not closed"))?;
    Ok(Layout {
        transition,
        owner,
        insertion: insertion.unwrap_or(root_close),
        prefix: root_prefix
            .ok_or_else(|| invalid("slide root prefix is unavailable"))?
            .into_boxed_str(),
        namespace: root_namespace
            .ok_or_else(|| invalid("slide root namespace is unavailable"))?
            .into_boxed_str(),
    })
}

fn child_rank(namespace: &ResolveResult<'_>, name: QName<'_>) -> Option<u8> {
    if is_markup_element(namespace) && name.local_name().as_ref() == b"AlternateContent" {
        return Some(3);
    }
    if !is_pml_namespace(namespace) {
        return None;
    }
    match name.local_name().as_ref() {
        b"cSld" => Some(1),
        b"clrMapOvr" => Some(2),
        b"transition" => Some(3),
        b"timing" => Some(4),
        b"extLst" => Some(5),
        _ => None,
    }
}

fn record_child(
    rank: u8,
    start: usize,
    previous_rank: &mut u8,
    insertion: &mut Option<usize>,
    counts: &mut [u8; 5],
) -> Result<()> {
    if rank < *previous_rank {
        return Err(invalid("slide direct children are outside schema order"));
    }
    let slot = usize::from(rank - 1);
    counts[slot] = counts[slot]
        .checked_add(1)
        .ok_or_else(|| invalid("slide direct-child count overflow"))?;
    if counts[slot] > 1 {
        return Err(invalid("slide contains a duplicate direct child"));
    }
    if rank > 3 && insertion.is_none() {
        *insertion = Some(start);
    }
    *previous_rank = rank;
    Ok(())
}

fn validate_direct_subtree(xml: &[u8], prefix: &str, expected_namespace: &str) -> Result<()> {
    validate_direct_subtree_with_extensions(xml, prefix, expected_namespace, false)
}

fn validate_direct_subtree_with_extensions(
    xml: &[u8],
    prefix: &str,
    expected_namespace: &str,
    allow_extensions: bool,
) -> Result<()> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut effect_seen = false;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let (resolved, event) = reader.resolver().resolve_event(event);
        match event {
            Event::Start(element) => {
                nodes = add_node(nodes)?;
                depth = enter_depth(depth)?;
                validate_transition_element(
                    &resolved,
                    &element,
                    depth,
                    &mut effect_seen,
                    prefix,
                    expected_namespace,
                    allow_extensions,
                    reader.resolver(),
                )?;
            },
            Event::Empty(element) => {
                nodes = add_node(nodes)?;
                let event_depth = enter_depth(depth)?;
                validate_transition_element(
                    &resolved,
                    &element,
                    event_depth,
                    &mut effect_seen,
                    prefix,
                    expected_namespace,
                    allow_extensions,
                    reader.resolver(),
                )?;
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid(
                        "transition subtree contains an unmatched end element",
                    ));
                }
                depth -= 1;
            },
            Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => {
                return Err(unsupported_transition());
            },
            Event::Comment(_) | Event::CData(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(unsupported_transition());
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is forbidden in transition XML")),
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("transition subtree is not closed"));
    }
    Ok(())
}

fn validate_transition_element(
    resolved: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    depth: usize,
    effect_seen: &mut bool,
    prefix: &str,
    namespace: &str,
    allow_extensions: bool,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<()> {
    let extension_attributes = if allow_extensions && depth == 2 {
        extension_effect_attributes(resolved, element.name())
    } else {
        None
    };
    if extension_attributes.is_none()
        && !is_expected_subtree_namespace(resolved, element.name(), prefix, namespace)
    {
        return Err(unsupported_transition());
    }
    let local = element.local_name();
    let allowed = if depth == 1 && local.as_ref() == b"transition" {
        &[
            b"spd".as_slice(),
            b"advClick".as_slice(),
            b"advTm".as_slice(),
        ][..]
    } else if let Some(attributes) = extension_attributes {
        if *effect_seen {
            return Err(invalid("transition contains more than one visual effect"));
        }
        *effect_seen = true;
        attributes
    } else if depth == 2 && is_standard_effect(local.as_ref()) {
        if *effect_seen {
            return Err(invalid("transition contains more than one visual effect"));
        }
        *effect_seen = true;
        effect_attributes(local.as_ref())
    } else {
        return Err(unsupported_transition());
    };
    let mut seen = Vec::<Vec<u8>>::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() == b"xmlns"
            || matches!(attribute.key.prefix(), Some(prefix) if prefix.as_ref() == b"xmlns")
        {
            continue;
        }
        let extension_duration = allow_extensions
            && depth == 1
            && attribute.key.local_name().as_ref() == b"dur"
            && matches!(
                resolver.resolve_attribute(attribute.key).0,
                ResolveResult::Bound(Namespace(value)) if value == b"http://schemas.microsoft.com/office/powerpoint/2010/main".as_slice()
            );
        if !extension_duration
            && (attribute.key.prefix().is_some()
                || !allowed
                    .iter()
                    .any(|name| *name == attribute.key.local_name().as_ref()))
        {
            return Err(unsupported_transition());
        }
        if extension_duration
            || attribute.key.prefix().is_none()
                && allowed
                    .iter()
                    .any(|name| *name == attribute.key.local_name().as_ref())
        {
            let local = attribute.key.local_name().as_ref().to_vec();
            if seen.iter().any(|previous| previous.as_slice() == local) {
                return Err(invalid("transition contains a duplicate typed attribute"));
            }
            seen.push(local);
        }
    }
    Ok(())
}

fn extension_effect_attributes(
    resolved: &ResolveResult<'_>,
    name: QName<'_>,
) -> Option<&'static [&'static [u8]]> {
    let local = name.local_name();
    match resolved {
        ResolveResult::Bound(Namespace(value))
            if *value == b"http://schemas.microsoft.com/office/powerpoint/2010/main" =>
        {
            match local.as_ref() {
                b"ripple" | b"conveyor" | b"ferris" | b"flip" | b"gallery" | b"switch"
                | b"doors" | b"window" | b"pan" | b"vortex" | b"warp" => Some(&[b"dir"]),
                b"flash" | b"honeycomb" => Some(&[]),
                b"flythrough" => Some(&[b"dir", b"hasBounce"]),
                b"glitter" => Some(&[b"dir", b"pattern"]),
                b"prism" => Some(&[b"dir", b"isContent", b"isInverted"]),
                b"reveal" => Some(&[b"thruBlk", b"dir"]),
                b"shred" => Some(&[b"pattern", b"dir"]),
                b"wheelReverse" => Some(&[b"spokes"]),
                _ => None,
            }
        },
        ResolveResult::Bound(Namespace(value))
            if *value == b"http://schemas.microsoft.com/office/powerpoint/2012/main"
                && local.as_ref() == b"prstTrans" =>
        {
            Some(&[b"prst", b"invX", b"invY"])
        },
        ResolveResult::Bound(Namespace(value))
            if *value == b"http://schemas.microsoft.com/office/powerpoint/2015/09/main"
                && local.as_ref() == b"morph" =>
        {
            Some(&[b"option"])
        },
        _ => None,
    }
}

fn is_expected_subtree_namespace(
    resolved: &ResolveResult<'_>,
    name: QName<'_>,
    prefix: &str,
    namespace: &str,
) -> bool {
    match resolved {
        ResolveResult::Bound(Namespace(value)) => *value == namespace.as_bytes(),
        ResolveResult::Unknown(value) => {
            !prefix.is_empty() && value.as_slice() == prefix.as_bytes()
        },
        ResolveResult::Unbound => prefix.is_empty() && name.prefix().is_none(),
    }
}

fn is_standard_effect(local: &[u8]) -> bool {
    matches!(
        local,
        b"cut"
            | b"fade"
            | b"push"
            | b"wipe"
            | b"split"
            | b"pull"
            | b"cover"
            | b"dissolve"
            | b"blinds"
            | b"checker"
            | b"randomBar"
            | b"circle"
            | b"diamond"
            | b"plus"
            | b"wedge"
            | b"zoom"
            | b"random"
            | b"wheel"
            | b"newsflash"
            | b"strips"
            | b"comb"
    )
}

fn effect_attributes(local: &[u8]) -> &'static [&'static [u8]] {
    match local {
        b"cut" | b"fade" => &[b"thruBlk"],
        b"push" | b"wipe" | b"pull" | b"cover" | b"blinds" | b"checker" | b"randomBar"
        | b"zoom" | b"strips" | b"comb" => &[b"dir"],
        b"split" => &[b"orient", b"dir"],
        b"wheel" => &[b"spokes"],
        _ => &[],
    }
}

fn presentation_is_protected(xml: &[u8]) -> Result<bool> {
    let text = std::str::from_utf8(xml)
        .map_err(|error| Error::Xml(format!("presentation XML is not UTF-8: {error}")))?;
    Ok(
        crate::presentation_properties::metadata::protection::Settings::parse_xml(text)?
            .is_protected(),
    )
}

fn unsupported_transition() -> Error {
    Error::UnsafeEdit {
        operation: "edit_transition",
        reason: "source-backed transition edits refuse extension, sound-action, or unknown transition markup",
    }
}

fn is_pml(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local && is_pml_namespace(namespace)
}

fn is_pml_element(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    is_pml(namespace, name, local)
        || (name.local_name().as_ref() == local
            && matches!(
                namespace,
                ResolveResult::Unknown(prefix) if prefix.as_slice() == b"p"
            ))
        || (name.local_name().as_ref() == local
            && matches!(namespace, ResolveResult::Unbound)
            && name.prefix().is_none())
}

fn is_markup_element(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
}

fn is_pml_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == PML || *value == STRICT_PML)
}

fn pml_namespace(namespace: &ResolveResult<'_>) -> Result<String> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == PML || *value == STRICT_PML => {
            std::str::from_utf8(value)
                .map(str::to_owned)
                .map_err(|error| Error::Xml(format!("slide namespace is not UTF-8: {error}")))
        },
        _ => Err(invalid("slide root has no PresentationML namespace")),
    }
}

fn qname_prefix(name: QName<'_>) -> Result<String> {
    let raw = name.as_ref();
    let prefix = raw
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |colon| &raw[..colon]);
    std::str::from_utf8(prefix)
        .map(str::to_owned)
        .map_err(|error| Error::Xml(format!("slide root prefix is not UTF-8: {error}")))
}

fn add_node(nodes: usize) -> Result<usize> {
    let nodes = nodes.checked_add(1).ok_or(Error::Limit {
        resource: "source-backed transition XML nodes",
        limit: MAX_NODES,
    })?;
    if nodes > MAX_NODES {
        Err(Error::Limit {
            resource: "source-backed transition XML nodes",
            limit: MAX_NODES,
        })
    } else {
        Ok(nodes)
    }
}

fn enter_depth(depth: usize) -> Result<usize> {
    let depth = depth.checked_add(1).ok_or(Error::Limit {
        resource: "source-backed transition XML depth",
        limit: MAX_DEPTH,
    })?;
    if depth > MAX_DEPTH {
        Err(Error::Limit {
            resource: "source-backed transition XML depth",
            limit: MAX_DEPTH,
        })
    } else {
        Ok(depth)
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| invalid("source-backed transition XML position exceeds usize"))
}

fn invalid(message: &str) -> Error {
    Error::Invalid(format!("source-backed slide transition: {message}"))
}
