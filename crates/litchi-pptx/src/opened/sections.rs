//! Section membership of slides a transaction adds, moves or removes
//! (record 0252, ported from the M2 L-PPTX-W4 reference work; removal and
//! the source-backed editor, record 0303).
//!
//! `PowerPoint` 2010 sections live in an extension of the presentation
//! part: `p:extLst/p:ext[@uri="{521415D9-…}"]/p14:sectionLst/p14:section/
//! p14:sldIdLst/p14:sldId@id`. `PowerPoint` keeps every slide in one section
//! and each section's slides in presentation order, so a slide that is
//! added or moved joins the section of the slide just before it in the new
//! order, or, at the start, the section of the slide just after it; with
//! neither (no other slide is in a section) it joins the first section.
//! A reorder that is not one move keeps every slide in its section and
//! re-sorts each section's entries into the new order. A removed slide
//! leaves its section; the section stays even when it becomes empty, as
//! `PowerPoint` keeps an empty section.
//!
//! Edits splice `p14:sldId` elements only: every other byte of the
//! presentation part, the section names and identifiers included, stays as
//! written. A presentation whose section extension the scan cannot read
//! exactly (markup-compatibility content inside it, two section lists, a
//! section with two slide lists, or an unexpected child of a slide list) is
//! refused rather than guessed at.

use std::collections::HashMap;
use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::xml::{MAX_XML_DEPTH, bump, is_presentation_namespace, position, span_bytes};
use crate::{Error, Result, SlideOrderRefusal};

/// The `p:ext@uri` of the `PowerPoint` 2010 section list.
pub(crate) const SECTION_EXTENSION_URI: &str = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
const P14: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/2010/main";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

/// The section lists of one presentation part, with the byte spans an edit
/// splices.
#[derive(Debug)]
pub(crate) struct Sections {
    sections: Vec<Section>,
}

#[derive(Debug, Default)]
struct Section {
    list: Option<List>,
    members: Vec<Member>,
}

/// One `p14:sldIdLst`.
#[derive(Debug)]
struct List {
    /// Its start tag, or the whole element when it is written empty.
    tag: Range<usize>,
    empty: bool,
    /// Where its end tag starts; `tag.end` when it is written empty.
    end_start: usize,
    /// Its qualified name, as written.
    name: Vec<u8>,
}

/// One `p14:sldId`.
#[derive(Debug, Clone)]
struct Member {
    id: u32,
    span: Range<usize>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Role {
    Root,
    ExtensionList,
    SectionExtension,
    SectionList,
    Section(usize),
    SlideList(usize),
    Member,
    Other,
    /// Any element inside the section extension that the scan does not
    /// model (a section's own `p14:extLst`, for instance).
    InsideOther,
}

impl Sections {
    /// Whether the presentation has at least one section.
    pub(crate) fn is_empty(&self) -> bool {
        self.sections.is_empty()
    }

    /// Every slide identifier the sections name, stale ones included.
    pub(crate) fn ids(&self) -> Vec<u32> {
        self.sections
            .iter()
            .flat_map(|section| section.members.iter().map(|member| member.id))
            .collect()
    }

    /// The section holding `id`, with the entry's position in it.
    fn membership(&self) -> HashMap<u32, (usize, usize)> {
        let mut membership = HashMap::new();
        for (section, entries) in self.sections.iter().enumerate() {
            for (position, member) in entries.members.iter().enumerate() {
                membership.entry(member.id).or_insert((section, position));
            }
        }
        membership
    }
}

/// Scan the section lists of a presentation part. `None` when it has no
/// section extension.
pub(crate) fn scan(xml: &[u8]) -> Result<Option<Sections>> {
    // An element name is never written with character references, so a part
    // without these bytes has no section list, and most parts are not read.
    if !mentions(xml, b"sectionLst") {
        return Ok(None);
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let origin = ReaderOrigin::of(xml);
    let mut stack: Vec<Role> = Vec::new();
    let mut sections: Vec<Section> = Vec::new();
    let mut extensions = 0usize;
    let mut section_lists = 0usize;
    let mut nodes = 0usize;
    // The open `p14:sldId` start offset, for an entry written with an end tag.
    let mut open_member: Option<(u32, usize)> = None;
    loop {
        let start = position(&reader, origin)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let space = Space::of(&namespace);
        let end = position(&reader, origin)?;
        let (element, empty) = match &event {
            Event::Start(element) => (Some(element), false),
            Event::Empty(element) => (Some(element), true),
            _ => (None, false),
        };
        if let Some(element) = element {
            bump(&mut nodes)?;
            if stack.len() >= MAX_XML_DEPTH {
                return Err(Error::Limit {
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
                    observed: stack.len().checked_add(1).unwrap_or(usize::MAX),
                    maximum: MAX_XML_DEPTH,
                });
            }
            let parent = stack.last().copied();
            let mut role = classify(parent, space, element, &reader)?;
            if matches!(role, Role::Section(_)) {
                role = Role::Section(sections.len());
                sections.push(Section::default());
            }
            match role {
                Role::SectionExtension => {
                    extensions += 1;
                    if extensions > 1 {
                        return Err(refusal("the presentation has two section extensions"));
                    }
                },
                Role::SectionList => {
                    section_lists += 1;
                    if section_lists > 1 {
                        return Err(refusal("the presentation has two section lists"));
                    }
                },
                Role::SlideList(index) => {
                    let section = sections
                        .get_mut(index)
                        .ok_or_else(|| refusal("a slide list lost its section"))?;
                    if section.list.is_some() {
                        return Err(refusal("a section has two slide lists"));
                    }
                    section.list = Some(List {
                        tag: start..end,
                        empty,
                        end_start: end,
                        name: element.name().as_ref().to_vec(),
                    });
                },
                Role::Member => {
                    let id = member_id(element, &reader)?;
                    if empty {
                        push_member(&mut sections, &stack, id, start..end)?;
                    } else {
                        open_member = Some((id, start));
                    }
                },
                _ => {},
            }
            if !empty {
                stack.push(role);
            }
            continue;
        }
        match event {
            Event::End(_) => {
                let role = stack
                    .pop()
                    .ok_or_else(|| Error::Xml("unbalanced presentation XML".to_owned()))?;
                match role {
                    Role::SlideList(index) => {
                        if let Some(list) = sections
                            .get_mut(index)
                            .and_then(|section| section.list.as_mut())
                        {
                            list.end_start = start;
                        }
                    },
                    Role::Member => {
                        let (id, member_start) = open_member
                            .take()
                            .ok_or_else(|| refusal("a section entry closed without opening"))?;
                        push_member(&mut sections, &stack, id, member_start..end)?;
                    },
                    _ => {},
                }
            },
            Event::Text(text) => {
                if matches!(stack.last(), Some(Role::SlideList(_)))
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(refusal("a section slide list contains text"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_)
                if matches!(stack.last(), Some(Role::SlideList(_))) =>
            {
                return Err(refusal("a section slide list contains character data"));
            },
            Event::DocType(_) => return Err(Error::Invalid("DTDs are rejected".to_owned())),
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(Error::Xml("unterminated presentation XML".to_owned()));
    }
    Ok((extensions > 0).then_some(Sections { sections }))
}

/// The namespaces the scan tells apart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Space {
    Presentation,
    PowerPoint2010,
    MarkupCompatibility,
    Other,
}

impl Space {
    fn of(namespace: &ResolveResult<'_>) -> Self {
        if is_presentation_namespace(namespace) {
            return Self::Presentation;
        }
        match namespace {
            ResolveResult::Bound(Namespace(value)) if *value == P14 => Self::PowerPoint2010,
            ResolveResult::Bound(Namespace(value)) if *value == MCE => Self::MarkupCompatibility,
            _ => Self::Other,
        }
    }
}

fn classify(
    parent: Option<Role>,
    space: Space,
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<Role> {
    let local = element.local_name();
    let local = local.as_ref();
    let p14 = space == Space::PowerPoint2010;
    let mce = space == Space::MarkupCompatibility;
    let presentation = space == Space::Presentation;
    let inside_extension = matches!(
        parent,
        Some(
            Role::SectionExtension
                | Role::SectionList
                | Role::Section(_)
                | Role::SlideList(_)
                | Role::Member
                | Role::InsideOther
        )
    );
    if mce && (inside_extension || matches!(parent, Some(Role::Root | Role::ExtensionList))) {
        return Err(refusal(
            "markup-compatibility content surrounds the presentation's section list",
        ));
    }
    Ok(match parent {
        None => Role::Root,
        Some(Role::Root) if presentation && local == b"extLst" => Role::ExtensionList,
        Some(Role::ExtensionList) if presentation && local == b"ext" => {
            let uri = litchi_ooxml_common::xml::unqualified_attribute_value(
                element,
                b"uri",
                reader.decoder(),
            )?;
            if uri.as_deref() == Some(SECTION_EXTENSION_URI) {
                Role::SectionExtension
            } else {
                Role::Other
            }
        },
        Some(Role::SectionExtension) if p14 && local == b"sectionLst" => Role::SectionList,
        // The caller numbers the section.
        Some(Role::SectionList) if p14 && local == b"section" => Role::Section(0),
        Some(Role::Section(index)) if p14 && local == b"sldIdLst" => Role::SlideList(index),
        Some(Role::SlideList(_)) if p14 && local == b"sldId" => Role::Member,
        Some(Role::SlideList(_)) => {
            return Err(refusal("a section slide list has an unexpected child"));
        },
        Some(_) if inside_extension => Role::InsideOther,
        Some(_) => Role::Other,
    })
}

fn member_id(element: &BytesStart<'_>, reader: &NsReader<&[u8]>) -> Result<u32> {
    let value =
        litchi_ooxml_common::xml::unqualified_attribute_value(element, b"id", reader.decoder())?
            .ok_or_else(|| refusal("a section entry has no slide identifier"))?;
    value
        .trim()
        .parse::<u32>()
        .map_err(|_error| refusal("a section entry has an invalid slide identifier"))
}

fn push_member(
    sections: &mut [Section],
    stack: &[Role],
    id: u32,
    span: Range<usize>,
) -> Result<()> {
    let Some(Role::SlideList(index)) = stack.last().copied() else {
        return Err(refusal("a section entry is outside a slide list"));
    };
    sections
        .get_mut(index)
        .ok_or_else(|| refusal("a section entry lost its section"))?
        .members
        .push(Member { id, span });
    Ok(())
}

fn refusal(_detail: &str) -> Error {
    Error::slide_order_refused(SlideOrderRefusal::Sections)
}

/// Place the slide `order[position]` in a section, as `PowerPoint` does
/// (see the module documentation): it leaves the section that holds it and
/// joins the section of its neighbour. Returns the rewritten presentation
/// part, or `None` when the part has no section.
pub(crate) fn place(xml: &[u8], order: &[u32], position: usize) -> Result<Option<Vec<u8>>> {
    let Some(sections) = scan(xml)? else {
        return Ok(None);
    };
    if sections.is_empty() {
        return Ok(None);
    }
    let id = *order
        .get(position)
        .ok_or_else(|| Error::Invalid("section placement position is out of range".into()))?;
    let membership = sections.membership();
    // The slide leaves the section that holds it.
    let mut edits: Vec<(Range<usize>, Vec<u8>)> = sections
        .sections
        .iter()
        .flat_map(|list| &list.members)
        .filter(|member| member.id == id)
        .map(|member| (member.span.clone(), Vec::new()))
        .collect();
    if edits.len() > 1 {
        return Err(refusal("a slide belongs to more than one section"));
    }
    let previous = position
        .checked_sub(1)
        .and_then(|before| order.get(before))
        .and_then(|before| membership.get(before).map(|place| (*before, *place)));
    let next = order
        .get(position + 1)
        .and_then(|after| membership.get(after).map(|place| (*after, *place)));
    // The target section and the neighbour the entry goes next to.
    let (section, neighbour) = match (previous, next) {
        (Some((before, (section, entry))), _) => {
            let member = &sections.sections[section].members[entry];
            edits.push((
                member.span.end..member.span.end,
                entry_xml(&sections.sections[section], id)?,
            ));
            (section, Neighbour::After(before))
        },
        (None, Some((after, (section, entry)))) => {
            let member = &sections.sections[section].members[entry];
            edits.push((
                member.span.start..member.span.start,
                entry_xml(&sections.sections[section], id)?,
            ));
            (section, Neighbour::Before(after))
        },
        (None, None) => {
            edits.push(append(xml, &sections.sections[0], id)?);
            (0, Neighbour::Last)
        },
    };
    edits.sort_by_key(|(range, _)| (range.start, range.end));
    let mut output = Vec::with_capacity(xml.len() + 32);
    let mut cursor = 0;
    for (range, bytes) in &edits {
        if range.start < cursor {
            return Err(refusal("section edits overlap"));
        }
        output.extend_from_slice(span_bytes(xml, cursor..range.start)?);
        output.extend_from_slice(bytes);
        cursor = range.end;
    }
    output.extend_from_slice(span_bytes(xml, cursor..xml.len())?);
    // The entry must read back once, in its section, next to its neighbour.
    let check = scan(&output)?.ok_or_else(|| refusal("the section list disappeared"))?;
    let placed: Vec<u32> = check
        .sections
        .get(section)
        .map(|list| list.members.iter().map(|member| member.id).collect())
        .unwrap_or_default();
    let at = placed.iter().position(|member| *member == id);
    let read_back = check.ids().iter().filter(|member| **member == id).count() == 1
        && match (neighbour, at) {
            (Neighbour::After(before), Some(at)) => {
                at.checked_sub(1).and_then(|index| placed.get(index)) == Some(&before)
            },
            (Neighbour::Before(after), Some(at)) => placed.get(at + 1) == Some(&after),
            (Neighbour::Last, Some(at)) => at + 1 == placed.len(),
            (_, None) => false,
        };
    if !read_back {
        return Err(refusal("the section entry did not read back"));
    }
    Ok(Some(output))
}

/// Which neighbour a placed entry follows or precedes.
#[derive(Debug, Clone, Copy)]
enum Neighbour {
    After(u32),
    Before(u32),
    Last,
}

/// Whether `haystack` holds `needle`.
fn mentions(haystack: &[u8], needle: &[u8]) -> bool {
    memchr::memmem::find(haystack, needle).is_some()
}

/// Re-sort every section's entries into `order`, keeping each slide in its
/// section. Entries that name no slide of `order` keep their place.
/// Returns `None` when nothing changes.
pub(crate) fn sort(xml: &[u8], order: &[u32]) -> Result<Option<Vec<u8>>> {
    let Some(sections) = scan(xml)? else {
        return Ok(None);
    };
    let rank: HashMap<u32, usize> = order
        .iter()
        .enumerate()
        .map(|(position, id)| (*id, position))
        .collect();
    // Splices of whole slide lists' contents, in document order.
    let mut splices: Vec<(Range<usize>, Vec<u8>)> = Vec::new();
    for section in &sections.sections {
        let ranked: Vec<usize> = section
            .members
            .iter()
            .filter_map(|member| rank.get(&member.id).copied())
            .collect();
        if ranked.windows(2).all(|pair| pair[0] < pair[1]) {
            continue;
        }
        // The slots of ranked entries take the ranked entries in order;
        // other entries keep their slots.
        let mut sorted: Vec<&Member> = section
            .members
            .iter()
            .filter(|member| rank.contains_key(&member.id))
            .collect();
        sorted.sort_by_key(|member| rank.get(&member.id).copied().unwrap_or(usize::MAX));
        let mut sorted = sorted.into_iter();
        let (Some(first), Some(last)) = (section.members.first(), section.members.last()) else {
            continue;
        };
        let mut bytes = Vec::with_capacity(last.span.end - first.span.start);
        for (index, member) in section.members.iter().enumerate() {
            let chosen = if rank.contains_key(&member.id) {
                sorted
                    .next()
                    .ok_or_else(|| refusal("section entries changed while sorting"))?
            } else {
                member
            };
            bytes.extend_from_slice(span_bytes(xml, chosen.span.clone())?);
            if let Some(next) = section.members.get(index + 1) {
                bytes.extend_from_slice(span_bytes(xml, member.span.end..next.span.start)?);
            }
        }
        splices.push((first.span.start..last.span.end, bytes));
    }
    if splices.is_empty() {
        return Ok(None);
    }
    let mut output = Vec::with_capacity(xml.len());
    let mut cursor = 0;
    for (range, bytes) in splices {
        output.extend_from_slice(span_bytes(xml, cursor..range.start)?);
        output.extend_from_slice(&bytes);
        cursor = range.end;
    }
    output.extend_from_slice(span_bytes(xml, cursor..xml.len())?);
    Ok(Some(output))
}

/// Whether the part may have a section list. A part without these bytes
/// has none.
pub(crate) fn present(xml: &[u8]) -> bool {
    mentions(xml, b"sectionLst")
}

/// The spans of the section entries naming the slide `id`.
pub(crate) fn member_spans(xml: &[u8], id: u32) -> Result<Vec<Range<usize>>> {
    Ok(scan(xml)?
        .map(|sections| {
            sections
                .sections
                .iter()
                .flat_map(|section| &section.members)
                .filter(|member| member.id == id)
                .map(|member| member.span.clone())
                .collect()
        })
        .unwrap_or_default())
}

/// Remove the section entry of the slide `id`. Returns the rewritten part,
/// or `None` when no section names the slide.
pub(crate) fn remove(xml: &[u8], id: u32) -> Result<Option<Vec<u8>>> {
    let Some(sections) = scan(xml)? else {
        return Ok(None);
    };
    let mut spans = sections
        .sections
        .iter()
        .flat_map(|section| &section.members)
        .filter(|member| member.id == id)
        .map(|member| member.span.clone());
    let Some(span) = spans.next() else {
        return Ok(None);
    };
    if spans.next().is_some() {
        return Err(refusal("a slide belongs to more than one section"));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len() - span.len())
        .map_err(|source| Error::Allocation {
            resource: "presentation section removal XML",
            source,
        })?;
    output.extend_from_slice(span_bytes(xml, 0..span.start)?);
    output.extend_from_slice(span_bytes(xml, span.end..xml.len())?);
    let check = scan(&output)?.ok_or_else(|| refusal("the section list disappeared"))?;
    if check.ids().contains(&id) || check.sections.len() != sections.sections.len() {
        return Err(refusal("the section entry did not leave"));
    }
    Ok(Some(output))
}

/// The part with every section entry removed, and each section's entries
/// in document order. `None` when the part has no section list.
///
/// Two parts whose stripped forms are equal differ only in section
/// membership: entry removal keeps the whitespace between entries, so an
/// entry added, removed or re-sorted by [`place`], [`remove`] or [`sort`]
/// leaves the stripped bytes unchanged.
pub(crate) fn strip(xml: &[u8]) -> Result<Option<(Vec<u8>, Vec<Vec<u32>>)>> {
    let Some(sections) = scan(xml)? else {
        return Ok(None);
    };
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len())
        .map_err(|source| Error::Allocation {
            resource: "presentation section proof XML",
            source,
        })?;
    let mut cursor = 0;
    for member in sections
        .sections
        .iter()
        .flat_map(|section| &section.members)
    {
        if member.span.start < cursor {
            return Err(refusal("section entries overlap"));
        }
        output.extend_from_slice(span_bytes(xml, cursor..member.span.start)?);
        cursor = member.span.end;
    }
    output.extend_from_slice(span_bytes(xml, cursor..xml.len())?);
    let membership = sections
        .sections
        .iter()
        .map(|section| section.members.iter().map(|member| member.id).collect())
        .collect();
    Ok(Some((output, membership)))
}

/// Prove that section membership after a slide catalog change is the one an
/// order change, insertion or removal keeps (record 0303):
///
/// - the section count is unchanged and no slide is in two sections;
/// - every entry names a slide of the new order, or a stale identifier the
///   part already held;
/// - each section's entries follow the new presentation order;
/// - a slide that kept its place in the catalog and was in a section is
///   still in one.
pub(crate) fn validate_membership(
    before: &[Vec<u32>],
    after: &[Vec<u32>],
    before_order: &[u32],
    after_order: &[u32],
) -> Result<()> {
    if before.len() != after.len() {
        return Err(refusal("a catalog change adds or removes a section"));
    }
    let rank: HashMap<u32, usize> = after_order
        .iter()
        .enumerate()
        .map(|(position, id)| (*id, position))
        .collect();
    let before_members: std::collections::HashSet<u32> = before.iter().flatten().copied().collect();
    let mut seen = std::collections::HashSet::new();
    for section in after {
        let mut previous = None;
        for id in section {
            if !seen.insert(*id) {
                return Err(refusal("a slide belongs to more than one section"));
            }
            match rank.get(id) {
                Some(position) => {
                    if previous.is_some_and(|previous| previous > *position) {
                        return Err(refusal("a section's entries leave presentation order"));
                    }
                    previous = Some(*position);
                },
                None if before_members.contains(id) && !before_order.contains(id) => {},
                None => return Err(refusal("a section names a slide outside the catalog")),
            }
        }
    }
    for id in before.iter().flatten() {
        if rank.contains_key(id) && !seen.contains(id) {
            return Err(refusal("a retained slide left every section"));
        }
    }
    Ok(())
}

/// One `p14:sldId` element for `id`, named as `section`'s list names its
/// entries.
fn entry_xml(section: &Section, id: u32) -> Result<Vec<u8>> {
    let list = section
        .list
        .as_ref()
        .ok_or_else(|| refusal("the target section has no slide list"))?;
    let element = entry_name(&list.name)?;
    let mut fragment = Vec::with_capacity(element.len() + 24);
    fragment.push(b'<');
    fragment.extend_from_slice(&element);
    fragment.extend_from_slice(b" id=\"");
    fragment.extend_from_slice(id.to_string().as_bytes());
    fragment.extend_from_slice(b"\"/>");
    Ok(fragment)
}

/// The edit that appends an entry for `id` to `section`'s list: before its
/// end tag, or, for a list written empty (`<p14:sldIdLst/>`), the list
/// opened with every attribute of its start tag kept.
fn append(xml: &[u8], section: &Section, id: u32) -> Result<(Range<usize>, Vec<u8>)> {
    let fragment = entry_xml(section, id)?;
    let list = section
        .list
        .as_ref()
        .ok_or_else(|| refusal("the target section has no slide list"))?;
    if !list.empty {
        return Ok((list.end_start..list.end_start, fragment));
    }
    let open = span_bytes(xml, list.tag.clone())?
        .strip_suffix(b"/>")
        .ok_or_else(|| refusal("an empty section slide list is malformed"))?;
    let mut bytes = Vec::with_capacity(open.len() + fragment.len() + list.name.len() + 4);
    bytes.extend_from_slice(open.trim_ascii_end());
    bytes.push(b'>');
    bytes.extend_from_slice(&fragment);
    bytes.extend_from_slice(b"</");
    bytes.extend_from_slice(&list.name);
    bytes.push(b'>');
    Ok((list.tag.clone(), bytes))
}

/// `p14:sldIdLst` → `p14:sldId`, keeping the list's prefix.
fn entry_name(list: &[u8]) -> Result<Vec<u8>> {
    let prefix = list
        .strip_suffix(b"sldIdLst")
        .ok_or_else(|| refusal("a section slide list has an unexpected name"))?;
    let mut name = prefix.to_vec();
    name.extend_from_slice(b"sldId");
    Ok(name)
}

#[cfg(test)]
mod tests {
    use super::{place, scan, sort};

    const HEAD: &str = r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst/>"#;

    fn deck(sections: &str) -> String {
        format!(
            r#"{HEAD}<p:extLst><p:ext uri="{{521415D9-36F7-43E2-AB2F-B90AF26B5E84}}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">{sections}</p14:sectionLst></p:ext></p:extLst></p:presentation>"#
        )
    }

    fn section(name: &str, ids: &[u32]) -> String {
        if ids.is_empty() {
            return format!(
                r#"<p14:section name="{name}" id="{{0}}"><p14:sldIdLst/></p14:section>"#
            );
        }
        let entries: String = ids
            .iter()
            .map(|id| format!(r#"<p14:sldId id="{id}"/>"#))
            .collect();
        format!(
            r#"<p14:section name="{name}" id="{{0}}"><p14:sldIdLst>{entries}</p14:sldIdLst></p14:section>"#
        )
    }

    fn members(xml: &[u8]) -> Vec<Vec<u32>> {
        scan(xml)
            .unwrap()
            .unwrap()
            .sections
            .iter()
            .map(|section| section.members.iter().map(|member| member.id).collect())
            .collect()
    }

    #[test]
    fn a_new_slide_joins_the_section_of_the_slide_before_it() {
        let xml = deck(&(section("a", &[256, 257]) + &section("b", &[258])));
        // After 257 (the last of "a"): "a".
        let placed = place(xml.as_bytes(), &[256, 257, 300, 258], 2)
            .unwrap()
            .unwrap();
        assert_eq!(members(&placed), [vec![256, 257, 300], vec![258]]);
        // At the start: the section of the slide after it.
        let placed = place(xml.as_bytes(), &[300, 256, 257, 258], 0)
            .unwrap()
            .unwrap();
        assert_eq!(members(&placed), [vec![300, 256, 257], vec![258]]);
        // At the end: the last slide's section.
        let placed = place(xml.as_bytes(), &[256, 257, 258, 300], 3)
            .unwrap()
            .unwrap();
        assert_eq!(members(&placed), [vec![256, 257], vec![258, 300]]);
        // Every byte outside the new entry is kept.
        let inserted = r#"<p14:sldId id="300"/>"#;
        let at = String::from_utf8(placed).unwrap();
        assert_eq!(at.replacen(inserted, "", 1), xml);
    }

    #[test]
    fn empty_lists_and_decks_without_sectioned_slides_take_the_first_section() {
        let xml = deck(&(section("empty", &[]) + &section("also", &[])));
        let placed = place(xml.as_bytes(), &[256], 0).unwrap().unwrap();
        assert_eq!(members(&placed), [vec![256], vec![]]);
        let text = String::from_utf8(placed).unwrap();
        assert!(text.contains(r#"<p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst>"#));
    }

    #[test]
    fn a_moved_slide_leaves_its_section_and_joins_its_new_neighbour() {
        let xml = deck(&(section("a", &[256, 257]) + &section("b", &[258, 259])));
        // 256 moves to the end, after 259.
        let placed = place(xml.as_bytes(), &[257, 258, 259, 256], 3)
            .unwrap()
            .unwrap();
        assert_eq!(members(&placed), [vec![257], vec![258, 259, 256]]);
        // 259 moves to the start, before 256.
        let placed = place(xml.as_bytes(), &[259, 256, 257, 258], 0)
            .unwrap()
            .unwrap();
        assert_eq!(members(&placed), [vec![259, 256, 257], vec![258]]);
    }

    #[test]
    fn a_reorder_keeps_membership_and_sorts_each_section() {
        let xml = deck(&(section("a", &[256, 257, 258]) + &section("b", &[259])));
        let sorted = sort(xml.as_bytes(), &[258, 259, 256, 257])
            .unwrap()
            .unwrap();
        assert_eq!(members(&sorted), [vec![258, 256, 257], vec![259]]);
        assert!(
            sort(xml.as_bytes(), &[256, 257, 259, 258])
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn presentations_without_sections_are_left_alone() {
        let xml = format!("{HEAD}</p:presentation>");
        assert!(scan(xml.as_bytes()).unwrap().is_none());
        assert!(place(xml.as_bytes(), &[256], 0).unwrap().is_none());
        assert!(sort(xml.as_bytes(), &[256]).unwrap().is_none());
    }

    #[test]
    fn unreadable_section_lists_are_refused() {
        let twice =
            deck(&section("a", &[256])).replace("</p14:section>", "<p14:sldIdLst/></p14:section>");
        assert!(scan(twice.as_bytes()).is_err());
        let foreign = deck(&section("a", &[256])).replace(
            r#"<p14:sldId id="256"/>"#,
            r#"<p14:sldId id="256"/><x:y xmlns:x="urn:x"/>"#,
        );
        assert!(scan(foreign.as_bytes()).is_err());
    }
}
