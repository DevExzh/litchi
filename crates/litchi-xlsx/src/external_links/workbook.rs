//! Ordered workbook reference ownership and conservative index-removal checks.
use std::collections::HashSet;
use std::ops::Range;

use litchi_core::edit::RefusalKind;
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::{OpcPackage, Part, SourceLimits};
use quick_xml::{
    XmlVersion,
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use super::{Entry, Link, invalid};
use crate::error::{Error, Result};

const SML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_REFERENCES: usize = 4096;
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_EVENTS: usize = 1_000_000;

struct Layout {
    ids: Vec<String>,
    owner: Option<Range<usize>>,
    owner_open_end: usize,
    owner_close_start: usize,
    entries: Vec<Range<usize>>,
    insertion: usize,
    root_open: Range<usize>,
    root_empty: bool,
    prefix: String,
    relationship_namespace: &'static str,
    opaque: bool,
}

fn is_sml(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value==SML || *value==STRICT)
}
fn refused(reason: &'static str) -> Error {
    Error::refused(RefusalKind::UnsupportedContent, "external_links", reason)
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    invalid(format!("external-reference XML: {error}"))
}

fn admit_workbook(workbook: &dyn Part, limits: &SourceLimits) -> Result<()> {
    let maximum = limits.max_part_bytes().min(MAX_XML_BYTES as u64) as usize;
    let size = workbook.payload_len()?;
    if size > maximum {
        return Err(Error::limit(super::LimitKind::XmlBytes, size, maximum));
    }
    Ok(())
}

fn layout(xml: &[u8], limits: &SourceLimits) -> Result<Layout> {
    let maximum = limits.max_part_bytes().min(MAX_XML_BYTES as u64) as usize;
    if xml.len() > maximum {
        return Err(Error::limit(super::LimitKind::XmlBytes, xml.len(), maximum));
    }
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    let mut events = 0usize;
    let mut owner_start = None;
    let mut entry_start = None;
    let mut result = Layout {
        ids: vec![],
        owner: None,
        owner_open_end: 0,
        owner_close_start: 0,
        entries: vec![],
        insertion: usize::MAX,
        root_open: 0..0,
        root_empty: false,
        prefix: String::new(),
        relationship_namespace: REL,
        opaque: false,
    };
    let mut root_seen = false;
    let mut root_closed = false;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("workbook XML offset overflow"))?;
        let event = reader.read_event().map_err(xml_error)?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("workbook XML offset overflow"))?;
        events += 1;
        if events > MAX_EVENTS {
            return Err(Error::limit(super::LimitKind::Events, events, MAX_EVENTS));
        }
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = matches!(event, Event::Empty(_));
                let level = depth + 1;
                if level > 128 {
                    return Err(Error::limit(super::LimitKind::Depth, level, 128));
                }
                let (ns, local) = reader.resolver().resolve_element(element.name());
                if level == 1 {
                    if root_seen || root_closed || !is_sml(&ns) || local.as_ref() != b"workbook" {
                        return Err(invalid("external links require one workbook root"));
                    }
                    root_seen = true;
                    result.root_open = start..end;
                    result.root_empty = empty;
                    result.relationship_namespace = if matches!(ns,ResolveResult::Bound(Namespace(value)) if value==STRICT)
                    {
                        STRICT_REL
                    } else {
                        REL
                    };
                    if let Some(prefix) = element.name().prefix() {
                        result.prefix = std::str::from_utf8(prefix.into_inner())
                            .map_err(xml_error)?
                            .to_owned();
                    }
                    if empty {
                        root_closed = true;
                        result.insertion = end;
                    }
                }
                if is_sml(&ns) && local.as_ref() == b"externalReferences" {
                    if level != 2 {
                        return Err(refused("externalReferences is not a direct workbook child"));
                    }
                    if result.owner.is_some() || owner_start.is_some() {
                        return Err(invalid("duplicate workbook externalReferences"));
                    }
                    if empty {
                        return Err(invalid("workbook externalReferences cannot be empty"));
                    }
                    owner_start = Some(start);
                    result.owner_open_end = end;
                    for a in element.checked_attributes() {
                        let a = a.map_err(xml_error)?;
                        if a.key.as_ref() != b"xmlns" && !a.key.as_ref().starts_with(b"xmlns:") {
                            result.opaque = true;
                        }
                    }
                } else if owner_start.is_some() {
                    if level == 3 && is_sml(&ns) && local.as_ref() == b"externalReference" {
                        let mut id = None;
                        for a in element.checked_attributes() {
                            let a = a.map_err(xml_error)?;
                            if a.key.as_ref() == b"xmlns" || a.key.as_ref().starts_with(b"xmlns:") {
                                continue;
                            }
                            let (ns, local) = reader.resolver().resolve_attribute(a.key);
                            if local.as_ref() == b"id"
                                && matches!(ns,ResolveResult::Bound(Namespace(n)) if n==REL.as_bytes()||n==STRICT_REL.as_bytes())
                            {
                                if id.is_some() {
                                    return Err(invalid(
                                        "duplicate externalReference relationship attribute",
                                    ));
                                }
                                id = Some(
                                    a.decoded_and_normalized_value(
                                        XmlVersion::Explicit1_0,
                                        reader.decoder(),
                                    )
                                    .map_err(xml_error)?
                                    .into_owned(),
                                );
                            } else {
                                result.opaque = true;
                            }
                        }
                        let id = id
                            .filter(|id| !id.is_empty())
                            .ok_or_else(|| invalid("externalReference has no relationship ID"))?;
                        if result.ids.len() >= MAX_REFERENCES {
                            return Err(Error::limit(
                                super::LimitKind::Items,
                                result.ids.len() + 1,
                                MAX_REFERENCES,
                            ));
                        }
                        if result.ids.contains(&id) {
                            return Err(invalid("duplicate externalReference relationship ID"));
                        }
                        result.ids.push(id);
                        if empty {
                            result.entries.push(start..end);
                        } else {
                            entry_start = Some(start);
                        }
                    } else {
                        result.opaque = true;
                    }
                }
                if level == 2
                    && is_sml(&ns)
                    && result.insertion == usize::MAX
                    && matches!(
                        local.as_ref(),
                        b"definedNames"
                            | b"calcPr"
                            | b"oleSize"
                            | b"customWorkbookViews"
                            | b"pivotCaches"
                            | b"smartTagPr"
                            | b"smartTagTypes"
                            | b"webPublishing"
                            | b"fileRecoveryPr"
                            | b"webPublishObjects"
                            | b"extLst"
                    )
                {
                    result.insertion = start;
                }
                if !empty {
                    depth = level;
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("workbook XML depth underflow"));
                }
                if depth == 3
                    && let Some(begin) = entry_start.take()
                {
                    result.entries.push(begin..end);
                }
                if depth == 2
                    && let Some(begin) = owner_start.take()
                {
                    result.owner = Some(begin..end);
                    result.owner_close_start = start;
                }
                if depth == 1 {
                    root_closed = true;
                    if result.insertion == usize::MAX {
                        result.insertion = start;
                    }
                }
                depth -= 1;
            },
            Event::Text(text) => {
                if owner_start.is_some() && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    result.opaque = true;
                }
            },
            Event::Comment(_) | Event::PI(_) | Event::CData(_) | Event::GeneralRef(_) => {
                if owner_start.is_some() {
                    result.opaque = true;
                }
            },
            Event::DocType(_) => {
                return Err(refused("workbook DTD declarations are not supported"));
            },
            Event::Eof => break,
            Event::Decl(_) => {},
        }
    }
    if !root_seen || !root_closed || depth != 0 || owner_start.is_some() {
        return Err(invalid("incomplete workbook external-reference source"));
    }
    if result.owner.is_some() && result.ids.is_empty() {
        return Err(invalid("workbook externalReferences cannot be empty"));
    }
    Ok(result)
}

pub(super) fn ordered_ids(workbook: &dyn Part, limits: &SourceLimits) -> Result<Vec<String>> {
    admit_workbook(workbook, limits)?;
    let source = layout(workbook.blob(), limits)?;
    let mut owned = workbook
        .rels()
        .iter()
        .filter(|r| super::package::is_owner_relationship(r.reltype()))
        .map(|r| r.r_id().to_owned())
        .collect::<Vec<_>>();
    if source.owner.is_none() {
        owned.sort();
        return Ok(owned);
    }
    let declared: HashSet<_> = source.ids.iter().map(String::as_str).collect();
    if source.ids.iter().any(|id| !owned.contains(id)) {
        return Err(invalid(
            "workbook externalReferences and external-link relationships disagree",
        ));
    }
    if owned.len() != declared.len() {
        return Err(refused(
            "externalReferences does not bind every external-link owner",
        ));
    }
    Ok(source.ids)
}

pub(super) fn update(package: &mut OpcPackage, before: &[Entry], after: &[Entry]) -> Result<()> {
    let old = before
        .iter()
        .map(|e| e.relationship_id.as_str())
        .collect::<Vec<_>>();
    let new = after
        .iter()
        .map(|e| e.relationship_id.as_str())
        .collect::<Vec<_>>();
    if old == new {
        return Ok(());
    }
    let limits = package.read_limits();
    let workbook = package.main_document_part()?;
    admit_workbook(workbook, &limits)?;
    let name = workbook.partname().clone();
    let source = workbook.blob();
    let parsed = layout(source, &limits)?;
    if parsed.opaque {
        return Err(refused(
            "externalReferences contains opaque content that cannot be rewritten",
        ));
    }
    if new.len() > MAX_REFERENCES {
        return Err(Error::limit(
            super::LimitKind::Items,
            new.len(),
            MAX_REFERENCES,
        ));
    }
    let mut affected = old
        .iter()
        .enumerate()
        .filter_map(|(i, id)| (new.get(i) != Some(id)).then_some(i + 1))
        .collect::<HashSet<_>>();
    // A relationship-only legacy catalog has no declared formula index
    // mapping. Do not silently bind pre-existing numeric/topic consumers.
    if parsed.owner.is_none() {
        affected.extend(1..=old.len());
    }
    if !affected.is_empty() {
        check_consumers(package, before, &affected)?;
    }
    let qualify = |local: &str| {
        if parsed.prefix.is_empty() {
            local.to_owned()
        } else {
            format!("{}:{local}", parsed.prefix)
        }
    };
    let maximum = limits.max_part_bytes().min(MAX_XML_BYTES as u64) as usize;
    let mut replacement = Vec::new();
    let mut append = |bytes: &[u8]| -> Result<()> {
        let size = replacement
            .len()
            .checked_add(bytes.len())
            .ok_or_else(|| invalid("external-reference replacement size overflow"))?;
        if size > maximum {
            return Err(Error::limit(super::LimitKind::XmlBytes, size, maximum));
        }
        replacement
            .try_reserve_exact(bytes.len())
            .map_err(|source| Error::Allocation {
                resource: "external-reference replacement",
                source,
            })?;
        replacement.extend_from_slice(bytes);
        Ok(())
    };
    if !new.is_empty() {
        let parent = qualify("externalReferences");
        let child = qualify("externalReference");
        if let Some(span) = &parsed.owner {
            append(&source[span.start..parsed.owner_open_end])?;
        } else {
            append(format!("<{parent}>").as_bytes())?;
        }
        for id in new {
            if let Some(index) = parsed.ids.iter().position(|old| old == id) {
                append(&source[parsed.entries[index].clone()])?;
            } else {
                append(
                    format!(
                        "<{child} xmlns:r=\"{}\" r:id=\"{}\"/>",
                        parsed.relationship_namespace,
                        litchi_core::xml::escape_xml(id)
                    )
                    .as_bytes(),
                )?;
            }
        }
        if let Some(span) = &parsed.owner {
            append(&source[parsed.owner_close_start..span.end])?;
        } else {
            append(format!("</{parent}>").as_bytes())?;
        }
    }
    let (removed, inserted) = if let Some(span) = &parsed.owner {
        (span.len(), replacement.len())
    } else if parsed.root_empty {
        let added = qualify("workbook")
            .len()
            .checked_add(2)
            .and_then(|n| n.checked_add(replacement.len()))
            .ok_or_else(|| invalid("workbook external-reference output size overflow"))?;
        (0, added)
    } else {
        (0, replacement.len())
    };
    let output_len = source
        .len()
        .checked_sub(removed)
        .and_then(|n| n.checked_add(inserted))
        .ok_or_else(|| invalid("workbook external-reference output size overflow"))?;
    if output_len > maximum {
        return Err(Error::limit(
            super::LimitKind::XmlBytes,
            output_len,
            maximum,
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "external-reference workbook output",
            source,
        })?;
    if let Some(span) = parsed.owner {
        output.extend_from_slice(&source[..span.start]);
        output.extend_from_slice(&replacement);
        output.extend_from_slice(&source[span.end..]);
    } else if parsed.root_empty {
        if !source[parsed.root_open.clone()].ends_with(b"/>") {
            return Err(invalid("empty workbook root has invalid framing"));
        }
        output.extend_from_slice(&source[..parsed.root_open.end - 2]);
        output.push(b'>');
        output.extend_from_slice(&replacement);
        output.extend_from_slice(format!("</{}>", qualify("workbook")).as_bytes());
        output.extend_from_slice(&source[parsed.root_open.end..]);
    } else {
        output.extend_from_slice(&source[..parsed.insertion]);
        output.extend_from_slice(&replacement);
        output.extend_from_slice(&source[parsed.insertion..]);
    }
    debug_assert_eq!(output.len(), output_len);
    package.get_part_mut(&name)?.set_blob(output);
    Ok(())
}

fn check_consumers(
    package: &OpcPackage,
    entries: &[Entry],
    affected: &HashSet<usize>,
) -> Result<()> {
    let owners = entries
        .iter()
        .map(|e| e.part_uri.as_str())
        .collect::<HashSet<_>>();
    let dde_or_ole = entries
        .iter()
        .enumerate()
        .any(|(i, e)| affected.contains(&(i + 1)) && !matches!(e.link, Link::Workbook(_)));
    let mut total = 0usize;
    for part in package.iter_parts().filter(|p| {
        !owners.contains(p.partname().as_str())
            && (p.content_type().ends_with("+xml")
                || p.content_type() == "application/xml"
                || p.content_type() == litchi_opc::constants::content_type::OFC_VML_DRAWING)
    }) {
        total = total
            .checked_add(part.payload_len()?)
            .ok_or_else(|| invalid("external-reference consumer size overflow"))?;
        if total > MAX_XML_BYTES {
            return Err(Error::limit(
                super::LimitKind::XmlBytes,
                total,
                MAX_XML_BYTES,
            ));
        }
        let payload = package.get_part(part.partname())?;
        let mut r = quick_xml::Reader::from_reader(payload.blob());
        let mut depth = 0usize;
        let mut formula = None;
        let mut text = String::new();
        let mut events = 0usize;
        loop {
            events += 1;
            if events > MAX_EVENTS {
                return Err(Error::limit(super::LimitKind::Events, events, MAX_EVENTS));
            }
            match r.read_event().map_err(xml_error)? {
                Event::Start(e) => {
                    depth += 1;
                    if depth > 128 {
                        return Err(Error::limit(super::LimitKind::Depth, depth, 128));
                    }
                    if formula.is_none()
                        && matches!(
                            e.local_name().as_ref(),
                            b"f" | b"definedName"
                                | b"formula"
                                | b"formula1"
                                | b"formula2"
                                | b"calculatedColumnFormula"
                                | b"totalsRowFormula"
                                | b"FmlaLink"
                                | b"FmlaRange"
                                | b"FmlaMacro"
                        )
                    {
                        formula = Some(depth);
                        text.clear();
                    }
                },
                Event::End(_) => {
                    if formula == Some(depth) {
                        if mentions_index(&text, affected) || (dde_or_ole && text.contains('|')) {
                            return Err(refused(
                                "external-link removal or reordering requires formula reference remapping",
                            ));
                        }
                        formula = None;
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("consumer XML depth underflow"))?;
                },
                Event::Text(t) if formula.is_some() => {
                    text.push_str(&t.decode().map_err(xml_error)?)
                },
                Event::CData(t) if formula.is_some() => {
                    text.push_str(&t.decode().map_err(xml_error)?)
                },
                Event::GeneralRef(t) if formula.is_some() => {
                    if let Some(c) = t.resolve_char_ref().map_err(xml_error)? {
                        text.push(c);
                    } else {
                        let c = match t.as_ref() {
                            b"amp" => '&',
                            b"lt" => '<',
                            b"gt" => '>',
                            b"apos" => '\'',
                            b"quot" => '"',
                            _ => return Err(refused("formula contains an unresolved XML entity")),
                        };
                        text.push(c);
                    }
                },
                Event::DocType(_) => return Err(refused("external-reference consumer uses a DTD")),
                Event::Eof => {
                    if depth != 0 || formula.is_some() {
                        return Err(refused("external-reference consumer XML is incomplete"));
                    }
                    break;
                },
                _ => {},
            }
        }
    }
    Ok(())
}
fn mentions_index(text: &str, affected: &HashSet<usize>) -> bool {
    let bytes = text.as_bytes();
    let mut at = 0;
    while at < bytes.len() {
        if bytes[at] != b'[' {
            at += 1;
            continue;
        }
        at += 1;
        let start = at;
        let mut n = Some(0usize);
        while at < bytes.len() && bytes[at].is_ascii_digit() {
            n = n
                .and_then(|v| v.checked_mul(10))
                .and_then(|v| v.checked_add(usize::from(bytes[at] - b'0')));
            at += 1;
        }
        if at > start && bytes.get(at) == Some(&b']') && n.is_some_and(|n| affected.contains(&n)) {
            return true;
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::external_links::{Conformance, Dde, Snapshot};
    use litchi_opc::{BlobPart, PackURI};

    fn package(xml: &str, ids: &[&str], conformance: Conformance) -> OpcPackage {
        let mut p = OpcPackage::new();
        let main = PackURI::new("/xl/workbook.xml").unwrap();
        p.add_part(Box::new(BlobPart::new(
            main.clone(),
            litchi_opc::constants::content_type::SML_SHEET_MAIN.into(),
            xml.as_bytes().to_vec(),
        )));
        p.relate_to(
            "xl/workbook.xml",
            litchi_opc::constants::relationship_type::OFFICE_DOCUMENT,
        );
        for (i, id) in ids.iter().enumerate() {
            let uri = PackURI::new(format!("/xl/externalLinks/link{i}.xml")).unwrap();
            let part = super::super::package::build_external_link_part_with_conformance(
                uri.clone(),
                &Link::Dde(Dde {
                    service: format!("Service{i}"),
                    topic: "Topic".into(),
                    items: vec![],
                }),
                conformance,
                &SourceLimits::default(),
            )
            .unwrap();
            p.add_part(Box::new(part));
            p.get_part_mut(&main).unwrap().rels_mut().add_relationship(
                conformance.external_link_relationship().into(),
                uri.relative_ref(main.base_uri()),
                (*id).into(),
                false,
            );
        }
        p
    }
    fn extra() -> Entry {
        Entry {
            index: 1,
            relationship_id: "rId12".into(),
            part_uri: PackURI::new("/xl/externalLinks/added.xml").unwrap(),
            link: Link::Dde(Dde {
                service: "Added".into(),
                topic: "T".into(),
                items: vec![],
            }),
        }
    }

    #[test]
    fn reference_order_uses_xml_in_both_dialects_with_aliased_relationship_attributes() {
        for (ns, rel, mode) in [
            (
                std::str::from_utf8(SML).unwrap(),
                REL,
                Conformance::Transitional,
            ),
            (
                std::str::from_utf8(STRICT).unwrap(),
                STRICT_REL,
                Conformance::Strict,
            ),
        ] {
            let xml = format!(
                r#"<s:workbook xmlns:s="{ns}" xmlns:z="{rel}"><s:externalReferences><s:externalReference z:id="rId2"/><s:externalReference z:id="rId12"/></s:externalReferences></s:workbook>"#
            );
            let p = package(&xml, &["rId12", "rId2"], mode);
            let s = Snapshot::load(&p, &SourceLimits::default()).unwrap();
            assert_eq!(
                s.entries()
                    .iter()
                    .map(|e| (e.index, e.relationship_id.as_str()))
                    .collect::<Vec<_>>(),
                [(0, "rId2"), (1, "rId12")]
            );
        }
    }
    #[test]
    fn append_and_remove_preserve_existing_reference_bytes_bom_and_opaque_workbook_children() {
        let xml = format!(
            "\u{feff}<?xml version=\"1.0\"?>\n<s:workbook xmlns:s=\"{}\" xmlns:z=\"{REL}\" xmlns:x=\"urn:future\" x:keep='yes'><s:sheets/><s:externalReferences><s:externalReference z:id='rId2'></s:externalReference></s:externalReferences><s:extLst><x:future/></s:extLst></s:workbook>",
            std::str::from_utf8(SML).unwrap()
        );
        let mut p = package(&xml, &["rId2"], Conformance::Transitional);
        let before =
            super::super::package::load_external_links(&p, &SourceLimits::default()).unwrap();
        let mut after = before.clone();
        after.push(extra());
        super::super::package::apply_entries(&mut p, &before, &after, Conformance::Transitional)
            .unwrap();
        let out = p.main_document_part().unwrap().blob();
        assert!(out.starts_with(b"\xef\xbb\xbf<?xml"));
        assert!(
            std::str::from_utf8(out)
                .unwrap()
                .contains("<s:externalReference z:id='rId2'></s:externalReference>")
        );
        assert!(
            std::str::from_utf8(out)
                .unwrap()
                .contains("<s:extLst><x:future/></s:extLst>")
        );
        super::super::package::apply_entries(&mut p, &after, &before, Conformance::Transitional)
            .unwrap();
        assert_eq!(p.main_document_part().unwrap().blob(), xml.as_bytes());
    }
    #[test]
    fn mismatched_duplicate_and_nondirect_reference_owners_are_rejected() {
        for body in [
            r#"<externalReferences><externalReference r:id="missing"/></externalReferences>"#,
            r#"<externalReferences><externalReference r:id="rId2"/><externalReference r:id="rId2"/></externalReferences>"#,
            r#"<extLst><externalReferences><externalReference r:id="rId2"/></externalReferences></extLst>"#,
        ] {
            let xml = format!(
                r#"<workbook xmlns="{}" xmlns:r="{REL}">{body}</workbook>"#,
                std::str::from_utf8(SML).unwrap()
            );
            let p = package(&xml, &["rId2"], Conformance::Transitional);
            assert!(Snapshot::load(&p, &SourceLimits::default()).is_err());
        }
    }
    #[test]
    fn opaque_reference_children_refuse_structural_changes_atomically() {
        let xml = format!(
            r#"<workbook xmlns="{}" xmlns:r="{REL}" xmlns:x="urn:future"><externalReferences><externalReference r:id="rId2"/><x:future/></externalReferences></workbook>"#,
            std::str::from_utf8(SML).unwrap()
        );
        let mut p = package(&xml, &["rId2"], Conformance::Transitional);
        let before =
            super::super::package::load_external_links(&p, &SourceLimits::default()).unwrap();
        let mut after = before.clone();
        after.push(extra());
        assert!(matches!(
            super::super::package::apply_entries(
                &mut p,
                &before,
                &after,
                Conformance::Transitional
            ),
            Err(Error::Refused(_))
        ));
        assert_eq!(p.main_document_part().unwrap().blob(), xml.as_bytes());
        assert_eq!(
            super::super::package::load_external_links(&p, &SourceLimits::default()).unwrap(),
            before
        );
    }
    #[test]
    fn removal_checks_split_character_references_and_vml_formula_consumers() {
        for (formula, kind, invalid) in [
            ("<f>'[&#49;]Sheet1'!A1</f>", "application/xml", true),
            (
                "<FmlaLink>'[1]Sheet1'!A1</FmlaLink>",
                litchi_opc::constants::content_type::OFC_VML_DRAWING,
                true,
            ),
            ("<f>Service|Topic!Item</f>", "application/xml", true),
            ("<t>[1] is only text</t>", "application/xml", false),
        ] {
            let xml = format!(
                r#"<workbook xmlns="{}" xmlns:r="{REL}"><externalReferences><externalReference r:id="rId2"/></externalReferences></workbook>"#,
                std::str::from_utf8(SML).unwrap()
            );
            let mut p = package(&xml, &["rId2"], Conformance::Transitional);
            let before =
                super::super::package::load_external_links(&p, &SourceLimits::default()).unwrap();
            p.add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/consumer.xml").unwrap(),
                kind.into(),
                format!("<root>{formula}</root>").into_bytes(),
            )));
            let result = super::super::package::apply_entries(
                &mut p,
                &before,
                &[],
                Conformance::Transitional,
            );
            assert_eq!(result.is_err(), invalid);
            if invalid {
                assert_eq!(p.main_document_part().unwrap().blob(), xml.as_bytes());
            }
        }
    }
    #[test]
    fn legacy_unbound_formula_consumers_are_not_silently_assigned_catalog_slots() {
        let xml = format!(
            r#"<workbook xmlns="{}"/>"#,
            std::str::from_utf8(SML).unwrap()
        );
        let mut p = package(&xml, &["rId2"], Conformance::Transitional);
        p.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/consumer.xml").unwrap(),
            "application/xml".into(),
            b"<root><f>'[1]Sheet1'!A1</f></root>".to_vec(),
        )));
        let before =
            super::super::package::load_external_links(&p, &SourceLimits::default()).unwrap();
        let mut after = before.clone();
        after.push(extra());
        assert!(matches!(
            super::super::package::apply_entries(
                &mut p,
                &before,
                &after,
                Conformance::Transitional
            ),
            Err(Error::Refused(_))
        ));
        assert_eq!(p.main_document_part().unwrap().blob(), xml.as_bytes());
    }
    #[test]
    fn workbook_output_limit_is_checked_before_atomic_publication() {
        let xml = format!(
            r#"<workbook xmlns="{}" xmlns:r="{REL}"><!--{}--><externalReferences><externalReference r:id="rId2"/></externalReferences></workbook>"#,
            std::str::from_utf8(SML).unwrap(),
            "padding".repeat(400)
        );
        let p = package(&xml, &["rId2"], Conformance::Transitional);
        let mut bytes = Vec::new();
        p.write_to(&mut bytes).unwrap();
        let limits = SourceLimits::builder()
            .max_part_bytes(xml.len() as u64)
            .unwrap()
            .build()
            .unwrap();
        let mut p = OpcPackage::from_bytes(&bytes, limits).unwrap();
        let before = super::super::package::load_external_links(&p, &limits).unwrap();
        let mut after = before.clone();
        after.push(extra());
        assert!(
            matches!(super::super::package::apply_entries(&mut p,&before,&after,Conformance::Transitional),Err(Error::Limit{maximum,..}) if maximum==xml.len() as u64)
        );
        assert_eq!(p.main_document_part().unwrap().blob(), xml.as_bytes());
        assert_eq!(
            super::super::package::load_external_links(&p, &limits).unwrap(),
            before
        );
    }
}
