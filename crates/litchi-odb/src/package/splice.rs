//! ODB adapter for provenance-bearing shared ODF XML publication.

use std::{collections::HashMap, ops::Range};

use litchi_core::{Error, Result};
use litchi_odf_common::core::{
    AuthoredXmlFragment, OwnedPackage, XmlSourcePart, XmlSplicePublication,
    rebuild_package_with_xml_splices,
};

const CONTENT_PATH: &str = "content.xml";
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

pub(super) fn rebuild_content(source: &OwnedPackage, content: &str) -> Result<Vec<u8>> {
    let part = XmlSourcePart::load(source, CONTENT_PATH)?;
    let exact_source = std::str::from_utf8(part.bytes()).map_err(|error| {
        Error::InvalidFormat(format!("ODB source content.xml is not UTF-8: {error}"))
    })?;
    let edits = differences(exact_source, content)?;
    let mut publication = XmlSplicePublication::new(part.clone());
    for edit in edits {
        let expected = part.bytes().get(edit.range.clone()).ok_or_else(|| {
            Error::InvalidFormat("ODB source splice range is invalid".to_string())
        })?;
        let proof = part.checked_range(edit.range, expected)?;
        let fragment = classify_fragment(edit.replacement)?;
        publication.replace(proof, fragment)?;
    }
    rebuild_package_with_xml_splices(source, vec![publication], MAX_OUTPUT_BYTES)
}

struct Edit {
    range: Range<usize>,
    replacement: Vec<u8>,
}

fn differences(source: &str, target: &str) -> Result<Vec<Edit>> {
    if source == target {
        return Ok(Vec::new());
    }
    let source_tokens = tokens(source)?;
    let target_tokens = tokens(target)?;
    let mut edits = Vec::new();
    let mut source_index = 0usize;
    let mut target_index = 0usize;
    while source_index < source_tokens.len() || target_index < target_tokens.len() {
        if source_index < source_tokens.len()
            && target_index < target_tokens.len()
            && token(source, &source_tokens[source_index])
                == token(target, &target_tokens[target_index])
        {
            source_index += 1;
            target_index += 1;
            continue;
        }
        let source_start = token_start(&source_tokens, source_index, source.len());
        let target_start = token_start(&target_tokens, target_index, target.len());
        let (next_source, next_target) = next_anchor(
            source,
            &source_tokens,
            source_index,
            target,
            &target_tokens,
            target_index,
        );
        let source_end = token_start(&source_tokens, next_source, source.len());
        let target_end = token_start(&target_tokens, next_target, target.len());
        edits.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "ODB XML splice edits",
            source,
        })?;
        edits.push(Edit {
            range: source_start..source_end,
            replacement: target.as_bytes()[target_start..target_end].to_vec(),
        });
        source_index = next_source;
        target_index = next_target;
    }
    Ok(edits)
}

fn next_anchor(
    source: &str,
    source_tokens: &[Range<usize>],
    source_index: usize,
    target: &str,
    target_tokens: &[Range<usize>],
    target_index: usize,
) -> (usize, usize) {
    let mut first_target = HashMap::<&str, usize>::new();
    for (index, range) in target_tokens.iter().enumerate().skip(target_index) {
        first_target.entry(token(target, range)).or_insert(index);
    }
    let mut selected = None;
    for (index, range) in source_tokens.iter().enumerate().skip(source_index) {
        let Some(candidate) = first_target.get(token(source, range)).copied() else {
            continue;
        };
        let distance = (index - source_index).saturating_add(candidate - target_index);
        if selected
            .as_ref()
            .is_none_or(|(best, _, _)| distance < *best)
        {
            selected = Some((distance, index, candidate));
        }
        if selected
            .as_ref()
            .is_some_and(|(best, _, _)| index - source_index > *best)
        {
            break;
        }
    }
    selected.map_or(
        (source_tokens.len(), target_tokens.len()),
        |(_, source, target)| (source, target),
    )
}

fn tokens(xml: &str) -> Result<Vec<Range<usize>>> {
    let bytes = xml.as_bytes();
    let mut ranges = Vec::new();
    let mut cursor = 0usize;
    while cursor < bytes.len() {
        let start = cursor;
        if bytes[cursor] != b'<' {
            cursor = bytes[cursor..]
                .iter()
                .position(|byte| *byte == b'<')
                .map_or(bytes.len(), |offset| cursor + offset);
        } else if bytes[cursor..].starts_with(b"<!--") {
            cursor = find_after(bytes, cursor + 4, b"-->")?;
        } else if bytes[cursor..].starts_with(b"<![CDATA[") {
            cursor = find_after(bytes, cursor + 9, b"]]>")?;
        } else if bytes[cursor..].starts_with(b"<?") {
            cursor = find_after(bytes, cursor + 2, b"?>")?;
        } else {
            cursor = tag_end(bytes, cursor + 1)?;
        }
        ranges.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "ODB XML lexical tokens",
            source,
        })?;
        ranges.push(start..cursor);
    }
    Ok(ranges)
}

fn find_after(bytes: &[u8], start: usize, marker: &[u8]) -> Result<usize> {
    bytes[start..]
        .windows(marker.len())
        .position(|window| window == marker)
        .map(|offset| start + offset + marker.len())
        .ok_or_else(|| Error::InvalidFormat("ODB XML lexical token is incomplete".to_string()))
}

fn tag_end(bytes: &[u8], mut cursor: usize) -> Result<usize> {
    let mut quote = None;
    while cursor < bytes.len() {
        match (bytes[cursor], quote) {
            (b'\'' | b'"', None) => quote = Some(bytes[cursor]),
            (value, Some(expected)) if value == expected => quote = None,
            (b'>', None) => return Ok(cursor + 1),
            (_, _) => {},
        }
        cursor += 1;
    }
    Err(Error::InvalidFormat(
        "ODB XML start tag is incomplete".to_string(),
    ))
}

fn classify_fragment(bytes: Vec<u8>) -> Result<AuthoredXmlFragment> {
    if bytes.is_empty() {
        return Ok(AuthoredXmlFragment::deletion());
    }
    if is_nonempty_start_tag(&bytes) {
        return AuthoredXmlFragment::start_tag(bytes);
    }
    if bytes.first() == Some(&b'<') {
        return AuthoredXmlFragment::markup(bytes);
    }
    AuthoredXmlFragment::text(bytes)
}

fn is_nonempty_start_tag(bytes: &[u8]) -> bool {
    bytes.first() == Some(&b'<')
        && bytes.last() == Some(&b'>')
        && !bytes.starts_with(b"</")
        && !bytes.starts_with(b"<!")
        && !bytes.starts_with(b"<?")
        && !bytes.ends_with(b"/>")
        && !bytes[1..bytes.len() - 1].contains(&b'<')
}

fn token<'a>(xml: &'a str, range: &Range<usize>) -> &'a str {
    &xml[range.clone()]
}

fn token_start(tokens: &[Range<usize>], index: usize, end: usize) -> usize {
    tokens.get(index).map_or(end, |range| range.start)
}

#[cfg(test)]
mod tests {
    use super::classify_fragment;

    #[test]
    fn shared_authored_fragment_classes_reject_raw_noncompact_bytes() {
        assert!(classify_fragment(b"<x  value=\"one\">".to_vec()).is_err());
        assert!(classify_fragment(b"<x>\n <y/>\n</x>".to_vec()).is_err());
        assert!(classify_fragment(b"   ".to_vec()).is_err());
        assert!(classify_fragment(b"<x>".to_vec()).is_ok());
        assert!(classify_fragment(b"one &amp; two".to_vec()).is_ok());
        assert!(classify_fragment(Vec::new()).is_ok());
    }
}
