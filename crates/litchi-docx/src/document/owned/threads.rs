//! Comment replies and resolution staged with a main-document edit (record
//! 0345).
//!
//! Word threads a reply under its parent and records a comment's resolved
//! state in `word/commentsExtended.xml`: one `w15:commentEx` per comment,
//! keyed by the `w14:paraId` of the comment's last paragraph, with
//! `w15:paraIdParent` naming the parent's paragraph and `w15:done` the
//! state. These helpers give a comment paragraph an identifier when it has
//! none, upsert the entries lexically (other entries keep their bytes),
//! create the part when the package has none, and keep an existing
//! `commentsIds` part in step for every identifier they assign.

use std::collections::BTreeSet;

use super::super::{Error, Refusal, Result};
use super::Edit;
use crate::Package;
use crate::modern_comments::{
    COMMENTS_EXTENDED_CONTENT_TYPE, COMMENTS_EXTENDED_RELATIONSHIP, COMMENTS_IDS_RELATIONSHIP,
    Conformance, Extended,
};
use crate::part_scan::{self, qualified};

/// The Word 2010 namespace `w14:paraId` lives in.
const WORD_2010_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
/// The largest `w14:paraId` Word accepts (`ST_LongHexNumber` below
/// `0x80000000`).
const MAX_PARAGRAPH_ID: u32 = 0x7FFF_FFFF;

/// One `commentsExtended` entry to write: the comment paragraph it is keyed
/// by, its parent's paragraph, and its resolved state (`None` keeps an
/// existing entry's state; a new entry is unresolved).
pub(super) struct Entry {
    pub(super) paragraph: u32,
    pub(super) parent: Option<u32>,
    pub(super) done: Option<bool>,
}

/// Paragraph identifiers no paragraph of the main document or the comments
/// part uses, handed out in ascending order.
pub(super) struct FreshIds {
    used: BTreeSet<u32>,
    cursor: u32,
    /// Every identifier handed out, in order.
    pub(super) assigned: Vec<u32>,
}

impl FreshIds {
    /// The next unused identifier.
    pub(super) fn next(&mut self) -> Result<u32> {
        loop {
            if self.cursor == 0 || self.cursor > MAX_PARAGRAPH_ID {
                // Past the largest identifier in use: take the smallest gap.
                self.cursor = (1..=MAX_PARAGRAPH_ID)
                    .find(|id| !self.used.contains(id))
                    .ok_or(Error::Refused {
                        position: 0,
                        reason: Refusal::InvalidPropertyValue,
                    })?;
            }
            let id = self.cursor;
            self.cursor = self.cursor.wrapping_add(1);
            if self.used.insert(id) {
                self.assigned.push(id);
                return Ok(id);
            }
        }
    }
}

impl Edit<Package> {
    /// Fresh paragraph identifiers: above every `paraId` of the projected
    /// main document and of `comments` (the comments part as staged).
    pub(super) fn fresh_paragraph_ids(&self, comments: Option<&[u8]>) -> Result<FreshIds> {
        let mut used: BTreeSet<u32> = self
            .inner_ref(|inner| inner.paragraph_ids())?
            .into_iter()
            .collect();
        if let Some(xml) = comments {
            used.extend(part_scan::paragraph_ids(xml));
        }
        let cursor = used
            .iter()
            .next_back()
            .map_or(1, |maximum| maximum.saturating_add(1));
        Ok(FreshIds {
            used,
            cursor,
            assigned: Vec::new(),
        })
    }

    /// Write `entries` to `commentsExtended`, creating the part when the
    /// package has none, and map every identifier of `assigned` in an
    /// existing `commentsIds` part. Returns whether anything changed.
    pub(super) fn upsert_comment_entries(
        &mut self,
        entries: &[Entry],
        assigned: &[u32],
    ) -> Result<bool> {
        let mut changed = false;
        match self.related_part(COMMENTS_EXTENDED_RELATIONSHIP)? {
            Some((name, _)) => {
                let xml = self
                    .part_payload(&name)?
                    .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?;
                let limits = crate::mail_merge::mce_limits_for_source(
                    &self.package.opc_package().read_limits(),
                );
                let updated = upsert_entries(&xml, entries, &limits)?;
                if updated.as_slice() != xml.as_slice() {
                    self.replace_part(&name, updated)?;
                    changed = true;
                }
            },
            None => {
                let items: Vec<Extended> = entries
                    .iter()
                    .map(|entry| Extended {
                        paragraph_id: entry.paragraph,
                        parent_paragraph_id: entry.parent,
                        done: entry.done.unwrap_or(false),
                    })
                    .collect();
                let conformance = if self.is_strict()? {
                    Conformance::Strict
                } else {
                    Conformance::Transitional
                };
                let limits = crate::mail_merge::mce_limits_for_source(
                    &self.package.opc_package().read_limits(),
                );
                let xml =
                    crate::modern_comments::write_comments_extended(&items, conformance, &limits)
                        .map_err(Error::Document)?;
                self.stage_related_part(
                    COMMENTS_EXTENDED_RELATIONSHIP,
                    ("", "commentsExtended", ".xml"),
                    true,
                    COMMENTS_EXTENDED_CONTENT_TYPE,
                    xml,
                )?;
                changed = true;
            },
        }
        if !assigned.is_empty()
            && let Some((name, _)) = self.related_part(COMMENTS_IDS_RELATIONSHIP)?
            && let Some(xml) = self.part_payload(&name)?
        {
            let limits =
                crate::mail_merge::mce_limits_for_source(&self.package.opc_package().read_limits());
            let updated = map_durable_ids(&xml, assigned, &limits)?;
            if updated.as_slice() != xml.as_slice() {
                self.replace_part(&name, updated)?;
                changed = true;
            }
        }
        Ok(changed)
    }
}

fn invalid(message: &str) -> Error {
    Error::Document(crate::Error::InvalidFormat(message.to_owned()))
}

fn hex(value: u32) -> String {
    format!("{value:08X}")
}

fn parse_hex(value: &str) -> Option<u32> {
    u32::from_str_radix(value.trim(), 16).ok()
}

/// ` w14:paraId="…"` for a new comment paragraph, declaring `w14` on the
/// paragraph when the comments part's root does not.
pub(super) fn paragraph_id_attributes(comments: &[u8], id: u32) -> String {
    let declared = root_tag(comments).is_some_and(|tag| tag.contains("xmlns:w14="));
    if declared {
        format!(r#" w14:paraId="{}""#, hex(id))
    } else {
        format!(
            r#" xmlns:w14="{WORD_2010_NAMESPACE}" w14:paraId="{}""#,
            hex(id)
        )
    }
}

/// The root start tag of a part, as text.
fn root_tag(xml: &[u8]) -> Option<&str> {
    let text = std::str::from_utf8(xml).ok()?;
    let mut at = 0usize;
    loop {
        let start = at + text[at..].find('<')?;
        let rest = &text[start..];
        if rest.starts_with("<?") || rest.starts_with("<!") {
            at = start + 2;
            continue;
        }
        let end = start + rest.find('>')?;
        return Some(&text[start..=end]);
    }
}

/// The `w14:paraId` of comment `id`'s last paragraph in the comments part
/// `xml` (the paragraph Word keys `commentsExtended` by, [MS-DOCX]
/// `CT_CommentEx`), giving the paragraph a fresh one (rewriting `xml`) when
/// it has none. `Ok(None)` for a comment without a paragraph.
///
/// # Errors
///
/// Returns [`Refusal::BlockNotFound`] for an unknown comment.
pub(super) fn keyed_paragraph_id(
    xml: &mut Vec<u8>,
    id: u32,
    fresh: &mut FreshIds,
) -> Result<Option<u32>> {
    let tree = part_scan::scan(xml, 2).map_err(Error::Document)?;
    let comment = tree
        .children(0)
        .find(|child| {
            let element = &tree.elements[*child];
            element.is(b"comment")
                && element
                    .attribute(b"id")
                    .and_then(|value| value.trim().parse::<u32>().ok())
                    == Some(id)
        })
        .ok_or(Error::Refused {
            position: 0,
            reason: Refusal::BlockNotFound,
        })?;
    let Some(paragraph) = tree
        .children(comment)
        .filter(|child| tree.elements[*child].is(b"p"))
        .last()
    else {
        return Ok(None);
    };
    let element = &tree.elements[paragraph];
    if let Some(existing) = element.attribute(b"paraId").and_then(parse_hex) {
        return Ok(Some(existing));
    }
    let assigned = fresh.next()?;
    let tag = &xml[element.start..element.tag_end];
    let close = if element.empty {
        tag.len() - 2
    } else {
        tag.len() - 1
    };
    let mut rewritten = tag[..close].to_vec();
    rewritten.extend_from_slice(paragraph_id_attributes(xml, assigned).as_bytes());
    rewritten.extend_from_slice(&tag[close..]);
    *xml = part_scan::splice(xml, element.start, element.tag_end, &rewritten);
    Ok(Some(assigned))
}

/// The `commentsExtended` part with each entry written: an existing
/// `commentEx` for the entry's paragraph has its parent and state
/// rewritten in place; a new one is appended.
fn upsert_entries(
    xml: &[u8],
    entries: &[Entry],
    limits: &litchi_ooxml_common::mce::Limits,
) -> Result<Vec<u8>> {
    ensure_output_bytes(xml.len(), limits)?;
    let tree = part_scan::scan(xml, 1).map_err(Error::Document)?;
    let root = tree.root().map_err(Error::Document)?;
    if root.empty {
        return Err(invalid("the commentsExtended part has an empty root"));
    }
    let prefix = root.prefix.clone();
    let close = root.close;
    let mut edits: Vec<(usize, usize, Vec<u8>)> = Vec::new();
    let mut appended = Vec::new();
    for entry in entries {
        let found = tree.children(0).find(|child| {
            let element = &tree.elements[*child];
            element.is(b"commentEx")
                && element.attribute(b"paraId").and_then(parse_hex) == Some(entry.paragraph)
        });
        match found {
            Some(child) => {
                let element = &tree.elements[child];
                let mut tag = xml[element.start..element.tag_end].to_vec();
                let mut set = |local: &str, value: String| -> Result<()> {
                    let key = crate::picture_spans::attribute_named(&tag, local)
                        .unwrap_or_else(|| qualified(&element.prefix, local));
                    tag = crate::picture_spans::with_attribute(&tag, &key, &value)
                        .ok_or_else(|| invalid("invalid commentEx tag"))?;
                    Ok(())
                };
                if let Some(parent) = entry.parent {
                    set("paraIdParent", hex(parent))?;
                }
                if let Some(done) = entry.done {
                    let current = element
                        .attribute(b"done")
                        .and_then(|value| part_scan::on_off(Some(value)));
                    if current != Some(done) {
                        set("done", if done { "1" } else { "0" }.to_owned())?;
                    }
                }
                if tag.as_slice() != &xml[element.start..element.tag_end] {
                    edits.push((element.start, element.tag_end, tag));
                }
            },
            None => {
                let name = qualified(&prefix, "commentEx");
                let attribute = |local: &str| qualified(&prefix, local);
                appended.extend_from_slice(
                    format!(
                        r#"<{name} {}="{}""#,
                        attribute("paraId"),
                        hex(entry.paragraph)
                    )
                    .as_bytes(),
                );
                if let Some(parent) = entry.parent {
                    appended.extend_from_slice(
                        format!(r#" {}="{}""#, attribute("paraIdParent"), hex(parent)).as_bytes(),
                    );
                }
                appended.extend_from_slice(
                    format!(
                        r#" {}="{}"/>"#,
                        attribute("done"),
                        if entry.done == Some(true) { "1" } else { "0" }
                    )
                    .as_bytes(),
                );
            },
        }
    }
    if !appended.is_empty() {
        edits.push((close, close, appended));
    }
    edits.sort_by_key(|(start, end, _)| (*start, *end));
    let output_len = edits
        .iter()
        .try_fold(xml.len(), |length, (start, end, bytes)| {
            length
                .checked_sub(end.saturating_sub(*start))
                .and_then(|length| length.checked_add(bytes.len()))
                .ok_or_else(|| invalid("commentsExtended output length overflow"))
        })?;
    ensure_output_bytes(output_len, limits)?;
    let mut output = Vec::with_capacity(output_len);
    let mut cursor = 0usize;
    for (start, end, bytes) in edits {
        output.extend_from_slice(&xml[cursor..start]);
        output.extend_from_slice(&bytes);
        cursor = end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok(output)
}

/// The `commentsIds` part with a `commentId` for every paragraph of
/// `assigned` it does not map, each with the next durable identifier.
fn map_durable_ids(
    xml: &[u8],
    assigned: &[u32],
    limits: &litchi_ooxml_common::mce::Limits,
) -> Result<Vec<u8>> {
    ensure_output_bytes(xml.len(), limits)?;
    let tree = part_scan::scan(xml, 1).map_err(Error::Document)?;
    let root = tree.root().map_err(Error::Document)?;
    if root.empty {
        return Err(invalid("the commentsIds part has an empty root"));
    }
    let mut mapped = BTreeSet::new();
    let mut durable = 0u32;
    for child in tree.children(0) {
        let element = &tree.elements[child];
        if !element.is(b"commentId") {
            continue;
        }
        if let Some(paragraph) = element.attribute(b"paraId").and_then(parse_hex) {
            mapped.insert(paragraph);
        }
        if let Some(id) = element.attribute(b"durableId").and_then(parse_hex) {
            durable = durable.max(id);
        }
    }
    let prefix = root.prefix.clone();
    let comment_id = qualified(&prefix, "commentId");
    let para_id = qualified(&prefix, "paraId");
    let durable_id = qualified(&prefix, "durableId");
    let initial_durable = durable;
    let mut appended_len = 0usize;
    for _ in assigned
        .iter()
        .filter(|paragraph| !mapped.contains(paragraph))
    {
        durable = durable
            .checked_add(1)
            .filter(|id| *id < MAX_PARAGRAPH_ID)
            .ok_or_else(|| invalid("commentsIds has no free durable identifier"))?;
        let element_len = 1usize
            .checked_add(comment_id.len())
            .and_then(|size| size.checked_add(1))
            .and_then(|size| size.checked_add(para_id.len()))
            .and_then(|size| size.checked_add(2))
            .and_then(|size| size.checked_add(8))
            .and_then(|size| size.checked_add(2))
            .and_then(|size| size.checked_add(durable_id.len()))
            .and_then(|size| size.checked_add(2))
            .and_then(|size| size.checked_add(8))
            .and_then(|size| size.checked_add(3))
            .ok_or_else(|| invalid("commentsIds output length overflow"))?;
        appended_len = appended_len
            .checked_add(element_len)
            .ok_or_else(|| invalid("commentsIds output length overflow"))?;
    }
    let output_len = xml
        .len()
        .checked_add(appended_len)
        .ok_or_else(|| invalid("commentsIds output length overflow"))?;
    ensure_output_bytes(output_len, limits)?;
    let mut appended = Vec::with_capacity(appended_len);
    durable = initial_durable;
    for paragraph in assigned
        .iter()
        .filter(|paragraph| !mapped.contains(paragraph))
    {
        durable = durable
            .checked_add(1)
            .filter(|id| *id < MAX_PARAGRAPH_ID)
            .ok_or_else(|| invalid("commentsIds has no free durable identifier"))?;
        appended.extend_from_slice(
            format!(
                r#"<{} {}="{}" {}="{}"/>"#,
                comment_id,
                para_id,
                hex(*paragraph),
                durable_id,
                hex(durable)
            )
            .as_bytes(),
        );
    }
    debug_assert_eq!(appended.len(), appended_len);
    Ok(part_scan::splice(xml, root.close, root.close, &appended))
}

fn ensure_output_bytes(actual: usize, limits: &litchi_ooxml_common::mce::Limits) -> Result<()> {
    let maximum = limits
        .max_output_bytes
        .min(crate::modern_comments::MAX_MODERN_COMMENT_PART_BYTES);
    if actual > maximum {
        return Err(Error::Document(crate::Error::ModernCommentLimit {
            kind: crate::modern_comments::LimitKind::OutputBytes,
            actual,
            maximum,
        }));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fresh(used: &[u32]) -> FreshIds {
        let used: BTreeSet<u32> = used.iter().copied().collect();
        let cursor = used.iter().next_back().map_or(1, |maximum| maximum + 1);
        FreshIds {
            used,
            cursor,
            assigned: Vec::new(),
        }
    }

    #[test]
    fn entries_are_upserted_and_appended() {
        let xml = br#"<w15:commentsEx xmlns:w15="urn:w15"><w15:commentEx w15:paraId="0000000A" w15:done="0"/></w15:commentsEx>"#;
        let limits = litchi_ooxml_common::mce::Limits::default();
        let updated = upsert_entries(
            xml,
            &[
                Entry {
                    paragraph: 0xA,
                    parent: None,
                    done: Some(true),
                },
                Entry {
                    paragraph: 0xB,
                    parent: Some(0xA),
                    done: Some(false),
                },
            ],
            &limits,
        )
        .unwrap();
        assert_eq!(
            String::from_utf8(updated).unwrap(),
            r#"<w15:commentsEx xmlns:w15="urn:w15"><w15:commentEx w15:paraId="0000000A" w15:done="1"/><w15:commentEx w15:paraId="0000000B" w15:paraIdParent="0000000A" w15:done="0"/></w15:commentsEx>"#
        );
        // An entry that already says so is unchanged.
        let same = upsert_entries(
            xml,
            &[Entry {
                paragraph: 0xA,
                parent: None,
                done: Some(false),
            }],
            &limits,
        )
        .unwrap();
        assert_eq!(same, xml);
    }

    #[test]
    fn existing_part_updates_refuse_over_budget_before_publication() {
        let xml = br#"<w15:commentsEx xmlns:w15="urn:w15"><w15:commentEx w15:paraId="0000000A" w15:done="0"/></w15:commentsEx>"#;
        let limits = litchi_ooxml_common::mce::Limits {
            max_output_bytes: xml.len() - 1,
            ..litchi_ooxml_common::mce::Limits::default()
        };
        let error = upsert_entries(
            xml,
            &[Entry {
                paragraph: 0xA,
                parent: None,
                done: Some(true),
            }],
            &limits,
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::Document(crate::Error::ModernCommentLimit {
                kind: crate::modern_comments::LimitKind::OutputBytes,
                ..
            })
        ));
    }

    #[test]
    fn paragraphs_get_identifiers_and_durable_ids_follow() {
        let mut xml = br#"<w:comments xmlns:w="urn:w"><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>x</w:t></w:r></w:p></w:comment></w:comments>"#.to_vec();
        let mut ids = fresh(&[0x10]);
        assert_eq!(
            keyed_paragraph_id(&mut xml, 0, &mut ids).unwrap(),
            Some(0x11)
        );
        let text = String::from_utf8(xml.clone()).unwrap();
        assert!(text.contains(r#"<w:p xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="00000011">"#), "{text}");
        // Read back without assigning another.
        assert_eq!(
            keyed_paragraph_id(&mut xml, 0, &mut ids).unwrap(),
            Some(0x11)
        );
        assert_eq!(ids.assigned, [0x11]);
        assert!(keyed_paragraph_id(&mut xml, 7, &mut ids).is_err());

        // A multi-paragraph comment is keyed by its last paragraph.
        let mut two = br#"<w:comments xmlns:w="urn:w" xmlns:w14="urn:w14"><w:comment w:id="1" w:author="A"><w:p w14:paraId="00000001"/><w:p w14:paraId="00000002"/></w:comment></w:comments>"#.to_vec();
        assert_eq!(keyed_paragraph_id(&mut two, 1, &mut ids).unwrap(), Some(2));

        let durable = br#"<w16cid:commentsIds xmlns:w16cid="urn:c"><w16cid:commentId w16cid:paraId="00000011" w16cid:durableId="00000005"/></w16cid:commentsIds>"#;
        let limits = litchi_ooxml_common::mce::Limits::default();
        let mapped = map_durable_ids(durable, &[0x11, 0x12], &limits).unwrap();
        assert!(String::from_utf8(mapped).unwrap().contains(
            r#"<w16cid:commentId w16cid:paraId="00000012" w16cid:durableId="00000006"/></w16cid:commentsIds>"#
        ));
        let expected = map_durable_ids(durable, &[0x11, 0x12], &limits).unwrap();
        let bounded = litchi_ooxml_common::mce::Limits {
            max_output_bytes: expected.len() - 1,
            ..limits
        };
        assert!(map_durable_ids(durable, &[0x11, 0x12], &bounded).is_err());
    }
}
