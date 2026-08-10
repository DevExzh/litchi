//! Failure-atomic hyperlink editing across main, header, and footer stories.

use super::super::model::{Error, Package, Result};
use super::super::story::{
    StoryHyperlinkTextReplacement, StoryInventory, StoryKind, StoryLimits, capture,
};
use crate::document::{
    HyperlinkTextReplacement, Refusal, Snapshot, TransactionError, TransactionResult,
};
use litchi_opc::PackURI;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::BTreeMap;

const MAX_STORY_HYPERLINK_REPLACEMENTS: usize = 4_096;

#[derive(Debug)]
struct PlannedStory {
    part: PackURI,
    source: Vec<u8>,
    replacement: Vec<u8>,
}

#[derive(Debug)]
struct RootLayout {
    open_end: usize,
    close_start: usize,
    qualified_name: Vec<u8>,
    root_start: usize,
}

impl Package {
    /// Atomically replace direct hyperlink text across the main document and
    /// any reachable header or footer story.
    ///
    /// Replacements must be non-empty, unique, and strictly increasing by
    /// part name and then by paragraph/hyperlink address. Story relationships,
    /// root attributes, declarations, and all unselected XML remain exact.
    ///
    /// # Errors
    ///
    /// Returns a typed document transaction error when a selector is
    /// ambiguous, a story is not an editable main/header/footer story, a leaf
    /// cannot be represented safely, or candidate package validation fails.
    pub fn replace_story_hyperlink_texts(
        &mut self,
        replacements: &[StoryHyperlinkTextReplacement],
    ) -> TransactionResult<StoryInventory> {
        validate_replacements(replacements)?;
        let limits = StoryLimits::default();
        let inventory = self
            .story_inventory_with_limits(limits)
            .map_err(TransactionError::from)?;
        let mut grouped = BTreeMap::<&str, Vec<HyperlinkTextReplacement>>::new();
        for replacement in replacements {
            grouped.entry(replacement.part_name()).or_default().push(
                HyperlinkTextReplacement::new(replacement.address(), replacement.text()),
            );
        }

        let mut planned = Vec::new();
        planned
            .try_reserve(grouped.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX story hyperlink plans",
                source,
            })?;
        for (part_name, story_replacements) in grouped {
            let part = PackURI::new(part_name).map_err(|error| {
                Error::InvalidUri(format!("story hyperlink part name: {error}"))
            })?;
            let story = inventory.get(&part).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "story hyperlink selector targets unreachable part '{part_name}'"
                ))
            })?;
            if !matches!(
                story.kind(),
                StoryKind::Main | StoryKind::Header | StoryKind::Footer
            ) {
                return Err(Error::InvalidFormat(format!(
                    "story hyperlink selector targets unsupported {:?} story '{part_name}'",
                    story.kind()
                ))
                .into());
            }
            let source = story.source().to_vec();
            let replacement =
                edit_story(story.kind(), story.source(), story_replacements.as_slice())?;
            if source != replacement {
                planned.push(PlannedStory {
                    part,
                    source,
                    replacement,
                });
            }
        }

        if planned.is_empty() {
            return Ok(inventory);
        }
        let topology = inventory.topology().as_bytes().to_vec();
        self.edit_semantic_opc("replace_story_hyperlink_texts", move |candidate| {
            let before = capture(candidate, limits)?;
            if before.topology().as_bytes() != topology.as_slice() {
                return Err(Error::InvalidFormat(
                    "DOCX story topology changed while staging hyperlink edits".into(),
                ));
            }
            for plan in &planned {
                let story = before.get(&plan.part).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "DOCX story '{}' disappeared while staging hyperlink edits",
                        plan.part
                    ))
                })?;
                if story.source() != plan.source {
                    return Err(Error::InvalidFormat(format!(
                        "DOCX story '{}' changed while staging hyperlink edits",
                        plan.part
                    )));
                }
            }
            for plan in planned {
                candidate
                    .get_part_mut(&plan.part)?
                    .set_blob(plan.replacement);
            }
            let after = capture(candidate, limits)?;
            if after.topology().as_bytes() != topology.as_slice() {
                return Err(Error::InvalidFormat(
                    "DOCX story hyperlink edit changed package ownership topology".into(),
                ));
            }
            Ok(after)
        })
        .map_err(TransactionError::from)
    }
}

fn validate_replacements(replacements: &[StoryHyperlinkTextReplacement]) -> TransactionResult<()> {
    if replacements.is_empty() || replacements.len() > MAX_STORY_HYPERLINK_REPLACEMENTS {
        return if replacements.len() > MAX_STORY_HYPERLINK_REPLACEMENTS {
            Err(TransactionError::Limit {
                resource: "story hyperlink replacements",
                max: MAX_STORY_HYPERLINK_REPLACEMENTS,
                actual: replacements.len(),
            })
        } else {
            Err(TransactionError::Refused {
                position: 0,
                reason: Refusal::AmbiguousCompositeSelector,
            })
        };
    }
    let mut previous = None;
    for replacement in replacements {
        if replacement.part_name().is_empty() {
            return Err(TransactionError::Refused {
                position: replacement.address().paragraph.get(),
                reason: Refusal::AmbiguousCompositeSelector,
            });
        }
        let key = (replacement.part_name(), replacement.address());
        if previous.is_some_and(|value| value >= key) {
            return Err(TransactionError::Refused {
                position: replacement.address().paragraph.get(),
                reason: Refusal::AmbiguousCompositeSelector,
            });
        }
        previous = Some(key);
    }
    Ok(())
}

fn edit_story(
    kind: StoryKind,
    source: &[u8],
    replacements: &[HyperlinkTextReplacement],
) -> TransactionResult<Vec<u8>> {
    if kind == StoryKind::Main {
        let snapshot = Snapshot::from_xml(source.to_vec())?;
        let mut edit = snapshot.edit();
        edit.replace_body_hyperlink_texts(replacements)?;
        return Ok(edit.projected().xml_bytes().to_vec());
    }

    let layout = root_layout(source, kind)?;
    let qualified_document = qualified_sibling(&layout.qualified_name, b"document");
    let qualified_body = qualified_sibling(&layout.qualified_name, b"body");
    let raw_open = source
        .get(layout.root_start..layout.open_end)
        .ok_or_else(|| Error::InvalidFormat("story root range is invalid".into()))?;
    let name_end = 1usize
        .checked_add(layout.qualified_name.len())
        .ok_or_else(|| Error::InvalidFormat("story root name range overflows".into()))?;
    if raw_open.get(..1) != Some(b"<") || name_end > raw_open.len() {
        return Err(Error::InvalidFormat("story root start tag is invalid".into()).into());
    }

    let mut prefix = Vec::new();
    prefix.extend_from_slice(b"<");
    prefix.extend_from_slice(&qualified_document);
    prefix.extend_from_slice(&raw_open[name_end..]);
    prefix.extend_from_slice(b"<");
    prefix.extend_from_slice(&qualified_body);
    prefix.extend_from_slice(b">");
    let mut suffix = Vec::new();
    suffix.extend_from_slice(b"</");
    suffix.extend_from_slice(&qualified_body);
    suffix.extend_from_slice(b"></");
    suffix.extend_from_slice(&qualified_document);
    suffix.extend_from_slice(b">");

    let inner = source
        .get(layout.open_end..layout.close_start)
        .ok_or_else(|| Error::InvalidFormat("story root content range is invalid".into()))?;
    let mut synthetic = Vec::new();
    synthetic.extend_from_slice(&prefix);
    synthetic.extend_from_slice(inner);
    synthetic.extend_from_slice(&suffix);
    let snapshot = Snapshot::from_xml(synthetic)?;
    let mut edit = snapshot.edit();
    edit.replace_body_hyperlink_texts(replacements)?;
    let projected = edit.projected().xml_bytes();
    if !projected.starts_with(&prefix)
        || !projected.ends_with(&suffix)
        || projected.len() < prefix.len() + suffix.len()
    {
        return Err(Error::InvalidFormat(
            "story hyperlink edit changed the synthetic ownership wrapper".into(),
        )
        .into());
    }
    let edited_inner = &projected[prefix.len()..projected.len() - suffix.len()];
    let mut output = Vec::new();
    output.extend_from_slice(&source[..layout.open_end]);
    output.extend_from_slice(edited_inner);
    output.extend_from_slice(&source[layout.close_start..]);
    Ok(output)
}

fn root_layout(source: &[u8], kind: StoryKind) -> Result<RootLayout> {
    let expected = match kind {
        StoryKind::Header => b"hdr".as_slice(),
        StoryKind::Footer => b"ftr".as_slice(),
        StoryKind::Main
        | StoryKind::Footnotes
        | StoryKind::Endnotes
        | StoryKind::Comments
        | StoryKind::Glossary => {
            return Err(Error::InvalidFormat(
                "synthetic story wrapper requires a header or footer".into(),
            ));
        },
    };
    let mut reader = Reader::from_reader(source);
    reader.config_mut().check_end_names = true;
    reader.config_mut().allow_unmatched_ends = false;
    let mut root = None;
    let mut depth = 0usize;
    loop {
        let before = usize::try_from(reader.buffer_position()).map_err(|_error| {
            Error::InvalidFormat("story XML position does not fit usize".into())
        })?;
        match reader.read_event().map_err(Error::from)? {
            Event::Start(element) => {
                depth += 1;
                if depth == 1 {
                    if element.local_name().as_ref() != expected {
                        return Err(Error::InvalidFormat(
                            "story root does not match its package role".into(),
                        ));
                    }
                    root = Some((
                        before,
                        usize::try_from(reader.buffer_position()).map_err(|_error| {
                            Error::InvalidFormat("story XML position does not fit usize".into())
                        })?,
                        element.name().as_ref().to_vec(),
                    ));
                }
            },
            Event::End(_) => {
                if depth == 1 {
                    let (root_start, open_end, qualified_name) = root.ok_or_else(|| {
                        Error::InvalidFormat("story XML has no root start".into())
                    })?;
                    return Ok(RootLayout {
                        open_end,
                        close_start: before,
                        qualified_name,
                        root_start,
                    });
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("story XML depth underflows".into()))?;
            },
            Event::Empty(element) if depth == 0 && element.local_name().as_ref() == expected => {
                return Err(Error::InvalidFormat(
                    "empty header/footer story has no hyperlink content".into(),
                ));
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "story XML ended before its root closed".into(),
                ));
            },
            Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn qualified_sibling(qualified_name: &[u8], local_name: &[u8]) -> Vec<u8> {
    let mut output = qualified_name
        .iter()
        .rposition(|byte| *byte == b':')
        .map_or_else(Vec::new, |colon| qualified_name[..=colon].to_vec());
    output.extend_from_slice(local_name);
    output
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::document::ParagraphHyperlinkAddress;
    use litchi_core::Position;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{BlobPart, Part};

    const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn hyperlink_story(root: &str, id: &str, text: &str, root_attribute: &str) -> Vec<u8> {
        format!(
            "<?xml version=\"1.0\"?><w:{root} xmlns:w=\"{WORD}\" xmlns:r=\"{REL}\" keep:root=\"{root_attribute}\" xmlns:keep=\"urn:keep\"><w:p keep:p=\"exact\"><w:hyperlink r:id=\"{id}\" w:tooltip=\"tip\"><w:r><w:rPr><w:b/></w:rPr><w:t>{text}</w:t></w:r></w:hyperlink></w:p></w:{root}>"
        )
        .into_bytes()
    }

    fn package() -> Package {
        let mut package = Package::new().unwrap();
        let main = package.opc.main_document_part().unwrap().partname().clone();
        package.opc.get_part_mut(&main).unwrap().set_blob(
            format!(
                "<w:document xmlns:w=\"{WORD}\" xmlns:r=\"{REL}\"><w:body><w:p><w:hyperlink r:id=\"rBodyLink\"><w:r><w:t>body old</w:t></w:r></w:hyperlink></w:p><w:sectPr/></w:body></w:document>"
            )
            .into_bytes(),
        );
        for (name, target, content_type, relationship_type, id, root, link, text) in [
            (
                "/word/footer1.xml",
                "footer1.xml",
                ct::WML_FOOTER,
                rt::FOOTER,
                "rFooter",
                "ftr",
                "rFooterLink",
                "footer old",
            ),
            (
                "/word/header1.xml",
                "header1.xml",
                ct::WML_HEADER,
                rt::HEADER,
                "rHeader",
                "hdr",
                "rHeaderLink",
                "header old",
            ),
        ] {
            let uri = PackURI::new(name).unwrap();
            let mut part = BlobPart::new(
                uri,
                content_type.to_owned(),
                hyperlink_story(root, link, text, root),
            );
            part.rels_mut().add_relationship(
                rt::HYPERLINK.to_owned(),
                format!("https://example.invalid/{root}"),
                link.to_owned(),
                true,
            );
            package.opc.add_part(Box::new(part));
            package
                .opc
                .get_part_mut(&main)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    relationship_type.to_owned(),
                    target.to_owned(),
                    id.to_owned(),
                    false,
                );
        }
        package
            .opc
            .get_part_mut(&main)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::HYPERLINK.to_owned(),
                "https://example.invalid/body".to_owned(),
                "rBodyLink".to_owned(),
                true,
            );
        package.mutable_doc = None;
        package
    }

    fn replacement(part: &str, text: &str) -> StoryHyperlinkTextReplacement {
        StoryHyperlinkTextReplacement::new(
            part,
            ParagraphHyperlinkAddress::new(Position::new(0), Position::new(0)),
            text,
        )
    }

    #[test]
    fn body_header_and_footer_hyperlinks_publish_atomically() {
        let mut package = package();
        let header_uri = PackURI::new("/word/header1.xml").unwrap();
        let header_relationships = package
            .opc
            .get_part(&header_uri)
            .unwrap()
            .rels()
            .iter()
            .map(|relationship| {
                (
                    relationship.r_id().to_owned(),
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.is_external(),
                )
            })
            .collect::<Vec<_>>();
        let inventory = package
            .replace_story_hyperlink_texts(&[
                replacement("/word/document.xml", "body changed"),
                replacement("/word/footer1.xml", "footer changed"),
                replacement("/word/header1.xml", "header changed"),
            ])
            .unwrap();

        for (part, expected) in [
            ("/word/document.xml", "body changed"),
            ("/word/footer1.xml", "footer changed"),
            ("/word/header1.xml", "header changed"),
        ] {
            let xml = std::str::from_utf8(
                inventory
                    .get(&PackURI::new(part).unwrap())
                    .unwrap()
                    .source(),
            )
            .unwrap();
            assert!(xml.contains(&format!("<w:t>{expected}</w:t>")));
        }
        let header = package.opc.get_part(&header_uri).unwrap();
        let header_xml = std::str::from_utf8(header.blob()).unwrap();
        assert!(header_xml.starts_with("<?xml version=\"1.0\"?><w:hdr"));
        assert!(header_xml.contains("keep:root=\"hdr\""));
        assert!(header_xml.contains("keep:p=\"exact\""));
        let relationships = header
            .rels()
            .iter()
            .map(|relationship| {
                (
                    relationship.r_id().to_owned(),
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.is_external(),
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(relationships, header_relationships);
    }

    #[test]
    fn late_story_failure_leaves_every_part_unchanged() {
        let mut package = package();
        let before = package.opc.clone();
        let result = package.replace_story_hyperlink_texts(&[
            replacement("/word/document.xml", "body would change"),
            StoryHyperlinkTextReplacement::new(
                "/word/header1.xml",
                ParagraphHyperlinkAddress::new(Position::new(1), Position::new(0)),
                "missing",
            ),
        ]);
        assert!(result.is_err());
        for name in [
            "/word/document.xml",
            "/word/footer1.xml",
            "/word/header1.xml",
        ] {
            let uri = PackURI::new(name).unwrap();
            assert_eq!(
                package.opc.get_part(&uri).unwrap().blob(),
                before.get_part(&uri).unwrap().blob()
            );
        }
    }
}
