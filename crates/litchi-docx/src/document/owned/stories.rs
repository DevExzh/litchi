//! Header and footer parts staged with a main-document edit (record 0319,
//! L-DOCX-R5).
//!
//! A new header or footer is a new part related from the main document and
//! a `w:headerReference`/`w:footerReference` in the body-final section
//! properties, staged in one commit; removal stages the inverse. Story text
//! and formatting edits replace the part exactly, splicing only the
//! paragraphs and runs they change.

use std::sync::Arc;

use super::super::transaction::GraphChange;
use super::super::{Error, Refusal, Result};
use super::Edit;
use crate::Package;
use crate::document::{TransferPart, TransferRelationship};
use crate::header_footer::catalog::Entry;
use crate::header_footer::{Kind, Role};
use crate::part_scan::{self, Child, Tree, escape, qualified, splice_error};
use crate::properties;

const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";

/// The `CT_SectPr` children after the header and footer references.
const SECTION_ORDER: &[&str] = &[
    "headerReference",
    "footerReference",
    "footnotePr",
    "endnotePr",
    "type",
    "pgSz",
    "pgMar",
    "paperSrc",
    "pgBorders",
    "lnNumType",
    "pgNumType",
    "cols",
    "formProt",
    "vAlign",
    "noEndnote",
    "titlePg",
    "textDirection",
    "bidi",
    "rtlGutter",
    "docGrid",
    "printerSettings",
    "sectPrChange",
];

/// The `CT_Settings` children, in schema order.
const SETTINGS_ORDER: &[&str] = &[
    "writeProtection",
    "view",
    "zoom",
    "removePersonalInformation",
    "removeDateAndTime",
    "doNotDisplayPageBoundaries",
    "displayBackgroundShape",
    "printPostScriptOverText",
    "printFractionalCharacterWidth",
    "printFormsData",
    "embedTrueTypeFonts",
    "embedSystemFonts",
    "saveSubsetFonts",
    "saveFormsData",
    "mirrorMargins",
    "alignBordersAndEdges",
    "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter",
    "gutterAtTop",
    "hideSpellingErrors",
    "hideGrammaticalErrors",
    "activeWritingStyle",
    "proofState",
    "formsDesign",
    "attachedTemplate",
    "linkStyles",
    "stylePaneFormatFilter",
    "stylePaneSortMethod",
    "documentType",
    "mailMerge",
    "revisionView",
    "trackRevisions",
    "doNotTrackMoves",
    "doNotTrackFormatting",
    "documentProtection",
    "autoFormatOverride",
    "styleLockTheme",
    "styleLockQFSet",
    "defaultTabStop",
    "autoHyphenation",
    "consecutiveHyphenLimit",
    "hyphenationZone",
    "doNotHyphenateCaps",
    "showEnvelope",
    "summaryLength",
    "clickAndTypeStyle",
    "defaultTableStyle",
    "evenAndOddHeaders",
    "bookFoldRevPrinting",
    "bookFoldPrinting",
    "bookFoldPrintingSheets",
    "drawingGridHorizontalSpacing",
    "drawingGridVerticalSpacing",
    "displayHorizontalDrawingGridEvery",
    "displayVerticalDrawingGridEvery",
    "doNotUseMarginsForDrawingGridOrigin",
    "drawingGridHorizontalOrigin",
    "drawingGridVerticalOrigin",
    "doNotShadeFormData",
    "noPunctuationKerning",
    "characterSpacingControl",
    "printTwoOnOne",
    "strictFirstAndLastChars",
    "noLineBreaksAfter",
    "noLineBreaksBefore",
    "savePreviewPicture",
    "doNotValidateAgainstSchema",
    "saveInvalidXml",
    "ignoreMixedContent",
    "alwaysShowPlaceholderText",
    "doNotDemarcateInvalidXml",
    "saveXmlDataOnly",
    "useXSLTWhenSaving",
    "saveThroughXslt",
    "showXMLTags",
    "alwaysMergeEmptyNamespace",
    "updateFields",
    "hdrShapeDefaults",
    "footnotePr",
    "endnotePr",
    "compat",
    "docVars",
    "rsids",
    "mathPr",
    "attachedSchema",
    "themeFontLang",
    "clrSchemeMapping",
    "doNotIncludeSubdocsInStats",
    "doNotAutoCompressPictures",
    "forceUpgrade",
    "captions",
    "readModeInkLockDown",
    "smartTagType",
    "schemaLibrary",
    "shapeDefaults",
    "doNotEmbedSmartTags",
    "decimalSymbol",
    "listSeparator",
];

/// The content of a new header or footer: one paragraph with `text`, its
/// paragraph properties and the run properties of its text.
///
/// ```
/// use litchi_docx::document::NewStory;
///
/// let story = NewStory::new("Confidential");
/// assert_eq!(story.text(), "Confidential");
/// ```
#[derive(Debug, Clone, PartialEq, Default)]
pub struct NewStory {
    text: String,
    paragraph: properties::Paragraph,
    run: properties::Run,
    field: Option<String>,
}

impl NewStory {
    /// A story of one paragraph holding `text`; a line feed becomes a
    /// break, a tab a tab.
    #[must_use]
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            ..Self::default()
        }
    }

    /// The paragraph's properties.
    #[must_use]
    pub fn with_paragraph(mut self, paragraph: properties::Paragraph) -> Self {
        self.paragraph = paragraph;
        self
    }

    /// The text run's properties.
    #[must_use]
    pub fn with_run(mut self, run: properties::Run) -> Self {
        self.run = run;
        self
    }

    /// Follow the text with a complex field of `instruction` (` PAGE `),
    /// whose cached result is `1`, as OfficeCLI writes a header field.
    #[must_use]
    pub fn with_field(mut self, instruction: impl Into<String>) -> Self {
        self.field = Some(instruction.into());
        self
    }

    /// The text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

impl Edit<Package> {
    /// The header and footer parts the main document relates, in
    /// relationship order, parts this edit added included.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error.
    pub fn header_footer_entries(&self) -> Result<Vec<Entry>> {
        let main = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?;
        let mut entries = crate::header_footer::catalog::entries_of(main.rels())?;
        let mut removed = Vec::new();
        self.inner_ref(|inner| {
            for (change, forward) in inner.staged_graph_changes() {
                if !forward {
                    removed.extend(change.parts.iter().map(|part| part.name.clone()));
                    continue;
                }
                for (relationship, part) in change.relationships.iter().zip(change.parts.iter()) {
                    let role = if relationship.relationship_type.ends_with("/header") {
                        Role::Header
                    } else if relationship.relationship_type.ends_with("/footer") {
                        Role::Footer
                    } else {
                        continue;
                    };
                    entries.push(Entry::new(role, relationship.id.clone(), part.name.clone()));
                }
            }
            entries.retain(|entry| !removed.iter().any(|name| name == entry.name()));
            entries
        })
    }

    /// Add a header or footer of `kind` holding `story` to the body-final
    /// section, and return its part name. A `first` story also marks the
    /// section as having a distinct first page (`w:titlePg`), and an `even`
    /// one turns on `w:evenAndOddHeaders` in the settings, as OfficeCLI
    /// does.
    ///
    /// ```
    /// use litchi_docx::document::NewStory;
    /// use litchi_docx::header_footer::{Kind, Role};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Body");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let name = edit.add_header_footer(Role::Header, Kind::Primary, &NewStory::new("Draft"))?;
    /// let next = edit.commit()?.into_package();
    /// let headers = next.document()?.headers()?;
    /// assert_eq!(headers[0].text()?, "Draft");
    /// assert!(next.header_footer_entries()?.iter().any(|entry| entry.name() == name));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidPropertyValue`] when the section already
    /// has a story of that role and kind, or for text XML cannot carry.
    pub fn add_header_footer(
        &mut self,
        role: Role,
        kind: Kind,
        story: &NewStory,
    ) -> Result<String> {
        let strict = self.is_strict()?;
        let section = self.inner_mut()?.body_section()?;
        if let Some((xml, _)) = &section
            && references(xml)?
                .iter()
                .any(|reference| reference.role == role && reference.kind == kind)
        {
            return Err(invalid_value());
        }
        let namespace = if strict {
            STRICT_WORD_NAMESPACE
        } else {
            WORD_NAMESPACE
        };
        let relationships = relationships_namespace(strict);
        let root = role.root();
        let mut xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{root} xmlns:w="{namespace}" xmlns:r="{relationships}">"#
        )
        .into_bytes();
        xml.extend_from_slice(&paragraph_markup(b"w", story, strict)?);
        xml.extend_from_slice(format!("</w:{root}>").as_bytes());
        let (relationship_type, stem, content_type) = match role {
            Role::Header => (
                if strict {
                    "http://purl.oclc.org/ooxml/officeDocument/relationships/header"
                } else {
                    litchi_opc::constants::relationship_type::HEADER
                },
                "header",
                litchi_opc::constants::content_type::WML_HEADER,
            ),
            Role::Footer => (
                if strict {
                    "http://purl.oclc.org/ooxml/officeDocument/relationships/footer"
                } else {
                    litchi_opc::constants::relationship_type::FOOTER
                },
                "footer",
                litchi_opc::constants::content_type::WML_FOOTER,
            ),
        };
        let (name, relationship_id) = self.stage_related_part(
            relationship_type,
            ("", stem, ".xml"),
            false,
            content_type,
            xml,
        )?;
        let reference = |prefix: &[u8], declare: Option<&str>| {
            let mut element = format!(
                r#"<{} {}="{}" "#,
                qualified(prefix, local_of(role)),
                qualified(prefix, "type"),
                kind.to_xml()
            );
            match declare {
                Some(namespace) => element.push_str(&format!(
                    r#"xmlns:r="{namespace}" r:id="{relationship_id}"/>"#
                )),
                None => element.push_str(&format!(r#"{}:id="{relationship_id}"/>"#, "r")),
            }
            element.into_bytes()
        };
        match section {
            Some((before, bindings)) => {
                let tree = part_scan::scan(&before, 1).map_err(Error::Document)?;
                let section_element = tree.root().map_err(Error::Document)?;
                let prefix = section_element.prefix.to_vec();
                let declared = relationships_prefix(&tree, &bindings, strict);
                let markup = match &declared {
                    Some(r) if r.as_slice() == b"r" => reference(&prefix, None),
                    Some(r) => {
                        let mut element =
                            String::from_utf8_lossy(&reference(&prefix, Some(relationships)))
                                .into_owned();
                        element = element.replace(
                            &format!(r#"xmlns:r="{relationships}" r:id"#),
                            &format!("{}:id", String::from_utf8_lossy(r)),
                        );
                        element.into_bytes()
                    },
                    None => reference(&prefix, Some(relationships)),
                };
                let mut after = insert_reference(&before, &tree, &markup)?;
                if kind == Kind::FirstPage {
                    after = with_title_page(&after)?;
                }
                if after != before.as_slice() {
                    self.inner_mut()?
                        .replace_body_section(&before, &Arc::new(after))?;
                }
            },
            None => {
                let mut markup = format!(r#"<w:sectPr xmlns:w="{namespace}">"#).into_bytes();
                markup.extend_from_slice(&reference(b"w", Some(relationships)));
                if kind == Kind::FirstPage {
                    markup.extend_from_slice(b"<w:titlePg/>");
                }
                markup.extend_from_slice(b"</w:sectPr>");
                self.inner_mut()?.append_body_section(&Arc::new(markup))?;
            },
        }
        if kind == Kind::EvenPage {
            self.turn_on_even_and_odd_headers(strict)?;
        }
        Ok(name)
    }

    /// Remove the header or footer part `name`: its part, its relationship
    /// and its references in the body-final section properties.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a part the document does not
    /// relate as a header or footer, and [`Refusal::StoryReference`] when a
    /// section other than the body-final one references it.
    pub fn remove_header_footer(&mut self, name: &str) -> Result<&mut Self> {
        let main = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?;
        let relationship = main
            .rels()
            .iter()
            .find(|relationship| {
                !relationship.is_external()
                    && (relationship.reltype().ends_with("/header")
                        || relationship.reltype().ends_with("/footer"))
                    && relationship
                        .target_partname()
                        .is_ok_and(|target| target.as_str() == name)
            })
            .ok_or_else(not_found)?;
        let relationship_id = relationship.r_id().to_owned();
        let transfer_relationship = TransferRelationship {
            id: relationship_id.clone(),
            relationship_type: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: false,
        };
        let uri = litchi_opc::PackURI::new(name.to_owned()).map_err(crate::Error::Invalid)?;
        let part = self
            .package
            .opc_package()
            .get_part(&uri)
            .map_err(crate::Error::from)?;
        let part_relationships: Vec<TransferRelationship> = part
            .rels()
            .iter()
            .map(|relationship| TransferRelationship {
                id: relationship.r_id().to_owned(),
                relationship_type: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                external: relationship.is_external(),
            })
            .collect();
        let transfer_part = TransferPart {
            name: name.to_owned(),
            content_type: part.content_type().to_owned(),
            blob: part.blob_arc(),
            relationships: part_relationships.into(),
        };
        let total = count_references(self.inner_mut()?.main_xml(), &relationship_id)?;
        let section = self.inner_mut()?.body_section()?;
        let in_section = match &section {
            Some((xml, _)) => count_references(xml, &relationship_id)?,
            None => 0,
        };
        if total != in_section {
            return Err(Error::Refused {
                position: 0,
                reason: Refusal::StoryReference,
            });
        }
        if let Some((before, _)) = section
            && in_section > 0
        {
            let after = without_references(&before, &relationship_id)?;
            self.inner_mut()?
                .replace_body_section(&before, &Arc::new(after))?;
        }
        self.inner_mut()?.stage_graph_change(
            Arc::new(GraphChange {
                relationships: Arc::new([transfer_relationship]),
                parts: Arc::new([transfer_part]),
                replaced: Arc::new([]),
            }),
            false,
        )?;
        Ok(self)
    }

    /// Change the kind (`default`, `first`, `even`) the body-final section
    /// references the header or footer part `name` as.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] when the body-final section does
    /// not reference the part, and [`Refusal::InvalidPropertyValue`] when
    /// it already has a story of that role and kind.
    pub fn set_header_footer_kind(&mut self, name: &str, kind: Kind) -> Result<&mut Self> {
        let entry = self
            .header_footer_entries()?
            .into_iter()
            .find(|entry| entry.name() == name)
            .ok_or_else(not_found)?;
        let (before, _) = self.inner_mut()?.body_section()?.ok_or_else(not_found)?;
        let found = references(&before)?;
        let own = found
            .iter()
            .find(|reference| reference.id == entry.relationship_id())
            .ok_or_else(not_found)?;
        if own.kind == kind {
            return Ok(self);
        }
        if found
            .iter()
            .any(|reference| reference.role == own.role && reference.kind == kind)
        {
            return Err(invalid_value());
        }
        let mut after =
            part_scan::splice(&before, own.start, own.end, &own.with_kind(&before, kind));
        if kind == Kind::FirstPage {
            after = with_title_page(&after)?;
        }
        self.inner_mut()?
            .replace_body_section(&before, &Arc::new(after))?;
        if kind == Kind::EvenPage {
            self.turn_on_even_and_odd_headers(self.is_strict()?)?;
        }
        Ok(self)
    }

    /// Replace the text of the story part `name`: its first paragraph's
    /// runs are replaced by one run with `text`, which keeps the first run's
    /// properties, as OfficeCLI writes a header's text.
    ///
    /// A story holding a field, or text outside its first paragraph's direct
    /// runs, is refused rather than left beside the new text (record 0304).
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing part,
    /// [`Refusal::ComplexContent`] for a story the set cannot replace
    /// whole, and [`Refusal::InvalidPropertyValue`] for text XML cannot
    /// carry.
    pub fn set_story_text(&mut self, name: &str, text: &str) -> Result<&mut Self> {
        let xml = self.part_payload(name)?.ok_or_else(not_found)?;
        let updated = story_with_text(&xml, text, self.is_strict()?)?;
        if updated != xml.as_slice() {
            self.replace_part(name, updated)?;
        }
        Ok(self)
    }

    /// Change the story part `name`'s formatting: `paragraph` applies to its
    /// first paragraph's properties and `run` to the properties of every
    /// direct run of its direct paragraphs and of the first paragraph's
    /// mark, as OfficeCLI's header and footer `set` does.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing part or a story
    /// without a paragraph, or a typed refusal for a property change that
    /// cannot be written.
    pub fn format_story(
        &mut self,
        name: &str,
        paragraph: Option<&dyn Fn(&mut properties::Paragraph)>,
        run: Option<&dyn Fn(&mut properties::Run)>,
    ) -> Result<&mut Self> {
        let xml = self.part_payload(name)?.ok_or_else(not_found)?;
        let updated = formatted_story(&xml, paragraph, run, self.is_strict()?)?;
        if updated != xml.as_slice() {
            self.replace_part(name, updated)?;
        }
        Ok(self)
    }

    /// Turn on `w:evenAndOddHeaders` in the settings part, creating the part
    /// when the document has none.
    fn turn_on_even_and_odd_headers(&mut self, strict: bool) -> Result<()> {
        let relationship = if strict {
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings"
        } else {
            litchi_opc::constants::relationship_type::SETTINGS
        };
        match self.related_part(relationship)? {
            Some((name, _)) => {
                let xml = self
                    .part_payload(&name)?
                    .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?;
                let tree = part_scan::scan(&xml, 1).map_err(Error::Document)?;
                if tree.child(0, b"evenAndOddHeaders").is_some() {
                    return Ok(());
                }
                let prefix = tree.root().map_err(Error::Document)?.prefix.to_vec();
                let updated = part_scan::rewrite_children(
                    &xml,
                    &tree,
                    0,
                    "settings",
                    &[Child {
                        local: "evenAndOddHeaders",
                        bytes: Some(
                            format!("<{}/>", qualified(&prefix, "evenAndOddHeaders")).into_bytes(),
                        ),
                    }],
                    SETTINGS_ORDER,
                )
                .map_err(Error::Document)?;
                self.replace_part(&name, updated)?;
            },
            None => {
                let namespace = if strict {
                    STRICT_WORD_NAMESPACE
                } else {
                    WORD_NAMESPACE
                };
                let xml = format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{namespace}"><w:evenAndOddHeaders/></w:settings>"#
                );
                self.stage_related_part(
                    relationship,
                    ("", "settings", ".xml"),
                    true,
                    litchi_opc::constants::content_type::WML_SETTINGS,
                    xml.into_bytes(),
                )?;
            },
        }
        Ok(())
    }
}

const fn local_of(role: Role) -> &'static str {
    match role {
        Role::Header => "headerReference",
        Role::Footer => "footerReference",
    }
}

const fn relationships_namespace(strict: bool) -> &'static str {
    if strict {
        STRICT_RELATIONSHIPS_NAMESPACE
    } else {
        RELATIONSHIPS_NAMESPACE
    }
}

/// The prefix bound to the relationships namespace in a section's scope.
fn relationships_prefix(
    tree: &Tree,
    bindings: &[(Option<Vec<u8>>, Vec<u8>)],
    strict: bool,
) -> Option<Vec<u8>> {
    let namespace = relationships_namespace(strict).as_bytes();
    let mut scope = bindings.to_vec();
    for declaration in tree.bindings(0) {
        scope.retain(|(prefix, _)| *prefix != declaration.0);
        scope.push(declaration);
    }
    scope
        .into_iter()
        .rev()
        .find(|(prefix, bound)| prefix.is_some() && bound.as_slice() == namespace)
        .and_then(|(prefix, _)| prefix)
}

/// One header or footer reference of a section.
struct Reference {
    role: Role,
    kind: Kind,
    id: String,
    start: usize,
    end: usize,
}

impl Reference {
    /// The reference's element with its `w:type` set to `kind`.
    fn with_kind(&self, xml: &[u8], kind: Kind) -> Vec<u8> {
        let element = &xml[self.start..self.end];
        let text = String::from_utf8_lossy(element);
        let current = format!(r#"type="{}""#, self.kind.to_xml());
        let replaced = if text.contains(&current) {
            text.replacen(&current, &format!(r#"type="{}""#, kind.to_xml()), 1)
        } else {
            text.replacen(
                local_of(self.role),
                &format!(r#"{} w:type="{}""#, local_of(self.role), kind.to_xml()),
                1,
            )
        };
        replaced.into_bytes()
    }
}

/// The header and footer references of a `w:sectPr` fragment.
fn references(xml: &[u8]) -> Result<Vec<Reference>> {
    let tree = part_scan::scan(xml, 1).map_err(Error::Document)?;
    let mut found = Vec::new();
    for child in tree.children(0) {
        let element = &tree.elements[child];
        let role = if element.is(b"headerReference") {
            Role::Header
        } else if element.is(b"footerReference") {
            Role::Footer
        } else {
            continue;
        };
        let kind = element
            .attribute(b"type")
            .map_or(Some(Kind::Primary), Kind::from_xml)
            .unwrap_or(Kind::Primary);
        found.push(Reference {
            role,
            kind,
            id: element.attribute(b"id").unwrap_or_default().to_owned(),
            start: element.start,
            end: element.end,
        });
    }
    Ok(found)
}

/// How many header or footer references in `xml` name relationship `id`.
fn count_references(xml: &[u8], id: &str) -> Result<usize> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut count = 0usize;
    loop {
        match reader
            .read_event()
            .map_err(|_| Error::Document(crate::Error::Xml("invalid Word XML".into())))?
        {
            quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element)
                if matches!(
                    element.local_name().as_ref(),
                    b"headerReference" | b"footerReference"
                ) =>
            {
                use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
                if element
                    .unchecked_attributes()
                    .filter_map(std::result::Result::ok)
                    .any(|attribute| {
                        attribute.key.local_name().as_ref() == b"id"
                            && attribute.value.as_ref() == id.as_bytes()
                    })
                {
                    count += 1;
                }
            },
            quick_xml::events::Event::Eof => return Ok(count),
            _ => {},
        }
    }
}

/// A `w:sectPr` fragment without its references to relationship `id`.
fn without_references(xml: &[u8], id: &str) -> Result<Vec<u8>> {
    let mut output = xml.to_vec();
    let mut found = references(xml)?;
    found.retain(|reference| reference.id == id);
    found.sort_by_key(|reference| std::cmp::Reverse(reference.start));
    for reference in found {
        output = part_scan::splice(&output, reference.start, reference.end, &[]);
    }
    Ok(output)
}

/// A `w:sectPr` fragment with `markup` after its last header or footer
/// reference, or as its first child.
fn insert_reference(xml: &[u8], tree: &Tree, markup: &[u8]) -> Result<Vec<u8>> {
    let root = tree.root().map_err(Error::Document)?;
    if root.empty {
        let mut replacement = xml[root.start..root.tag_end - 2].to_vec();
        while replacement.last().is_some_and(u8::is_ascii_whitespace) {
            replacement.pop();
        }
        replacement.push(b'>');
        replacement.extend_from_slice(markup);
        replacement
            .extend_from_slice(format!("</{}>", qualified(&root.prefix, "sectPr")).as_bytes());
        return Ok(part_scan::splice(xml, root.start, root.end, &replacement));
    }
    let children: Vec<usize> = tree.children(0).collect();
    let offset = children
        .iter()
        .rev()
        .find(|child| {
            let element = &tree.elements[**child];
            element.is(b"headerReference") || element.is(b"footerReference")
        })
        .map(|child| tree.elements[*child].end)
        .or_else(|| children.first().map(|child| tree.elements[*child].start))
        .unwrap_or(root.close);
    Ok(part_scan::splice(xml, offset, offset, markup))
}

/// A `w:sectPr` fragment with `w:titlePg` on.
fn with_title_page(xml: &[u8]) -> Result<Vec<u8>> {
    let tree = part_scan::scan(xml, 1).map_err(Error::Document)?;
    if tree.child(0, b"titlePg").is_some() {
        return Ok(xml.to_vec());
    }
    let prefix = tree.root().map_err(Error::Document)?.prefix.to_vec();
    part_scan::rewrite_children(
        xml,
        &tree,
        0,
        "sectPr",
        &[Child {
            local: "titlePg",
            bytes: Some(format!("<{}/>", qualified(&prefix, "titlePg")).into_bytes()),
        }],
        SECTION_ORDER,
    )
    .map_err(Error::Document)
}

/// The markup of one paragraph of `story`, written with `prefix`.
fn paragraph_markup(prefix: &[u8], story: &NewStory, strict: bool) -> Result<Vec<u8>> {
    let mut xml = format!("<{}>", qualified(prefix, "p")).into_bytes();
    if let Some(element) =
        properties::paragraph_properties_element(prefix, &story.paragraph, strict)
            .map_err(splice_error)
            .map_err(Error::Document)?
    {
        xml.extend_from_slice(&element);
    }
    let run = properties::run_properties_element(prefix, &story.run)
        .map_err(splice_error)
        .map_err(Error::Document)?;
    if !story.text.is_empty() || story.field.is_none() {
        xml.extend_from_slice(&text_run(prefix, run.as_deref(), &story.text)?);
    }
    if let Some(instruction) = &story.field {
        if instruction
            .chars()
            .any(|character| character.is_control() && character != '\t')
        {
            return Err(invalid_value());
        }
        let q = |local: &str| qualified(prefix, local);
        let properties = run
            .as_deref()
            .map(String::from_utf8_lossy)
            .unwrap_or_default();
        let marker = |kind: &str| {
            format!(
                r#"<{r}>{properties}<{c} {t}="{kind}"/></{r}>"#,
                r = q("r"),
                c = q("fldChar"),
                t = q("fldCharType"),
            )
        };
        xml.extend_from_slice(marker("begin").as_bytes());
        xml.extend_from_slice(
            format!(
                r#"<{r}>{properties}<{i} xml:space="preserve">{}</{i}></{r}>"#,
                escape(instruction),
                r = q("r"),
                i = q("instrText"),
            )
            .as_bytes(),
        );
        xml.extend_from_slice(marker("separate").as_bytes());
        xml.extend_from_slice(&text_run(prefix, run.as_deref(), "1")?);
        xml.extend_from_slice(marker("end").as_bytes());
    }
    xml.extend_from_slice(format!("</{}>", qualified(prefix, "p")).as_bytes());
    Ok(xml)
}

/// One run holding `text`, with `properties` (a complete `w:rPr`) first:
/// line feeds become breaks and tabs tabs.
pub(crate) fn text_run(prefix: &[u8], properties: Option<&[u8]>, text: &str) -> Result<Vec<u8>> {
    if text
        .chars()
        .any(|character| character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    {
        return Err(invalid_value());
    }
    let mut xml = format!("<{}>", qualified(prefix, "r")).into_bytes();
    if let Some(properties) = properties {
        xml.extend_from_slice(properties);
    }
    let normalized = text.replace("\r\n", "\n").replace('\r', "\n");
    let mut first = true;
    for line in normalized.split('\n') {
        if !first {
            xml.extend_from_slice(format!("<{}/>", qualified(prefix, "br")).as_bytes());
        }
        first = false;
        let mut first_segment = true;
        for segment in line.split('\t') {
            if !first_segment {
                xml.extend_from_slice(format!("<{}/>", qualified(prefix, "tab")).as_bytes());
            }
            first_segment = false;
            if !segment.is_empty() {
                let t = qualified(prefix, "t");
                xml.extend_from_slice(
                    format!(r#"<{t} xml:space="preserve">{}</{t}>"#, escape(segment)).as_bytes(),
                );
            }
        }
    }
    xml.extend_from_slice(format!("</{}>", qualified(prefix, "r")).as_bytes());
    Ok(xml)
}

/// A story part with its first paragraph's text replaced (see
/// [`Edit::set_story_text`]).
fn story_with_text(xml: &[u8], text: &str, strict: bool) -> Result<Vec<u8>> {
    // A text set replaces the story's text and keeps or refuses everything
    // else (record 0304): a field is refused, and so is text the first
    // paragraph's direct runs do not hold, which the set would leave
    // beside the new text.
    if part_scan::holds_any(xml, part_scan::FIELD_MARKUP).map_err(Error::Document)? {
        return Err(complex_content());
    }
    let tree = part_scan::scan(xml, 3).map_err(Error::Document)?;
    let total = part_scan::run_text(xml, &[]).map_err(Error::Document)?;
    let mut direct = String::new();
    if let Some(first) = tree.child(0, b"p") {
        for run in tree
            .children(first)
            .filter(|child| tree.elements[*child].is(b"r"))
        {
            let element = &tree.elements[run];
            direct.push_str(
                &part_scan::run_text(&xml[element.start..element.end], &[])
                    .map_err(Error::Document)?,
            );
        }
    }
    if total != direct {
        return Err(complex_content());
    }
    let root = tree.root().map_err(Error::Document)?;
    let prefix = root.prefix.to_vec();
    let Some(paragraph) = tree.child(0, b"p") else {
        let story = NewStory::new(text);
        let markup = paragraph_markup(&prefix, &story, strict)?;
        if root.empty {
            return Err(not_found());
        }
        let offset = tree
            .children(0)
            .find(|child| tree.elements[*child].is(b"sectPr"))
            .map_or(root.close, |child| tree.elements[child].start);
        return Ok(part_scan::splice(xml, offset, offset, &markup));
    };
    let runs: Vec<usize> = tree
        .children(paragraph)
        .filter(|child| tree.elements[*child].is(b"r"))
        .collect();
    let mut in_field = false;
    let mut field_runs = Vec::new();
    for run in &runs {
        let char_type = tree.child(*run, b"fldChar").and_then(|child| {
            tree.elements[child]
                .attribute(b"fldCharType")
                .map(str::to_owned)
        });
        let instruction = tree.child(*run, b"instrText").is_some();
        if char_type.is_some() || instruction {
            field_runs.push(*run);
            match char_type.as_deref() {
                Some("begin") => in_field = true,
                Some("end") => in_field = false,
                _ => {},
            }
        } else if in_field {
            field_runs.push(*run);
        }
    }
    let static_runs: Vec<usize> = runs
        .iter()
        .copied()
        .filter(|run| !field_runs.contains(run))
        .collect();
    let run_properties = static_runs.first().and_then(|run| {
        tree.child(*run, b"rPr").map(|child| {
            let element = &tree.elements[child];
            xml[element.start..element.end].to_vec()
        })
    });
    let paragraph_prefix = tree.elements[paragraph].prefix.to_vec();
    let markup = text_run(&paragraph_prefix, run_properties.as_deref(), text)?;
    let node = &tree.elements[paragraph];
    let mut edits: Vec<(usize, usize, Vec<u8>)> = static_runs
        .iter()
        .map(|run| {
            (
                tree.elements[*run].start,
                tree.elements[*run].end,
                Vec::new(),
            )
        })
        .collect();
    if node.empty {
        let mut replacement = xml[node.start..node.tag_end - 2].to_vec();
        replacement.push(b'>');
        replacement.extend_from_slice(&markup);
        replacement
            .extend_from_slice(format!("</{}>", qualified(&paragraph_prefix, "p")).as_bytes());
        return Ok(part_scan::splice(xml, node.start, node.end, &replacement));
    }
    let offset = field_runs
        .first()
        .map_or(node.close, |run| tree.elements[*run].start);
    edits.push((offset, offset, markup));
    edits.sort_by_key(|(start, end, _)| (*start, *end != *start));
    apply(xml, edits)
}

/// A story part with its formatting changed (see [`Edit::format_story`]).
fn formatted_story(
    xml: &[u8],
    paragraph_change: Option<&dyn Fn(&mut properties::Paragraph)>,
    run_change: Option<&dyn Fn(&mut properties::Run)>,
    strict: bool,
) -> Result<Vec<u8>> {
    let tree = part_scan::scan(xml, 2).map_err(Error::Document)?;
    let paragraphs: Vec<usize> = tree
        .children(0)
        .filter(|child| tree.elements[*child].is(b"p"))
        .collect();
    let first = *paragraphs.first().ok_or_else(not_found)?;
    let mut edits = Vec::new();
    for paragraph in &paragraphs {
        let bindings = tree.bindings(*paragraph);
        if *paragraph == first {
            let element = &tree.elements[*paragraph];
            let bytes = &xml[element.start..element.end];
            let current = properties::parse_paragraph(bytes, &bindings).map_err(Error::Document)?;
            let mut target = current.clone();
            if let Some(change) = paragraph_change {
                change(&mut target);
            }
            if let Some(change) = run_change {
                let mut mark = target.mark.take().map(|mark| *mark).unwrap_or_default();
                change(&mut mark);
                target.mark = (mark != properties::Run::default()).then(|| Box::new(mark));
            }
            if let Some(spliced) =
                properties::splice_paragraph(bytes, &current, &target, strict, &bindings)
                    .map_err(splice_error)
                    .map_err(Error::Document)?
            {
                // The paragraph's runs are edited below over the original
                // bytes; splice the properties as a separate edit when the
                // run set is untouched by it.
                edits.push(ParagraphEdit::Properties { bytes: spliced });
            }
        }
        if let Some(change) = run_change {
            for run in tree
                .children(*paragraph)
                .filter(|child| tree.elements[*child].is(b"r"))
            {
                let element = &tree.elements[run];
                let bytes = &xml[element.start..element.end];
                let run_bindings = tree.bindings(run);
                let current =
                    properties::parse_run(bytes, &run_bindings).map_err(Error::Document)?;
                let mut target = current.clone();
                change(&mut target);
                if let Some(spliced) =
                    properties::splice_run(bytes, &current, &target, &run_bindings)
                        .map_err(splice_error)
                        .map_err(Error::Document)?
                {
                    edits.push(ParagraphEdit::Run {
                        run,
                        bytes: spliced,
                    });
                }
            }
        }
    }
    // A paragraph-properties edit rewrites the whole first paragraph; apply
    // its run edits inside that rewrite instead of over the part.
    let mut part_edits: Vec<(usize, usize, Vec<u8>)> = Vec::new();
    let first_element = &tree.elements[first];
    let properties_edit = edits.iter().find_map(|edit| match edit {
        ParagraphEdit::Properties { bytes, .. } => Some(bytes.clone()),
        ParagraphEdit::Run { .. } => None,
    });
    for edit in &edits {
        if let ParagraphEdit::Run { run, bytes } = edit {
            let element = &tree.elements[*run];
            let inside_first =
                element.start >= first_element.start && element.end <= first_element.end;
            if !(inside_first && properties_edit.is_some()) {
                part_edits.push((element.start, element.end, bytes.clone()));
            }
        }
    }
    if let Some(spliced) = properties_edit {
        // Re-apply the first paragraph's run edits to the spliced paragraph.
        let rescanned = part_scan::scan(&spliced, 1).map_err(Error::Document)?;
        let originals: Vec<usize> = tree
            .children(first)
            .filter(|child| tree.elements[*child].is(b"r"))
            .collect();
        let spliced_runs: Vec<usize> = rescanned
            .children(0)
            .filter(|child| rescanned.elements[*child].is(b"r"))
            .collect();
        let mut inner = Vec::new();
        for edit in &edits {
            if let ParagraphEdit::Run { run, bytes } = edit
                && let Some(index) = originals.iter().position(|original| original == run)
                && let Some(target) = spliced_runs.get(index)
            {
                let element = &rescanned.elements[*target];
                inner.push((element.start, element.end, bytes.clone()));
            }
        }
        let paragraph_bytes = apply(&spliced, inner)?;
        part_edits.push((first_element.start, first_element.end, paragraph_bytes));
    }
    part_edits.sort_by_key(|(start, _, _)| *start);
    apply(xml, part_edits)
}

enum ParagraphEdit {
    Properties { bytes: Vec<u8> },
    Run { run: usize, bytes: Vec<u8> },
}

/// `xml` with non-overlapping `(start, end, bytes)` edits applied.
fn apply(xml: &[u8], edits: Vec<(usize, usize, Vec<u8>)>) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(xml.len());
    let mut cursor = 0usize;
    for (start, end, bytes) in edits {
        if start < cursor || end < start || end > xml.len() {
            return Err(Error::Document(crate::Error::InvalidFormat(
                "story edits overlap".into(),
            )));
        }
        output.extend_from_slice(&xml[cursor..start]);
        output.extend_from_slice(&bytes);
        cursor = end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok(output)
}

const fn complex_content() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::ComplexContent,
    }
}

const fn invalid_value() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::InvalidPropertyValue,
    }
}

const fn not_found() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::BlockNotFound,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const HEADER: &str = r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Old</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>tail</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:hdr>"#;

    #[test]
    fn text_refuses_fields_and_text_beside_the_first_paragraph() {
        assert!(story_with_text(HEADER.as_bytes(), "New", false).is_err());
        let plain = r#"<w:hdr xmlns:w="w"><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Old</w:t></w:r><w:r><w:t>tail</w:t></w:r></w:p><w:p/></w:hdr>"#;
        let updated =
            String::from_utf8(story_with_text(plain.as_bytes(), "New\tText", false).unwrap())
                .unwrap();
        assert!(
            updated.contains(r#"<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">New</w:t><w:tab/><w:t xml:space="preserve">Text</w:t></w:r></w:p>"#),
            "{updated}"
        );
        assert!(!updated.contains("Old") && !updated.contains("tail"));
        assert!(story_with_text(plain.as_bytes(), "bad\u{1}", false).is_err());
        let beside = r#"<w:hdr xmlns:w="w"><w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p></w:hdr>"#;
        assert!(story_with_text(beside.as_bytes(), "C", false).is_err());
    }

    #[test]
    fn formatting_reaches_runs_marks_and_the_first_paragraph() {
        let center = |paragraph: &mut properties::Paragraph| {
            paragraph.justification = Some(properties::Justification::Center);
        };
        let italic = |run: &mut properties::Run| run.italic = Some(true);
        let updated =
            formatted_story(HEADER.as_bytes(), Some(&center), Some(&italic), false).unwrap();
        let text = String::from_utf8(updated).unwrap();
        assert!(text.contains(r#"<w:jc w:val="center"/>"#), "{text}");
        assert_eq!(text.matches("<w:i/>").count(), 9, "{text}");
    }

    #[test]
    fn section_references_insert_count_and_strip() {
        let section = br#"<w:sectPr xmlns:w="w" xmlns:r="r"><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="1"/></w:sectPr>"#;
        let tree = part_scan::scan(section, 1).unwrap();
        let inserted = insert_reference(
            section,
            &tree,
            br#"<w:footerReference w:type="default" r:id="rId2"/>"#,
        )
        .unwrap();
        let text = String::from_utf8(inserted.clone()).unwrap();
        assert!(text.find("headerReference").unwrap() < text.find("footerReference").unwrap());
        assert!(text.find("footerReference").unwrap() < text.find("pgSz").unwrap());
        assert_eq!(count_references(&inserted, "rId2").unwrap(), 1);
        let titled = String::from_utf8(with_title_page(&inserted).unwrap()).unwrap();
        assert!(titled.ends_with("<w:titlePg/></w:sectPr>"), "{titled}");
        let stripped = without_references(&inserted, "rId1").unwrap();
        assert_eq!(count_references(&stripped, "rId1").unwrap(), 0);
        let found = references(&stripped).unwrap();
        assert_eq!(found.len(), 1);
        assert_eq!(found[0].role, Role::Footer);
        let retyped = found[0].with_kind(&stripped, Kind::FirstPage);
        assert!(
            String::from_utf8(retyped)
                .unwrap()
                .contains(r#"w:type="first""#)
        );
    }

    #[test]
    fn new_story_paragraph_carries_its_properties() {
        let mut run = properties::Run::default();
        run.bold = Some(true);
        let markup = paragraph_markup(b"w", &NewStory::new("A\nB").with_run(run), false).unwrap();
        assert_eq!(
            String::from_utf8(markup).unwrap(),
            r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">A</w:t><w:br/><w:t xml:space="preserve">B</w:t></w:r></w:p>"#
        );
    }
}
