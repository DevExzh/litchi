//! Text nodes: the `w:t` elements of one paragraph that a text search reads
//! and rewrites (record 0273).
//!
//! [`Edit::text_nodes`] lists the searchable text nodes of a paragraph in
//! document order, grouped so that a search never joins text across a
//! boundary the paragraph's markup draws:
//!
//! - the runs of the paragraph itself and of the inline containers whose
//!   runs carry ordinary visible text (`w:hyperlink`, inline `w:sdt`
//!   content, `w:smartTag`, `w:customXml`, `w:ins`, `w:moveTo`, `w:dir`,
//!   `w:bdo`) are read; entering or leaving a container starts a new group;
//! - every other run child is a boundary: tabs, breaks, symbols, drawings,
//!   field characters and instructions, note references. Bookmarks,
//!   proofing marks, permission and comment ranges and
//!   `w:lastRenderedPageBreak` are layout or annotation markers that carry
//!   no text, and do not separate text;
//! - the text of a complex field (between its `begin` and `end` field
//!   characters) and of every other element (`w:del`, `w:moveFrom`,
//!   `w:fldSimple`, math, markup-compatibility choices) is not listed: it is
//!   a cached or deleted result, not authored text.
//!
//! [`Edit::rewrite_text_nodes`] replaces the text of listed nodes, keeping
//! every other byte of the paragraph. A rewrite may also *isolate* byte
//! ranges of its new text: each range becomes a run of its own carrying a
//! copy of the original run's start tag and properties, so a caller can
//! format exactly that text. The paths of the isolated runs are returned in
//! order.

use std::ops::Range;
use std::sync::Arc;

use litchi_core::Position;
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

use super::super::super::{
    Edit, FragmentPrefix, Operation, ParagraphSplice, Refusal, Snapshot, SplicePolicy,
    TransactionNamespaceState, checked_end, checked_slice, checked_start, decode_text_fragment,
    find_tag_end, is_fragment_word_element_with_default_fallback, replace_ranges,
    seed_transaction_namespaces, tag_name_from_start,
};
use super::{ElementPath, ElementStep, Result, bindings_at, invalid, locate, refusal};

#[cfg(test)]
mod tests;

/// The most text nodes and runs one paragraph may report.
const MAX_NODES: usize = 65_536;

/// One searchable `w:t` text node of a paragraph, as
/// [`Edit::text_nodes`] lists it.
#[derive(Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct TextNode {
    /// The node's decoded text.
    pub text: String,
    /// Nodes with the same group are adjacent text that no boundary
    /// separates; groups ascend in document order.
    pub group: usize,
    /// The path of the node's run, when an element path addresses it (a run
    /// of the paragraph, of a direct hyperlink or of a direct inline
    /// content control); `None` inside other inline containers.
    pub run: Option<ElementPath>,
}

impl std::fmt::Debug for TextNode {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // Node text is authored document content.
        formatter
            .debug_struct("TextNode")
            .field("bytes", &self.text.len())
            .field("group", &self.group)
            .finish()
    }
}

/// New text for one text node of a paragraph.
#[derive(Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct TextNodeRewrite {
    node: usize,
    text: String,
    isolate: Vec<Range<usize>>,
}

impl std::fmt::Debug for TextNodeRewrite {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("TextNodeRewrite")
            .field("node", &self.node)
            .field("bytes", &self.text.len())
            .field("isolate", &self.isolate.len())
            .finish()
    }
}

impl TextNodeRewrite {
    /// Give node `node` (an index into [`Edit::text_nodes`]) the text
    /// `text`.
    #[must_use]
    pub fn new(node: usize, text: impl Into<String>) -> Self {
        Self {
            node,
            text: text.into(),
            isolate: Vec::new(),
        }
    }

    /// Also isolate `ranges` of the new text: ascending, disjoint, non-empty
    /// byte ranges on character boundaries, each of which becomes a run of
    /// its own.
    #[must_use]
    pub fn isolating(mut self, ranges: Vec<Range<usize>>) -> Self {
        self.isolate = ranges;
        self
    }
}

/// The text-node rewrites of one paragraph, for [`Edit::rewrite_text_nodes`].
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct ParagraphRewrite {
    path: ElementPath,
    slots: Vec<TextNodeRewrite>,
}

impl ParagraphRewrite {
    /// Rewrite text nodes of the paragraph at `path`; node indexes ascend.
    #[must_use]
    pub fn new(path: ElementPath, slots: Vec<TextNodeRewrite>) -> Self {
        Self { path, slots }
    }
}

impl Edit {
    /// The searchable text nodes of the paragraph at `path`, in document
    /// order. See the [module documentation](self) for what is listed.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{ElementPath, ElementStep, Snapshot};
    ///
    /// let snapshot = Snapshot::from_xml(
    ///     r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hel</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:b/></w:rPr><w:t>lo</w:t><w:tab/><w:t>there</w:t></w:r></w:p></w:body></w:document>"#,
    /// )?;
    /// let edit = snapshot.edit();
    /// let path = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// let slots = edit.text_nodes(&path)?;
    /// let texts: Vec<_> = slots.iter().map(|slot| (slot.text.as_str(), slot.group)).collect();
    /// assert_eq!(texts[0].0, "Hel");
    /// assert_eq!(texts[1].0, "lo");
    /// assert_eq!(texts[0].1, texts[1].1);
    /// assert_ne!(texts[1].1, texts[2].1);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::BlockNotFound`] for a path that names no paragraph, and a
    /// typed refusal for malformed paragraph markup.
    pub fn text_nodes(&self, path: &ElementPath) -> Result<Vec<TextNode>> {
        let (xml, inherited) = paragraph_at(&self.projected, path)?;
        let scan = scan(xml, &inherited).map_err(refusal)?;
        let runs = &scan.runs;
        Ok(scan
            .slots
            .into_iter()
            .map(|slot| TextNode {
                text: slot.text,
                group: slot.group,
                run: runs.get(slot.run).and_then(|run| run_path(path, run)),
            })
            .collect())
    }

    /// Rewrite text nodes of several paragraphs in one atomic step and
    /// return the paths of the isolated runs, paragraph by paragraph, in
    /// order.
    ///
    /// Each paragraph is listed at most once. Direct-body paragraphs share
    /// one splice of the document; others are replaced one by one. Every
    /// byte outside the rewritten text nodes (and outside the runs split to
    /// isolate text) is kept.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{ElementPath, ElementStep, ParagraphRewrite, TextNodeRewrite, Snapshot};
    ///
    /// let snapshot = Snapshot::from_xml(
    ///     r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>say hi</w:t></w:r></w:p></w:body></w:document>"#,
    /// )?;
    /// let path = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// let mut edit = snapshot.edit();
    /// let isolated = edit.rewrite_text_nodes(&[ParagraphRewrite::new(
    ///     path.clone(),
    ///     vec![TextNodeRewrite::new(0, "say hello").isolating(vec![4..9])],
    /// )])?;
    /// assert_eq!(isolated.len(), 1);
    /// assert_eq!(isolated[0], path.child(ElementStep::Run(Position::new(1))));
    /// let commit = edit.commit()?;
    /// let xml = std::str::from_utf8(commit.snapshot().xml_bytes())?;
    /// assert!(xml.contains(
    ///     "<w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\">say </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>hello</w:t></w:r>"
    /// ));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Refuses a repeated paragraph, a node index out of range or not
    /// ascending, an isolated range that is empty, unordered or off a
    /// character boundary, text a run cannot hold (control characters), and
    /// isolation inside a container that no element path addresses
    /// (`w:ins`, `w:smartTag` and the like: [`Refusal::ComplexRun`]). The
    /// edit is unchanged on failure.
    pub fn rewrite_text_nodes(
        &mut self,
        rewrites: &[ParagraphRewrite],
    ) -> Result<Vec<ElementPath>> {
        self.ensure_unmanaged("rewrite_text_nodes")?;
        for (index, rewrite) in rewrites.iter().enumerate() {
            if rewrites[..index]
                .iter()
                .any(|earlier| earlier.path == rewrite.path)
            {
                return Err(refusal(Refusal::AmbiguousCompositeSelector));
            }
        }
        let mut candidate = self.try_clone()?;
        let mut isolated: Vec<Vec<ElementPath>> = vec![Vec::new(); rewrites.len()];
        // Paragraphs below the body first: each replaces its element, which
        // moves later offsets but no paragraph position.
        for (index, rewrite) in rewrites.iter().enumerate() {
            if rewrite.path.body_paragraph().is_some() {
                continue;
            }
            let (before, inherited) = paragraph_at(&candidate.projected, &rewrite.path)?;
            let (after, paths) = rewrite_paragraph(before, &inherited, rewrite)?;
            isolated[index] = paths;
            if after != before {
                let before = Arc::new(before.to_vec());
                candidate.replace_element(&rewrite.path, &before, &Arc::new(after))?;
            }
        }
        // Direct-body paragraphs: one splice and one layout derivation.
        let mut ranges: Vec<(usize, usize, Vec<u8>)> = Vec::new();
        let mut order: Vec<usize> = Vec::new();
        for (index, rewrite) in rewrites.iter().enumerate() {
            let Some(position) = rewrite.path.body_paragraph() else {
                continue;
            };
            let range = candidate.range(position)?;
            let start = checked_start(range, "paragraph")?;
            let end = checked_end(range, "paragraph")?;
            let before =
                checked_slice(candidate.projected.xml_bytes(), start, end, "paragraph")?.to_vec();
            let inherited = candidate
                .projected
                .paragraph_namespaces
                .get(position.get())
                .cloned()
                .ok_or_else(|| invalid("a paragraph has no namespace snapshot"))?;
            let (after, paths) = rewrite_paragraph(&before, &inherited, rewrite)?;
            isolated[index] = paths;
            if after == before {
                continue;
            }
            candidate.reserve_operation()?;
            let replacement_text_bytes = candidate.checked_text_total(after.len())?;
            candidate.operations.push(Operation::ReplaceElement {
                path: rewrite.path.clone(),
                before: Arc::new(before),
                after: Arc::new(after.clone()),
            });
            candidate.replacement_text_bytes = replacement_text_bytes;
            ranges.push((start, end, after));
            order.push(index);
        }
        if !ranges.is_empty() {
            // Splices must ascend; rewrites may arrive in any order.
            ranges.sort_unstable_by_key(|range| range.0);
            let splices: Vec<ParagraphSplice<'_>> = ranges
                .iter()
                .map(|(start, end, replacement)| ParagraphSplice {
                    start: *start,
                    end: *end,
                    replacement: replacement.as_slice(),
                })
                .collect();
            let xml = replace_ranges(candidate.projected.xml_bytes(), &ranges)?;
            let projected = candidate.projected.with_spliced_paragraphs_under(
                xml,
                &splices,
                SplicePolicy::BoundedGrowth,
            )?;
            for &index in &order {
                let Some(position) = rewrites[index].path.body_paragraph() else {
                    continue;
                };
                let range = projected.range(position)?;
                let start = checked_start(range, "paragraph")?;
                let end = checked_end(range, "paragraph")?;
                let written = checked_slice(projected.xml_bytes(), start, end, "paragraph")?;
                let expected = candidate
                    .operations
                    .iter()
                    .rev()
                    .find_map(|operation| match operation {
                        Operation::ReplaceElement { path, after, .. }
                            if *path == rewrites[index].path =>
                        {
                            Some(after)
                        },
                        _ => None,
                    })
                    .ok_or_else(|| invalid("a text-node rewrite lost its operation"))?;
                if written != expected.as_slice() {
                    return Err(invalid("text-node rewrite failed exact readback"));
                }
            }
            candidate.projected = projected;
        }
        *self = candidate;
        Ok(isolated.into_iter().flatten().collect())
    }
}

/// The bytes and inherited namespaces of the paragraph at `path`.
fn paragraph_at<'s>(
    snapshot: &'s Snapshot,
    path: &ElementPath,
) -> Result<(&'s [u8], Vec<(Option<Vec<u8>>, Vec<u8>)>)> {
    if !matches!(path.steps().last(), Some(ElementStep::Paragraph(_))) {
        return Err(refusal(Refusal::BlockNotFound));
    }
    if let Some(position) = path.body_paragraph() {
        let range = snapshot.range(position)?;
        let start = checked_start(range, "paragraph")?;
        let end = checked_end(range, "paragraph")?;
        let xml = checked_slice(snapshot.xml_bytes(), start, end, "paragraph")?;
        let inherited = snapshot
            .paragraph_namespaces
            .get(position.get())
            .map(|bindings| bindings.to_vec())
            .ok_or_else(|| invalid("a paragraph has no namespace snapshot"))?;
        return Ok((xml, inherited));
    }
    let located = locate(snapshot, path.steps())?;
    let inherited = bindings_at(snapshot.xml_bytes(), located.start)?;
    let xml = checked_slice(
        snapshot.xml_bytes(),
        located.start,
        located.end,
        "paragraph",
    )?;
    Ok((xml, inherited))
}

/// The path of `run`, a run of the paragraph at `paragraph`.
fn run_path(paragraph: &ElementPath, run: &ScannedRun) -> Option<ElementPath> {
    run_path_at(paragraph, run.container, run.ordinal)
}

fn run_path_at(
    paragraph: &ElementPath,
    container: Container,
    ordinal: usize,
) -> Option<ElementPath> {
    let base = match container {
        Container::Paragraph => paragraph.clone(),
        Container::Hyperlink(position) => {
            paragraph.child(ElementStep::Hyperlink(Position::new(position)))
        },
        Container::Control(position) => {
            paragraph.child(ElementStep::ContentControl(Position::new(position)))
        },
        Container::Other => return None,
    };
    Some(base.child(ElementStep::Run(Position::new(ordinal))))
}

/// Where a run sits, for the path of a run split out of it.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Container {
    /// A direct run of the paragraph.
    Paragraph,
    /// A run of the paragraph's `n`-th direct hyperlink.
    Hyperlink(usize),
    /// A run of the content of the paragraph's `n`-th direct inline control.
    Control(usize),
    /// A run no element path addresses.
    Other,
}

struct ScannedSlot {
    text: String,
    group: usize,
    /// The `w:t` element, relative to the paragraph.
    element: Range<usize>,
    /// Index into [`Scan::runs`].
    run: usize,
}

struct ScannedRun {
    /// The `w:r` element.
    element: Range<usize>,
    /// End of the run's start tag.
    open_end: usize,
    /// Start of the run's end tag.
    close_start: usize,
    /// The run's direct `w:rPr`.
    properties: Option<Range<usize>>,
    container: Container,
    /// Position among the container's direct runs.
    ordinal: usize,
}

struct Scan {
    slots: Vec<ScannedSlot>,
    runs: Vec<ScannedRun>,
}

/// What an open element is to the scanner.
enum Frame {
    /// The paragraph or an inline container whose runs are read.
    Container { container: Container, runs: usize },
    /// An inline `w:sdt`; its `w:sdtContent` is read.
    Control { ordinal: Option<usize> },
    /// A run being read.
    Run { index: usize },
    /// A run's `w:rPr`.
    RunProperties { index: usize },
    /// A `w:t` being read, from its start offset.
    Text { start: usize },
    /// A subtree that is not read.
    Skip,
}

/// Inline containers whose runs carry visible authored text.
const CONTAINERS: [&[u8]; 7] = [
    b"hyperlink",
    b"smartTag",
    b"customXml",
    b"ins",
    b"moveTo",
    b"dir",
    b"bdo",
];

/// Run children that carry no text and separate nothing.
const RUN_MARKERS: [&[u8]; 1] = [b"lastRenderedPageBreak"];

fn scan(
    xml: &[u8],
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> std::result::Result<Scan, Refusal> {
    let mut reader = NsReader::from_reader(xml);
    seed_transaction_namespaces(&mut reader, inherited)?;
    let mut namespaces = TransactionNamespaceState::new(inherited);
    let origin = ReaderOrigin::of(xml);
    let mut prefix = FragmentPrefix::Unseen;
    let mut frames: Vec<Frame> = Vec::new();
    let mut result = Scan {
        slots: Vec::new(),
        runs: Vec::new(),
    };
    let mut group = 0usize;
    // Open complex fields: text inside one is a cached result.
    let mut fields = 0usize;
    let mut hyperlinks = 0usize;
    let mut controls = 0usize;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or(Refusal::ComplexContent)?;
        let raw = reader
            .read_event()
            .map_err(|_error| Refusal::ComplexContent)?;
        let (namespace, event) = reader.resolver().resolve_event(raw);
        let fallback = namespaces.fallback_for_event(&event);
        let end = origin
            .offset(reader.buffer_position())
            .ok_or(Refusal::ComplexContent)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = matches!(event, Event::Empty(_));
                if frames.is_empty() {
                    prefix = FragmentPrefix::from_name(element.name());
                    if !is_word(&namespace, &prefix, fallback)
                        || element.local_name().as_ref() != b"p"
                    {
                        return Err(Refusal::ComplexContent);
                    }
                    if empty {
                        return Ok(result);
                    }
                    frames.push(Frame::Container {
                        container: Container::Paragraph,
                        runs: 0,
                    });
                    continue;
                }
                let word = is_word(&namespace, &prefix, fallback);
                let local = element.local_name();
                let local = local.as_ref();
                let top = frames.len() - 1;
                let frame = match &mut frames[top] {
                    Frame::Skip | Frame::Text { .. } | Frame::RunProperties { .. } => Frame::Skip,
                    Frame::Container { container, runs } => {
                        let direct = matches!(container, Container::Paragraph);
                        if word && local == b"r" {
                            let ordinal = *runs;
                            *runs += 1;
                            let container = *container;
                            if result.runs.len() >= MAX_NODES {
                                return Err(Refusal::ComplexContent);
                            }
                            result.runs.push(ScannedRun {
                                element: start..end,
                                open_end: end,
                                close_start: end,
                                properties: None,
                                container,
                                ordinal,
                            });
                            Frame::Run {
                                index: result.runs.len() - 1,
                            }
                        } else if word && local == b"pPr" && direct {
                            Frame::Skip
                        } else if word && local == b"hyperlink" {
                            group += 1;
                            let container = if direct {
                                hyperlinks += 1;
                                Container::Hyperlink(hyperlinks - 1)
                            } else {
                                Container::Other
                            };
                            Frame::Container { container, runs: 0 }
                        } else if word && local == b"sdt" {
                            group += 1;
                            let ordinal = direct.then(|| {
                                controls += 1;
                                controls - 1
                            });
                            Frame::Control { ordinal }
                        } else if word && CONTAINERS.contains(&local) {
                            group += 1;
                            Frame::Container {
                                container: Container::Other,
                                runs: 0,
                            }
                        } else if empty {
                            // A marker: bookmarks, proofing, permissions,
                            // comment ranges. It carries no text.
                            Frame::Skip
                        } else {
                            // Deleted text, simple fields, math, choices.
                            group += 1;
                            Frame::Skip
                        }
                    },
                    Frame::Control { ordinal } => {
                        if word && local == b"sdtContent" {
                            Frame::Container {
                                container: ordinal.map_or(Container::Other, Container::Control),
                                runs: 0,
                            }
                        } else {
                            Frame::Skip
                        }
                    },
                    Frame::Run { index } => {
                        let index = *index;
                        if word && local == b"rPr" && result.runs[index].properties.is_none() {
                            result.runs[index].properties = Some(start..end);
                            Frame::RunProperties { index }
                        } else if word && local == b"t" && fields == 0 {
                            if result.slots.len() >= MAX_NODES {
                                return Err(Refusal::ComplexContent);
                            }
                            if empty {
                                result.slots.push(ScannedSlot {
                                    text: String::new(),
                                    group,
                                    element: start..end,
                                    run: index,
                                });
                                Frame::Skip
                            } else {
                                Frame::Text { start }
                            }
                        } else if word && RUN_MARKERS.contains(&local) {
                            Frame::Skip
                        } else {
                            if word && local == b"fldChar" {
                                match field_character(element) {
                                    Some(b"begin") => fields += 1,
                                    Some(b"end") => fields = fields.saturating_sub(1),
                                    _ => {},
                                }
                            }
                            group += 1;
                            Frame::Skip
                        }
                    },
                };
                if let Frame::Run { index } = frame {
                    if empty {
                        // `<w:r/>`: a run without content.
                        let run = &mut result.runs[index];
                        run.open_end = end;
                        run.close_start = end;
                        continue;
                    }
                    result.runs[index].open_end = end;
                }
                if !empty {
                    frames.push(frame);
                } else if let Frame::Container { .. } | Frame::Control { .. } = frame {
                    // An empty container closes at once.
                    group += 1;
                }
            },
            Event::End(_) => {
                let frame = frames.pop().ok_or(Refusal::ComplexContent)?;
                match frame {
                    Frame::Text { start: text_start } => {
                        let Some(Frame::Run { index }) = frames.last() else {
                            return Err(Refusal::ComplexContent);
                        };
                        let text = decode_text_fragment(
                            xml.get(text_start..end).ok_or(Refusal::ComplexRun)?,
                        )?;
                        result.slots.push(ScannedSlot {
                            text,
                            group,
                            element: text_start..end,
                            run: *index,
                        });
                    },
                    Frame::Run { index } => {
                        let run = &mut result.runs[index];
                        run.element.end = end;
                        run.close_start = start;
                    },
                    Frame::Container {
                        container: Container::Paragraph,
                        ..
                    } if frames.is_empty() => {
                        return Ok(result);
                    },
                    Frame::Container { .. } | Frame::Control { .. } => group += 1,
                    Frame::RunProperties { index } => {
                        if let Some(properties) = &mut result.runs[index].properties {
                            properties.end = end;
                        }
                    },
                    Frame::Skip => {},
                }
            },
            Event::Text(text) => {
                if !matches!(
                    frames.last(),
                    Some(Frame::Text { .. } | Frame::Skip | Frame::RunProperties { .. })
                ) && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(Refusal::ComplexContent);
                }
            },
            Event::CData(_) | Event::GeneralRef(_) => {
                if !matches!(
                    frames.last(),
                    Some(Frame::Text { .. } | Frame::Skip | Frame::RunProperties { .. })
                ) {
                    return Err(Refusal::ComplexContent);
                }
            },
            Event::Eof => return Err(Refusal::ComplexContent),
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {},
        }
    }
}

fn is_word(
    namespace: &quick_xml::name::ResolveResult<'_>,
    prefix: &FragmentPrefix,
    fallback: bool,
) -> bool {
    is_fragment_word_element_with_default_fallback(namespace, prefix, fallback)
}

/// The `w:fldCharType` of a field character.
fn field_character<'e>(element: &'e BytesStart<'_>) -> Option<&'e [u8]> {
    element
        .unchecked_attributes()
        .filter_map(std::result::Result::ok)
        .find(|attribute| attribute.key.local_name().as_ref() == b"fldCharType")
        .map(|attribute| match attribute.value {
            std::borrow::Cow::Borrowed(value) => value,
            std::borrow::Cow::Owned(_) => b"".as_slice(),
        })
}

/// Rewrite one paragraph's text nodes, returning its new bytes and the paths of
/// its isolated runs.
fn rewrite_paragraph(
    xml: &[u8],
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
    rewrite: &ParagraphRewrite,
) -> Result<(Vec<u8>, Vec<ElementPath>)> {
    let scanned = scan(xml, inherited).map_err(refusal)?;
    let mut previous = None;
    for slot in &rewrite.slots {
        if slot.node >= scanned.slots.len() || previous.is_some_and(|index| slot.node <= index) {
            return Err(refusal(Refusal::AmbiguousCompositeSelector));
        }
        previous = Some(slot.node);
        check_text(&slot.text)?;
        let mut end = 0;
        for range in &slot.isolate {
            if range.start < end
                || range.start >= range.end
                || range.end > slot.text.len()
                || !slot.text.is_char_boundary(range.start)
                || !slot.text.is_char_boundary(range.end)
            {
                return Err(refusal(Refusal::AmbiguousCompositeSelector));
            }
            end = range.end;
        }
    }
    let rewrite_of = |index: usize| rewrite.slots.iter().find(|slot| slot.node == index);
    let mut splices: Vec<(usize, usize, Vec<u8>)> = Vec::new();
    // Runs split to isolate text: (run index, outputs, isolated outputs).
    let mut splits: Vec<(usize, usize, Vec<usize>)> = Vec::new();
    let mut run = 0;
    while run < scanned.runs.len() {
        let slots: Vec<usize> = (0..scanned.slots.len())
            .filter(|&index| scanned.slots[index].run == run)
            .collect();
        let isolates = slots
            .iter()
            .any(|&index| rewrite_of(index).is_some_and(|slot| !slot.isolate.is_empty()));
        if !isolates {
            for &index in &slots {
                if let Some(slot) = rewrite_of(index) {
                    let element = scanned.slots[index].element.clone();
                    let bytes = text_element(&xml[element.clone()], &slot.text)?;
                    splices.push((element.start, element.end, bytes));
                }
            }
            run += 1;
            continue;
        }
        let scanned_run = &scanned.runs[run];
        if scanned_run.container == Container::Other {
            return Err(refusal(Refusal::ComplexRun));
        }
        let open = &xml[scanned_run.element.start..scanned_run.open_end];
        let close = &xml[scanned_run.close_start..scanned_run.element.end];
        let properties = scanned_run
            .properties
            .clone()
            .map_or(&[][..], |range| &xml[range]);
        let mut cursor = scanned_run
            .properties
            .as_ref()
            .map_or(scanned_run.open_end, |range| range.end);
        let mut outputs: Vec<Vec<u8>> = Vec::new();
        let mut isolated: Vec<usize> = Vec::new();
        let mut current: Vec<u8> = Vec::new();
        let flush = |current: &mut Vec<u8>, outputs: &mut Vec<Vec<u8>>| {
            if !current.is_empty() {
                let mut run =
                    Vec::with_capacity(open.len() + properties.len() + current.len() + close.len());
                run.extend_from_slice(open);
                run.extend_from_slice(properties);
                run.append(current);
                run.extend_from_slice(close);
                outputs.push(run);
            }
        };
        for &index in &slots {
            let element = scanned.slots[index].element.clone();
            current.extend_from_slice(&xml[cursor..element.start]);
            cursor = element.end;
            let original = &xml[element];
            let Some(slot) = rewrite_of(index) else {
                current.extend_from_slice(original);
                continue;
            };
            let mut at = 0;
            for range in &slot.isolate {
                if range.start > at {
                    current
                        .extend_from_slice(&text_element(original, &slot.text[at..range.start])?);
                }
                flush(&mut current, &mut outputs);
                current.extend_from_slice(&text_element(original, &slot.text[range.clone()])?);
                flush(&mut current, &mut outputs);
                isolated.push(outputs.len() - 1);
                at = range.end;
            }
            if at < slot.text.len() || slot.isolate.is_empty() {
                current.extend_from_slice(&text_element(original, &slot.text[at..])?);
            }
        }
        current.extend_from_slice(&xml[cursor..scanned_run.close_start]);
        flush(&mut current, &mut outputs);
        splices.push((
            scanned_run.element.start,
            scanned_run.element.end,
            outputs.concat(),
        ));
        splits.push((run, outputs.len(), isolated));
        run += 1;
    }
    splices.sort_unstable_by_key(|splice| splice.0);
    let after = if splices.is_empty() {
        xml.to_vec()
    } else {
        replace_ranges(xml, &splices)?
    };
    // Paths of the isolated runs: each split shifts the later runs of its
    // container by its extra outputs.
    let mut paths = Vec::new();
    for (run, outputs, isolated) in &splits {
        let scanned_run = &scanned.runs[*run];
        let shift: usize = splits
            .iter()
            .filter(|(other, _, _)| {
                *other < *run && scanned.runs[*other].container == scanned_run.container
            })
            .map(|(_, outputs, _)| outputs.saturating_sub(1))
            .sum();
        let _ = outputs;
        for &output in isolated {
            paths.push(
                run_path_at(
                    &rewrite.path,
                    scanned_run.container,
                    scanned_run.ordinal + shift + output,
                )
                .ok_or_else(|| refusal(Refusal::ComplexRun))?,
            );
        }
    }
    // Read back: the searchable text must be the old text with the
    // rewrites applied, in order.
    let reread = scan(&after, inherited).map_err(refusal)?;
    let expected: String = scanned
        .slots
        .iter()
        .enumerate()
        .map(|(index, slot)| {
            rewrite_of(index).map_or(slot.text.as_str(), |slot| slot.text.as_str())
        })
        .collect();
    let actual: String = reread.slots.iter().map(|slot| slot.text.as_str()).collect();
    if actual != expected {
        return Err(invalid("text-node rewrite failed semantic readback"));
    }
    Ok((after, paths))
}

/// Text a run's `w:t` can hold: XML characters without tabs or breaks,
/// which are elements of their own.
fn check_text(text: &str) -> Result<()> {
    if text.chars().any(|character| {
        !matches!(character, '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
    }) {
        return Err(refusal(Refusal::StructuralText));
    }
    Ok(())
}

/// `original` (one `w:t` element) with `text` as its content, keeping its
/// start tag; `xml:space="preserve"` is added when the text needs it.
fn text_element(original: &[u8], text: &str) -> Result<Vec<u8>> {
    let tag_end = find_tag_end(original).ok_or_else(|| refusal(Refusal::ComplexRun))?;
    let tag = &original[..tag_end];
    let empty = tag[..tag.len() - 1].trim_ascii_end().ends_with(b"/");
    let name = tag_name_from_start(tag)?;
    let mut open = if empty {
        let slash = tag[..tag.len() - 1]
            .iter()
            .rposition(|byte| *byte == b'/')
            .ok_or_else(|| refusal(Refusal::ComplexRun))?;
        tag[..slash].trim_ascii_end().to_vec()
    } else {
        tag[..tag.len() - 1].to_vec()
    };
    let preserve = text.starts_with(char::is_whitespace) || text.ends_with(char::is_whitespace);
    if preserve && !open.windows(9).any(|window| window == b"xml:space") {
        open.extend_from_slice(b" xml:space=\"preserve\"");
    }
    open.push(b'>');
    let mut output = Vec::with_capacity(open.len() + text.len() + name.len() + 3);
    output.extend_from_slice(&open);
    for character in text.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            _ => {
                let mut buffer = [0u8; 4];
                output.extend_from_slice(character.encode_utf8(&mut buffer).as_bytes());
            },
        }
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(name);
    output.push(b'>');
    Ok(output)
}
