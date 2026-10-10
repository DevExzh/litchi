//! The runs a positional run address (`r[N]`) counts.

use litchi_core::xml::ReaderOrigin;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use smallvec::SmallVec;

use crate::error::{Error, Result};

use super::super::model::{Paragraph, Run};
use super::xml::is_fragment_word_name;

/// Deepest element nesting accepted, as [`Paragraph::inlines`] accepts.
const MAX_DEPTH: usize = 128;
/// Most elements visited, as [`Paragraph::inlines`] visits.
const MAX_ELEMENTS: usize = 1_000_000;

/// An open element of the paragraph.
#[derive(Clone, Copy)]
enum Frame {
    /// A candidate run, by its index in the candidate list.
    Run(usize),
    /// An element whose runs are not addressed as the paragraph's runs.
    Excluding,
    /// Anything else.
    Other,
}

/// A run found in document order, with whether it is addressable.
struct Candidate {
    start: usize,
    end: usize,
    addressable: bool,
    namespaces: Option<crate::namespace::NamespaceBindings>,
}

/// How the walk treats an element.
enum Role {
    Run,
    Excluding,
    CommentReference,
    Other,
}

#[allow(
    clippy::ref_option,
    reason = "the fragment prefix is the shared tri-state of the paragraph scanners"
)]
fn role(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Role {
    let local = element.local_name();
    match local.as_ref() {
        // `OfficeCLI` matches these by local name in any namespace.
        b"AlternateContent" | b"Choice" | b"Fallback" => Role::Excluding,
        name @ (b"r" | b"sdt" | b"fldSimple" | b"txbxContent" | b"commentReference") => {
            let (namespace, _) = reader.resolver().resolve_element(element.name());
            if !is_fragment_word_name(&namespace, element.name(), name, fragment_prefix) {
                return Role::Other;
            }
            match name {
                b"r" => Role::Run,
                b"commentReference" => Role::CommentReference,
                _ => Role::Excluding,
            }
        },
        _ => Role::Other,
    }
}

impl Paragraph {
    /// Return the runs a positional run address (`r[N]`) counts, in
    /// document order.
    ///
    /// This is the run set `OfficeCLI` resolves `/…/p[N]/r[M]` against: every
    /// WordprocessingML run inside the paragraph, at any depth, except
    ///
    /// - a run with a direct `w:commentReference` child (a comment's
    ///   reference mark);
    /// - a run inside an inline content control (`w:sdt`);
    /// - a run inside a simple field (`w:fldSimple`);
    /// - a run inside a text box (`w:txbxContent`) within the paragraph;
    /// - a run inside markup-compatibility alternatives
    ///   (`AlternateContent`, `Choice` or `Fallback`, matched by local name
    ///   as `OfficeCLI` matches them).
    ///
    /// Runs inside hyperlinks, tracked changes, smart tags, custom XML,
    /// bidirectional wrappers and Office Math are counted, and so are the
    /// base and guide runs of a phonetic guide (`w:ruby`), each after the
    /// run that contains it. Only the paragraph's own XML is consulted: a
    /// paragraph inside a text box that itself sits in an inline content
    /// control keeps its runs here, where `OfficeCLI`, which also inspects
    /// the ancestors outside the paragraph, would drop them.
    ///
    /// [`Self::runs`] is a different set, the outermost runs of the
    /// paragraph; it keeps comment-reference runs and runs inside content
    /// controls, simple fields and alternatives, and never returns a run
    /// nested in another run.
    ///
    /// ```
    /// use litchi_docx::Paragraph;
    ///
    /// let paragraph = Paragraph::new(
    ///     br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>one</w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>2</w:t></w:r></w:fldSimple><w:hyperlink w:anchor="a"><w:r><w:t>three</w:t></w:r></w:hyperlink><w:r><w:commentReference w:id="0"/></w:r></w:p>"#.to_vec(),
    /// );
    /// let texts: Vec<String> = paragraph
    ///     .addressable_runs()?
    ///     .iter()
    ///     .map(|run| run.text())
    ///     .collect::<Result<_, _>>()?;
    /// assert_eq!(texts, ["one", "three"]);
    /// assert_eq!(paragraph.runs()?.len(), 4);
    /// # Ok::<(), litchi_docx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for paragraph XML that is malformed, unterminated,
    /// nested deeper than 128 elements, larger than a million elements, or
    /// whose root is not a WordprocessingML `w:p`, or for an execution limit
    /// of a source-backed paragraph.
    pub fn addressable_runs(&self) -> Result<SmallVec<[Run; 8]>> {
        let xml = self.xml_bytes();
        let source = self.xml_data.xml_ref()?;
        let _parser_admission = source.parser_admission()?;
        let mut reader = NsReader::from_reader(xml);
        for (prefix, namespace) in self.xml_data.inherited_namespaces()?.iter() {
            let prefix = prefix.as_deref().map_or(
                quick_xml::name::PrefixDeclaration::Default,
                quick_xml::name::PrefixDeclaration::Named,
            );
            reader
                .resolver_mut()
                .add(prefix, quick_xml::name::Namespace(namespace))
                .map_err(|_| Error::Xml("invalid inherited paragraph namespace".into()))?;
        }
        let origin = ReaderOrigin::of(xml);
        let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
        let mut stack: SmallVec<[Frame; 32]> = SmallVec::new();
        let mut candidates: SmallVec<[Candidate; 8]> = SmallVec::new();
        let mut excluding = 0usize;
        let mut namespaces = crate::namespace::NamespaceCapture::default();
        let mut elements = 0usize;
        let mut saw_root = false;

        loop {
            let event_start = origin.offset(reader.buffer_position()).ok_or_else(|| {
                Error::InvalidFormat("Word paragraph offset does not fit usize".into())
            })?;
            let event = reader
                .read_event()
                .map_err(|_| Error::Xml("malformed addressable-run XML".into()))?;
            let event_end = origin.offset(reader.buffer_position()).ok_or_else(|| {
                Error::InvalidFormat("Word paragraph offset does not fit usize".into())
            })?;
            let (element, empty) = match event {
                Event::Start(element) => (element, false),
                Event::Empty(element) => (element, true),
                Event::End(_) => {
                    match stack.pop() {
                        Some(Frame::Run(index)) => candidates[index].end = event_end,
                        Some(Frame::Excluding) => excluding -= 1,
                        Some(Frame::Other) => {},
                        None => {
                            return Err(Error::InvalidFormat(
                                "invalid Word paragraph nesting".into(),
                            ));
                        },
                    }
                    continue;
                },
                Event::Eof if !stack.is_empty() => {
                    return Err(Error::InvalidFormat(
                        "unterminated Word paragraph XML".into(),
                    ));
                },
                Event::Eof => break,
                Event::DocType(_) => {
                    return Err(Error::Xml("DOCTYPE is forbidden in paragraph XML".into()));
                },
                Event::Text(ref text)
                    if stack.is_empty() && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
                {
                    return Err(Error::Xml("text outside the paragraph XML root".into()));
                },
                Event::CData(_) | Event::GeneralRef(_) if stack.is_empty() => {
                    return Err(Error::Xml("content outside the paragraph XML root".into()));
                },
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::GeneralRef(_) => continue,
            };

            crate::properties::audit_attributes(&element, reader.resolver(), reader.decoder())?;
            elements += 1;
            if elements > MAX_ELEMENTS {
                return Err(Error::InvalidFormat(format!(
                    "Word paragraph exceeds {MAX_ELEMENTS} elements"
                )));
            }
            if stack.is_empty() {
                if saw_root {
                    return Err(Error::InvalidFormat(
                        "paragraph fragment contains multiple roots".into(),
                    ));
                }
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                if !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix = Some(
                        element
                            .name()
                            .prefix()
                            .map(|prefix| prefix.into_inner().to_vec()),
                    );
                }
                if !is_fragment_word_name(&namespace, element.name(), b"p", &fragment_prefix) {
                    return Err(Error::InvalidFormat(
                        "paragraph XML has no WordprocessingML root".into(),
                    ));
                }
                saw_root = true;
                if !empty {
                    stack.push(Frame::Other);
                }
                continue;
            }
            if !empty && stack.len() >= MAX_DEPTH {
                return Err(Error::InvalidFormat(format!(
                    "Word paragraph nesting exceeds the {MAX_DEPTH} depth limit"
                )));
            }

            let frame = match role(&reader, &element, &fragment_prefix) {
                Role::Run => {
                    candidates.push(Candidate {
                        start: event_start,
                        end: event_end,
                        addressable: excluding == 0,
                        namespaces: if excluding == 0 {
                            Some(namespaces.capture(reader.resolver())?)
                        } else {
                            None
                        },
                    });
                    Frame::Run(candidates.len() - 1)
                },
                Role::Excluding => {
                    if !empty {
                        excluding += 1;
                    }
                    Frame::Excluding
                },
                Role::CommentReference => {
                    if let Some(Frame::Run(index)) = stack.last() {
                        candidates[*index].addressable = false;
                    }
                    Frame::Other
                },
                Role::Other => Frame::Other,
            };
            if !empty {
                stack.push(frame);
            }
        }

        if !saw_root {
            return Err(Error::InvalidFormat(
                "paragraph XML has no WordprocessingML root".into(),
            ));
        }
        let mut runs = SmallVec::new();
        for candidate in candidates.iter().filter(|candidate| candidate.addressable) {
            Self::push_run_slice(&mut runs, &source, candidate.start, candidate.end)?;
            if let Some(run) = runs.last_mut() {
                run.namespaces.clone_from(&candidate.namespaces);
            }
        }
        Ok(runs)
    }
}

#[cfg(test)]
#[path = "addressable_tests.rs"]
mod tests;
