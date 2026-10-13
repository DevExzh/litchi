//! One-pass construction of the body [`Tree`].

use litchi_core::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use super::xml::{Ns, Prefixes};
use super::{AutomaticInsert, Kind, NONE, NodeData, StyleData, StylePart, Tree};

/// The most nodes one body index holds.
const MAX_NODES: usize = 4_000_000;
/// The deepest element nesting the index follows.
const MAX_DEPTH: usize = 4_096;

/// What an open element contributes to the index.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Open {
    /// Outside `office:text`, or an element without index meaning.
    Plain,
    /// An indexed element.
    Node(u32),
    /// A subtree the index skips (tracked changes, forms, declarations).
    Excluded,
    /// `text:s`, `text:tab` or `text:line-break` written with an end tag.
    Space,
    /// `office:automatic-styles` of `content.xml`.
    AutomaticStyles,
    /// A `style:style` declaration being read.
    Style(u32),
    /// The `office:text` root.
    Body,
}

struct Builder<'x> {
    xml: &'x str,
    prefixes: Prefixes,
    nodes: Vec<NodeData>,
    styles: Vec<StyleData>,
    stack: Vec<Open>,
    /// The open character-data run and its container.
    segment: Option<(u32, u32)>,
    automatic_insert: Option<AutomaticInsert>,
    body_start: Option<u32>,
}

pub(super) fn build(content: &str, styles: Option<&str>) -> Result<Tree> {
    if u32::try_from(content.len()).is_err() {
        return Err(Error::InvalidFormat(
            "ODT content.xml exceeds the body index's 4 GiB address space".to_string(),
        ));
    }
    let mut builder = Builder {
        xml: content,
        prefixes: Prefixes::of_root(content),
        nodes: Vec::new(),
        styles: Vec::new(),
        stack: Vec::new(),
        segment: None,
        automatic_insert: None,
        body_start: None,
    };
    builder.run()?;
    if builder.nodes.is_empty() {
        return Err(Error::InvalidFormat(
            "ODT content.xml has no office:text body".to_string(),
        ));
    }
    let automatic_insert = match builder.automatic_insert {
        Some(insert) => insert,
        None => AutomaticInsert::Create(builder.body_start.ok_or_else(|| {
            Error::InvalidFormat("ODT content.xml has no office:body".to_string())
        })?),
    };
    let mut automatic_names = Vec::new();
    for (index, style) in builder.styles.iter().enumerate() {
        let tag = &content[style.start as usize..style.open_end as usize];
        if let Some(name) = attribute(tag, &builder.prefixes, Ns::Style, "name") {
            automatic_names.push((name.into_owned().into_boxed_str(), index as u32));
        }
    }
    automatic_names.sort_by(|left, right| left.0.cmp(&right.0));
    automatic_names.dedup_by(|later, earlier| later.0 == earlier.0);

    let mut styles_out = builder.styles;
    let style_prefixes = styles.map(Prefixes::of_root).unwrap_or_default();
    let mut common_names = Vec::new();
    if let Some(styles_xml) = styles {
        let first = styles_out.len();
        common_styles(styles_xml, &style_prefixes, &mut styles_out)?;
        for (index, style) in styles_out.iter().enumerate().skip(first) {
            let tag = &styles_xml[style.start as usize..style.open_end as usize];
            if let Some(name) = attribute(tag, &style_prefixes, Ns::Style, "name") {
                common_names.push((name.into_owned().into_boxed_str(), index as u32));
            }
        }
        common_names.sort_by(|left, right| left.0.cmp(&right.0));
        common_names.dedup_by(|later, earlier| later.0 == earlier.0);
    }
    Ok(Tree {
        prefixes: builder.prefixes,
        style_prefixes,
        nodes: builder.nodes,
        styles: styles_out,
        automatic_names,
        common_names,
        automatic_insert,
    })
}

fn attribute<'t>(
    tag: &'t str,
    prefixes: &Prefixes,
    ns: Ns,
    local: &str,
) -> Option<std::borrow::Cow<'t, str>> {
    super::xml::attributes(tag).find_map(|(name, value)| {
        prefixes
            .is(name.as_bytes(), ns, local)
            .then(|| super::xml::unescape(value))
    })
}

fn position<R>(reader: &Reader<R>) -> Result<u32> {
    u32::try_from(reader.buffer_position())
        .map_err(|_error| Error::InvalidFormat("ODT XML offset overflow".to_string()))
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!("ODT content.xml is not well-formed: {error}"))
}

impl Builder<'_> {
    fn run(&mut self) -> Result<()> {
        let mut reader = Reader::from_str(self.xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().expand_empty_elements = false;
        loop {
            let start = position(&reader)?;
            let event = reader.read_event().map_err(xml_error)?;
            let end = position(&reader)?;
            match event {
                Event::Start(tag) => {
                    if self.stack.len() >= MAX_DEPTH {
                        return Err(Error::InvalidFormat(format!(
                            "ODT content.xml nests deeper than {MAX_DEPTH} elements"
                        )));
                    }
                    let open = self.open(&tag, start, end, false)?;
                    self.stack.push(open);
                },
                Event::Empty(tag) => {
                    self.open(&tag, start, end, true)?;
                },
                Event::End(_) => {
                    let open = self.stack.pop().unwrap_or(Open::Plain);
                    self.close(open, start, end);
                },
                Event::Text(text) => {
                    let whitespace = text.iter().all(u8::is_ascii_whitespace);
                    self.character_data(start, end, whitespace && text.contains(&b'\n'));
                },
                Event::CData(_) | Event::GeneralRef(_) => {
                    self.character_data(start, end, false);
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(())
    }

    /// The innermost indexed node and whether it is the direct parent.
    fn container(&self) -> Option<(u32, bool)> {
        let mut direct = true;
        for open in self.stack.iter().rev() {
            match open {
                Open::Node(id) => return Some((*id, direct)),
                Open::Excluded => return None,
                Open::Body => return Some((0, direct)),
                Open::Plain | Open::Space => direct = false,
                Open::AutomaticStyles | Open::Style(_) => return None,
            }
        }
        None
    }

    fn in_body(&self) -> bool {
        self.stack.iter().any(|open| matches!(open, Open::Body))
            && !self.stack.iter().any(|open| matches!(open, Open::Excluded))
    }

    fn character_data(&mut self, start: u32, end: u32, formatting: bool) {
        if !self.in_body() {
            return;
        }
        let Some((container, direct)) = self.container() else {
            return;
        };
        if !direct || !self.nodes[container as usize].kind.holds_runs() {
            return;
        }
        match self.segment {
            Some((segment, owner)) if owner == container => {
                let data = &mut self.nodes[segment as usize];
                data.end = end;
                data.close_start = end;
            },
            _ => {
                if formatting {
                    return;
                }
                if let Ok(id) = self.push_node(Kind::Text, container, start, start, end, end) {
                    self.segment = Some((id, container));
                }
            },
        }
    }

    fn open(&mut self, tag: &BytesStart<'_>, start: u32, end: u32, empty: bool) -> Result<Open> {
        let name = tag.name();
        let name = name.as_ref();
        let resolved = self.prefixes.resolve(name);
        // Root bookkeeping outside the body.
        if !self.in_body() {
            return Ok(self.open_outside(resolved, start, end, empty));
        }
        let Some((ns, local)) = resolved else {
            self.segment = None;
            return Ok(Open::Plain);
        };
        // White space elements extend the current character-data run.
        if ns == Ns::Text && matches!(local, b"s" | b"tab" | b"line-break" | b"soft-page-break") {
            self.character_data(start, end, false);
            return Ok(if empty { Open::Plain } else { Open::Space });
        }
        self.segment = None;
        if excluded(ns, local) {
            return Ok(Open::Excluded);
        }
        let Some(kind) = self.kind(ns, local, start, end) else {
            return Ok(Open::Plain);
        };
        let Some((mut parent, _)) = self.container() else {
            return Ok(Open::Plain);
        };
        if self.nodes[parent as usize].kind == Kind::Span {
            if kind == Kind::Span {
                // Nested spans belong to the outer run.
                return Ok(Open::Plain);
            }
            parent = self.nodes[parent as usize].parent;
        }
        let id = if empty {
            self.push_node(kind, parent, start, end, end, end)?
        } else {
            self.push_node(kind, parent, start, end, NONE, NONE)?
        };
        Ok(if empty { Open::Plain } else { Open::Node(id) })
    }

    fn open_outside(
        &mut self,
        resolved: Option<(Ns, &[u8])>,
        start: u32,
        end: u32,
        empty: bool,
    ) -> Open {
        let Some((ns, local)) = resolved else {
            return self.inside_style(None, start, end);
        };
        match (ns, local) {
            (Ns::Office, b"automatic-styles") if self.stack.len() == 1 => {
                if empty {
                    self.automatic_insert = Some(AutomaticInsert::Expand(start, end));
                    Open::Plain
                } else {
                    Open::AutomaticStyles
                }
            },
            (Ns::Office, b"body") if self.stack.len() == 1 => {
                self.body_start = Some(start);
                Open::Plain
            },
            (Ns::Office, b"text")
                if self.stack.len() == 2 && self.nodes.is_empty() && self.body_start.is_some() =>
            {
                let close = if empty { end } else { NONE };
                self.nodes.push(NodeData {
                    kind: Kind::Body,
                    parent: NONE,
                    first_child: NONE,
                    last_child: NONE,
                    next_sibling: NONE,
                    start,
                    open_end: end,
                    close_start: close,
                    end: close,
                });
                if empty { Open::Plain } else { Open::Body }
            },
            (Ns::Style, b"style") if matches!(self.stack.last(), Some(Open::AutomaticStyles)) => {
                let index = self.styles.len() as u32;
                self.styles.push(StyleData {
                    part: StylePart::Automatic,
                    start,
                    open_end: end,
                    end,
                    text: None,
                    paragraph: None,
                    cell: None,
                });
                if empty {
                    Open::Plain
                } else {
                    Open::Style(index)
                }
            },
            _ => self.inside_style(Some((ns, local)), start, end),
        }
    }

    /// Record a property child of the style being read.
    fn inside_style(&mut self, resolved: Option<(Ns, &[u8])>, start: u32, end: u32) -> Open {
        if let (Some(Open::Style(index)), Some((Ns::Style, local))) = (self.stack.last(), resolved)
        {
            let style = &mut self.styles[*index as usize];
            match local {
                b"text-properties" => style.text = Some(start..end),
                b"paragraph-properties" => style.paragraph = Some(start..end),
                b"table-cell-properties" => style.cell = Some(start..end),
                _ => {},
            }
        }
        Open::Plain
    }

    fn kind(&self, ns: Ns, local: &[u8], start: u32, end: u32) -> Option<Kind> {
        Some(match (ns, local) {
            (Ns::Text, b"p") => Kind::Paragraph,
            (Ns::Text, b"h") => Kind::Heading,
            (Ns::Text, b"span") => Kind::Span,
            (Ns::Text, b"a") => Kind::Hyperlink,
            (Ns::Text, b"list") => Kind::List,
            (Ns::Text, b"list-item" | b"list-header") => Kind::ListItem,
            (Ns::Text, b"section") => Kind::Section,
            (Ns::Text, b"note") => {
                let tag = &self.xml[start as usize..end as usize];
                match attribute(tag, &self.prefixes, Ns::Text, "note-class").as_deref() {
                    Some("endnote") => Kind::Endnote,
                    _ => Kind::Footnote,
                }
            },
            (Ns::Text, b"bookmark" | b"bookmark-start") => Kind::Bookmark,
            (Ns::Table, b"table") => Kind::Table,
            (Ns::Table, b"table-row") => Kind::Row,
            (Ns::Table, b"table-cell") => Kind::Cell,
            (Ns::Table, b"covered-table-cell") => Kind::CoveredCell,
            (Ns::Draw, b"frame") => Kind::Frame,
            (Ns::Draw, b"image") => Kind::Image,
            (Ns::Draw, b"text-box") => Kind::TextBox,
            (Ns::Office, b"annotation") => Kind::Annotation,
            _ => return None,
        })
    }

    fn push_node(
        &mut self,
        kind: Kind,
        parent: u32,
        start: u32,
        open_end: u32,
        close_start: u32,
        end: u32,
    ) -> Result<u32> {
        if self.nodes.len() >= MAX_NODES {
            return Err(Error::InvalidFormat(format!(
                "ODT body holds more than {MAX_NODES} indexed elements"
            )));
        }
        let id = self.nodes.len() as u32;
        self.nodes.push(NodeData {
            kind,
            parent,
            first_child: NONE,
            last_child: NONE,
            next_sibling: NONE,
            start,
            open_end,
            close_start,
            end,
        });
        let parent_data = &mut self.nodes[parent as usize];
        let previous = parent_data.last_child;
        parent_data.last_child = id;
        if previous == NONE {
            parent_data.first_child = id;
        } else {
            self.nodes[previous as usize].next_sibling = id;
        }
        Ok(id)
    }

    fn close(&mut self, open: Open, start: u32, end: u32) {
        match open {
            Open::Node(id) => {
                let data = &mut self.nodes[id as usize];
                data.close_start = start;
                data.end = end;
                if self.segment.is_some_and(|(_, owner)| owner == id) {
                    self.segment = None;
                }
            },
            Open::Body => {
                let data = &mut self.nodes[0];
                data.close_start = start;
                data.end = end;
                self.segment = None;
            },
            Open::Space => {
                if let Some((segment, _)) = self.segment {
                    let data = &mut self.nodes[segment as usize];
                    data.end = end;
                    data.close_start = end;
                }
            },
            Open::AutomaticStyles => {
                self.automatic_insert = Some(AutomaticInsert::Before(start));
            },
            Open::Style(index) => {
                self.styles[index as usize].end = end;
            },
            Open::Plain => {
                if self.in_body() {
                    self.segment = None;
                }
            },
            Open::Excluded => {},
        }
    }
}

/// Subtrees whose content is not part of the visible body.
fn excluded(ns: Ns, local: &[u8]) -> bool {
    matches!(
        (ns, local),
        (
            Ns::Text,
            b"tracked-changes"
                | b"sequence-decls"
                | b"variable-decls"
                | b"user-field-decls"
                | b"dde-connection-decls"
                | b"note-citation"
        ) | (Ns::Office, b"forms" | b"binary-data")
            | (Ns::Draw, b"object" | b"object-ole")
            | (Ns::Svg, b"title" | b"desc")
            | (Ns::Dc, _)
            | (Ns::Meta, b"date-string")
    )
}

/// Index the `style:style` declarations of `office:styles` in `styles.xml`.
fn common_styles(xml: &str, prefixes: &Prefixes, out: &mut Vec<StyleData>) -> Result<()> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut in_styles = false;
    let mut current: Option<usize> = None;
    loop {
        let start = position(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("ODT styles.xml: {error}")))?;
        let end = position(&reader)?;
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                let empty = matches!(xml.as_bytes().get(end as usize - 2), Some(b'/'));
                let name = tag.name();
                let resolved = prefixes.resolve(name.as_ref());
                match resolved {
                    Some((Ns::Office, b"styles")) if depth == 1 && !empty => in_styles = true,
                    Some((Ns::Style, b"style")) if in_styles && depth == 2 => {
                        out.push(StyleData {
                            part: StylePart::Common,
                            start,
                            open_end: end,
                            end,
                            text: None,
                            paragraph: None,
                            cell: None,
                        });
                        if !empty {
                            current = Some(out.len() - 1);
                        }
                    },
                    Some((Ns::Style, local)) if current.is_some() && depth == 3 => {
                        if let Some(index) = current {
                            match local {
                                b"text-properties" => out[index].text = Some(start..end),
                                b"paragraph-properties" => {
                                    out[index].paragraph = Some(start..end);
                                },
                                b"table-cell-properties" => out[index].cell = Some(start..end),
                                _ => {},
                            }
                        }
                    },
                    _ => {},
                }
                if !empty {
                    depth += 1;
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 2 {
                    if let Some(index) = current.take() {
                        out[index].end = end;
                    }
                }
                if depth == 1 {
                    in_styles = false;
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(())
}
