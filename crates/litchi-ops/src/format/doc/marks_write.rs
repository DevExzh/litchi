//! Unified writes of the DOC marks of record 0324: hyperlinks (`set` of
//! `url` and `text`, `add`, `remove`), bookmarks (`add`, `remove`, renaming
//! through `set name`), tab stops (`set`, `add`, `remove`) and section page
//! setup (`set`).
//!
//! Every call stages into one native [`litchi_doc::Edit`] and publishes one
//! commit, as the paragraph and run writes do. A value a target already has
//! is not staged, so a call that changes nothing returns the source handle.

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_doc::body_text::{FieldPart, HyperlinkInsert, SectionValue, TextTarget};
use litchi_doc::parts::pap::{TabAlignment, TabLeader, TabStop};

use super::node::Item;
use super::spec::Kind;
use super::{Doc, Nav, Node, WriteCommit};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::placement::Placement;
use crate::props::{PropValue, Props};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::{LengthSpec, Scalar, ValueType};

fn unsupported(message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message).with_format(FileFormat::Doc)
}

fn invalid(message: &'static str, key: &str) -> Error {
    Error::new(Code::InvalidValue, message)
        .with_property(key.to_owned())
        .with_format(FileFormat::Doc)
}

fn text<'a>(value: &'a PropValue<'_>, key: &str) -> Result<&'a str> {
    value
        .as_text()
        .ok_or_else(|| invalid("the DOC property requires a text value", key))
}

fn int(value: &PropValue<'_>, key: &str) -> Result<i64> {
    value
        .with_coerced(ValueType::Int, |value| match value {
            Some(Scalar::Int(value)) => Ok(value),
            _ => Err(invalid("the DOC property requires an integer", key)),
        })
        .map_err(|error| {
            error
                .with_property(key.to_owned())
                .with_format(FileFormat::Doc)
        })
}

fn boolean(value: &PropValue<'_>, key: &str) -> Result<bool> {
    value
        .with_coerced(ValueType::Bool, |value| match value {
            Some(Scalar::Bool(value)) => Ok(value),
            _ => Err(invalid("the DOC property requires a boolean", key)),
        })
        .map_err(|error| {
            error
                .with_property(key.to_owned())
                .with_format(FileFormat::Doc)
        })
}

fn token(value: &PropValue<'_>, ty: ValueType, key: &str) -> Result<&'static str> {
    value
        .with_coerced(ty, |value| match value {
            Some(Scalar::Token(value)) => Ok(value),
            _ => Err(invalid("the DOC property requires one of its tokens", key)),
        })
        .map_err(|error| {
            error
                .with_property(key.to_owned())
                .with_format(FileFormat::Doc)
        })
}

/// A length in twips, rounded to the nearest twip.
fn twips(value: &PropValue<'_>, spec: LengthSpec, key: &str) -> Result<i64> {
    value
        .with_coerced(ValueType::Length(spec), |value| match value {
            Some(Scalar::Length(length)) => {
                let emu = length.emu();
                Ok((emu + emu.signum() * 317) / 635)
            },
            _ => Err(invalid("the DOC property requires a length", key)),
        })
        .map_err(|error| {
            error
                .with_property(key.to_owned())
                .with_format(FileFormat::Doc)
        })
}

fn unsigned_twips(value: &PropValue<'_>, key: &str) -> Result<u16> {
    u16::try_from(twips(value, super::spec::EXTENT, key)?)
        .ok()
        .filter(|twips| *twips <= 31_680)
        .ok_or_else(|| invalid("the DOC length is outside 0 through 22 inches", key))
}

fn signed_twips(value: &PropValue<'_>, key: &str) -> Result<i16> {
    i16::try_from(twips(value, super::spec::SIGNED_POINTS, key)?)
        .ok()
        .filter(|twips| (-31_680..=31_680).contains(twips))
        .ok_or_else(|| invalid("the DOC length is outside ±22 inches", key))
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(TransactionError::Limit {
            kind,
            observed,
            maximum,
        }
        .into());
    }
    Ok(())
}

/// A `HYPERLINK` instruction quoting `value`, which must hold no quote.
fn quoted(value: &str, key: &str) -> Result<String> {
    if value.is_empty() || value.contains('"') || value.chars().any(char::is_control) {
        return Err(invalid(
            "DOC hyperlink targets are non-empty and hold no quotes or control characters",
            key,
        ));
    }
    Ok(format!("\"{value}\""))
}

/// `instruction` with its target argument (the first argument after the
/// keyword) replaced by `url`, or the target inserted after the keyword.
fn with_url(instruction: &str, url: &str) -> Result<String> {
    let url = quoted(url, "url")?;
    let leading = instruction.len() - instruction.trim_start().len();
    let body = &instruction[leading..];
    let keyword_end = body.find(char::is_whitespace).unwrap_or(body.len());
    let after_keyword = &body[keyword_end..];
    let gap = after_keyword.len() - after_keyword.trim_start().len();
    let rest = &after_keyword[gap..];
    let prefix = &instruction[..leading + keyword_end + gap];
    if let Some(quoted_rest) = rest.strip_prefix('"') {
        let end = quoted_rest
            .find('"')
            .ok_or_else(|| invalid("the DOC hyperlink instruction is malformed", "url"))?;
        return Ok(format!("{prefix}{url}{}", &quoted_rest[end + 1..]));
    }
    if !rest.is_empty() && !rest.starts_with('\\') {
        let end = rest.find(char::is_whitespace).unwrap_or(rest.len());
        return Ok(format!("{prefix}{url}{}", &rest[end..]));
    }
    let separator = if gap == 0 { " " } else { "" };
    Ok(format!("{prefix}{separator}{url} {rest}"))
}

/// A tab stop's alignment for a `val` token.
fn tab_alignment(value: &str) -> TabAlignment {
    match value {
        "center" => TabAlignment::Center,
        "right" => TabAlignment::Right,
        "decimal" => TabAlignment::Decimal,
        "bar" => TabAlignment::Bar,
        "num" => TabAlignment::List,
        _ => TabAlignment::Left,
    }
}

/// A tab stop's leader for a `leader` token.
fn tab_leader(value: &str) -> TabLeader {
    match value {
        "dot" => TabLeader::Dots,
        "hyphen" => TabLeader::Hyphens,
        "underscore" => TabLeader::Underline,
        "heavy" => TabLeader::Heavy,
        "middleDot" => TabLeader::MiddleDot,
        _ => TabLeader::None,
    }
}

/// Apply `props` to `stop`.
fn tab_props(stop: &mut TabStop, props: &Props<'_>) -> Result<()> {
    for (key, value) in props.iter() {
        match key.as_str() {
            "pos" => {
                stop.position = i32::try_from(int(value, "pos")?)
                    .ok()
                    .filter(|position| (-31_680..=31_680).contains(position))
                    .ok_or_else(|| invalid("a DOC tab position is within ±31680 twips", "pos"))?;
            },
            "val" => {
                stop.alignment = tab_alignment(token(
                    value,
                    ValueType::Token(&super::spec::TAB_VAL),
                    "val",
                )?);
            },
            "leader" => {
                stop.leader = tab_leader(token(
                    value,
                    ValueType::Token(&super::spec::TAB_LEADER),
                    "leader",
                )?);
            },
            other => {
                return Err(unsupported("the DOC tab property is not writable")
                    .with_property(other.to_owned()));
            },
        }
    }
    Ok(())
}

/// The section values `props` set.
fn section_values(props: &Props<'_>) -> Result<Vec<SectionValue>> {
    use litchi_doc::section::{BreakKind, VerticalJustification};
    let mut values = Vec::with_capacity(props.len());
    for (key, value) in props.iter() {
        let key = key.as_str();
        values.push(match key {
            "type" => SectionValue::Break(
                match token(value, ValueType::Token(&super::spec::SECTION_TYPE), key)? {
                    "continuous" => BreakKind::Continuous,
                    "nextColumn" => BreakKind::NewColumn,
                    "evenPage" => BreakKind::EvenPage,
                    "oddPage" => BreakKind::OddPage,
                    _ => BreakKind::NewPage,
                },
            ),
            "pageWidth" => SectionValue::PageWidth(unsigned_twips(value, key)?),
            "pageHeight" => SectionValue::PageHeight(unsigned_twips(value, key)?),
            "orientation" => SectionValue::Landscape(
                token(value, ValueType::Token(&super::spec::ORIENTATION), key)? == "landscape",
            ),
            "marginTop" => SectionValue::MarginTop(signed_twips(value, key)?),
            "marginBottom" => SectionValue::MarginBottom(signed_twips(value, key)?),
            "marginLeft" => SectionValue::MarginLeft(unsigned_twips(value, key)?),
            "marginRight" => SectionValue::MarginRight(unsigned_twips(value, key)?),
            "marginHeader" => SectionValue::HeaderDistance(unsigned_twips(value, key)?),
            "marginFooter" => SectionValue::FooterDistance(unsigned_twips(value, key)?),
            "marginGutter" => SectionValue::Gutter(unsigned_twips(value, key)?),
            "columns" => SectionValue::Columns(
                u16::try_from(int(value, key)?)
                    .ok()
                    .filter(|count| (1..=45).contains(count))
                    .ok_or_else(|| invalid("a DOC section has 1 through 45 columns", key))?,
            ),
            "columnSpace" => SectionValue::ColumnSpace(unsigned_twips(value, key)?),
            "titlePage" => SectionValue::TitlePage(boolean(value, key)?),
            "vAlign" => SectionValue::VerticalAlignment(
                match token(value, ValueType::Token(&super::spec::SECTION_VALIGN), key)? {
                    "center" => VerticalJustification::Center,
                    "both" => VerticalJustification::Justified,
                    "bottom" => VerticalJustification::Bottom,
                    _ => VerticalJustification::Top,
                },
            ),
            other => {
                return Err(unsupported("the DOC section property is not writable")
                    .with_property(other.to_owned()));
            },
        });
    }
    Ok(values)
}

/// The path of the `index`th child of `kind` under body paragraph
/// `paragraph`.
fn paragraph_child_path(paragraph: usize, kind: &str, index: usize) -> Path<'static> {
    Path::from_steps([
        Step::new("body"),
        Step::new("p").with(Predicate::Index(Position::new(paragraph))),
        Step::new(kind.to_owned()).with(Predicate::Index(Position::new(index))),
    ])
}

/// Whether the last step of `path` may name a hyperlink, bookmark, tab stop
/// or section.
fn names_mark(path: &crate::path::PathArg<'_>) -> bool {
    path.path().last().is_some_and(|step| {
        ["hyperlink", "link", "bookmark", "tab", "section"]
            .iter()
            .any(|name| step.name().eq_ignore_ascii_case(name))
    })
}

fn top_path(kind: &str, index: usize) -> Path<'static> {
    Path::from_steps([Step::new(kind.to_owned()).with(Predicate::Index(Position::new(index)))])
}

impl Doc {
    /// `set` on a hyperlink, bookmark, tab stop or section path; `None` for
    /// every other target, which the paragraph and run writer handles.
    pub(super) fn set_mark(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<WriteCommit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        // Only a path whose last step names a mark can resolve to one; every
        // other path keeps the paragraph and run writer's single resolution.
        if !names_mark(path) {
            return Ok(None);
        }
        let nav = Nav::new(self);
        let node = crate::path::resolve::get(&nav, path)?;
        let kind = node.element_kind();
        if !matches!(
            kind,
            Kind::Hyperlink | Kind::Bookmark | Kind::Tab | Kind::Section
        ) {
            return Ok(None);
        }
        if props.is_empty() {
            return Err(Error::new(
                Code::InvalidArgument,
                "a DOC write needs at least one property",
            )
            .with_format(FileFormat::Doc));
        }
        limit(LimitKind::Operations, props.len(), limits.max_operations())?;
        let mut edit = self.native_edit(limits)?;
        let mut operations = 0usize;
        match (kind, node.item()) {
            (Kind::Hyperlink, Item::Hyperlink(link)) => {
                let start = link.field.field.start_cp;
                for (key, value) in props.iter() {
                    match key.as_str() {
                        "url" => {
                            let url = text(value, "url")?;
                            let instruction = with_url(&link.field.instruction, url)?;
                            limit(
                                LimitKind::StagedBytes,
                                instruction.len(),
                                limits.max_staged_bytes(),
                            )?;
                            edit.replace_text(
                                TextTarget::Field {
                                    start,
                                    part: FieldPart::Instruction,
                                },
                                &instruction,
                            )?;
                        },
                        "text" => {
                            let replacement = text(value, "text")?;
                            limit(
                                LimitKind::StagedBytes,
                                replacement.len(),
                                limits.max_staged_bytes(),
                            )?;
                            edit.replace_text(
                                TextTarget::Field {
                                    start,
                                    part: FieldPart::Result,
                                },
                                replacement,
                            )?;
                        },
                        other => {
                            return Err(unsupported(
                                "DOC hyperlinks set `url` and `text`; their formatting lives on their runs",
                            )
                            .with_property(other.to_owned()));
                        },
                    }
                    operations += 1;
                }
            },
            (Kind::Bookmark, Item::Bookmark(bookmark)) => {
                for (key, value) in props.iter() {
                    if key.as_str() != "name" {
                        return Err(unsupported("DOC bookmarks set `name`")
                            .with_property(key.as_str().to_owned()));
                    }
                    let name = text(value, "name")?;
                    edit.rename_bookmark(&bookmark.name, name)?;
                    operations += 1;
                }
            },
            (Kind::Tab, Item::Tab(stop)) => {
                let Some(paragraph) = node.body_paragraph_of() else {
                    return Err(unsupported(
                        "DOC tab stops are written on ordinary main-story paragraphs",
                    ));
                };
                let mut stops = edit.paragraph_tabs(paragraph)?;
                let current = stops
                    .iter_mut()
                    .find(|existing| existing.position == stop.position)
                    .ok_or_else(|| {
                        Error::new(Code::PathNotFound, "the DOC tab stop is gone")
                            .with_format(FileFormat::Doc)
                    })?;
                tab_props(current, props)?;
                edit.set_paragraph_tabs(paragraph, &stops)?;
                operations += 1;
            },
            (Kind::Section, _) => {
                let section = node.own_position().unwrap_or(Position::new(0));
                let values = section_values(props)?;
                edit.set_section_values(section, &values)?;
                operations += values.len();
            },
            _ => return Ok(None),
        }
        self.publish(edit, operations, limits).map(Some)
    }

    /// `add` of a hyperlink or tab stop under a body paragraph, or of a
    /// bookmark under a body paragraph or the body; `None` for every other
    /// kind, which the paragraph writer handles.
    pub(super) fn add_mark(
        &self,
        parent: &Node<'_>,
        node: &NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<(WriteCommit, Path<'static>)>> {
        let kind = node.kind();
        let kind = if kind.eq_ignore_ascii_case("hyperlink") || kind.eq_ignore_ascii_case("link") {
            Kind::Hyperlink
        } else if kind.eq_ignore_ascii_case("bookmark") {
            Kind::Bookmark
        } else if kind.eq_ignore_ascii_case("tab") || kind.eq_ignore_ascii_case("tabstop") {
            Kind::Tab
        } else {
            return Ok(None);
        };
        limit(LimitKind::Operations, 1, limits.max_operations())?;
        let props = node.props();
        match kind {
            Kind::Hyperlink => self.add_hyperlink(parent, props, node.placement(), limits),
            Kind::Bookmark => self.add_bookmark(parent, props, limits),
            _ => self.add_tab(parent, props, limits),
        }
        .map(Some)
    }

    fn add_hyperlink(
        &self,
        parent: &Node<'_>,
        props: &Props<'_>,
        placement: Option<&Placement<'_>>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        let Some(position) = parent.paragraph_position() else {
            return Err(Error::new(
                Code::UnsupportedType,
                "DOC hyperlinks are added to ordinary main-story paragraphs",
            )
            .with_format(FileFormat::Doc));
        };
        let (mut url, mut anchor, mut tooltip, mut display) = (None, None, None, None);
        for (key, value) in props.iter() {
            match key.as_str() {
                "url" | "href" | "link" => url = Some(text(value, "url")?),
                "anchor" | "bookmark" => anchor = Some(text(value, "anchor")?),
                "tooltip" => tooltip = Some(text(value, "tooltip")?),
                "text" => display = Some(text(value, "text")?),
                other => {
                    return Err(unsupported(
                        "DOC hyperlink additions accept `url`, `anchor`, `tooltip` and `text`",
                    )
                    .with_property(other.to_owned()));
                },
            }
        }
        let mut instruction = String::from("HYPERLINK");
        match (url, anchor) {
            (Some(url), None) => {
                instruction.push(' ');
                instruction.push_str(&quoted(url, "url")?);
            },
            (None, Some(anchor)) => {
                instruction.push_str(" \\l ");
                instruction.push_str(&quoted(anchor, "anchor")?);
            },
            _ => {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "a DOC hyperlink names exactly one of `url` and `anchor`",
                )
                .with_format(FileFormat::Doc));
            },
        }
        if let Some(tooltip) = tooltip {
            instruction.push_str(" \\o ");
            instruction.push_str(&quoted(tooltip, "tooltip")?);
        }
        let text = display.or(url).or(anchor).unwrap_or_default().to_owned();
        limit(
            LimitKind::StagedBytes,
            text.len() + instruction.len(),
            limits.max_staged_bytes(),
        )?;
        let paragraph = parent.native_paragraph()?.ok_or_else(|| {
            Error::new(Code::PathNotFound, "the DOC paragraph is gone").with_format(FileFormat::Doc)
        })?;
        let length = u32::try_from(
            parent
                .paragraph_text()?
                .map_or(0, |text| text.encode_utf16().count()),
        )
        .unwrap_or(u32::MAX);
        let start = paragraph
            .content_cp_range()
            .map_or(0, |range| range.start());
        let nav = Nav::new(self);
        let run_offset = |path: &crate::path::PathArg<'_>, end: bool| -> Result<u32> {
            let run = crate::path::resolve::get(&nav, path)?;
            match run.body_run() {
                Some((at, from, to)) if at == position => Ok(if end { to } else { from }),
                _ => Err(Error::new(
                    Code::UnsupportedType,
                    "a DOC hyperlink is placed before or after a run of its paragraph",
                )
                .with_format(FileFormat::Doc)),
            }
        };
        let offset = match placement {
            None | Some(Placement::End) => length,
            Some(Placement::Start) => 0,
            Some(Placement::Before(path)) => run_offset(path, false)?,
            Some(Placement::After(path)) => run_offset(path, true)?,
            Some(Placement::Position(_) | Placement::Into(..)) => {
                return Err(unsupported(
                    "DOC hyperlinks are placed at the start or end of a paragraph, or beside one of its runs",
                ));
            },
        };
        // The new link's ordinal among the paragraph's links.
        let at = start.saturating_add(offset);
        let index = super::marks::links(self, paragraph)?
            .iter()
            .filter(|link| link.field.field.start_cp < at)
            .count();
        let mut edit = self.native_edit(limits)?;
        edit.insert_hyperlink(position, offset, &HyperlinkInsert { instruction, text })?;
        Ok((
            self.publish(edit, 1, limits)?,
            paragraph_child_path(position.get(), "hyperlink", index),
        ))
    }

    fn add_bookmark(
        &self,
        parent: &Node<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        let (mut name, mut covered) = (None, None);
        for (key, value) in props.iter() {
            match key.as_str() {
                "name" => name = Some(text(value, "name")?),
                "text" => covered = Some(text(value, "text")?),
                other => {
                    return Err(
                        unsupported("DOC bookmark additions accept `name` and `text`")
                            .with_property(other.to_owned()),
                    );
                },
            }
        }
        let name = name.ok_or_else(|| {
            Error::new(Code::InvalidArgument, "a DOC bookmark needs a `name`")
                .with_format(FileFormat::Doc)
        })?;
        let mut edit = self.native_edit(limits)?;
        let mut operations = 1;
        match (parent.element_kind(), parent.paragraph_position(), covered) {
            // Under the body (or the document): a new last paragraph holding
            // the covered text.
            (Kind::Body | Kind::Document, _, text) => {
                let text = text.unwrap_or_default();
                limit(
                    LimitKind::StagedBytes,
                    text.len(),
                    limits.max_staged_bytes(),
                )?;
                let position = Position::new(self.body_paragraph_count()?);
                edit.insert_paragraph(position, text)?;
                let units = u32::try_from(text.encode_utf16().count()).unwrap_or(u32::MAX);
                edit.add_bookmark(name, position, 0, units)?;
                operations += 1;
            },
            // Under a paragraph: over the first occurrence of `text`, or an
            // empty bookmark at the paragraph's end.
            (Kind::Paragraph, Some(position), covered) => {
                let paragraph = parent.paragraph_text()?.unwrap_or_default().into_owned();
                let (start, end) = match covered {
                    Some(covered) => {
                        let byte = paragraph.find(covered).ok_or_else(|| {
                            invalid("the DOC paragraph does not hold the bookmark text", "text")
                        })?;
                        let start = paragraph[..byte].encode_utf16().count();
                        (start, start + covered.encode_utf16().count())
                    },
                    None => {
                        let end = paragraph.encode_utf16().count();
                        (end, end)
                    },
                };
                edit.add_bookmark(
                    name,
                    position,
                    u32::try_from(start).unwrap_or(u32::MAX),
                    u32::try_from(end).unwrap_or(u32::MAX),
                )?;
            },
            _ => {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "DOC bookmarks are added under the body or an ordinary main-story paragraph",
                )
                .with_format(FileFormat::Doc));
            },
        }
        let index = edit
            .bookmarks()?
            .iter()
            .position(|(existing, ..)| existing == name)
            .unwrap_or(0);
        Ok((
            self.publish(edit, operations, limits)?,
            top_path("bookmark", index),
        ))
    }

    fn add_tab(
        &self,
        parent: &Node<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<(WriteCommit, Path<'static>)> {
        let Some(position) = parent.paragraph_position() else {
            return Err(Error::new(
                Code::UnsupportedType,
                "DOC tab stops are added to ordinary main-story paragraphs",
            )
            .with_format(FileFormat::Doc));
        };
        if !props.iter().any(|(key, _)| key.as_str() == "pos") {
            return Err(
                Error::new(Code::InvalidArgument, "a DOC tab stop needs a `pos`")
                    .with_format(FileFormat::Doc),
            );
        }
        let mut stop = TabStop {
            position: 0,
            alignment: TabAlignment::Left,
            leader: TabLeader::None,
        };
        tab_props(&mut stop, props)?;
        let mut edit = self.native_edit(limits)?;
        let mut stops = edit.paragraph_tabs(position)?;
        if stops
            .iter()
            .any(|existing| existing.position == stop.position)
        {
            return Err(Error::new(
                Code::DuplicateName,
                "the DOC paragraph already has a tab stop at that position",
            )
            .with_format(FileFormat::Doc));
        }
        stops.push(stop);
        stops.sort_by_key(|stop| stop.position);
        let index = stops
            .iter()
            .position(|existing| existing.position == stop.position)
            .unwrap_or(0);
        edit.set_paragraph_tabs(position, &stops)?;
        Ok((
            self.publish(edit, 1, limits)?,
            paragraph_child_path(position.get(), "tab", index),
        ))
    }

    /// `remove` of a hyperlink (the whole field), a bookmark (its text
    /// stays) or a tab stop; `None` for every other target.
    pub(super) fn remove_mark(
        &self,
        target: &TargetArg<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<WriteCommit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !names_mark(path) {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let mut edit = match node.element_kind() {
            Kind::Hyperlink | Kind::Bookmark | Kind::Tab => self.native_edit(limits)?,
            _ => return Ok(None),
        };
        limit(LimitKind::Operations, 1, limits.max_operations())?;
        match node.item() {
            Item::Hyperlink(link) => edit.remove_field(link.field.field.start_cp)?,
            Item::Bookmark(bookmark) => edit.remove_bookmark(&bookmark.name)?,
            Item::Tab(stop) => {
                let Some(paragraph) = node.body_paragraph_of() else {
                    return Err(unsupported(
                        "DOC tab stops are written on ordinary main-story paragraphs",
                    ));
                };
                let mut stops = edit.paragraph_tabs(paragraph)?;
                stops.retain(|existing| existing.position != stop.position);
                edit.set_paragraph_tabs(paragraph, &stops)?;
            },
            _ => return Ok(None),
        }
        self.publish(edit, 1, limits).map(Some)
    }
}

#[cfg(test)]
mod tests {
    use super::with_url;

    #[test]
    fn a_url_replaces_the_target_and_keeps_the_switches() {
        assert_eq!(
            with_url(r#" HYPERLINK "http://old/" \o "tip" "#, "http://new/").unwrap(),
            r#" HYPERLINK "http://new/" \o "tip" "#
        );
        assert_eq!(
            with_url(" HYPERLINK bare.htm ", "x.htm").unwrap(),
            r#" HYPERLINK "x.htm" "#
        );
        assert_eq!(
            with_url(r#"HYPERLINK \l "top""#, "http://a/").unwrap(),
            r#"HYPERLINK "http://a/" \l "top""#
        );
        assert!(with_url("HYPERLINK \"a\"", "has\"quote").is_err());
    }
}
