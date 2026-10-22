//! Unified writes over the leaf ODT body edits.
//!
//! Every verb resolves its targets against one immutable handle, translates
//! the unified properties into typed leaf body edits
//! (`litchi_odt::transaction::Edit::{set_text, set_attribute,
//! set_text_properties, set_paragraph_properties, insert, remove, relocate,
//! swap, set_metadata}`), and commits them in one leaf transaction: one
//! audited splice publication of the exact source parts, so bytes the edit
//! does not touch keep their source spelling and an exact no-op returns the
//! source package itself.

use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_odt::body::{
    Attribute, Kind as LeafKind, Location, NewNode as LeafNew, Node as LeafNode, NodeId,
    ParagraphPropertyUpdate, Placement as LeafPlacement, TextPropertyUpdate,
};
use litchi_odt::metadata::Field as MetadataField;
use litchi_odt::transaction::Edit;

/// The leaf proof of one ODT write: the primary or the metadata family's.
pub(crate) enum Patch {
    Primary(litchi_odt::transaction::Patch),
    Metadata(litchi_odt::metadata::Patch),
}

impl Patch {
    pub(crate) fn is_noop(&self) -> bool {
        match self {
            Self::Primary(patch) => patch.is_noop(),
            Self::Metadata(patch) => patch.is_noop(),
        }
    }
}

use super::spec::{self, Kind};
use super::{Nav, Node, Odt, core_error, odf_error};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
use crate::placement::Placement;
use crate::props::{IntoProps, PropValue, Props};
use crate::target::{IntoTarget, TargetArg};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::{Base, Color, Length, Scalar, ValueType};

pub(crate) mod coalesce;
pub(crate) mod find;
pub(crate) mod refresh;

/// The leaf proof and immutable endpoints of one ODT write.
pub(crate) struct Commit {
    pub(crate) before: Odt,
    pub(crate) after: Odt,
    pub(crate) patch: Patch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

fn refusal(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Odt)
}

/// A leaf staging refusal: the caller's request, not the file, is at fault.
pub(crate) fn staging_error(error: litchi_odf_common::Error) -> Error {
    match error {
        litchi_odf_common::Error::Core(error) => {
            core_error(error).caused_by_argument(Code::InvalidValue)
        },
        litchi_odf_common::Error::Refused(_) => {
            odf_error(error).caused_by_argument(Code::InvalidValue)
        },
        other => odf_error(other),
    }
}

/// A text set's staging error: the refusal of content a text set can
/// neither keep nor replace (record 0304) stays `unsupported_capability`.
fn text_error(error: litchi_odf_common::Error) -> Error {
    match error {
        litchi_odf_common::Error::Core(
            litchi_core::Error::Unsupported(_) | litchi_core::Error::Refused(_),
        )
        | litchi_odf_common::Error::Refused(_) => odf_error(error),
        other => staging_error(other),
    }
}

pub(crate) fn bound(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
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

/// Commit `edit` and adopt its package, charging the retained endpoint.
pub(crate) fn finish(
    file: &Odt,
    edit: Edit,
    operations: usize,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let native = edit.commit().map_err(odf_error)?;
    let changed = native.diagnostics().changed();
    let (package, patch) = native.into_parts();
    let snapshot_bytes = if changed { package.as_bytes().len() } else { 0 };
    bound(
        LimitKind::SnapshotBytes,
        snapshot_bytes,
        limits.max_snapshot_bytes(),
    )?;
    let after = if changed {
        file.adopt(package)?
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok(Commit {
        before: file.clone(),
        after,
        patch: Patch::Primary(patch),
        operations,
        snapshot_bytes,
    })
}

/// An ODF length: points with up to four decimals (`12pt`, `10.5pt`).
pub(crate) fn odf_length(length: Length) -> String {
    #[allow(
        clippy::cast_precision_loss,
        reason = "EMU values of document lengths are far below 2^52"
    )]
    let points = length.emu() as f64 / 12_700.0;
    let mut text = format!("{points:.4}");
    while text.ends_with('0') {
        text.pop();
    }
    if text.ends_with('.') {
        text.pop();
    }
    if text == "-0" {
        "0".clone_into(&mut text);
    }
    text.push_str("pt");
    text
}

/// An ODF `#rrggbb` colour from an opaque sRGB colour.
fn odf_color(color: Color, key: &'static str) -> Result<String> {
    match color.base() {
        Base::Rgb(rgba) if rgba.is_opaque() && color.transforms().is_empty() => Ok(format!(
            "#{:02x}{:02x}{:02x}",
            rgba.red(),
            rgba.green(),
            rgba.blue()
        )),
        _ => Err(
            refusal(Code::InvalidValue, "ODT colours are opaque #RRGGBB values").with_property(key),
        ),
    }
}

/// The edits one set call stages on one node.
#[derive(Default)]
pub(crate) struct Staged {
    text: Option<String>,
    text_properties: TextPropertyUpdate,
    paragraph_properties: ParagraphPropertyUpdate,
    attributes: Vec<(Attribute, Option<String>)>,
}

impl Staged {
    fn operations(&self) -> usize {
        usize::from(self.text.is_some())
            + usize::from(!self.text_properties.is_empty())
            + usize::from(!self.paragraph_properties.is_empty())
            + self.attributes.len()
    }

    /// Whether the staged edits replace the node's content.
    pub(crate) const fn replaces_content(&self) -> bool {
        self.text.is_some()
    }

    fn bytes(&self) -> usize {
        self.text.as_ref().map_or(0, String::len)
            + self
                .attributes
                .iter()
                .map(|(_, value)| value.as_ref().map_or(0, String::len))
                .sum::<usize>()
            + size_of::<Self>()
    }
}

/// An owned coerced value.
enum Coerced {
    Bool(bool),
    Int(i64),
    Text(String),
    Token(&'static str),
    Length(Length),
    Color(Color),
    Other,
}

fn coerce(value: &PropValue<'_>, ty: ValueType, key: &'static str) -> Result<Option<Coerced>> {
    value
        .with_coerced(ty, |scalar| {
            Ok(scalar.map(|scalar| match scalar {
                Scalar::Bool(value) => Coerced::Bool(value),
                Scalar::Int(value) => Coerced::Int(value),
                Scalar::Text(text) => Coerced::Text(text.to_owned()),
                Scalar::Token(token) => Coerced::Token(token),
                Scalar::Length(length) => Coerced::Length(length),
                Scalar::Color(color) => Coerced::Color(color),
                Scalar::Number(_) | Scalar::List(_) => Coerced::Other,
            }))
        })
        .map_err(|error| error.with_property(key).with_format(FileFormat::Odt))
}

fn coerce_text(value: &PropValue<'_>, key: &'static str) -> Result<Option<String>> {
    match coerce(value, ValueType::Text, key)? {
        None => Ok(None),
        Some(Coerced::Text(text)) => Ok(Some(text)),
        Some(_) => Err(wrong(key)),
    }
}

fn required<T>(value: Option<T>, key: &'static str) -> Result<T> {
    value.ok_or_else(|| {
        refusal(
            Code::InvalidValue,
            "clearing this ODT property is not supported",
        )
        .with_property(key)
    })
}

fn wrong(key: &'static str) -> Error {
    refusal(Code::InvalidValue, "the property value has the wrong type").with_property(key)
}

fn bool_of(coerced: Option<Coerced>, key: &'static str) -> Result<bool> {
    match required(coerced, key)? {
        Coerced::Bool(value) => Ok(value),
        _ => Err(wrong(key)),
    }
}

fn length_of(coerced: Option<Coerced>, key: &'static str) -> Result<String> {
    match required(coerced, key)? {
        Coerced::Length(length) => Ok(odf_length(length)),
        _ => Err(wrong(key)),
    }
}

fn color_of(coerced: Option<Coerced>, key: &'static str) -> Result<String> {
    match required(coerced, key)? {
        Coerced::Color(color) => odf_color(color, key),
        _ => Err(wrong(key)),
    }
}

fn text_of(coerced: Option<Coerced>, key: &'static str) -> Result<Option<String>> {
    match coerced {
        None => Ok(None),
        Some(Coerced::Text(text)) => Ok(Some(text)),
        Some(Coerced::Token(token)) => Ok(Some(token.to_owned())),
        Some(Coerced::Int(value)) => Ok(Some(value.to_string())),
        Some(_) => Err(wrong(key)),
    }
}

fn check_style(body: litchi_odt::body::Body<'_>, style: &str, key: &'static str) -> Result<()> {
    if body.common_style(style).is_none() && body.automatic_style(style).is_none() {
        return Err(refusal(Code::InvalidValue, "the ODT style does not exist")
            .with_property(key)
            .with_suggestion("name a style listed under /styles"));
    }
    Ok(())
}

/// Translate one set call's properties for a node of `kind`.
#[allow(
    clippy::too_many_lines,
    reason = "one flat key table, kept in one place so it reads against the spec tables"
)]
fn translate(kind: Kind, props: &Props<'_>, body: litchi_odt::body::Body<'_>) -> Result<Staged> {
    let mut staged = Staged::default();
    let specs = spec::SPEC.elements[kind as usize].props;
    for (key, value) in props.iter() {
        let name = key.as_str();
        let Some(spec) = specs
            .iter()
            .find(|spec| spec.name.eq_ignore_ascii_case(name))
        else {
            return Err(refusal(
                Code::UnsupportedProperty,
                "the property is not an ODT property of this element",
            )
            .with_property(name.to_owned()));
        };
        if !spec.access.set {
            return Err(
                refusal(Code::UnsupportedCapability, "the ODT property is read-only")
                    .with_property(spec.name),
            );
        }
        let key = spec.name;
        let coerced = coerce(value, spec.ty, key)?;
        let paragraph = &mut staged.paragraph_properties;
        let character = &mut staged.text_properties;
        match key {
            "text" => staged.text = Some(required(text_of(coerced, key)?, key)?),
            "style" => {
                let value = text_of(coerced, key)?;
                if let Some(style) = &value {
                    check_style(body, style, key)?;
                }
                staged.attributes.push((Attribute::StyleName, value));
            },
            "align" => paragraph.align = Some(required(text_of(coerced, key)?, key)?),
            "spaceBefore" => paragraph.space_before = Some(length_of(coerced, key)?),
            "spaceAfter" => paragraph.space_after = Some(length_of(coerced, key)?),
            "indent" => paragraph.indent = Some(length_of(coerced, key)?),
            "rightIndent" => paragraph.right_indent = Some(length_of(coerced, key)?),
            "firstLineIndent" => paragraph.first_line_indent = Some(length_of(coerced, key)?),
            "lineSpacing" => paragraph.line_height = Some(required(text_of(coerced, key)?, key)?),
            "keepNext" => paragraph.keep_next = Some(bool_of(coerced, key)?),
            "pageBreakBefore" => paragraph.page_break_before = Some(bool_of(coerced, key)?),
            "shading" => paragraph.shading = Some(color_of(coerced, key)?),
            "bold" => character.bold = Some(bool_of(coerced, key)?),
            "italic" => character.italic = Some(bool_of(coerced, key)?),
            "underline" => character.underline = Some(bool_of(coerced, key)?),
            "strike" => character.strike = Some(bool_of(coerced, key)?),
            "superscript" => character.superscript = Some(bool_of(coerced, key)?),
            "subscript" => character.subscript = Some(bool_of(coerced, key)?),
            "caps" => character.caps = Some(bool_of(coerced, key)?),
            "smallcaps" => character.small_caps = Some(bool_of(coerced, key)?),
            "size" => character.size = Some(length_of(coerced, key)?),
            "color" => character.color = Some(color_of(coerced, key)?),
            "highlight" => character.highlight = Some(color_of(coerced, key)?),
            "font" => character.font = Some(required(text_of(coerced, key)?, key)?),
            "level" => {
                let level = match required(coerced, key)? {
                    Coerced::Int(level) if (1..=10).contains(&level) => level,
                    Coerced::Int(_) => {
                        return Err(refusal(
                            Code::OutOfRange,
                            "ODT heading levels run from 1 to 10",
                        )
                        .with_property(key));
                    },
                    _ => return Err(wrong(key)),
                };
                staged
                    .attributes
                    .push((Attribute::OutlineLevel, Some(level.to_string())));
            },
            "url" => staged.attributes.push((
                Attribute::Href,
                Some(required(text_of(coerced, key)?, key)?),
            )),
            "tooltip" => staged
                .attributes
                .push((Attribute::Title, text_of(coerced, key)?)),
            "tgtFrame" => staged
                .attributes
                .push((Attribute::TargetFrame, text_of(coerced, key)?)),
            "name" => staged.attributes.push((
                Attribute::Name,
                Some(required(text_of(coerced, key)?, key)?),
            )),
            "protected" => {
                let value = bool_of(coerced, key)?;
                staged
                    .attributes
                    .push((Attribute::Protected, Some(value.to_string())));
            },
            "hidden" => {
                let value = bool_of(coerced, key)?;
                staged
                    .attributes
                    .push((Attribute::Display, value.then(|| "none".to_owned())));
            },
            "width" => staged
                .attributes
                .push((Attribute::Width, Some(length_of(coerced, key)?))),
            "height" => staged
                .attributes
                .push((Attribute::Height, Some(length_of(coerced, key)?))),
            "x" => staged
                .attributes
                .push((Attribute::X, Some(length_of(coerced, key)?))),
            "y" => staged
                .attributes
                .push((Attribute::Y, Some(length_of(coerced, key)?))),
            "anchor" => staged.attributes.push((
                Attribute::AnchorType,
                Some(required(text_of(coerced, key)?, key)?),
            )),
            _ => {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "the ODT property has no writer",
                )
                .with_property(key));
            },
        }
    }
    Ok(staged)
}

/// Stage `staged` on node `id`.
fn stage(edit: &mut Edit, id: NodeId, staged: &Staged) -> Result<()> {
    if let Some(text) = &staged.text {
        edit.set_text(id, text.clone()).map_err(text_error)?;
    }
    for (attribute, value) in &staged.attributes {
        edit.set_attribute(id, *attribute, value.as_deref())
            .map_err(staging_error)?;
    }
    if !staged.text_properties.is_empty() {
        edit.set_text_properties(id, &staged.text_properties)
            .map_err(staging_error)?;
    }
    if !staged.paragraph_properties.is_empty() {
        edit.set_paragraph_properties(id, &staged.paragraph_properties)
            .map_err(staging_error)?;
    }
    Ok(())
}

/// The nodes a set or remove targets: one path, or every match of a
/// scoped query.
fn resolve_targets<'f>(
    file: &'f Odt,
    target: &TargetArg<'_>,
    maximum: usize,
) -> Result<Vec<Node<'f>>> {
    let mut nodes = Vec::new();
    if let Some(path) = target.path() {
        nodes.push(crate::path::resolve::get(&Nav::new(file), path)?);
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::new(file), query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(refusal(
                Code::InvalidSelector,
                "the ODT mutation selector contains an unsupported predicate",
            ));
        }
        for node in walk {
            nodes.push(node?);
            bound(LimitKind::ExpandedTargets, nodes.len(), maximum)?;
        }
        if nodes.is_empty() {
            return Err(refusal(
                Code::ZeroMatches,
                "the ODT mutation selector matched nothing",
            ));
        }
    } else {
        return Err(refusal(
            Code::InvalidArgument,
            "the ODT mutation target is empty",
        ));
    }
    bound(LimitKind::ExpandedTargets, nodes.len(), maximum)?;
    Ok(nodes)
}

/// Set properties on every target through one leaf edit.
pub(crate) fn set_with_limits<'p, 'v>(
    file: &Odt,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.check_source()?;
    let props = Props::from_props(props)?;
    let nodes = resolve_targets(file, &target, limits.max_expanded_targets())?;
    if nodes
        .iter()
        .any(|node| node.element_kind() == Kind::Metadata)
    {
        if nodes.len() != 1 {
            return Err(refusal(
                Code::InvalidSelector,
                "a metadata set cannot be combined with element targets",
            ));
        }
        return set_metadata(file, &props, limits);
    }
    if nodes.iter().any(|node| node.element_kind().is_part()) {
        return set_parts(file, &nodes, &props, limits);
    }
    let body = file.body()?;
    let mut edit = file.package()?.edit();
    let mut operations = 0usize;
    let mut staged_bytes = 0usize;
    for node in &nodes {
        let Some(id) = node.id() else {
            return Err(refusal(
                Code::UnsupportedType,
                "the ODT element has no writable properties",
            ));
        };
        let staged = translate(node.element_kind(), &props, body)?;
        operations = operations.saturating_add(staged.operations());
        staged_bytes = staged_bytes.saturating_add(staged.bytes());
        bound(LimitKind::Operations, operations, limits.max_operations())?;
        bound(
            LimitKind::StagedBytes,
            staged_bytes,
            limits.max_staged_bytes(),
        )?;
        stage(&mut edit, id, &staged)?;
    }
    finish(file, edit, operations, limits)
}

/// One tab stop as the unified keys write it: the leaf edit and its
/// position in twips (for ordering).
type TabList = Vec<(Option<i64>, litchi_odt::body::TabStopSpec)>;

/// Apply one tab stop's unified `pos`, `val` and `leader` to `stop`.
fn tab_props(
    props: &Props<'_>,
    stop: &mut (Option<i64>, litchi_odt::body::TabStopSpec),
    adding: bool,
) -> Result<usize> {
    let mut operations = 0;
    for (key, value) in props.iter() {
        let name = key.as_str();
        let Some(spec) = spec::TAB_PROPS
            .iter()
            .find(|spec| spec.name.eq_ignore_ascii_case(name))
        else {
            return Err(refusal(
                Code::UnsupportedProperty,
                "the property is not an ODT tab stop property",
            )
            .with_property(name.to_owned()));
        };
        let key = spec.name;
        let coerced = required(coerce(value, spec.ty, key)?, key)?;
        match (key, coerced) {
            ("pos", Coerced::Int(twips)) => {
                let twips = i32::try_from(twips)
                    .ok()
                    .filter(|twips| (0..=31_680).contains(twips))
                    .ok_or_else(|| {
                        refusal(
                            Code::OutOfRange,
                            "an ODT tab stop position runs from 0 to 31680 twips",
                        )
                        .with_property(key)
                    })?;
                // The same position keeps its spelling (an exact no-op).
                if stop.0 != Some(i64::from(twips)) {
                    stop.0 = Some(i64::from(twips));
                    stop.1.position = odf_length(Length::from_twips(twips));
                }
            },
            ("val", Coerced::Token(token)) => {
                let (kind, character) = match token {
                    "center" => (Some("center"), None),
                    "right" => (Some("right"), None),
                    "decimal" => (Some("char"), Some(".")),
                    _ => (None, None),
                };
                // A left stop keeps an explicit `left` it was written with.
                if kind.is_some() || stop.1.kind.as_deref() != Some("left") {
                    stop.1.kind = kind.map(str::to_owned);
                }
                stop.1.character = character.map(str::to_owned);
            },
            ("leader", Coerced::Token(token)) => {
                let (style, text) = match token {
                    "dot" => (Some("dotted"), Some(".")),
                    "hyphen" => (Some("dash"), Some("-")),
                    "underscore" => (Some("solid"), Some("_")),
                    "middleDot" => (Some("dotted"), Some("·")),
                    _ => (None, None),
                };
                stop.1.leader_style = style.map(str::to_owned);
                stop.1.leader_text = text.map(str::to_owned);
            },
            _ => return Err(wrong(key)),
        }
        operations += 1;
    }
    if adding && stop.0.is_none() {
        return Err(missing("pos", "an ODT tab stop needs a `pos` in twips"));
    }
    Ok(operations)
}

/// Stage the tab stops of each paragraph whose list changed.
fn stage_tabs(edit: &mut Edit, lists: Vec<(NodeId, TabList, TabList)>) -> Result<()> {
    for (paragraph, before, after) in lists {
        let unchanged = before.len() == after.len()
            && before
                .iter()
                .zip(&after)
                .all(|(left, right)| left.1 == right.1);
        if unchanged {
            continue;
        }
        let mut properties = ParagraphPropertyUpdate::new();
        properties.tab_stops = Some(after.into_iter().map(|(_, stop)| stop).collect());
        edit.set_paragraph_properties(paragraph, &properties)
            .map_err(staging_error)?;
    }
    Ok(())
}

/// The tab list of `paragraph` among the lists one call edits, read from
/// the source on first use.
fn tab_list<'l>(
    file: &Odt,
    lists: &'l mut Vec<(NodeId, TabList, TabList)>,
    paragraph: NodeId,
) -> Result<&'l mut TabList> {
    let index = match lists.iter().position(|(held, _, _)| *held == paragraph) {
        Some(index) => index,
        None => {
            let stops = super::parts::tab_edits(file, paragraph)?;
            lists.push((paragraph, stops.clone(), stops));
            lists.len() - 1
        },
    };
    Ok(&mut lists[index].2)
}

/// Set properties of part targets (record 0328): a header or footer's
/// text, tab stops. Every other part is read-only.
fn set_parts(
    file: &Odt,
    nodes: &[Node<'_>],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    use super::node::Part;
    if nodes.iter().any(|node| !node.element_kind().is_part()) {
        return Err(refusal(
            Code::InvalidSelector,
            "an ODT set cannot combine part targets with body elements",
        ));
    }
    let mut edit = file.package()?.edit();
    let mut operations = 0usize;
    let mut lists: Vec<(NodeId, TabList, TabList)> = Vec::new();
    for node in nodes {
        match (node.element_kind(), node.part()) {
            (Kind::Header | Kind::Footer, Some(Part::Region(index))) => {
                let mut text = None;
                for (key, value) in props.iter() {
                    let name = key.as_str();
                    if !name.eq_ignore_ascii_case("text") {
                        let readable = spec::STORY_PROPS
                            .iter()
                            .any(|spec| spec.name.eq_ignore_ascii_case(name));
                        return Err(if readable {
                            refusal(
                                Code::UnsupportedCapability,
                                "the ODT header or footer property is read-only",
                            )
                        } else {
                            refusal(
                                Code::UnsupportedProperty,
                                "the property is not an ODT header or footer property",
                            )
                        }
                        .with_property(name.to_owned()));
                    }
                    text = Some(required(coerce_text(value, "text")?, "text")?);
                }
                let Some(text) = text else {
                    continue;
                };
                let region = super::parts::region(file, index)?.ok_or_else(|| {
                    refusal(Code::InternalError, "an ODT region no longer resolves")
                })?;
                edit.set_region_text(&region, &text).map_err(text_error)?;
                operations += 1;
            },
            (Kind::Tab, Some(Part::Tab(paragraph, index))) => {
                let list = tab_list(file, &mut lists, paragraph)?;
                let stop = list.get_mut(index as usize).ok_or_else(|| {
                    refusal(Code::InternalError, "an ODT tab stop no longer resolves")
                })?;
                operations += tab_props(props, stop, false)?;
                list.sort_by_key(|(position, _)| position.unwrap_or(i64::MAX));
            },
            _ => {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "this ODT element is read-only",
                ));
            },
        }
        bound(LimitKind::Operations, operations, limits.max_operations())?;
    }
    stage_tabs(&mut edit, lists)?;
    finish(file, edit, operations, limits)
}

/// Add one tab stop under a paragraph or heading, in position order.
fn add_tab(
    file: &Odt,
    parent: &Node<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let Some(paragraph) = parent
        .id()
        .filter(|_| matches!(parent.element_kind(), Kind::Paragraph | Kind::Heading))
    else {
        return Err(refusal(
            Code::UnsupportedType,
            "an ODT tab stop is added under a paragraph or heading",
        ));
    };
    let mut stop = (None, litchi_odt::body::TabStopSpec::new("0pt"));
    let operations = tab_props(props, &mut stop, true)?;
    bound(LimitKind::Operations, operations, limits.max_operations())?;
    let mut lists = Vec::new();
    let list = tab_list(file, &mut lists, paragraph)?;
    let position = stop.0;
    if list.iter().any(|(held, _)| *held == position) {
        return Err(refusal(
            Code::InvalidValue,
            "the ODT paragraph already has a tab stop at this position",
        )
        .with_property("pos"));
    }
    let ordinal = list
        .iter()
        .take_while(|(held, _)| held.unwrap_or(i64::MAX) < position.unwrap_or(0))
        .count();
    list.insert(ordinal, stop);
    let mut edit = file.package()?.edit();
    stage_tabs(&mut edit, lists)?;
    let commit = finish(file, edit, operations.max(1), limits)?;
    Ok((commit, child_path(parent, Kind::Tab, ordinal)))
}

/// Remove tab stops (record 0328).
fn remove_tabs(file: &Odt, nodes: &[Node<'_>], limits: &SemanticLimits) -> Result<Commit> {
    use super::node::Part;
    let mut lists: Vec<(NodeId, TabList, TabList)> = Vec::new();
    // Remove the highest index of each paragraph first, so indexes stay
    // valid.
    let mut targets: Vec<(NodeId, u32)> = Vec::new();
    for node in nodes {
        match node.part() {
            Some(Part::Tab(paragraph, index)) => targets.push((paragraph, index)),
            _ => {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "this ODT element cannot be removed",
                ));
            },
        }
    }
    targets.sort_by_key(|target| std::cmp::Reverse(target.1));
    targets.dedup();
    for (paragraph, index) in &targets {
        let list = tab_list(file, &mut lists, *paragraph)?;
        if (*index as usize) < list.len() {
            list.remove(*index as usize);
        }
    }
    bound(
        LimitKind::Operations,
        targets.len(),
        limits.max_operations(),
    )?;
    let mut edit = file.package()?.edit();
    stage_tabs(&mut edit, lists)?;
    finish(file, edit, targets.len(), limits)
}

/// The `/metadata` keys and the fields they write.
const METADATA_KEYS: [(&str, MetadataField); 7] = [
    ("title", MetadataField::Title),
    ("subject", MetadataField::Subject),
    ("description", MetadataField::Description),
    ("author", MetadataField::Author),
    ("lastModifiedBy", MetadataField::LastModifiedBy),
    ("language", MetadataField::Language),
    ("keywords", MetadataField::Keywords),
];

/// Commit `/metadata` assignments through the leaf metadata family.
fn set_metadata(file: &Odt, props: &Props<'_>, limits: &SemanticLimits) -> Result<Commit> {
    if props.is_empty() {
        return Err(refusal(
            Code::InvalidArgument,
            "a metadata set needs at least one property",
        ));
    }
    bound(LimitKind::Operations, props.len(), limits.max_operations())?;
    let mut edit = file.package()?.edit_metadata();
    for (key, value) in props.iter() {
        let name = key.as_str();
        let Some((key, field)) = METADATA_KEYS
            .iter()
            .find(|(key, _)| key.eq_ignore_ascii_case(name))
        else {
            let readable = spec::METADATA_PROPS
                .iter()
                .any(|spec| spec.name.eq_ignore_ascii_case(name));
            return Err(if readable {
                refusal(
                    Code::UnsupportedCapability,
                    "the ODT metadata property is read-only",
                )
            } else {
                refusal(
                    Code::UnsupportedProperty,
                    "the metadata key is not an ODT document property",
                )
                .with_suggestion(
                    "use title, subject, description, author, lastModifiedBy, language or keywords",
                )
            }
            .with_property(name.to_owned()));
        };
        if *field == MetadataField::Keywords {
            let keywords = crate::props::keyword_list(value).ok_or_else(|| wrong(key))?;
            edit.set_keywords(&keywords).map_err(metadata_error)?;
            continue;
        }
        let value = coerce_text(value, key)?;
        edit.set(*field, value.as_deref()).map_err(metadata_error)?;
    }
    let native = edit.commit().map_err(odf_error)?;
    let changed = native.diagnostics().changed();
    let (package, patch) = native.into_parts();
    let snapshot_bytes = if changed { package.as_bytes().len() } else { 0 };
    bound(
        LimitKind::SnapshotBytes,
        snapshot_bytes,
        limits.max_snapshot_bytes(),
    )?;
    let after = if changed {
        file.adopt(package)?
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok(Commit {
        before: file.clone(),
        after,
        patch: Patch::Metadata(patch),
        operations: props.len(),
        snapshot_bytes,
    })
}

/// A metadata staging refusal: the caller's value, not the file, is at fault.
fn metadata_error(error: litchi_odf_common::Error) -> Error {
    match error {
        litchi_odf_common::Error::Core(error) => staging_error(error.into()),
        other => odf_error(other),
    }
}

/// The unified node at a body-backed path.
fn body_node<'f>(file: &'f Odt, path: &PathArg<'_>) -> Result<Node<'f>> {
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    if node.id().is_none() {
        return Err(refusal(
            Code::UnsupportedType,
            "the ODT element is not part of the text body",
        ));
    }
    Ok(node)
}

fn leaf<'f>(node: &Node<'f>) -> Result<LeafNode<'f>> {
    node.leaf()
        .ok_or_else(|| refusal(Code::InternalError, "an ODT body node lost its leaf"))
}

/// The leaf placement of a unified placement within `parent`, and the
/// same-kind ordinal the placed node takes.
fn placement_in<'f>(
    file: &'f Odt,
    parent: &Node<'f>,
    kind: Kind,
    placement: Option<&Placement<'_>>,
    moving: Option<NodeId>,
) -> Result<(LeafPlacement, usize)> {
    let parent_leaf = leaf(parent)?;
    let siblings: Vec<LeafNode<'f>> = parent_leaf
        .children()
        .filter(|child| Some(child.id()) != moving)
        .collect();
    let same = |child: &LeafNode<'_>| Kind::of(child.kind()) == Some(kind);
    let count = siblings.iter().filter(|child| same(child)).count();
    let before_ordinal = |anchor: NodeId| {
        siblings
            .iter()
            .take_while(|child| child.id() != anchor)
            .filter(|child| same(child))
            .count()
    };
    let anchor_of = |path: &PathArg<'_>| -> Result<NodeId> {
        let anchor = body_node(file, path)?;
        let id = anchor.id().unwrap_or(NodeId::BODY);
        if !siblings.iter().any(|child| child.id() == id) {
            return Err(refusal(
                Code::UnsupportedType,
                "an ODT placement anchor must be a sibling under the same parent",
            ));
        }
        Ok(id)
    };
    Ok(match placement {
        None | Some(Placement::End | Placement::Into(_, None)) => (LeafPlacement::End, count),
        Some(Placement::Start) => (LeafPlacement::Start, 0),
        Some(Placement::Before(path)) => {
            let anchor = anchor_of(path)?;
            (LeafPlacement::Before(anchor), before_ordinal(anchor))
        },
        Some(Placement::After(path)) => {
            let anchor = anchor_of(path)?;
            let anchored = siblings
                .iter()
                .any(|child| child.id() == anchor && same(child));
            (
                LeafPlacement::After(anchor),
                before_ordinal(anchor) + usize::from(anchored),
            )
        },
        Some(Placement::Position(position) | Placement::Into(_, Some(position))) => {
            let position = position.get();
            if position > count {
                return Err(refusal(
                    Code::OutOfRange,
                    "the ODT insertion position is outside the parent",
                ));
            }
            match siblings.iter().filter(|child| same(child)).nth(position) {
                Some(anchor) => (LeafPlacement::Before(anchor.id()), position),
                None => (LeafPlacement::End, count),
            }
        },
    })
}

/// The path of a node with `kind` at `ordinal` under `parent`.
fn child_path(parent: &Node<'_>, kind: Kind, ordinal: usize) -> Path<'static> {
    use crate::path::navigate::ElementKind;
    let mut steps: Vec<Step<'static>> = parent
        .address()
        .steps()
        .iter()
        .map(|step| match step.kind {
            Kind::Body | Kind::Styles | Kind::Metadata => Step::new(step.kind.name()),
            other => {
                Step::new(other.name()).with(Predicate::Index(Position::new(step.ordinal as usize)))
            },
        })
        .collect();
    steps.push(Step::new(kind.name()).with(Predicate::Index(Position::new(ordinal))));
    Path::from_steps(steps)
}

fn find_prop<'a, 'b>(props: &'b Props<'a>, key: &str) -> Option<&'b PropValue<'a>> {
    props
        .iter()
        .find(|(name, _)| name.as_str().eq_ignore_ascii_case(key))
        .map(|(_, value)| value)
}

/// Read one text-valued add property.
fn add_text(props: &Props<'_>, key: &'static str) -> Result<Option<String>> {
    find_prop(props, key).map_or(Ok(None), |value| coerce_text(value, key))
}

fn add_int(props: &Props<'_>, key: &'static str) -> Result<Option<i64>> {
    match find_prop(props, key) {
        Some(value) => match coerce(value, ValueType::Int, key)? {
            Some(Coerced::Int(value)) => Ok(Some(value)),
            None => Ok(None),
            Some(_) => Err(wrong(key)),
        },
        None => Ok(None),
    }
}

/// Refuse add properties outside `allowed`.
fn admit_add(props: &Props<'_>, allowed: &[&str]) -> Result<()> {
    for (key, _) in props.iter() {
        if !allowed
            .iter()
            .any(|name| name.eq_ignore_ascii_case(key.as_str()))
        {
            return Err(refusal(
                Code::UnsupportedProperty,
                "the property is not accepted when adding this ODT element",
            )
            .with_property(key.as_str().to_owned()));
        }
    }
    Ok(())
}

/// The kind a unified add names.
fn add_kind(name: &str) -> Option<Kind> {
    use crate::path::navigate::ElementKind;
    let canonical = Kind::VOCABULARY.canonical(name)?;
    Kind::from_name(canonical)
}

fn missing(key: &'static str, message: &'static str) -> Error {
    refusal(Code::MissingProperty, message).with_property(key)
}

/// Translate a unified new node into the leaf node.
#[allow(
    clippy::too_many_lines,
    reason = "one arm per insertable element, kept together for review against the leaf"
)]
fn new_node(kind: Kind, props: &Props<'_>, body: litchi_odt::body::Body<'_>) -> Result<LeafNew> {
    let text = || add_text(props, "text").map(Option::unwrap_or_default);
    Ok(match kind {
        Kind::Paragraph | Kind::Heading => {
            for (key, _) in props.iter() {
                let name = key.as_str();
                let known = spec::SPEC.elements[kind as usize]
                    .props
                    .iter()
                    .any(|spec| spec.access.set && spec.name.eq_ignore_ascii_case(name));
                if !known {
                    return Err(refusal(
                        Code::UnsupportedProperty,
                        "the property is not accepted when adding this ODT element",
                    )
                    .with_property(name.to_owned()));
                }
            }
            let style = add_text(props, "style")?;
            if let Some(style) = &style {
                check_style(body, style, "style")?;
            }
            if kind == Kind::Heading {
                let level = add_int(props, "level")?.unwrap_or(1);
                let level = u8::try_from(level)
                    .ok()
                    .filter(|level| (1..=10).contains(level))
                    .ok_or_else(|| {
                        refusal(Code::OutOfRange, "ODT heading levels run from 1 to 10")
                            .with_property("level")
                    })?;
                LeafNew::Heading {
                    text: text()?,
                    level,
                    style,
                }
            } else {
                LeafNew::Paragraph {
                    text: text()?,
                    style,
                }
            }
        },
        Kind::Run => {
            let mut formatting = Props::new();
            for (key, value) in props.iter() {
                if !key.as_str().eq_ignore_ascii_case("text") {
                    formatting.push(key.clone(), value.clone());
                }
            }
            let staged = translate(Kind::Run, &formatting, body)?;
            if !staged.attributes.is_empty() || !staged.paragraph_properties.is_empty() {
                return Err(refusal(
                    Code::UnsupportedProperty,
                    "only character formatting is accepted when adding an ODT run",
                ));
            }
            LeafNew::Span {
                text: text()?,
                properties: staged.text_properties,
            }
        },
        Kind::Hyperlink => {
            admit_add(props, &["text", "url"])?;
            LeafNew::Hyperlink {
                href: add_text(props, "url")?
                    .ok_or_else(|| missing("url", "an ODT hyperlink needs a `url`"))?,
                text: text()?,
            }
        },
        Kind::Bookmark => {
            admit_add(props, &["name"])?;
            LeafNew::Bookmark {
                name: add_text(props, "name")?
                    .ok_or_else(|| missing("name", "an ODT bookmark needs a `name`"))?,
            }
        },
        Kind::Table => {
            admit_add(props, &["rows", "cols"])?;
            let size = |key: &'static str| -> Result<usize> {
                let value = add_int(props, key)?.unwrap_or(1);
                usize::try_from(value)
                    .ok()
                    .filter(|value| *value >= 1)
                    .ok_or_else(|| {
                        refusal(Code::OutOfRange, "ODT table dimensions are positive")
                            .with_property(key)
                    })
            };
            LeafNew::Table {
                rows: size("rows")?,
                columns: size("cols")?,
            }
        },
        Kind::Row => {
            admit_add(props, &[])?;
            LeafNew::Row
        },
        Kind::List => {
            admit_add(props, &["text"])?;
            LeafNew::List {
                items: text()?.split('\n').map(str::to_owned).collect(),
            }
        },
        Kind::ListItem => {
            admit_add(props, &["text"])?;
            LeafNew::ListItem { text: text()? }
        },
        Kind::Section => {
            admit_add(props, &["name"])?;
            LeafNew::Section {
                name: add_text(props, "name")?
                    .ok_or_else(|| missing("name", "an ODT section needs a `name`"))?,
            }
        },
        Kind::Footnote => {
            admit_add(props, &["text"])?;
            LeafNew::Footnote { text: text()? }
        },
        Kind::Endnote => {
            admit_add(props, &["text"])?;
            LeafNew::Endnote { text: text()? }
        },
        Kind::Comment => {
            admit_add(props, &["text", "author"])?;
            LeafNew::Annotation {
                text: text()?,
                author: add_text(props, "author")?,
            }
        },
        _ => {
            return Err(refusal(
                Code::UnsupportedElement,
                "this ODT element cannot be added",
            ));
        },
    })
}

/// The leaf kind of an insertable unified kind.
const fn leaf_kind(kind: Kind) -> Option<LeafKind> {
    Some(match kind {
        Kind::Paragraph => LeafKind::Paragraph,
        Kind::Heading => LeafKind::Heading,
        Kind::Run => LeafKind::Span,
        Kind::Hyperlink => LeafKind::Hyperlink,
        Kind::List => LeafKind::List,
        Kind::ListItem => LeafKind::ListItem,
        Kind::Table => LeafKind::Table,
        Kind::Row => LeafKind::Row,
        Kind::Section => LeafKind::Section,
        Kind::Footnote => LeafKind::Footnote,
        Kind::Endnote => LeafKind::Endnote,
        Kind::Comment => LeafKind::Annotation,
        Kind::Bookmark => LeafKind::Bookmark,
        _ => return None,
    })
}

/// Add one element; returns its positional path.
pub(crate) fn add_with_limits<'p, 'n>(
    file: &Odt,
    parent: PathArg<'p>,
    node: NewNode<'n>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    file.check_source()?;
    let kind = add_kind(node.kind()).ok_or_else(|| {
        refusal(
            Code::UnsupportedElement,
            "the element type is not an ODT element",
        )
    })?;
    if kind == Kind::Tab {
        let parent = match node.placement() {
            Some(Placement::Into(path, _)) => body_node(file, path)?,
            _ => body_node(file, &parent)?,
        };
        return add_tab(file, &parent, node.props(), limits);
    }
    let leaf_child = leaf_kind(kind)
        .ok_or_else(|| refusal(Code::UnsupportedElement, "this ODT element cannot be added"))?;
    let parent = match node.placement() {
        Some(Placement::Into(path, _)) => body_node(file, path)?,
        _ => body_node(file, &parent)?,
    };
    if !litchi_odt::body::admits(leaf(&parent)?.kind(), leaf_child) {
        return Err(refusal(
            Code::UnsupportedType,
            "the parent cannot hold this ODT element",
        ));
    }
    let body = file.body()?;
    let new = new_node(kind, node.props(), body)?;
    let (placement, ordinal) = placement_in(file, &parent, kind, node.placement(), None)?;
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    let parent_id = parent.id().unwrap_or(NodeId::BODY);
    let mut edit = file.package()?.edit();
    edit.insert(Location::new(parent_id, placement), &new)
        .map_err(staging_error)?;
    let commit = finish(file, edit, 1, limits)?;
    let path = child_path(&parent, kind, ordinal);
    // Formatting given with a new paragraph or heading names a node the
    // source tree does not have; it is set on the committed node.
    let formatting = paragraph_formatting(kind, node.props());
    if formatting.is_empty() {
        return Ok((commit, path));
    }
    let followup = set_with_limits(
        &commit.after,
        path.clone().into_target()?,
        formatting,
        limits,
    )?;
    Ok((
        Commit {
            before: commit.before,
            after: followup.after,
            patch: followup.patch,
            operations: commit.operations + followup.operations,
            snapshot_bytes: commit
                .snapshot_bytes
                .saturating_add(followup.snapshot_bytes),
        },
        path,
    ))
}

/// The formatting properties of a paragraph or heading add.
fn paragraph_formatting<'a>(kind: Kind, props: &Props<'a>) -> Props<'a> {
    let mut out = Props::new();
    if matches!(kind, Kind::Paragraph | Kind::Heading) {
        for (key, value) in props.iter() {
            let name = key.as_str();
            if !["text", "style", "level"]
                .iter()
                .any(|skip| skip.eq_ignore_ascii_case(name))
            {
                out.push(key.clone(), value.clone());
            }
        }
    }
    out
}

/// Remove every target in one leaf edit.
pub(crate) fn remove_with_limits<'p>(
    file: &Odt,
    target: TargetArg<'p>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.check_source()?;
    let nodes = resolve_targets(file, &target, limits.max_expanded_targets())?;
    if nodes.iter().any(|node| node.element_kind().is_part()) {
        if nodes.iter().all(|node| node.element_kind() == Kind::Tab) {
            return remove_tabs(file, &nodes, limits);
        }
        return Err(refusal(
            Code::UnsupportedCapability,
            "this ODT element cannot be removed",
        ));
    }
    bound(LimitKind::Operations, nodes.len(), limits.max_operations())?;
    let mut edit = file.package()?.edit();
    for node in &nodes {
        let id = node.id().filter(|id| *id != NodeId::BODY).ok_or_else(|| {
            refusal(
                Code::UnsupportedType,
                "only elements of the ODT body can be removed",
            )
        })?;
        edit.remove(id).map_err(staging_error)?;
    }
    finish(file, edit, nodes.len(), limits)
}

/// The unified node of a node's parent, re-resolved by its address.
fn parent_of<'f>(file: &'f Odt, node: &Node<'f>) -> Result<Node<'f>> {
    use crate::path::navigate::{Lookup, Navigate, Resolved, Select};
    let steps = node.address().steps();
    let Some((_, parent_steps)) = steps.split_last() else {
        return Err(refusal(Code::UnsupportedType, "the ODT root has no parent"));
    };
    let nav = Nav::new(file);
    let mut current = nav.root()?;
    for step in parent_steps {
        let resolved = Resolved {
            kind: step.kind,
            select: Select::Ordinal(Position::new(step.ordinal as usize)),
        };
        current = match nav.child(&current, &resolved)? {
            Lookup::Found(node) => node,
            _ => {
                return Err(refusal(
                    Code::InternalError,
                    "an ODT node's parent no longer resolves",
                ));
            },
        };
    }
    Ok(current)
}

/// Move one element, keeping its bytes.
pub(crate) fn move_to_with_limits<'p, 'q>(
    file: &Odt,
    path: PathArg<'p>,
    placement: Placement<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.check_source()?;
    let node = body_node(file, &path)?;
    let id = node.id().unwrap_or(NodeId::BODY);
    let parent = match &placement {
        Placement::Before(anchor) | Placement::After(anchor) => {
            parent_of(file, &body_node(file, anchor)?)?
        },
        Placement::Into(parent, _) => body_node(file, parent)?,
        _ => parent_of(file, &node)?,
    };
    let (placement, _) = placement_in(
        file,
        &parent,
        node.element_kind(),
        Some(&placement),
        Some(id),
    )?;
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    let mut edit = file.package()?.edit();
    edit.relocate(
        id,
        Location::new(parent.id().unwrap_or(NodeId::BODY), placement),
    )
    .map_err(staging_error)?;
    finish(file, edit, 1, limits)
}

/// Exchange two elements, keeping both elements' bytes.
pub(crate) fn swap_with_limits<'p, 'q>(
    file: &Odt,
    left: PathArg<'p>,
    right: PathArg<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    file.check_source()?;
    let left = body_node(file, &left)?;
    let right = body_node(file, &right)?;
    bound(LimitKind::Operations, 1, limits.max_operations())?;
    let mut edit = file.package()?.edit();
    edit.swap(
        left.id().unwrap_or(NodeId::BODY),
        right.id().unwrap_or(NodeId::BODY),
    )
    .map_err(staging_error)?;
    finish(file, edit, 1, limits)
}

/// A fresh empty text document, with create-time metadata committed through
/// the leaf metadata edit.
pub(crate) fn create(options: &crate::CreateOptions) -> Result<Odt> {
    create_from(options, FileFormat::Odt, || {
        litchi_odt::Package::new().map_err(odf_error)
    })
}

/// A fresh empty document of an ODF text-family `format` whose minimal
/// package `base` builds (the ODT body machinery then owns it, record
/// 0282), with create-time metadata committed through the leaf metadata
/// edit.
pub(crate) fn create_from(
    options: &crate::CreateOptions,
    format: FileFormat,
    base: impl FnOnce() -> Result<litchi_odt::Package>,
) -> Result<Odt> {
    create_tagged(options, format, base).map_err(|error| error.with_format(format))
}

fn create_tagged(
    options: &crate::CreateOptions,
    format: FileFormat,
    base: impl FnOnce() -> Result<litchi_odt::Package>,
) -> Result<Odt> {
    use crate::{Conformance, Flavor};
    if options.flavor_value() != Flavor::Document {
        return Err(refusal(
            Code::UnsupportedCapability,
            "ODT template creation is not supported",
        )
        .with_suggestion("use default CreateOptions for an ordinary empty ODT document"));
    }
    if options.conformance_value() == Conformance::Strict {
        return Err(refusal(
            Code::UnsupportedCapability,
            "ODT does not support OOXML Strict conformance",
        ));
    }
    let package = base()?;
    let Some(metadata) = options.metadata_value() else {
        return Odt::from_package(package, None, format);
    };
    let unsupported = metadata.identifier.is_some()
        || metadata.template.is_some()
        || metadata.revision.is_some()
        || metadata.created.is_some()
        || metadata.modified.is_some()
        || metadata.category.is_some()
        || metadata.company.is_some()
        || metadata.manager.is_some()
        || metadata.content_status.is_some()
        || metadata.content_type.is_some()
        || metadata.version.is_some()
        || metadata.last_printed.is_some()
        || metadata.hyperlink_base.is_some()
        || metadata.application.is_some();
    if unsupported {
        return Err(refusal(
            Code::UnsupportedCapability,
            "ODT creation writes only title, subject, description, author, lastModifiedBy, language and keywords",
        ));
    }
    let mut edit = package.edit_metadata();
    for (field, value) in [
        (MetadataField::Title, &metadata.title),
        (MetadataField::Subject, &metadata.subject),
        (MetadataField::Description, &metadata.description),
        (MetadataField::Author, &metadata.author),
        (MetadataField::LastModifiedBy, &metadata.last_modified_by),
        (MetadataField::Language, &metadata.language),
    ] {
        if let Some(value) = value {
            edit.set(field, Some(value)).map_err(metadata_error)?;
        }
    }
    if !metadata.keywords.is_empty() {
        edit.set_keywords(&metadata.keywords)
            .map_err(metadata_error)?;
    }
    let package = edit.commit().map_err(odf_error)?.into_package();
    Odt::from_package(package, None, format)
}

/// The columns a table's first row spans (`table:number-columns-repeated`
/// counted).
pub(crate) fn table_columns(table: LeafNode<'_>) -> usize {
    let Some(row) = table.children().find(|child| child.kind() == LeafKind::Row) else {
        return 0;
    };
    row.children()
        .map(|cell| {
            cell.attribute(Attribute::ColumnsRepeated)
                .and_then(|repeat| repeat.parse::<usize>().ok())
                .unwrap_or(1)
                .max(1)
        })
        .fold(0usize, usize::saturating_add)
}

/// One direct-path set of one body node, planned against a run's base
/// handle.
pub(crate) struct Planned {
    pub(crate) id: NodeId,
    /// The node's source byte range.
    pub(crate) range: std::ops::Range<usize>,
    pub(crate) staged: Staged,
}

impl Planned {
    pub(crate) fn operations(&self) -> usize {
        self.staged.operations()
    }

    pub(crate) fn staged_bytes(&self) -> usize {
        self.staged.bytes()
    }

    /// Stage the plan on a run's leaf edit.
    pub(crate) fn stage(&self, edit: &mut Edit) -> Result<()> {
        stage(edit, self.id, &self.staged)
    }
}

/// Plan one direct-path set against `file` without staging it; `None` when
/// the target is not one body element (`/metadata` commits through its own
/// leaf family), or when the edit changes later positional paths (clearing a
/// character-data run, replacing a run that holds notes or bookmarks).
pub(crate) fn plan(
    file: &Odt,
    target: &TargetArg<'_>,
    props: &Props<'_>,
) -> Result<Option<Planned>> {
    let Some(path) = target.path() else {
        return Ok(None);
    };
    if props.is_empty() {
        return Ok(None);
    }
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    let (Some(id), Some(leaf)) = (node.id(), node.leaf()) else {
        return Ok(None);
    };
    let staged = translate(node.element_kind(), props, file.body()?)?;
    if leaf.kind() == LeafKind::Text && staged.text.as_deref() == Some("") {
        return Ok(None);
    }
    if staged.text.is_some() && leaf.kind() == LeafKind::Span && leaf.encloses_indexed() {
        return Ok(None);
    }
    Ok(Some(Planned {
        id,
        range: leaf.range(),
        staged,
    }))
}
