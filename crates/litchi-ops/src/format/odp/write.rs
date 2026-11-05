//! Unified writes for OpenDocument presentations (record 0278).
//!
//! Targets resolve through the read navigator; properties are admitted by
//! the ODP write registry; the complete call becomes one list of leaf
//! `scene::Change`s staged by `litchi_odp::edit::Edit::splice_scene` on one
//! `Package::edit()` and published by its commit. Only the changed byte
//! ranges of `content.xml` (and `meta.xml` for metadata) are rewritten;
//! every other byte and member keeps its source.

pub(crate) mod coalesce;
pub(crate) mod find;
pub(crate) mod refresh;

use std::collections::HashSet;

use litchi_core::{FileFormat, Position};
use litchi_odp::scene::{
    Attribute, Change, NewPage, NewTextBox, Scene, Scope, StoryRef, TextPropertyUpdate,
};

use super::node::{Item, Node, object_kind};
use super::registry::{
    CELL_WRITE_TABLE, FRAME_WRITE_TABLE, GROUP_WRITE_TABLE, NOTES_WRITE_TABLE,
    PARAGRAPH_WRITE_TABLE, RUN_WRITE_TABLE, SHAPE_WRITE_TABLE, SLIDE_WRITE_TABLE, Setter,
    TABLE_WRITE_TABLE,
};
use super::spec::Kind;
use super::{Nav, Odp, core_error, odf_error};
use crate::error::{Code, Error, Result};
use crate::path::PathArg;
use crate::placement::Placement;
use crate::props::{IntoProps, Props};
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::{Base, Color, Length, Scalar};

/// The native proof of one unified ODP write.
pub(crate) enum Patch {
    /// A primary presentation transaction.
    Edit(litchi_odp::edit::Patch),
    /// A `/metadata` commit of the metadata family (record 0288).
    Metadata(litchi_odp::metadata::Patch),
}

impl Patch {
    pub(crate) fn is_noop(&self) -> bool {
        match self {
            Self::Edit(patch) => patch.is_noop(),
            Self::Metadata(patch) => patch.is_noop(),
        }
    }
}

/// The native proof and immutable endpoints of one unified ODP write.
pub(crate) struct WriteCommit {
    pub(crate) before: Odp,
    pub(crate) after: Odp,
    pub(crate) patch: Patch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

/// One planned write: the scene changes of one call.
#[derive(Debug, Clone, Default)]
pub(crate) struct Plan {
    pub(crate) changes: Vec<Change>,
    /// Text bytes the changes stage.
    pub(crate) staged: usize,
}

/// Values one call assigns.
#[derive(Debug, Default)]
struct Values {
    text: Option<String>,
    name: Option<String>,
    geometry: [Option<String>; 4],
    /// Character formatting (paragraphs and runs, record 0308).
    format: TextPropertyUpdate,
}

/// A font size as ODF points (`18pt`, `10.5pt`).
fn points(length: Length) -> String {
    #[allow(
        clippy::cast_precision_loss,
        reason = "EMU values of font sizes are far below 2^52"
    )]
    let points = length.emu() as f64 / 12_700.0;
    let mut text = format!("{points:.4}");
    while text.ends_with('0') {
        text.pop();
    }
    if text.ends_with('.') {
        text.pop();
    }
    text.push_str("pt");
    text
}

/// An opaque colour as `#rrggbb`.
fn rgb(color: Color) -> Result<String> {
    match color.base() {
        Base::Rgb(rgba) if rgba.is_opaque() && color.transforms().is_empty() => Ok(format!(
            "#{:02x}{:02x}{:02x}",
            rgba.red(),
            rgba.green(),
            rgba.blue()
        )),
        _ => Err(
            Error::new(Code::InvalidValue, "ODP colours are opaque #RRGGBB values")
                .with_format(FileFormat::Odp)
                .with_property("color"),
        ),
    }
}

/// Set properties on a path, a scoped query or `/metadata`.
pub(crate) fn set_with_limits<'p, 'v>(
    file: &Odp,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let props = Props::from_props(props)?;
    if crate::format::odf_metadata::is_target(&target) {
        return set_metadata(file, &props, limits);
    }
    let nodes = resolve_nodes(file, &target, limits)?;
    if nodes.iter().any(|node| node.element() == Kind::Hyperlink) {
        let changes = plan_links(file, &nodes, &props)?;
        return commit_changes(file, &changes, limits);
    }
    let plan = plan_set(&nodes, &props, limits)?;
    commit_changes(file, &plan.changes, limits)
}

/// The `link` and `tooltip` a hyperlink write assigns (record 0328).
struct LinkValues {
    link: Option<String>,
    tooltip: Option<String>,
}

fn link_values(props: &Props<'_>) -> Result<LinkValues> {
    let mut values = LinkValues {
        link: None,
        tooltip: None,
    };
    for (key, value) in props.iter() {
        let slot = match key.as_str().to_ascii_lowercase().as_str() {
            "link" | "url" | "href" => &mut values.link,
            "tooltip" => &mut values.tooltip,
            _ => {
                return Err(Error::new(
                    Code::UnsupportedProperty,
                    "an ODP hyperlink writes `link` and `tooltip`",
                )
                .with_format(FileFormat::Odp)
                .with_property(key.as_str().to_owned()));
            },
        };
        *slot = Some(text_value(value, "link")?);
    }
    if values.link.is_none() && values.tooltip.is_none() {
        return Err(Error::new(
            Code::InvalidArgument,
            "the ODP hyperlink write has no writable property",
        )
        .with_format(FileFormat::Odp));
    }
    Ok(values)
}

fn link_refusal(message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message)
        .with_format(FileFormat::Odp)
        .with_property("link")
}

/// The scene change that sets a shape's click to `link` (`none` removes
/// it).
fn shape_click_change(
    file: &Odp,
    object: &litchi_odp::scene::ObjectRef,
    link: &str,
) -> Result<Change> {
    let click = if link.eq_ignore_ascii_case("none") {
        None
    } else {
        Some(super::parts::shape_action(file.scene()?, link).map_err(link_refusal)?)
    };
    Ok(Change::SetShapeClick {
        object: object.clone(),
        click,
    })
}

/// Plan hyperlink sets: a shape's click action, a link run's target and
/// title (record 0328).
fn plan_links(file: &Odp, nodes: &[Node<'_>], props: &Props<'_>) -> Result<Vec<Change>> {
    let values = link_values(props)?;
    let mut changes = Vec::new();
    for node in nodes {
        match node.item() {
            Item::ShapeLink(object) => {
                if values.tooltip.is_some() {
                    return Err(Error::new(
                        Code::UnsupportedCapability,
                        "an ODP shape's click action has no tooltip; set a link run's",
                    )
                    .with_format(FileFormat::Odp)
                    .with_property("tooltip"));
                }
                if let Some(link) = &values.link {
                    changes.push(shape_click_change(file, object, link)?);
                }
            },
            Item::RunLink(story, paragraph, run) => {
                let href = match values.link.as_deref() {
                    Some(link) if link.eq_ignore_ascii_case("none") => None,
                    Some(link) => {
                        super::parts::check_url(link).map_err(link_refusal)?;
                        Some(link.to_owned())
                    },
                    None => node
                        .scene()
                        .story(story)
                        .and_then(|story| story.paragraphs().get(*paragraph))
                        .and_then(|paragraph| paragraph.runs().get(*run))
                        .and_then(|run| run.href().map(str::to_owned)),
                };
                let title = values
                    .tooltip
                    .as_ref()
                    .map(|tooltip| (!tooltip.is_empty()).then(|| tooltip.clone()));
                changes.push(Change::SetRunLink {
                    story: story.clone(),
                    paragraph: *paragraph,
                    run: *run,
                    href,
                    title,
                });
            },
            _ => {
                return Err(Error::new(
                    Code::InvalidSelector,
                    "an ODP set cannot combine hyperlinks with other elements",
                )
                .with_format(FileFormat::Odp));
            },
        }
    }
    Ok(changes)
}

/// Plan a set on resolved nodes.
pub(crate) fn plan_set(
    nodes: &[Node<'_>],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Plan> {
    let mut plan = Plan::default();
    for node in nodes {
        let values = admit(node, props, limits)?;
        plan_node(node, values, &mut plan)?;
        bound(
            LimitKind::StagedBytes,
            plan.staged,
            limits.max_staged_bytes(),
        )?;
    }
    Ok(plan)
}

/// Admit `props` for `node` through its registry table.
fn admit(node: &Node<'_>, props: &Props<'_>, limits: &SemanticLimits) -> Result<Values> {
    match node.element() {
        Kind::Slide => admit_table(props, &SLIDE_WRITE_TABLE, limits),
        Kind::Shape | Kind::Placeholder => admit_table(props, &SHAPE_WRITE_TABLE, limits),
        Kind::Picture | Kind::Chart | Kind::Ole | Kind::Media => {
            admit_table(props, &FRAME_WRITE_TABLE, limits)
        },
        Kind::Table => admit_table(props, &TABLE_WRITE_TABLE, limits),
        Kind::Group | Kind::Connector => admit_table(props, &GROUP_WRITE_TABLE, limits),
        Kind::Paragraph => admit_table(props, &PARAGRAPH_WRITE_TABLE, limits),
        Kind::Run => admit_table(props, &RUN_WRITE_TABLE, limits),
        Kind::Cell => admit_table(props, &CELL_WRITE_TABLE, limits),
        Kind::Notes => admit_table(props, &NOTES_WRITE_TABLE, limits),
        _ => Err(unsupported_target(
            "this ODP element has no writable properties",
        )),
    }
}

fn admit_table<const N: usize, const A: usize>(
    props: &Props<'_>,
    table: &crate::registry::Table<Setter, N, A>,
    limits: &SemanticLimits,
) -> Result<Values> {
    let mut values = Values::default();
    let mut seen = 0_u128;
    for (key, value) in props.iter() {
        let (index, entry) = table.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Odp)
                .with_property(key.as_str().to_owned())
        })?;
        let bit = 1_u128 << (index % 128);
        if seen & bit != 0 {
            return Err(Error::new(
                Code::InvalidArgument,
                "an ODP property was assigned more than once",
            )
            .with_format(FileFormat::Odp)
            .with_property(entry.spec.name));
        }
        seen |= bit;
        entry
            .prepare(Operation::Set, value, |setter, value| {
                match setter {
                    Setter::Text => {
                        let text = match value {
                            Some(Scalar::Text(text)) => text,
                            None => "",
                            _ => unreachable!("ODP text coercion returned a non-text scalar"),
                        };
                        bound(
                            LimitKind::StagedBytes,
                            text.len(),
                            limits.max_staged_bytes(),
                        )?;
                        values.text = Some(text.to_owned());
                    },
                    Setter::Name => {
                        let Some(Scalar::Text(name)) = value else {
                            unreachable!("ODP name coercion returned a non-text scalar");
                        };
                        values.name = Some(name.to_owned());
                    },
                    Setter::X | Setter::Y | Setter::Width | Setter::Height => {
                        let Some(Scalar::Length(length)) = value else {
                            unreachable!("ODP geometry coercion returned a non-length scalar");
                        };
                        let index = match setter {
                            Setter::X => 0,
                            Setter::Y => 1,
                            Setter::Width => 2,
                            _ => 3,
                        };
                        values.geometry[index] = Some(crate::format::odf_odp::odf_length(length));
                    },
                    Setter::Bold | Setter::Italic | Setter::Underline => {
                        let Some(Scalar::Bool(flag)) = value else {
                            unreachable!("ODP boolean coercion returned a non-boolean scalar");
                        };
                        let slot = match setter {
                            Setter::Bold => &mut values.format.bold,
                            Setter::Italic => &mut values.format.italic,
                            _ => &mut values.format.underline,
                        };
                        *slot = Some(flag);
                    },
                    Setter::Size => {
                        let Some(Scalar::Length(length)) = value else {
                            unreachable!("ODP size coercion returned a non-length scalar");
                        };
                        if length.emu() <= 0 {
                            return Err(Error::new(
                                Code::OutOfRange,
                                "ODP font sizes are positive",
                            )
                            .with_property("size"));
                        }
                        values.format.size = Some(points(length));
                    },
                    Setter::Color => {
                        let Some(Scalar::Color(color)) = value else {
                            unreachable!("ODP colour coercion returned a non-colour scalar");
                        };
                        values.format.color = Some(rgb(color)?);
                    },
                    Setter::Font => {
                        let Some(Scalar::Text(font)) = value else {
                            unreachable!("ODP font coercion returned a non-text scalar");
                        };
                        values.format.font = Some(font.to_owned());
                    },
                }
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Odp))?;
    }
    if values.text.is_none()
        && values.name.is_none()
        && values.geometry.iter().all(Option::is_none)
        && values.format.is_empty()
    {
        return Err(Error::new(
            Code::InvalidArgument,
            "the ODP write has no writable property",
        )
        .with_format(FileFormat::Odp));
    }
    Ok(values)
}

/// The scene changes one node's values make.
fn plan_node(node: &Node<'_>, values: Values, plan: &mut Plan) -> Result<()> {
    let text_len = values.text.as_ref().map_or(0, String::len);
    plan.staged = plan.staged.saturating_add(text_len);
    match node.item() {
        Item::Slide(page) => {
            if let Some(name) = values.name {
                plan.changes.push(Change::SetPageName { page: *page, name });
            }
        },
        Item::Object(object, kind) => {
            if let Some(text) = values.text {
                if *kind != Kind::Shape || node.own_story().is_none() {
                    return Err(unsupported_target(
                        "this ODP shape has no text story to replace",
                    ));
                }
                plan.changes.push(Change::SetStoryText {
                    story: StoryRef::object(object.clone()),
                    text,
                });
            }
            if !values.format.is_empty() {
                let story = node
                    .own_story()
                    .filter(|_| *kind == Kind::Shape)
                    .ok_or_else(|| {
                        unsupported_target("this ODP shape has no text story to format")
                    })?;
                format_story(node, &story, &values.format, plan)?;
            }
            if let Some(name) = values.name {
                plan.changes.push(Change::SetAttribute {
                    object: object.clone(),
                    attribute: Attribute::Name,
                    value: Some(name),
                });
            }
            for (attribute, value) in [
                Attribute::X,
                Attribute::Y,
                Attribute::Width,
                Attribute::Height,
            ]
            .into_iter()
            .zip(values.geometry)
            {
                if let Some(value) = value {
                    plan.changes.push(Change::SetAttribute {
                        object: object.clone(),
                        attribute,
                        value: Some(value),
                    });
                }
            }
        },
        Item::Paragraph(story, paragraph) => {
            if let Some(text) = values.text {
                plan.changes.push(Change::SetParagraphText {
                    story: story.clone(),
                    paragraph: *paragraph,
                    text,
                });
            }
            if !values.format.is_empty() {
                plan.changes.push(Change::SetTextProperties {
                    story: story.clone(),
                    paragraph: *paragraph,
                    run: None,
                    properties: values.format,
                });
            }
        },
        Item::Run(story, paragraph, run) => {
            if let Some(text) = values.text {
                let kind = node
                    .scene()
                    .story(story)
                    .and_then(|story| story.paragraphs().get(*paragraph))
                    .and_then(|paragraph| paragraph.runs().get(*run))
                    .map(litchi_odp::scene::Run::kind);
                if kind == Some(litchi_odp::scene::RunKind::Field) {
                    return Err(unsupported_target(
                        "ODP field runs (page numbers, dates) are computed and cannot be set",
                    ));
                }
                plan.changes.push(Change::SetRunText {
                    story: story.clone(),
                    paragraph: *paragraph,
                    run: *run,
                    text,
                });
            }
            if !values.format.is_empty() {
                plan.changes.push(Change::SetTextProperties {
                    story: story.clone(),
                    paragraph: *paragraph,
                    run: Some(*run),
                    properties: values.format,
                });
            }
        },
        Item::Cell(..) | Item::Notes(_) => {
            let story = || {
                node.own_story()
                    .ok_or_else(|| unsupported_target("this ODP element has no text story"))
            };
            if let Some(text) = values.text {
                plan.changes.push(Change::SetStoryText {
                    story: story()?,
                    text,
                });
            }
            if !values.format.is_empty() {
                format_story(node, &story()?, &values.format, plan)?;
            }
        },
        _ => {
            return Err(unsupported_target(
                "this ODP element has no writable properties",
            ));
        },
    }
    Ok(())
}

/// Character formatting of a whole story (a shape, table cell or notes,
/// record 0316): every paragraph takes it through its paragraph style, and
/// every run whose own style chain sets a different value takes it too, so
/// each run reads the new value.
fn format_story(
    node: &Node<'_>,
    story: &StoryRef,
    format: &TextPropertyUpdate,
    plan: &mut Plan,
) -> Result<()> {
    let paragraphs = node
        .scene()
        .story(story)
        .ok_or_else(|| unsupported_target("this ODP element has no text story"))?
        .paragraphs();
    let styles = node.file().automatic_styles()?;
    let common = node.file().common_styles()?;
    for (index, paragraph) in paragraphs.iter().enumerate() {
        plan.changes.push(Change::SetTextProperties {
            story: story.clone(),
            paragraph: index,
            run: None,
            properties: format.clone(),
        });
        for (run, item) in paragraph.runs().iter().enumerate() {
            if overrides(&styles.resolve_run(common, item), format) {
                plan.changes.push(Change::SetTextProperties {
                    story: story.clone(),
                    paragraph: index,
                    run: Some(run),
                    properties: format.clone(),
                });
            }
        }
    }
    Ok(())
}

/// Whether a run's own properties set a value `format` changes.
fn overrides(own: &litchi_odp::scene::TextProperties, format: &TextPropertyUpdate) -> bool {
    let differs =
        |set: Option<bool>, held: Option<bool>| set.zip(held).is_some_and(|(a, b)| a != b);
    differs(format.bold, own.bold())
        || differs(format.italic, own.italic())
        || differs(format.underline, own.underline())
        || format
            .size
            .as_deref()
            .zip(own.size())
            .is_some_and(|(set, held)| set != held)
        || format
            .color
            .as_deref()
            .zip(own.color())
            .is_some_and(|(set, held)| !set.eq_ignore_ascii_case(held))
        || format
            .font
            .as_deref()
            .zip(own.font())
            .is_some_and(|(set, held)| set != held)
}

/// Set `/metadata` through the leaf metadata family (record 0288).
fn set_metadata(file: &Odp, props: &Props<'_>, limits: &SemanticLimits) -> Result<WriteCommit> {
    let package = file.package()?;
    let mut edit = package.edit_metadata();
    crate::format::odf_metadata::stage(&mut edit, props, FileFormat::Odp, limits)?;
    let (next, patch, changed) = crate::format::odf_metadata::commit(edit, FileFormat::Odp)?;
    let snapshot_bytes = package
        .as_bytes()
        .len()
        .saturating_add(next.as_bytes().len());
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let after = if changed {
        file.with_package(next)?
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok(WriteCommit {
        before: file.clone(),
        after,
        patch: Patch::Metadata(patch),
        operations: if changed { props.len() } else { 0 },
        snapshot_bytes,
    })
}

/// Commit `changes` as one checked splice on one leaf edit.
pub(crate) fn commit_changes(
    file: &Odp,
    changes: &[Change],
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let package = file.package()?;
    check_snapshot(package, limits)?;
    let mut edit = package
        .edit()
        .map_err(|error| file.flat_refusal(odf_error(error)))?;
    edit.splice_scene(changes).map_err(leaf_error)?;
    finish(file, edit, changes.len(), limits)
}

fn check_snapshot(package: &litchi_odp::Package, limits: &SemanticLimits) -> Result<()> {
    let before = package.as_bytes().len();
    if before > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(before, limits.max_snapshot_bytes()));
    }
    Ok(())
}

fn finish(
    file: &Odp,
    edit: litchi_odp::edit::Edit,
    operations: usize,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    let commit = edit
        .commit()
        .map_err(|error| file.flat_refusal(odf_error(error)))?;
    let changed = commit.diagnostics().changed();
    let (package, patch) = commit.into_parts();
    let before_bytes = file.package()?.as_bytes().len();
    let snapshot_bytes = before_bytes.saturating_add(package.as_bytes().len());
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let after = if changed {
        file.with_package(package)?
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok(WriteCommit {
        before: file.clone(),
        after,
        patch: Patch::Edit(patch),
        operations: if changed { operations.max(1) } else { 0 },
        snapshot_bytes,
    })
}

/// Map a leaf planning or publication refusal: capability refusals stay
/// `unsupported_capability`; a request the scene refuses (an address,
/// overlap or reference conflict) is `invalid_argument`.
fn leaf_error(error: litchi_odf_common::Error) -> Error {
    match error {
        litchi_odf_common::Error::Core(error @ litchi_core::Error::InvalidFormat(_)) => {
            core_error(error).caused_by_argument(Code::InvalidArgument)
        },
        other => odf_error(other),
    }
}

// ---------------------------------------------------------------------------
// add, remove, move, swap
// ---------------------------------------------------------------------------

/// Add a slide under `/`, a text box under a slide, or a paragraph under a
/// text story (shape, notes, table cell).
pub(crate) fn add_with_limits(
    file: &Odp,
    parent: PathArg<'_>,
    node: crate::NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(WriteCommit, crate::path::Path<'static>)> {
    file.package()?;
    let nav = Nav::new(file);
    let parent = crate::path::resolve::get(&nav, &parent)?;
    let kind = node.kind().to_ascii_lowercase();
    let scene = file.scene()?;
    match (parent.item(), kind.as_str()) {
        (Item::Presentation, "slide") => {
            let mut new_page = NewPage::default();
            for (key, value) in node.props().iter() {
                match key.as_str().to_ascii_lowercase().as_str() {
                    "name" => new_page.name = Some(text_value(value, "name")?),
                    _ => return Err(add_property(key.as_str(), "a new ODP slide accepts `name` only")),
                }
            }
            if scene.pages().is_empty() {
                new_page.master = file.masters()?.first().map(|master| master.name.clone());
            }
            let count = scene.pages().len();
            let index = insertion_index(file, count, node.placement(), Kind::Slide)?;
            let commit = commit_changes(
                file,
                &[Change::InsertPage {
                    index,
                    page: new_page,
                }],
                limits,
            )?;
            let path = crate::path::Path::from_steps([crate::path::Step::new("slide")
                .with(crate::path::Predicate::Index(Position::new(index)))]);
            Ok((commit, path))
        },
        (Item::Slide(page), "shape" | "textbox") => {
            if !matches!(node.placement(), None | Some(Placement::End)) {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    "ODP text boxes are appended at the end of the slide's objects",
                )
                .with_format(FileFormat::Odp)
                .with_suggestion("add without a placement, then move the shape"));
            }
            let text_box = new_text_box(&node)?;
            bound(LimitKind::StagedBytes, text_box.text.len(), limits.max_staged_bytes())?;
            let ordinal = scene
                .page(*page)
                .map_or(0, |slide| {
                    slide
                        .objects()
                        .iter()
                        .filter(|object| object_kind(object) == Kind::Shape)
                        .count()
                });
            let commit = commit_changes(
                file,
                &[Change::InsertTextBox {
                    page: *page,
                    text_box,
                }],
                limits,
            )?;
            let mut steps = parent.steps().to_vec();
            steps.push(crate::node::PathStep::at("shape", Position::new(ordinal)));
            Ok((commit, owned_path(&steps)))
        },
        (Item::Object(_, Kind::Shape) | Item::Notes(_) | Item::Cell(..), "paragraph" | "p") => {
            let story = parent.own_story().ok_or_else(|| {
                unsupported_target("this ODP element has no text story to add a paragraph to")
            })?;
            let mut text = String::new();
            for (key, value) in node.props().iter() {
                match key.as_str().to_ascii_lowercase().as_str() {
                    "text" => text = text_value(value, "text")?,
                    _ => {
                        return Err(add_property(
                            key.as_str(),
                            "a new ODP paragraph accepts `text` only",
                        ));
                    },
                }
            }
            bound(LimitKind::StagedBytes, text.len(), limits.max_staged_bytes())?;
            let count = scene
                .story(&story)
                .map_or(0, |story| story.paragraphs().len());
            let index = insertion_index(file, count, node.placement(), Kind::Paragraph)?;
            let commit = commit_changes(
                file,
                &[Change::InsertParagraph {
                    story,
                    index,
                    text,
                }],
                limits,
            )?;
            let mut steps = parent.steps().to_vec();
            steps.push(crate::node::PathStep::at("paragraph", Position::new(index)));
            Ok((commit, owned_path(&steps)))
        },
        (Item::Object(object, Kind::Shape), "hyperlink" | "hlink") => {
            let values = link_values(node.props())?;
            if values.tooltip.is_some() {
                return Err(Error::new(
                    Code::UnsupportedCapability,
                    "an ODP shape's click action has no tooltip; add the link to a run instead",
                )
                .with_format(FileFormat::Odp)
                .with_property("tooltip"));
            }
            let Some(link) = values.link.filter(|link| !link.eq_ignore_ascii_case("none")) else {
                return Err(missing_link());
            };
            let shape = scene.object(object).ok_or_else(|| {
                Error::new(Code::InternalError, "the ODP shape vanished").with_format(FileFormat::Odp)
            })?;
            if litchi_odp::scene::parts::shape_click(file.content_xml()?, shape).is_some() {
                return Err(Error::new(
                    Code::InvalidArgument,
                    "the ODP shape already has a click action; set its hyperlink instead",
                )
                .with_format(FileFormat::Odp));
            }
            let change = shape_click_change(file, object, &link)?;
            let commit = commit_changes(file, &[change], limits)?;
            let mut steps = parent.steps().to_vec();
            steps.push(crate::node::PathStep::new("hyperlink"));
            Ok((commit, owned_path(&steps)))
        },
        (Item::Run(story, paragraph, run), "hyperlink" | "hlink") => {
            let values = link_values(node.props())?;
            let Some(link) = values.link.filter(|link| !link.eq_ignore_ascii_case("none")) else {
                return Err(missing_link());
            };
            super::parts::check_url(&link).map_err(link_refusal)?;
            let commit = commit_changes(
                file,
                &[Change::LinkRun {
                    story: story.clone(),
                    paragraph: *paragraph,
                    run: *run,
                    href: link,
                    title: values.tooltip.filter(|tooltip| !tooltip.is_empty()),
                }],
                limits,
            )?;
            let mut steps = parent.steps().to_vec();
            steps.push(crate::node::PathStep::new("hyperlink"));
            Ok((commit, owned_path(&steps)))
        },
        (Item::Paragraph(story, paragraph), "linebreak" | "br") => {
            if let Some((key, _)) = node.props().iter().next() {
                return Err(add_property(key.as_str(), "a new ODP line break takes no properties"));
            }
            let count = super::parts::line_break_count(file, scene, story, *paragraph)?;
            let commit = commit_changes(
                file,
                &[Change::InsertLineBreak {
                    story: story.clone(),
                    paragraph: *paragraph,
                }],
                limits,
            )?;
            let mut steps = parent.steps().to_vec();
            steps.push(crate::node::PathStep::at("linebreak", Position::new(count)));
            Ok((commit, owned_path(&steps)))
        },
        _ => Err(Error::new(
            Code::UnsupportedElement,
            "ODP add supports a slide under /, a shape (text box) under a slide, a paragraph under a text story, a hyperlink under a shape or run, and a line break under a paragraph",
        )
        .with_format(FileFormat::Odp)),
    }
}

fn missing_link() -> Error {
    Error::new(Code::MissingProperty, "an ODP hyperlink needs a `link`")
        .with_format(FileFormat::Odp)
        .with_property("link")
}

fn owned_path(steps: &[crate::node::PathStep<'_>]) -> crate::path::Path<'static> {
    crate::node::NodePath::new(steps).to_path().into_owned()
}

fn add_property(key: &str, message: &'static str) -> Error {
    Error::new(Code::UnsupportedProperty, message)
        .with_format(FileFormat::Odp)
        .with_property(key.to_owned())
}

fn text_value(value: &crate::props::PropValue<'_>, key: &'static str) -> Result<String> {
    value.as_text().map(str::to_owned).ok_or_else(|| {
        Error::new(Code::InvalidValue, "the property takes text")
            .with_format(FileFormat::Odp)
            .with_property(key)
    })
}

/// The text box an `add` of a shape describes.
fn new_text_box(node: &crate::NewNode<'_>) -> Result<NewTextBox> {
    let mut geometry: [Option<String>; 4] = Default::default();
    let mut text = String::new();
    let mut name = None;
    for (key, value) in node.props().iter() {
        let (index, entry) = SHAPE_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Odp)
                .with_property(key.as_str().to_owned())
        })?;
        let _ = index;
        entry
            .prepare(Operation::Set, value, |setter, value| {
                match (setter, value) {
                    (Setter::Text, Some(Scalar::Text(value))) => text = value.to_owned(),
                    (Setter::Text, None) => text.clear(),
                    (Setter::Name, Some(Scalar::Text(value))) => name = Some(value.to_owned()),
                    (setter, Some(Scalar::Length(length))) => {
                        let index = match setter {
                            Setter::X => 0,
                            Setter::Y => 1,
                            Setter::Width => 2,
                            _ => 3,
                        };
                        geometry[index] = Some(crate::format::odf_odp::odf_length(length));
                    },
                    _ => {
                        return Err(Error::new(
                            Code::InvalidValue,
                            "the property value does not fit a new ODP text box",
                        ));
                    },
                }
                Ok(())
            })
            .map_err(|error| error.with_format(FileFormat::Odp))?;
    }
    let [x, y, width, height] = geometry;
    let (Some(width), Some(height)) = (width, height) else {
        return Err(Error::new(
            Code::InvalidArgument,
            "a new ODP text box requires positive `width` and `height`",
        )
        .with_format(FileFormat::Odp)
        .with_suggestion("add the shape with width and height lengths such as 8cm and 2cm"));
    };
    Ok(NewTextBox {
        name,
        geometry: [
            x.unwrap_or_else(|| "0cm".to_owned()),
            y.unwrap_or_else(|| "0cm".to_owned()),
            width,
            height,
        ],
        text,
    })
}

/// The final position an `add` placement names among `count` siblings.
fn insertion_index(
    file: &Odp,
    count: usize,
    placement: Option<&Placement<'_>>,
    kind: Kind,
) -> Result<usize> {
    let index = match placement {
        None | Some(Placement::End) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => position.get(),
        Some(Placement::Before(path)) => sibling_index(file, path, kind)?,
        Some(Placement::After(path)) => sibling_index(file, path, kind)? + 1,
        Some(Placement::Into(_, position)) => position.map_or(count, Position::get),
    };
    if index > count {
        return Err(out_of_range(
            "the ODP insertion position is outside its parent",
        ));
    }
    Ok(index)
}

fn sibling_index(file: &Odp, path: &PathArg<'_>, kind: Kind) -> Result<usize> {
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    match (node.item(), kind) {
        (Item::Slide(page), Kind::Slide) => Ok(*page),
        (Item::Paragraph(_, index), Kind::Paragraph) => Ok(*index),
        _ => Err(Error::new(
            Code::UnsupportedType,
            "the ODP placement anchor is not a sibling of the new element",
        )
        .with_format(FileFormat::Odp)),
    }
}

/// Remove slides, drawing objects, paragraphs or runs.
pub(crate) fn remove_with_limits(
    file: &Odp,
    target: TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.package()?;
    let nodes = resolve_nodes(file, &target, limits)?;
    let scene = file.scene()?;
    let mut changes = Vec::new();
    let mut ranges: Vec<std::ops::Range<usize>> = Vec::new();
    let mut pages = 0_usize;
    for node in &nodes {
        let (change, range) = removal(scene, node)?;
        if matches!(change, Change::RemovePage { .. }) {
            pages += 1;
        }
        // A target inside another removed target goes with it.
        if ranges
            .iter()
            .any(|other| other.start <= range.start && range.end <= other.end)
        {
            continue;
        }
        ranges.retain(|other| !(range.start <= other.start && other.end <= range.end));
        changes.retain(|existing: &(std::ops::Range<usize>, Change)| {
            !(range.start <= existing.0.start && existing.0.end <= range.end)
        });
        ranges.push(range.clone());
        changes.push((range, change));
    }
    if pages > 0 && pages >= scene.pages().len() {
        return Err(Error::new(
            Code::InvalidArgument,
            "an ODP presentation must keep at least one slide",
        )
        .with_format(FileFormat::Odp));
    }
    let changes: Vec<Change> = changes.into_iter().map(|(_, change)| change).collect();
    commit_changes(file, &changes, limits)
}

fn removal(scene: &Scene, node: &Node<'_>) -> Result<(Change, std::ops::Range<usize>)> {
    let missing = || {
        Error::new(
            Code::InternalError,
            "the ODP removal target vanished from its scene",
        )
        .with_format(FileFormat::Odp)
    };
    Ok(match node.item() {
        Item::Slide(page) => (
            Change::RemovePage { page: *page },
            scene.page(*page).ok_or_else(missing)?.range(),
        ),
        Item::Object(object, _) if object.scope == Scope::Page => (
            Change::RemoveObject {
                object: object.clone(),
            },
            scene.object(object).ok_or_else(missing)?.range(),
        ),
        Item::Paragraph(story, paragraph) => (
            Change::RemoveParagraph {
                story: story.clone(),
                paragraph: *paragraph,
            },
            scene
                .story(story)
                .and_then(|story| story.paragraphs().get(*paragraph))
                .ok_or_else(missing)?
                .range(),
        ),
        Item::Run(story, paragraph, run) => (
            Change::RemoveRun {
                story: story.clone(),
                paragraph: *paragraph,
                run: *run,
            },
            scene
                .story(story)
                .and_then(|story| story.paragraphs().get(*paragraph))
                .and_then(|paragraph| paragraph.runs().get(*run))
                .ok_or_else(missing)?
                .range(),
        ),
        Item::ShapeLink(object) => {
            let shape = scene.object(object).ok_or_else(missing)?;
            if litchi_odp::scene::parts::shape_click(node.file().content_xml()?, shape).is_none() {
                return Err(unsupported_target(
                    "the ODP shape's hyperlink is its first link run's; remove that run's hyperlink",
                ));
            }
            (
                Change::SetShapeClick {
                    object: object.clone(),
                    click: None,
                },
                shape.range(),
            )
        },
        Item::RunLink(story, paragraph, run) => (
            Change::SetRunLink {
                story: story.clone(),
                paragraph: *paragraph,
                run: *run,
                href: None,
                title: None,
            },
            scene
                .story(story)
                .and_then(|story| story.paragraphs().get(*paragraph))
                .and_then(|paragraph| paragraph.runs().get(*run))
                .ok_or_else(missing)?
                .range(),
        ),
        Item::LineBreak(story, paragraph, index) => {
            let paragraph_node = scene
                .story(story)
                .and_then(|story| story.paragraphs().get(*paragraph))
                .ok_or_else(missing)?;
            let range =
                litchi_odp::scene::parts::line_breaks(node.file().content_xml()?, paragraph_node)
                    .get(*index)
                    .cloned()
                    .ok_or_else(missing)?;
            (
                Change::RemoveLineBreak {
                    story: story.clone(),
                    paragraph: *paragraph,
                    index: *index,
                },
                range,
            )
        },
        _ => {
            return Err(unsupported_target(
                "ODP remove supports slides, drawing objects, paragraphs, runs, hyperlinks and line breaks",
            ));
        },
    })
}

/// Move a slide, or a top-level object within its slide's z-order.
pub(crate) fn move_to_with_limits<'p, 'q>(
    file: &Odp,
    path: PathArg<'p>,
    placement: Placement<'q>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.package()?;
    let nav = Nav::new(file);
    let node = crate::path::resolve::get(&nav, &path)?;
    let scene = file.scene()?;
    let change = match node.item() {
        Item::Slide(page) => {
            let count = scene.pages().len();
            let to = destination(file, *page, count, &placement, Kind::Slide)?;
            Change::MovePage { page: *page, to }
        },
        Item::Object(object, _) if object.scope == Scope::Page && object.path.len() == 1 => {
            let count = scene
                .page(object.page)
                .map_or(0, |page| page.objects().len());
            let from = object.path[0];
            let to = destination(file, from, count, &placement, Kind::Shape)?;
            Change::MoveObject {
                object: object.clone(),
                to,
            }
        },
        _ => {
            return Err(unsupported_target(
                "ODP move supports slides and top-level slide objects",
            ));
        },
    };
    commit_changes(file, &[change], limits)
}

fn destination(
    file: &Odp,
    from: usize,
    count: usize,
    placement: &Placement<'_>,
    kind: Kind,
) -> Result<usize> {
    let last = count.saturating_sub(1);
    let anchor = |path: &PathArg<'_>| -> Result<usize> {
        let node = crate::path::resolve::get(&Nav::new(file), path)?;
        match (node.item(), kind) {
            (Item::Slide(page), Kind::Slide) => Ok(*page),
            (Item::Object(object, _), Kind::Shape) if object.path.len() == 1 => Ok(object.path[0]),
            _ => Err(Error::new(
                Code::UnsupportedType,
                "the ODP move anchor is not a sibling of the moved element",
            )
            .with_format(FileFormat::Odp)),
        }
    };
    let to = match placement {
        Placement::Start => 0,
        Placement::End => last,
        Placement::Position(position) => position.get(),
        Placement::Before(path) => {
            let anchor = anchor(path)?;
            if from < anchor { anchor - 1 } else { anchor }
        },
        Placement::After(path) => {
            let anchor = anchor(path)?;
            if from < anchor { anchor } else { anchor + 1 }
        },
        Placement::Into(_, position) => position.map_or(last, Position::get),
    };
    if to >= count {
        return Err(out_of_range("the ODP move position is outside its parent"));
    }
    Ok(to)
}

/// Exchange two slides, or two top-level objects of one slide.
pub(crate) fn swap_with_limits<'p, 'q>(
    file: &Odp,
    left: PathArg<'p>,
    right: PathArg<'q>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.package()?;
    let nav = Nav::new(file);
    let left = crate::path::resolve::get(&nav, &left)?;
    let right = crate::path::resolve::get(&nav, &right)?;
    let change = match (left.item(), right.item()) {
        (Item::Slide(first), Item::Slide(second)) => Change::SwapPages {
            first: *first,
            second: *second,
        },
        (Item::Object(first, _), Item::Object(second, _))
            if first.scope == Scope::Page
                && second.scope == Scope::Page
                && first.page == second.page
                && first.path.len() == second.path.len() =>
        {
            Change::SwapObjects {
                first: first.clone(),
                second: second.clone(),
            }
        },
        _ => {
            return Err(unsupported_target(
                "ODP swap supports two slides or two objects of one slide",
            ));
        },
    };
    commit_changes(file, &[change], limits)
}

// ---------------------------------------------------------------------------
// Target resolution
// ---------------------------------------------------------------------------

/// Resolve a path or a scoped query to distinct nodes.
pub(crate) fn resolve_nodes<'f>(
    file: &'f Odp,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<Node<'f>>> {
    let nav = Nav::new(file);
    if let Some(path) = target.path() {
        return Ok(vec![crate::path::resolve::get(&nav, path)?]);
    }
    let mut nodes = Vec::new();
    let mut seen = HashSet::new();
    let mut push = |node: Node<'f>| -> Result<()> {
        if !seen.insert(node.item().clone()) {
            return Ok(());
        }
        if nodes.len() >= limits.max_expanded_targets() {
            return Err(expanded_limit(
                nodes.len() + 1,
                limits.max_expanded_targets(),
            ));
        }
        nodes.push(node);
        Ok(())
    };
    if let Some(query) = target.query() {
        let walk = crate::query::run(nav, query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(Error::new(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            )
            .with_format(FileFormat::Odp));
        }
        for node in walk {
            push(node?)?;
        }
    } else {
        return Err(
            Error::new(Code::InvalidArgument, "the mutation target is empty")
                .with_format(FileFormat::Odp),
        );
    }
    if nodes.is_empty() {
        return Err(
            Error::new(Code::ZeroMatches, "mutation selector matched no element")
                .with_format(FileFormat::Odp),
        );
    }
    Ok(nodes)
}

fn bound(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
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

fn expanded_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::ExpandedTargets,
        observed,
        maximum,
    })
    .with_format(FileFormat::Odp)
}

fn snapshot_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::SnapshotBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Odp)
}

fn unsupported_target(message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message).with_format(FileFormat::Odp)
}

fn out_of_range(message: &'static str) -> Error {
    Error::new(Code::OutOfRange, message).with_format(FileFormat::Odp)
}
