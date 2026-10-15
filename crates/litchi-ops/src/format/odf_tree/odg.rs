//! OpenDocument drawings (`.odg`, `.otg`): pages, their shapes and layers,
//! read from the `litchi_odg::Drawing` projection and written through
//! `litchi_odg::edit::Edit`.

use litchi_core::selector::Path;
use litchi_core::{FileFormat, Position};
use litchi_odg::shape::ShapeKind;

use super::tree::{count, length, odf_length, text};
use super::write::{
    Commit, Patch, added_path, bound, element_refusal, finish, key_refusal, refusal, required,
    resolve, staging_error,
};
use super::{Flavor, OdfTree, Owner, TreeBuilder, addable, element};
use crate::Result;
use crate::error::Code;
use crate::new_node::NewNode;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::path::PathArg;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::{LengthSpec, ValueType};

kinds! {
    /// The element kinds of an ODG drawing.
    Kind: &crate::path::vocabulary::DRAWING, root Drawing {
        Drawing => "drawing",
        Page => "page",
        Shape => "shape",
        Layer => "layer",
    }
}

const DRAWING_PROPS: &[PropSpec] = &[
    PropSpec::new("pageCount", ValueType::Int),
    PropSpec::new("isTemplate", ValueType::Bool),
    PropSpec::new("title", ValueType::Text),
    PropSpec::new("author", ValueType::Text),
    PropSpec::new("subject", ValueType::Text),
    PropSpec::new("description", ValueType::Text),
];

const PAGE_PROPS: &[PropSpec] = &[
    addable("name", ValueType::Text),
    PropSpec::new("style", ValueType::Text),
    PropSpec::new("master", ValueType::Text),
    PropSpec::new("shapeCount", ValueType::Int),
];

const SHAPE_PROPS: &[PropSpec] = &[
    addable("name", ValueType::Text),
    PropSpec::new("shapeType", ValueType::Text),
    addable("text", ValueType::Text),
    addable("x", ValueType::Length(LengthSpec::POSITION)),
    addable("y", ValueType::Length(LengthSpec::POSITION)),
    addable("width", ValueType::Length(LengthSpec::EXTENT)),
    addable("height", ValueType::Length(LengthSpec::EXTENT)),
    PropSpec::new("layer", ValueType::Text),
    PropSpec::new("style", ValueType::Text),
    PropSpec::new("alt", ValueType::Text),
    PropSpec::new("description", ValueType::Text),
];

const LAYER_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("display", ValueType::Text),
];

static OPERATION_GAPS: [GapSpec; 5] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "shape",
            context: "a shape other than a text rectangle",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 ODG typed shape creation",
        reason: "the adapter adds text rectangles; other geometries need typed shape values",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "shape",
            context: "a group member, or a shape whose identity a connector or link on its page names, into another page",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0292 ODG cross-page reference rewriting",
        reason: "a shape moves to another page with its bytes unchanged; group members move with their group, and an identity dependency that would cross pages is refused rather than left dangling",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "swap",
            element: "shape",
            context: "two groups on different pages",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0292 ODG cross-page group swap",
        reason: "a cross-page swap is two transfers addressed by position; two group subtrees shift each other's positions, so move them one at a time",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "page",
            context: "a page another page's `draw:page-name` names",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0288 ODG reference rewriting",
        reason: "the leaf refuses removing a page a link or thumbnail references rather than leaving it dangling",
    },
    super::METADATA_PATH_GAP,
];

/// The ODG spec table.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Drawing,
    elements: &[
        element(
            Kind::Drawing,
            "drawing",
            &["drawing"],
            &[],
            &[Kind::Page],
            &["/"],
            DRAWING_PROPS,
        ),
        element(
            Kind::Page,
            "page",
            &["page"],
            &[Kind::Drawing],
            &[Kind::Shape, Kind::Layer],
            &["/page[N]", "/page[NAME]"],
            PAGE_PROPS,
        ),
        element(
            Kind::Shape,
            "shape",
            &["shape"],
            &[Kind::Page],
            &[],
            &["/page[N]/shape[N]"],
            SHAPE_PROPS,
        ),
        element(
            Kind::Layer,
            "layer",
            &["layer"],
            &[Kind::Page],
            &[],
            &["/page[N]/layer[N]"],
            LAYER_PROPS,
        ),
    ],
    gaps: &[],
};

crate::node::spec::check_format_spec!(SPEC);

/// The ODG flavor.
pub(crate) struct Odg;

impl Flavor for Odg {
    type Kind = Kind;
    const FORMAT: FileFormat = FileFormat::Odg;
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Drawing, Kind::Page];
    const ABSENT: &'static [&'static str] = &["slide"];
}

/// The elements `set` writes, `add` creates and `remove` deletes.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &["page", "shape"],
    addable: &["page", "shape"],
    removable: &["page", "shape"],
    gaps: &OPERATION_GAPS,
};

/// The unified spelling of a shape family.
const fn shape_type(kind: ShapeKind) -> &'static str {
    match kind {
        ShapeKind::Caption => "caption",
        ShapeKind::Circle => "circle",
        ShapeKind::Connector => "connector",
        ShapeKind::Control => "control",
        ShapeKind::Custom => "customShape",
        ShapeKind::Ellipse => "ellipse",
        ShapeKind::Frame => "frame",
        ShapeKind::Group => "group",
        ShapeKind::Line => "line",
        ShapeKind::Measure => "measure",
        ShapeKind::Path => "path",
        ShapeKind::PageThumbnail => "pageThumbnail",
        ShapeKind::Polygon => "polygon",
        ShapeKind::Polyline => "polyline",
        ShapeKind::Rectangle => "rect",
        ShapeKind::RegularPolygon => "regularPolygon",
        _ => "scene3d",
    }
}

/// Build the tree of `package`.
pub(crate) fn build(package: &litchi_odg::Package) -> Result<super::Tree> {
    let drawing = package.drawing();
    let mut tree = TreeBuilder::new(SPEC.elements, &[]);
    let metadata = drawing.metadata();
    let field = |pick: fn(&litchi_core::Metadata) -> Option<&String>| {
        text(metadata.and_then(|metadata| pick(metadata).map(String::as_str)))
    };
    tree.root()
        .prop("pageCount", count(drawing.pages().len()))
        .prop("isTemplate", Some(drawing.is_template().into()))
        .prop("title", field(|metadata| metadata.title.as_ref()))
        .prop("author", field(|metadata| metadata.author.as_ref()))
        .prop("subject", field(|metadata| metadata.subject.as_ref()))
        .prop(
            "description",
            field(|metadata| metadata.description.as_ref()),
        );
    for (page_index, page) in drawing.pages().iter().enumerate() {
        let parent = tree
            .push(super::tree::ROOT, Kind::Page)
            .prop("name", text(page.name()))
            .prop("style", text(page.style_name()))
            .prop("master", text(page.master_page_name()))
            .prop("shapeCount", count(page.shapes().len()))
            .label(page.name())
            .leaf(page_index, 0)
            .done();
        for (shape_index, shape) in page.shapes().iter().enumerate() {
            let value = shape.text();
            let draft = tree
                .push(parent, Kind::Shape)
                .prop("name", text(shape.name()))
                .prop("shapeType", text(Some(shape_type(shape.kind()))))
                .prop("text", text(Some(value).filter(|value| !value.is_empty())))
                .prop("x", length(shape.x()))
                .prop("y", length(shape.y()))
                .prop("width", length(shape.width()))
                .prop("height", length(shape.height()))
                .prop("layer", text(shape.layer()))
                .prop("style", text(shape.style_name()))
                .prop("alt", text(shape.title()))
                .prop("description", text(shape.description()))
                .label(shape.name())
                .leaf(page_index, shape_index);
            if value.is_empty() {
                draft.done();
            } else {
                draft.text(value).done();
            }
        }
        for layer in page.layers() {
            tree.push(parent, Kind::Layer)
                .prop("name", text(Some(layer.name())))
                .prop("display", text(layer.display()))
                .label(Some(layer.name()))
                .done();
        }
    }
    Ok(tree.build())
}

fn staging(error: litchi_core::Error) -> crate::Error {
    staging_error(error, FileFormat::Odg)
}

/// Set page names and shape text, names and geometry.
pub(crate) fn set(
    file: &OdfTree,
    package: &litchi_odg::Package,
    groups: &[super::write::Group<'_, '_, '_>],
    limits: &SemanticLimits,
) -> Result<Commit> {
    let drawing = package.drawing();
    let mut edit = package.edit();
    let mut operations = 0usize;
    for (node, props) in groups
        .iter()
        .flat_map(|(nodes, props)| nodes.iter().map(move |node| (node, *props)))
    {
        let [page, shape] = node.item().leaf();
        let mut geometry: [Option<String>; 4] = [None, None, None, None];
        for (key, value) in props.iter() {
            let key = key.as_str();
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            match (node.item().type_name(), key) {
                ("page", "name") => {
                    let name = required(super::write::text(value, "name")?, "name")?;
                    if node.item().label() != Some(name.as_str()) {
                        edit.set_page_name(page, name).map_err(staging)?;
                    }
                },
                ("shape", "text") => {
                    let value = super::write::text(value, "text")?.unwrap_or_default();
                    if node.item().text().unwrap_or_default() != value {
                        edit.set_shape_text(page, shape, value).map_err(staging)?;
                    }
                },
                ("shape", "name") => {
                    let name = required(super::write::text(value, "name")?, "name")?;
                    if node.item().label() != Some(name.as_str()) {
                        edit.set_shape_name(page, shape, name).map_err(staging)?;
                    }
                },
                ("shape", key @ ("x" | "y" | "width" | "height")) => {
                    let (slot, spec, name) = match key {
                        "x" => (0, LengthSpec::POSITION, "x"),
                        "y" => (1, LengthSpec::POSITION, "y"),
                        "width" => (2, LengthSpec::EXTENT, "width"),
                        _ => (3, LengthSpec::EXTENT, "height"),
                    };
                    let length = required(super::write::length(value, name, spec)?, name)?;
                    geometry[slot] = Some(odf_length(length));
                },
                ("page" | "shape" | "drawing" | "layer", other) => {
                    return Err(key_refusal(node, other));
                },
                _ => return Err(element_refusal("set", node)),
            }
        }
        if geometry.iter().any(Option::is_some) {
            let source = drawing
                .pages()
                .get(page)
                .and_then(|value| value.shapes().get(shape))
                .ok_or_else(|| refusal(Code::InternalError, "the ODG shape disappeared"))?;
            let current = [source.x(), source.y(), source.width(), source.height()];
            let mut merged = Vec::with_capacity(4);
            for (wanted, held) in geometry.into_iter().zip(current) {
                match wanted.or_else(|| held.map(str::to_owned)) {
                    Some(value) => merged.push(value),
                    None => {
                        return Err(refusal(
                            Code::UnsupportedCapability,
                            "the ODG shape has no complete x/y/width/height geometry to update",
                        ));
                    },
                }
            }
            if merged
                .iter()
                .map(String::as_str)
                .ne(current.iter().map(|value| value.unwrap_or_default()))
            {
                let [x, y, width, height]: [String; 4] = merged
                    .try_into()
                    .map_err(|_| refusal(Code::InternalError, "ODG geometry arity"))?;
                edit.set_shape_geometry(page, shape, x, y, width, height)
                    .map_err(staging)?;
            }
        }
    }
    commit(file, edit, operations, limits)
}

fn commit(
    file: &OdfTree,
    edit: litchi_odg::edit::Edit,
    operations: usize,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let commit = edit
        .commit()
        .map_err(|error| super::write::commit_error(error, FileFormat::Odg))?;
    let changed = commit.diagnostics().changed();
    let (package, patch) = commit.into_parts();
    finish(
        file,
        changed,
        Owner::Odg(package),
        Patch::Odg(patch),
        operations,
        limits,
    )
}

/// Append a page, or a text rectangle to a page.
pub(crate) fn add(
    file: &OdfTree,
    package: &litchi_odg::Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let parent_node = super::write::parent::<Odg>(file, parent)?;
    let mut props = std::collections::HashMap::new();
    for (key, value) in node.props().iter() {
        props.insert(key.as_str().to_owned(), value);
    }
    let mut edit = package.edit();
    let kind = node.kind().to_ascii_lowercase();
    let parent_path = parent_node.item().path().to_vec();
    match (kind.as_str(), parent_node.item().type_name()) {
        ("page", "drawing") => {
            let name = match props.remove("name") {
                Some(value) => required(super::write::text(value, "name")?, "name")?,
                None => format!("page{}", package.drawing().pages().len() + 1),
            };
            reject_rest(&props, "page")?;
            let mut page = litchi_odg::page::Page::new(name);
            if let Some(last) = package.drawing().pages().last() {
                if let Some(master) = last.master_page_name() {
                    page = page.with_master_page_name(master);
                }
                if let Some(style) = last.style_name() {
                    page = page.with_style_name(style);
                }
            }
            let count = package.drawing().pages().len();
            let position = insertion(file, node.placement(), count, Kind::Page)?;
            edit.insert_page(position, page).map_err(staging)?;
            let commit = commit(file, edit, 1, limits)?;
            let path = added_path(&commit.after, &parent_path, "page", Some(position))?;
            Ok((commit, path))
        },
        ("shape" | "rect" | "rectangle" | "textbox", "page") => {
            let mut geometry = Vec::with_capacity(4);
            for (key, spec) in [
                ("x", LengthSpec::POSITION),
                ("y", LengthSpec::POSITION),
                ("width", LengthSpec::EXTENT),
                ("height", LengthSpec::EXTENT),
            ] {
                let value = props.remove(key).ok_or_else(|| {
                    refusal(Code::InvalidArgument, format!("an ODG shape needs `{key}`"))
                        .with_property(key.to_owned())
                })?;
                geometry.push(odf_length(required(
                    super::write::length(value, key, spec)?,
                    key,
                )?));
            }
            let mut shape = litchi_odg::shape::Shape::new(ShapeKind::Rectangle).with_geometry(
                geometry[0].clone(),
                geometry[1].clone(),
                geometry[2].clone(),
                geometry[3].clone(),
            );
            if let Some(value) = props.remove("text") {
                if let Some(value) = super::write::text(value, "text")? {
                    shape = shape.with_text(value);
                }
            }
            if let Some(value) = props.remove("name") {
                if let Some(value) = super::write::text(value, "name")? {
                    shape = shape.with_name(value);
                }
            }
            reject_rest(&props, "shape")?;
            let page = parent_node.item().leaf()[0];
            let count = package
                .drawing()
                .pages()
                .get(page)
                .map_or(0, |value| value.shapes().len());
            let position = insertion(file, node.placement(), count, Kind::Shape)?;
            edit.insert_shape(page, position, shape).map_err(staging)?;
            let commit = commit(file, edit, 1, limits)?;
            let path = added_path(&commit.after, &parent_path, "shape", Some(position))?;
            Ok((commit, path))
        },
        _ => Err(refusal(
            Code::UnsupportedCapability,
            format!(
                "adding `{}` under `{}` is not supported for ODG drawings",
                node.kind(),
                parent_node.item().type_name()
            ),
        )
        .with_suggestion("add a page under `/` or a shape under `/page[N]`")),
    }
}

fn reject_rest(
    props: &std::collections::HashMap<String, &crate::props::PropValue<'_>>,
    element: &str,
) -> Result<()> {
    if let Some(key) = props.keys().next() {
        return Err(refusal(
            Code::UnsupportedProperty,
            format!("`{key}` is not a property `add` accepts for an ODG {element}"),
        )
        .with_property(key.clone()));
    }
    Ok(())
}

/// Remove shapes.
pub(crate) fn remove(
    file: &OdfTree,
    package: &litchi_odg::Package,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Odg>(file, target, limits.max_expanded_targets())?;
    bound(LimitKind::Operations, nodes.len(), limits.max_operations())?;
    let mut pages = Vec::new();
    let mut shapes = Vec::with_capacity(nodes.len());
    for node in &nodes {
        match node.item().type_name() {
            "page" => pages.push(node.item().leaf()[0]),
            "shape" => shapes.push(node.item().leaf()),
            _ => return Err(element_refusal("remove", node)),
        }
    }
    // Shapes of removed pages go with their page.
    shapes.retain(|[page, _]| !pages.contains(page));
    // Later elements first, so earlier removals do not shift later
    // positions.
    shapes.sort_unstable_by(|left, right| right.cmp(left));
    shapes.dedup();
    pages.sort_unstable_by(|left, right| right.cmp(left));
    pages.dedup();
    if !pages.is_empty() && pages.len() >= package.drawing().pages().len() {
        return Err(refusal(
            Code::InvalidArgument,
            "an ODG drawing keeps at least one page",
        ));
    }
    let mut edit = package.edit();
    for [page, shape] in shapes {
        edit.remove_shape(page, shape).map_err(staging)?;
    }
    for page in pages {
        edit.remove_page(page).map_err(staging)?;
    }
    commit(file, edit, nodes.len(), limits)
}

/// The final position an `add` placement names among `count` siblings.
fn insertion(
    file: &OdfTree,
    placement: Option<&crate::Placement<'_>>,
    count: usize,
    kind: Kind,
) -> Result<usize> {
    use crate::Placement;
    let index = match placement {
        None | Some(Placement::End) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => position.get(),
        Some(Placement::Before(path)) => anchor(file, path, kind)?,
        Some(Placement::After(path)) => anchor(file, path, kind)? + 1,
        Some(Placement::Into(_, position)) => position.map_or(count, Position::get),
    };
    if index > count {
        return Err(refusal(
            Code::OutOfRange,
            "the ODG insertion position is outside its parent",
        ));
    }
    Ok(index)
}

/// The sibling position of the placement anchor at `path`.
fn anchor(file: &OdfTree, path: &PathArg<'_>, kind: Kind) -> Result<usize> {
    let node = super::write::parent::<Odg>(file, path)?;
    let [page, shape] = node.item().leaf();
    match (node.item().type_name(), kind) {
        ("page", Kind::Page) => Ok(page),
        ("shape", Kind::Shape) => Ok(shape),
        _ => Err(refusal(
            Code::UnsupportedType,
            "the ODG placement anchor is not a sibling of the element",
        )),
    }
}

/// Move a page within the drawing, or a shape within its page or into
/// another page (record 0292).
pub(crate) fn move_to(
    file: &OdfTree,
    package: &litchi_odg::Package,
    path: &PathArg<'_>,
    placement: &crate::Placement<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    use super::write::Destination;
    let node = super::write::parent::<Odg>(file, path)?;
    let [page, shape] = node.item().leaf();
    let kind = node.item().type_name();
    if !matches!(kind, "page" | "shape") {
        return Err(element_refusal("move", &node));
    }
    let siblings = super::write::siblings(file, &node)?;
    let destination = super::write::destination::<Odg>(file, &node, &siblings, placement)?;
    let mut edit = package.edit();
    match (kind, destination) {
        ("page", Destination::Within(to)) => edit.move_page(page, to).map_err(staging)?,
        ("shape", Destination::Within(to)) => {
            edit.move_shape(page, shape, to).map_err(staging)?;
        },
        ("shape", Destination::Across { parent, position }) => {
            if parent.item().type_name() != "page" {
                return Err(refusal(
                    Code::UnsupportedType,
                    "an ODG shape moves into a page",
                ));
            }
            let to_page = parent.item().leaf()[0];
            let count = package
                .drawing()
                .pages()
                .get(to_page)
                .map_or(0, |value| value.shapes().len());
            let to = position.unwrap_or(count);
            if to > count {
                return Err(refusal(
                    Code::OutOfRange,
                    "the ODG move position is outside the destination page",
                ));
            }
            edit.move_shape_to_page(page, shape, to_page, to)
                .map_err(staging)?;
        },
        _ => {
            return Err(refusal(
                Code::UnsupportedType,
                "an ODG page moves among the drawing's pages",
            ));
        },
    }
    commit(file, edit, 1, limits)
}

/// Exchange two pages, or two shapes (of one page or of two pages).
pub(crate) fn swap(
    file: &OdfTree,
    package: &litchi_odg::Package,
    left: &PathArg<'_>,
    right: &PathArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let (left, right) = super::write::swap_pair::<Odg>(file, left, right)?;
    let [left_page, left_shape] = left.item().leaf();
    let [right_page, right_shape] = right.item().leaf();
    let mut edit = package.edit();
    match left.item().type_name() {
        "page" => {
            let (first, second) = (left_page.min(right_page), left_page.max(right_page));
            if first != second {
                edit.move_page(first, second).map_err(staging)?;
                edit.move_page(second - 1, first).map_err(staging)?;
            }
        },
        "shape" if left_page == right_page => {
            let (first, second) = (left_shape.min(right_shape), left_shape.max(right_shape));
            if first != second {
                edit.move_shape(left_page, first, second).map_err(staging)?;
                edit.move_shape(left_page, second - 1, first)
                    .map_err(staging)?;
            }
        },
        "shape" => {
            // The first transfer lands just before the other shape, which
            // shifts by one; a group would shift it by its whole subtree.
            let is_group = |node: &super::Node<'_>| {
                node.item()
                    .value("shapeType")
                    .and_then(crate::node::Value::as_str)
                    == Some("group")
            };
            let ((a_page, a_shape), (b_page, b_shape)) = if !is_group(&left) {
                ((left_page, left_shape), (right_page, right_shape))
            } else if !is_group(&right) {
                ((right_page, right_shape), (left_page, left_shape))
            } else {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "swapping two ODG groups across pages is not supported; move them one at a time",
                )
                .with_suggestion("follow-up: 0292 ODG cross-page group swap"));
            };
            edit.move_shape_to_page(a_page, a_shape, b_page, b_shape)
                .map_err(staging)?;
            edit.move_shape_to_page(b_page, b_shape + 1, a_page, a_shape)
                .map_err(staging)?;
        },
        _ => return Err(element_refusal("swap", &left)),
    }
    commit(file, edit, 2, limits)
}
