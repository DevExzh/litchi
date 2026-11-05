//! OpenDocument formulas (`.odf`): the MathML element tree and its
//! StarMath annotation, read from `litchi_odf_formula::Formula` and written
//! through `litchi_odf_formula::Edit`.

use litchi_core::FileFormat;
use litchi_core::selector::Path;
use litchi_odf_formula::{Element, NodePath, OpaqueStarMath, StarMathVersion};

use super::tree::{count, text};
use super::write::{
    Commit, Patch, bound, element_refusal, finish, key_refusal, refusal, required, resolve,
    staging_error,
};
use super::{Flavor, OdfTree, Owner, TreeBuilder, element, odf_staging_error, writable};
use crate::Result;
use crate::error::Code;
use crate::new_node::NewNode;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::path::PathArg;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::ValueType;

kinds! {
    /// The element kinds of an ODF formula.
    Kind: &crate::path::vocabulary::MATH, root Math {
        Math => "math",
        Element => "element",
    }
}

/// A readable property `add` accepts but `set` does not write.
const fn add_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.add = true;
    spec
}

const MATH_PROPS: &[PropSpec] = &[
    writable("starmath", ValueType::Text),
    PropSpec::new("starmathVersion", ValueType::Text),
    PropSpec::new("elementCount", ValueType::Int),
    PropSpec::new("display", ValueType::Text),
];

const ELEMENT_PROPS: &[PropSpec] = &[
    add_only("tag", ValueType::Text),
    super::addable("text", ValueType::Text),
    add_only("mathvariant", ValueType::Text),
    PropSpec::new("childCount", ValueType::Int),
];

static OPERATION_GAPS: [GapSpec; 3] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "element",
            context: "a MathML element with attributes other than `mathvariant`, or a subtree",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0292 formula subtree authoring",
        reason: "`add` creates one element from `tag`, `text` and `mathvariant`; nested subtrees are built by further adds, and other attributes need typed values",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "element",
            context: "the `text` of a layout (non-token) element",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 formula layout text",
        reason: "only MathML token elements hold character data",
    },
    super::METADATA_PATH_GAP,
];

/// The formula spec table.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Math,
    elements: &[
        element(
            Kind::Math,
            "math",
            &["math"],
            &[],
            &[Kind::Element],
            &["/"],
            MATH_PROPS,
        ),
        element(
            Kind::Element,
            "element",
            &["element", "node"],
            &[Kind::Math, Kind::Element],
            &[Kind::Element],
            &["/element[N]", "/element[N]/element[N]"],
            ELEMENT_PROPS,
        ),
    ],
    gaps: &[],
};

crate::node::spec::check_format_spec!(SPEC);

/// The formula flavor.
pub(crate) struct Formula;

impl Flavor for Formula {
    type Kind = Kind;
    const FORMAT: FileFormat = FileFormat::Odf;
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Math, Kind::Element];
}

/// The elements `set` writes, `add` creates and `remove` deletes.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &["math", "element"],
    addable: &["element"],
    removable: &["element"],
    gaps: &OPERATION_GAPS,
};

/// Whether `element` holds character data directly (a MathML token).
fn is_token(element: &Element) -> bool {
    matches!(
        element.local_name(),
        "mi" | "mn" | "mo" | "mtext" | "ms" | "mspace" | "mglyph" | "annotation"
    )
}

/// Append `element`'s element children under `parent`, depth first.
fn push_children(
    tree: &mut TreeBuilder<Kind>,
    parent: u32,
    element: &Element,
    route: &mut Vec<usize>,
) -> usize {
    let mut total = 0;
    for (index, child) in element.children().enumerate() {
        route.push(index);
        let token = is_token(child);
        let draft = tree
            .push(parent, Kind::Element)
            .prop("tag", text(Some(child.local_name())))
            .prop("mathvariant", text(child.attribute(None, "mathvariant")))
            .prop("childCount", count(child.children().count()))
            .route(route);
        let draft = if token {
            let value = child.all_text();
            draft.prop("text", text(Some(&value))).text(value)
        } else {
            draft
        };
        let index = draft.done();
        total += 1 + push_children(tree, index, child, route);
        route.pop();
    }
    total
}

/// Build the tree of `package`.
pub(crate) fn build(package: &litchi_odf_formula::Package) -> Result<super::Tree> {
    let formula = package.formula();
    let root = formula.root();
    let mut tree = TreeBuilder::new(SPEC.elements, &[]);
    let mut route = Vec::new();
    let total = push_children(&mut tree, super::tree::ROOT, root, &mut route);
    let starmath = formula.starmath();
    let version = starmath
        .as_ref()
        .map(|annotation| annotation.opaque().version().encoding());
    let source = starmath
        .as_ref()
        .map(|annotation| annotation.opaque().source().to_owned());
    tree.root()
        .prop("starmath", text(source.as_deref()))
        .prop("starmathVersion", text(version))
        .prop("elementCount", count(total))
        .prop("display", text(root.attribute(None, "display")));
    if let Some(source) = source {
        tree.root().text(source);
    }
    Ok(tree.build())
}

/// Set element text or the StarMath annotation.
pub(crate) fn set(
    file: &OdfTree,
    package: &litchi_odf_formula::Package,
    groups: &[super::write::Group<'_, '_, '_>],
    limits: &SemanticLimits,
) -> Result<Commit> {
    let mut edit = package.edit();
    let mut operations = 0usize;
    let staging = |error| odf_staging_error(error, FileFormat::Odf);
    for (node, props) in groups
        .iter()
        .flat_map(|(nodes, props)| nodes.iter().map(move |node| (node, *props)))
    {
        for (key, value) in props.iter() {
            let key = key.as_str();
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            match (node.item().type_name(), key) {
                ("element", "text") => {
                    let value = super::write::text(value, "text")?.unwrap_or_default();
                    if node.item().text() == Some(value.as_str()) {
                        continue;
                    }
                    if node.item().text().is_none() {
                        return Err(refusal(
                            Code::UnsupportedCapability,
                            "only MathML token elements hold text",
                        )
                        .with_property("text".to_owned()));
                    }
                    edit.set_text(&NodePath::new(node.item().route()), &value)
                        .map_err(staging)?;
                },
                ("math", "starmath") => {
                    let source = required(super::write::text(value, "starmath")?, "starmath")?;
                    if node.item().text() == Some(source.as_str()) {
                        continue;
                    }
                    let version = package
                        .formula()
                        .starmath()
                        .map_or(StarMathVersion::V5, |annotation| {
                            annotation.opaque().version()
                        });
                    let opaque = OpaqueStarMath::new(version, source).map_err(staging)?;
                    edit.set_starmath(&opaque).map_err(staging)?;
                },
                ("element" | "math", other) => return Err(key_refusal(node, other)),
                _ => return Err(element_refusal("set", node)),
            }
        }
    }
    commit(file, edit, operations, limits)
}

fn commit(
    file: &OdfTree,
    edit: litchi_odf_formula::Edit<'_>,
    operations: usize,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let commit = edit
        .commit()
        .map_err(|error| super::write::commit_error(error, FileFormat::Odf))?;
    let changed = commit.diagnostics().changed();
    let (package, patch) = commit.into_parts();
    finish(
        file,
        changed,
        Owner::Formula(package),
        Patch::Formula(patch),
        operations,
        limits,
    )
}

/// The element at `route` under `root`.
fn element_at<'e>(root: &'e Element, route: &[usize]) -> Result<&'e Element> {
    let mut element = root;
    for index in route {
        element = element
            .children()
            .nth(*index)
            .ok_or_else(|| refusal(Code::InternalError, "the MathML element disappeared"))?;
    }
    Ok(element)
}

/// The insertion index an `add` placement names among `count` children of
/// `parent` (whose route is `parent_route`).
fn insertion(
    file: &OdfTree,
    parent_route: &[usize],
    placement: Option<&crate::Placement<'_>>,
    count: usize,
) -> Result<usize> {
    use crate::Placement;
    let anchor = |path: &PathArg<'_>| -> Result<usize> {
        let anchor = super::write::parent::<Formula>(file, path)?;
        match anchor.item().route().split_last() {
            Some((index, route)) if route == parent_route => Ok(*index),
            _ => Err(refusal(
                Code::UnsupportedType,
                "the placement anchor is not a child of the parent element",
            )),
        }
    };
    let index = match placement {
        None | Some(Placement::End) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => position.get(),
        Some(Placement::Before(path)) => anchor(path)?,
        Some(Placement::After(path)) => anchor(path)? + 1,
        Some(Placement::Into(_, position)) => position.map_or(count, litchi_core::Position::get),
    };
    if index > count {
        return Err(refusal(
            Code::OutOfRange,
            "the MathML insertion position is outside its parent",
        ));
    }
    Ok(index)
}

/// Add one MathML element (`tag` required; `text` for token elements;
/// `mathvariant`) under `/` or another element (record 0292).
pub(crate) fn add(
    file: &OdfTree,
    package: &litchi_odf_formula::Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    if !matches!(
        node.kind().to_ascii_lowercase().as_str(),
        "element" | "node"
    ) {
        return Err(refusal(
            Code::UnsupportedCapability,
            format!(
                "adding `{}` elements to an ODF formula is not supported",
                node.kind()
            ),
        )
        .with_suggestion("add a MathML element: NewNode::new(\"element\").prop(\"tag\", \"mi\")"));
    }
    let parent_node = super::write::parent::<Formula>(file, parent)?;
    let parent_route = parent_node.item().route().to_vec();
    let mut tag = None;
    let mut content = None;
    let mut variant = None;
    for (key, value) in node.props().iter() {
        match key.as_str() {
            "tag" => tag = super::write::text(value, "tag")?,
            "text" => content = super::write::text(value, "text")?,
            "mathvariant" => variant = super::write::text(value, "mathvariant")?,
            other => {
                return Err(refusal(
                    Code::UnsupportedProperty,
                    format!("`{other}` is not a property `add` accepts for a MathML element"),
                )
                .with_property(other.to_owned()));
            },
        }
    }
    let tag = required(tag, "tag")?;
    let staging = |error| odf_staging_error(error, FileFormat::Odf);
    let mut element = Element::new(&tag)
        .map_err(|error| staging_error(error, FileFormat::Odf))
        .map_err(|error| error.with_property("tag".to_owned()))?;
    if let Some(content) = content.filter(|content| !content.is_empty()) {
        if !is_token(&element) {
            return Err(refusal(
                Code::UnsupportedCapability,
                "only MathML token elements hold text",
            )
            .with_property("text".to_owned()));
        }
        element.push_text(&content);
    }
    if let Some(variant) = variant {
        element
            .set_attribute(None, "mathvariant", &variant)
            .map_err(|error| staging_error(error, FileFormat::Odf))?;
    }
    let count = element_at(package.formula().root(), &parent_route)?
        .children()
        .count();
    let index = insertion(file, &parent_route, node.placement(), count)?;
    let mut edit = package.edit();
    edit.insert_child(&NodePath::new(parent_route.clone()), index, element)
        .map_err(staging)?;
    let commit = commit(file, edit, 1, limits)?;
    let parent_path = parent_node.item().path().to_vec();
    let path = super::write::added_path(&commit.after, &parent_path, "element", Some(index))?;
    Ok((commit, path))
}

/// Move one MathML element among its siblings or under another element
/// (record 0292). The leaf validates the content model of the result.
pub(crate) fn move_to(
    file: &OdfTree,
    package: &litchi_odf_formula::Package,
    path: &PathArg<'_>,
    placement: &crate::Placement<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    use super::write::Destination;
    let node = super::write::parent::<Formula>(file, path)?;
    if node.item().type_name() != "element" {
        return Err(element_refusal("move", &node));
    }
    let route = node.item().route().to_vec();
    let Some((&from, parent_route)) = route.split_last() else {
        return Err(element_refusal("move", &node));
    };
    let siblings = super::write::siblings(file, &node)?;
    let (target, index) =
        match super::write::destination::<Formula>(file, &node, &siblings, placement)? {
            Destination::Within(to) => {
                if to == from {
                    return commit(file, package.edit(), 1, limits);
                }
                (parent_route.to_vec(), Some(to))
            },
            Destination::Across { parent, position } => {
                let mut target = parent.item().route().to_vec();
                if parent.item().type_name() == "element" && target.starts_with(&route) {
                    return Err(refusal(
                        Code::InvalidArgument,
                        "a MathML element cannot move into its own subtree",
                    ));
                }
                // Removing the element shifts a later sibling subtree's route.
                let depth = parent_route.len();
                if target.len() > depth && target.starts_with(parent_route) && target[depth] > from
                {
                    target[depth] -= 1;
                }
                (target, position)
            },
        };
    let staging = |error| odf_staging_error(error, FileFormat::Odf);
    let mut edit = package.edit();
    let element = edit
        .remove(&NodePath::new(route.clone()))
        .map_err(staging)?;
    let count = element_at(edit.root(), &target)?.children().count();
    let index = index.unwrap_or(count);
    if index > count {
        return Err(refusal(
            Code::OutOfRange,
            "the MathML move position is outside its parent",
        ));
    }
    edit.insert_child(&NodePath::new(target), index, element)
        .map_err(staging)?;
    commit(file, edit, 1, limits)
}

/// Exchange two MathML elements, neither inside the other.
pub(crate) fn swap(
    file: &OdfTree,
    package: &litchi_odf_formula::Package,
    left: &PathArg<'_>,
    right: &PathArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let (left, right) = super::write::swap_pair::<Formula>(file, left, right)?;
    if left.item().type_name() != "element" {
        return Err(element_refusal("swap", &left));
    }
    let (left, right) = (left.item().route().to_vec(), right.item().route().to_vec());
    let mut edit = package.edit();
    if left != right {
        if left.starts_with(&right) || right.starts_with(&left) {
            return Err(refusal(
                Code::InvalidArgument,
                "a MathML element cannot swap with its own ancestor",
            ));
        }
        let root = package.formula().root();
        let (first, second) = (
            element_at(root, &left)?.clone(),
            element_at(root, &right)?.clone(),
        );
        let staging = |error| odf_staging_error(error, FileFormat::Odf);
        edit.replace(&NodePath::new(left), second)
            .map_err(staging)?;
        edit.replace(&NodePath::new(right), first)
            .map_err(staging)?;
    }
    commit(file, edit, 2, limits)
}

/// Remove MathML elements.
pub(crate) fn remove(
    file: &OdfTree,
    package: &litchi_odf_formula::Package,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Formula>(file, target, limits.max_expanded_targets())?;
    bound(LimitKind::Operations, nodes.len(), limits.max_operations())?;
    let mut routes = Vec::with_capacity(nodes.len());
    for node in &nodes {
        if node.item().type_name() != "element" {
            return Err(element_refusal("remove", node));
        }
        routes.push(node.item().route().to_vec());
    }
    // Later siblings first, so earlier removals do not shift later paths.
    routes.sort_unstable_by(|left, right| right.cmp(left));
    let mut edit = package.edit();
    for route in routes {
        edit.remove(&NodePath::new(route))
            .map_err(|error| odf_staging_error(error, FileFormat::Odf))?;
    }
    commit(file, edit, nodes.len(), limits)
}
