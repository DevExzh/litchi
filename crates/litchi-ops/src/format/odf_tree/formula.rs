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
use super::{Flavor, OdfTree, Owner, TreeBuilder, element, writable};
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

const MATH_PROPS: &[PropSpec] = &[
    writable("starmath", ValueType::Text),
    PropSpec::new("starmathVersion", ValueType::Text),
    PropSpec::new("elementCount", ValueType::Int),
    PropSpec::new("display", ValueType::Text),
];

const ELEMENT_PROPS: &[PropSpec] = &[
    PropSpec::new("tag", ValueType::Text),
    writable("text", ValueType::Text),
    PropSpec::new("mathvariant", ValueType::Text),
    PropSpec::new("childCount", ValueType::Int),
];

static OPERATION_GAPS: [GapSpec; 3] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "element",
            context: "a new MathML element",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 formula element authoring",
        reason: "inserting MathML needs a typed element value the unified registry does not have yet",
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
    addable: &[],
    removable: &["element"],
    movable: &[],
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
    let staging = |error| staging_error(error, FileFormat::Odf);
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

/// Adding MathML needs typed element values.
pub(crate) fn add(
    file: &OdfTree,
    package: &litchi_odf_formula::Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let _ = (file, package, parent, limits);
    Err(refusal(
        Code::UnsupportedCapability,
        format!(
            "adding `{}` elements to an ODF formula is not supported",
            node.kind()
        ),
    ))
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
            .map_err(|error| staging_error(error, FileFormat::Odf))?;
    }
    commit(file, edit, nodes.len(), limits)
}
