//! Apple Keynote presentations (`.key`): the slides, their title and body
//! placeholders and speaker notes, read from `litchi_keynote::Package` and
//! written through its slide-text, slide-skip, slide-order and
//! slide-deletion transactions (record 0296).

use litchi_core::{FileFormat, Position};
use litchi_keynote::slide::placeholder::Kind as Role;
use litchi_keynote::{Package, SlideSelector};

use super::tree::{count, text};
use super::write::{
    Chain, Commit, LeafPatch, bound, destination, element_refusal, finish, key_refusal,
    leaf_refusal, node_at, refusal, resolve,
};
use super::{Flavor, Iwork, LeafError, Owner, TreeBuilder, element, writable};
use crate::Result;
use crate::error::Code;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::path::PathArg;
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::ValueType;

kinds! {
    /// The element kinds of a Keynote presentation.
    Kind: &crate::path::vocabulary::PRESENTATION, root Presentation {
        Presentation => "presentation",
        Slide => "slide",
        Placeholder => "placeholder",
        Notes => "notes",
    }
}

const PRESENTATION_PROPS: &[PropSpec] = &[PropSpec::new("slideCount", ValueType::Int)];

pub(crate) const SLIDE_PROPS: &[PropSpec] = &[
    PropSpec::new("name", ValueType::Text),
    PropSpec::new("title", ValueType::Text),
    writable("hidden", ValueType::Bool),
];

pub(crate) const PLACEHOLDER_PROPS: &[PropSpec] = &[
    PropSpec::new("type", ValueType::Text),
    writable("text", ValueType::Text),
];

const NOTES_PROPS: &[PropSpec] = &[PropSpec::new("text", ValueType::Text)];

static GAPS: [GapSpec; 3] = [
    GapSpec {
        subject: Subject::Element("shape"),
        code: Code::UnsupportedCapability,
        follow_up: "0298 Keynote shapes (IWORK_API_REQUESTS.md request 9)",
        reason: "the Keynote leaf has no public shape model beyond the title and body placeholders",
    },
    GapSpec {
        subject: Subject::Element("run"),
        code: Code::UnsupportedCapability,
        follow_up: "0298 Keynote text runs (IWORK_API_REQUESTS.md request 10)",
        reason: "the Keynote leaf exposes placeholder text as plain text; character runs and their formatting have no public model yet",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "notes",
            context: "speaker notes text",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0298 Keynote notes writes",
        reason: "speaker notes are read-only here; the leaf's slide-notes transaction is not wired yet",
    },
];

static OPERATION_GAPS: [GapSpec; 2] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "slide",
            context: "a new slide",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 8",
        reason: "the Keynote leaf has no slide creation transaction",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "create",
            element: "presentation",
            context: "a presentation from scratch",
        },
        code: Code::UnsupportedCapability,
        follow_up: "IWORK_API_REQUESTS.md request 1",
        reason: "the Keynote leaf has no from-scratch package builder",
    },
];

/// The Keynote spec table.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Presentation,
    elements: &[
        element(
            Kind::Presentation,
            "presentation",
            &["presentation"],
            &[],
            &[Kind::Slide],
            &["/"],
            PRESENTATION_PROPS,
        ),
        element(
            Kind::Slide,
            "slide",
            &["slide"],
            &[Kind::Presentation],
            &[Kind::Placeholder, Kind::Notes],
            &["/slide[N]"],
            SLIDE_PROPS,
        ),
        element(
            Kind::Placeholder,
            "placeholder",
            &["ph", "placeholder"],
            &[Kind::Slide],
            &[],
            &[
                "/slide[N]/placeholder[title]",
                "/slide[N]/placeholder[body]",
            ],
            PLACEHOLDER_PROPS,
        ),
        element(
            Kind::Notes,
            "notes",
            &["notes"],
            &[Kind::Slide],
            &[],
            &["/slide[N]/notes"],
            NOTES_PROPS,
        ),
    ],
    gaps: &GAPS,
};

crate::node::spec::check_format_spec!(SPEC);

/// The Keynote flavor.
pub(crate) struct Keynote;

impl Flavor for Keynote {
    type Kind = Kind;
    const FORMAT: FileFormat = FileFormat::Keynote;
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Presentation];
    const SINGLETONS: &'static [Kind] = &[Kind::Notes];
}

/// The elements `set` writes, `remove` deletes and `move` reorders.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &["placeholder", "slide"],
    addable: &[],
    removable: &["slide"],
    movable: &["slide"],
    gaps: &OPERATION_GAPS,
};

/// The leaf role a placeholder item's `leaf[1]` records.
const TITLE: usize = 0;
const BODY: usize = 1;

/// Build the tree of `package`.
pub(crate) fn build(package: &Package) -> Result<super::Tree> {
    let read = |error| leaf_refusal(LeafError::Keynote(error), FileFormat::Keynote);
    let slides = package.slides().map_err(read)?;
    let mut tree = TreeBuilder::new(SPEC.elements, Keynote::SINGLETONS);
    tree.root().prop("slideCount", count(slides.len())).done();
    for (index, slide) in slides.iter().enumerate() {
        let item = tree
            .push(super::tree::ROOT, Kind::Slide)
            .prop("name", text(slide.name()))
            .prop("title", text(slide.title()))
            .prop("hidden", Some(slide.is_skipped().into()))
            .label(slide.name())
            .leaf([index, 0, 0, 0])
            .done();
        // The placeholders the leaf's slide-text transaction addresses; a
        // slide without one has no such child.
        let roles = [(TITLE, Role::Title, "title"), (BODY, Role::Body, "body")];
        for (slot, role, name) in roles {
            let selector = SlideSelector::index(index);
            let Ok(Some(value)) = package.slide_text(selector, role) else {
                continue;
            };
            tree.push(item, Kind::Placeholder)
                .prop("type", text(Some(name)))
                .prop("text", text(Some(&value)))
                .label(Some(name))
                .text(value)
                .leaf([index, slot, 0, 0])
                .done();
        }
        if let Some(notes) = slide.notes() {
            tree.push(item, Kind::Notes)
                .prop("text", text(Some(notes)))
                .text(notes)
                .leaf([index, 0, 0, 0])
                .done();
        }
    }
    Ok(tree.build())
}

/// Set slide visibility or placeholder text on every target.
pub(crate) fn set(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Keynote>(file, target, limits.max_expanded_targets())?;
    let mut chain = Chain::new(package.clone());
    let mut operations = 0_usize;
    for node in &nodes {
        for (key, value) in props.iter() {
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            let key = key.as_str();
            let [slide, slot, _, _] = node.item().leaf().map(|value| value as usize);
            let selector = SlideSelector::index(slide);
            match (node.item().type_name(), key) {
                ("slide", "hidden") => {
                    let hidden = super::write::boolean(value, "hidden")?;
                    let error =
                        |error| leaf_refusal(LeafError::KeynoteEdit(error), FileFormat::Keynote);
                    let mut edit = chain.package.edit();
                    edit.set_slide_skipped(selector, hidden).map_err(error)?;
                    let commit = edit.commit().map_err(error)?;
                    let patch = commit.patch().clone();
                    let changed = !patch.is_noop();
                    chain.push(
                        commit.into_package(),
                        LeafPatch::KeynoteSkip(patch),
                        changed,
                    );
                },
                ("placeholder", "text") => {
                    let value = super::write::text(value, "text")?.unwrap_or_default();
                    let role = if slot == TITLE {
                        Role::Title
                    } else {
                        Role::Body
                    };
                    let error =
                        |error| leaf_refusal(LeafError::KeynoteText(error), FileFormat::Keynote);
                    let mut edit = chain
                        .package
                        .edit_slide_text(selector, role)
                        .map_err(error)?;
                    edit.set(&value).map_err(error)?;
                    let commit = edit.commit().map_err(error)?;
                    let patch = commit.patch().clone();
                    let changed = !patch.is_noop();
                    chain.push(
                        commit.into_package(),
                        LeafPatch::KeynoteText(patch),
                        changed,
                    );
                },
                ("slide" | "placeholder", _) => return Err(key_refusal(node, key)),
                _ => return Err(element_refusal("set", node)),
            }
        }
    }
    finish(
        file,
        chain.changed,
        Owner::Keynote(chain.package),
        chain.patches,
        operations,
        limits,
    )
}

/// Remove slides. Several targets are removed from the last to the first,
/// so each one's position still holds when its deletion runs.
pub(crate) fn remove(
    file: &Iwork,
    package: &Package,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Keynote>(file, target, limits.max_expanded_targets())?;
    let mut slides = Vec::with_capacity(nodes.len());
    for node in &nodes {
        if node.item().type_name() != "slide" {
            return Err(element_refusal("remove", node));
        }
        slides.push(node.item().leaf()[0] as usize);
    }
    slides.sort_unstable();
    slides.dedup();
    let mut chain = Chain::new(package.clone());
    let error = |error| leaf_refusal(LeafError::KeynoteDeletion(error), FileFormat::Keynote);
    for (operations, slide) in slides.iter().rev().enumerate() {
        bound(
            LimitKind::Operations,
            operations + 1,
            limits.max_operations(),
        )?;
        let mut edit = chain.package.edit_slide_deletion();
        edit.remove_slide(SlideSelector::index(*slide))
            .map_err(error)?;
        let commit = edit.commit().map_err(error)?;
        let patch = commit.patch().clone();
        chain.push(
            commit.into_package(),
            LeafPatch::KeynoteDeletion(patch),
            true,
        );
    }
    finish(
        file,
        chain.changed,
        Owner::Keynote(chain.package),
        chain.patches,
        slides.len(),
        limits,
    )
}

/// Move one slide among the presentation's slides.
pub(crate) fn move_to(
    file: &Iwork,
    package: &Package,
    path: &PathArg<'_>,
    placement: &crate::Placement<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let node = node_at::<Keynote>(file, path)?;
    if node.item().type_name() != "slide" {
        return Err(element_refusal("move", &node));
    }
    if matches!(placement, crate::Placement::Into(..)) {
        return Err(refusal(
            Code::UnsupportedCapability,
            "a Keynote slide moves among the presentation's slides",
        ));
    }
    let from = node.item().leaf()[0] as usize;
    let slides = file.tree().item(super::tree::ROOT).children().len();
    let to = destination(from, slides, placement, |anchor| {
        let anchor = node_at::<Keynote>(file, anchor)?;
        if anchor.item().type_name() == "slide" {
            Ok(anchor.item().leaf()[0] as usize)
        } else {
            Err(refusal(
                Code::InvalidArgument,
                "a slide moves before or after another slide",
            ))
        }
    })?;
    let error = |error| leaf_refusal(LeafError::KeynoteOrder(error), FileFormat::Keynote);
    let mut edit = package.edit_slide_order();
    edit.move_slide(SlideSelector::index(from), Position::new(to))
        .map_err(error)?;
    let commit = edit.commit().map_err(error)?;
    let patch = commit.patch().clone();
    let changed = !patch.is_noop();
    let mut chain = Chain::new(package.clone());
    chain.push(
        commit.into_package(),
        LeafPatch::KeynoteOrder(patch),
        changed,
    );
    finish(
        file,
        chain.changed,
        Owner::Keynote(chain.package),
        chain.patches,
        1,
        limits,
    )
}
