//! OpenDocument images (`.odi`): the image frame and `/metadata`, read from
//! `litchi_odi::Image` and written through `litchi_odi::Edit`.

use litchi_core::FileFormat;
use litchi_core::selector::Path;

use super::tree::{length, text};
use super::write::{Commit, Patch, bound, element_refusal, finish, key_refusal, refusal};
use super::{Flavor, Node, OdfTree, Owner, TreeBuilder, element, odf_staging_error, writable};
use crate::Result;
use crate::error::Code;
use crate::new_node::NewNode;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::path::PathArg;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::{LengthSpec, ValueType};

kinds! {
    /// The element kinds of an ODI image.
    Kind: &crate::path::vocabulary::IMAGE, root Image {
        Image => "image",
        Frame => "frame",
        Metadata => "metadata",
    }
}

const IMAGE_PROPS: &[PropSpec] = &[
    PropSpec::new("frameCount", ValueType::Int),
    PropSpec::new("resourceCount", ValueType::Int),
];

pub(crate) const FRAME_PROPS: &[PropSpec] = &[
    writable("name", ValueType::Text),
    writable("title", ValueType::Text),
    writable("description", ValueType::Text),
    writable("x", ValueType::Length(LengthSpec::POSITION)),
    writable("y", ValueType::Length(LengthSpec::POSITION)),
    writable("width", ValueType::Length(LengthSpec::EXTENT)),
    writable("height", ValueType::Length(LengthSpec::EXTENT)),
    PropSpec::new("source", ValueType::Text),
    PropSpec::new("embedded", ValueType::Bool),
    PropSpec::new("mediaType", ValueType::Text),
    writable("anchorType", ValueType::Text),
    writable("zIndex", ValueType::Int),
    writable("layer", ValueType::Text),
    writable("style", ValueType::Text),
];

pub(crate) const METADATA_PROPS: &[PropSpec] = &[
    writable("title", ValueType::Text),
    writable("author", ValueType::Text),
    writable("subject", ValueType::Text),
    writable("description", ValueType::Text),
    writable("keywords", ValueType::List),
];

static OPERATION_GAPS: [GapSpec; 4] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "frame",
            context: "an ODI image holds exactly one frame",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 ODI single frame",
        reason: "the ODI leaf admits exactly one image frame",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "remove",
            element: "frame",
            context: "an ODI image holds exactly one frame",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 ODI single frame",
        reason: "an image package without its frame is not an ODI image",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "move",
            element: "frame",
            context: "a position other than the frame's own",
        },
        code: Code::OutOfRange,
        follow_up: "none: an ODI image holds one frame",
        reason: "the frame and the metadata are singletons, so a move or swap names only their own position (an exact no-op)",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "frame",
            context: "the `source` property (the image payload)",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 ODI image payload writes",
        reason: "replacing the image payload needs a typed source value the unified registry does not have yet",
    },
];

/// The ODI spec table.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Image,
    elements: &[
        element(
            Kind::Image,
            "image",
            &["image"],
            &[],
            &[Kind::Frame, Kind::Metadata],
            &["/"],
            IMAGE_PROPS,
        ),
        element(
            Kind::Frame,
            "frame",
            &["frame", "picture"],
            &[Kind::Image],
            &[],
            &["/frame[N]"],
            FRAME_PROPS,
        ),
        element(
            Kind::Metadata,
            "metadata",
            &["metadata"],
            &[Kind::Image],
            &[],
            &["/metadata"],
            METADATA_PROPS,
        ),
    ],
    gaps: &[],
};

crate::node::spec::check_format_spec!(SPEC);

/// The ODI flavor.
pub(crate) struct Odi;

impl Flavor for Odi {
    type Kind = Kind;
    const FORMAT: FileFormat = FileFormat::Odi;
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Image];
    const SINGLETONS: &'static [Kind] = &[Kind::Metadata];
}

/// The elements `set` writes, `add` creates and `remove` deletes.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &["frame", "metadata"],
    addable: &[],
    removable: &["metadata"],
    gaps: &OPERATION_GAPS,
};

/// Build the tree of `package`.
pub(crate) fn build(package: &litchi_odi::Package) -> Result<super::Tree> {
    let image = package.image();
    let mut tree = TreeBuilder::new(SPEC.elements, Odi::SINGLETONS);
    tree.root()
        .prop("frameCount", super::tree::count(image.frames().len()))
        .prop("resourceCount", super::tree::count(image.resources().len()));
    for (index, frame) in image.frames().iter().enumerate() {
        let (source, embedded) = match frame.source() {
            litchi_odi::source::Source::Linked(href) => (Some(href.as_str()), false),
            litchi_odi::source::Source::Embedded(_) => (None, true),
            _ => (None, false),
        };
        tree.push(super::tree::ROOT, Kind::Frame)
            .prop("name", text(frame.name()))
            .prop("title", text(frame.title()))
            .prop("description", text(frame.description()))
            .prop("x", length(frame.x()))
            .prop("y", length(frame.y()))
            .prop("width", length(frame.width()))
            .prop("height", length(frame.height()))
            .prop("source", text(source))
            .prop("embedded", Some(embedded.into()))
            .prop("mediaType", text(frame.media_type()))
            .prop("anchorType", text(frame.anchor_type()))
            .prop("zIndex", frame.z_index().map(Into::into))
            .prop("layer", text(frame.layer()))
            .prop("style", text(frame.style_name()))
            .label(frame.name())
            .leaf(index, 0)
            .done();
    }
    let metadata = image.metadata();
    let field = |pick: fn(&litchi_core::Metadata) -> Option<&String>| {
        text(metadata.and_then(|metadata| pick(metadata).map(String::as_str)))
    };
    tree.push(super::tree::ROOT, Kind::Metadata)
        .prop("title", field(|metadata| metadata.title.as_ref()))
        .prop("author", field(|metadata| metadata.author.as_ref()))
        .prop("subject", field(|metadata| metadata.subject.as_ref()))
        .prop(
            "description",
            field(|metadata| metadata.description.as_ref()),
        )
        .prop(
            "keywords",
            metadata.and_then(|metadata| crate::node::Value::owned_list(&metadata.keywords)),
        )
        .done();
    Ok(tree.build())
}

/// Set frame or metadata properties.
pub(crate) fn set(
    file: &OdfTree,
    package: &litchi_odi::Package,
    groups: &[super::write::Group<'_, '_, '_>],
    limits: &SemanticLimits,
) -> Result<Commit> {
    let image = package.image();
    let mut edit = package.edit();
    let mut operations = 0usize;
    let staging = |error| odf_staging_error(error, FileFormat::Odi);
    for (node, props) in groups
        .iter()
        .flat_map(|(nodes, props)| nodes.iter().map(move |node| (node, *props)))
    {
        let kind = node.item().type_name();
        // Position and size stage together: the leaf audits a frame's
        // geometry as one set of four values.
        let mut geometry: [Option<Option<String>>; 4] = [None, None, None, None];
        for (key, value) in props.iter() {
            let key = key.as_str();
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            match kind {
                "frame" => {
                    let frame = node.item().leaf()[0];
                    let slot = match key {
                        "x" => Some((0, LengthSpec::POSITION)),
                        "y" => Some((1, LengthSpec::POSITION)),
                        "width" => Some((2, LengthSpec::EXTENT)),
                        "height" => Some((3, LengthSpec::EXTENT)),
                        _ => None,
                    };
                    if let Some((slot, spec)) = slot {
                        geometry[slot] = Some(
                            super::write::length(value, static_key(key, FRAME_PROPS), spec)?
                                .map(super::tree::odf_length),
                        );
                        continue;
                    }
                    stage_frame(&mut edit, frame, node, key, value).map_err(
                        |error| match error {
                            Staged::Leaf(error) => staging(error),
                            Staged::Unified(error) => error,
                        },
                    )?;
                },
                "metadata" if key == "keywords" => {
                    let keywords = crate::props::keyword_list(value).ok_or_else(|| {
                        crate::Error::new(
                            Code::InvalidValue,
                            "the keywords take a list, or text separated by ';' or ','",
                        )
                        .with_property("keywords")
                        .with_format(FileFormat::Odi)
                    })?;
                    edit.set_keywords(keywords).map_err(staging)?;
                },
                "metadata" => {
                    let value = super::write::text(value, static_key(key, METADATA_PROPS))?;
                    match key {
                        "title" => edit.set_title(value),
                        "author" => edit.set_author(value),
                        "subject" => edit.set_subject(value),
                        "description" => edit.set_description(value),
                        other => return Err(key_refusal(node, other)),
                    }
                    .map_err(staging)?;
                },
                _ => return Err(element_refusal("set", node)),
            }
        }
        if geometry.iter().any(Option::is_some) {
            let frame = node.item().leaf()[0];
            let current = image
                .frames()
                .get(frame)
                .map(|frame| [frame.x(), frame.y(), frame.width(), frame.height()])
                .unwrap_or_default();
            let [x, y, width, height] = std::array::from_fn(|index| {
                geometry[index]
                    .clone()
                    .unwrap_or_else(|| current[index].map(str::to_owned))
            });
            edit.set_geometry(frame, x, y, width, height)
                .map_err(staging)?;
        }
    }
    let commit = edit
        .commit()
        .map_err(|error| super::write::commit_error(error, FileFormat::Odi))?;
    let changed = commit.diagnostics().changed();
    let (package, patch) = commit.into_parts();
    finish(
        file,
        changed,
        Owner::Odi(package),
        Patch::Odi(patch),
        operations,
        limits,
    )
}

/// A frame staging failure: the leaf's or the unified layer's.
enum Staged {
    Leaf(litchi_odf_common::Error),
    Unified(crate::Error),
}

impl From<crate::Error> for Staged {
    fn from(error: crate::Error) -> Self {
        Self::Unified(error)
    }
}

fn stage_frame(
    edit: &mut litchi_odi::Edit<'_>,
    frame: usize,
    node: &Node<'_>,
    key: &str,
    value: &crate::props::PropValue<'_>,
) -> std::result::Result<(), Staged> {
    let result = match key {
        "name" => edit.set_frame_name(frame, super::write::text(value, "name")?),
        "title" => edit.set_frame_title(frame, super::write::text(value, "title")?),
        "description" => {
            edit.set_frame_description(frame, super::write::text(value, "description")?)
        },
        "anchorType" => edit.set_anchor_type(frame, super::write::text(value, "anchorType")?),
        "zIndex" => {
            let index = super::write::int(value, "zIndex")?
                .map(|value| {
                    u32::try_from(value).map_err(|_| {
                        refusal(Code::OutOfRange, "zIndex must be between 0 and 4294967295")
                            .with_property("zIndex".to_owned())
                    })
                })
                .transpose()?;
            edit.set_z_index(frame, index)
        },
        "layer" => edit.set_layer(frame, super::write::text(value, "layer")?),
        "style" => edit.set_style_name(frame, super::write::text(value, "style")?),
        other => return Err(Staged::Unified(key_refusal(node, other))),
    };
    result.map_err(Staged::Leaf)
}

/// The static spelling of a declared key.
fn static_key(key: &str, props: &'static [PropSpec]) -> &'static str {
    props
        .iter()
        .find(|prop| prop.name == key)
        .map_or("value", |prop| prop.name)
}

/// ODI images hold one frame and one metadata node: nothing can be added.
pub(crate) fn add(
    file: &OdfTree,
    package: &litchi_odi::Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let _ = (package, limits);
    super::write::parent::<Odi>(file, parent)?;
    let kind = node.kind().to_ascii_lowercase();
    Err(match kind.as_str() {
        "frame" | "picture" => refusal(
            Code::UnsupportedCapability,
            "an ODI image holds exactly one frame; set the existing `/frame[1]` instead",
        )
        .with_suggestion("follow-up: 0282 ODI single frame"),
        "metadata" => refusal(
            Code::InvalidArgument,
            "an ODI image always has `/metadata`; set its properties instead",
        ),
        _ => refusal(
            Code::UnsupportedType,
            format!("an ODI image has no `{}` elements", node.kind()),
        ),
    })
}

/// Remove `/metadata`: clear the fields it exposes (title, author,
/// subject, description, keywords) through the ODF metadata family, which
/// keeps every other `meta.xml` node. The image's one frame cannot be
/// removed.
pub(crate) fn remove(
    file: &OdfTree,
    package: &litchi_odi::Package,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    use litchi_odf_common::metadata::Field;
    let nodes = super::write::resolve::<Odi>(file, target, limits.max_expanded_targets())?;
    bound(LimitKind::Operations, nodes.len(), limits.max_operations())?;
    for node in &nodes {
        match node.item().type_name() {
            "metadata" => {},
            "frame" => {
                return Err(refusal(
                    Code::UnsupportedCapability,
                    "an ODI image holds exactly one frame, which cannot be removed",
                )
                .with_suggestion("follow-up: 0282 ODI single frame"));
            },
            _ => return Err(element_refusal("remove", node)),
        }
    }
    let mut edit = litchi_odf_common::metadata::Edit::new(package);
    for field in [
        Field::Title,
        Field::Author,
        Field::Subject,
        Field::Description,
        Field::Keywords,
    ] {
        edit.set(field, None).map_err(|error| {
            super::odf_error(error, FileFormat::Odi).caused_by_argument(Code::InvalidValue)
        })?;
    }
    let (next, patch, changed) = crate::format::odf_metadata::commit(edit, FileFormat::Odi)?;
    finish(
        file,
        changed,
        Owner::Odi(next),
        Patch::Metadata(super::write::MetadataPatch::Odi(patch)),
        nodes.len(),
        limits,
    )
}

/// The image's elements are singletons: a move to the element's own
/// position is an exact no-op; any other position is out of range.
pub(crate) fn move_to(
    file: &OdfTree,
    path: &PathArg<'_>,
    placement: &crate::Placement<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let node = super::write::parent::<Odi>(file, path)?;
    if node.item().type_name() == "image" {
        return Err(element_refusal("move", &node));
    }
    let siblings = super::write::siblings(file, &node)?;
    match super::write::destination::<Odi>(file, &node, &siblings, placement)? {
        super::write::Destination::Within(_) => finish_unchanged(file, limits),
        super::write::Destination::Across { .. } => Err(refusal(
            Code::UnsupportedType,
            "an ODI frame stays in its image",
        )),
    }
}

/// Swapping an image element with itself is an exact no-op; an ODI image
/// has no two elements of one kind to exchange.
pub(crate) fn swap(
    file: &OdfTree,
    left: &PathArg<'_>,
    right: &PathArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let (left, right) = super::write::swap_pair::<Odi>(file, left, right)?;
    if left.index() != right.index() {
        return Err(refusal(
            Code::InternalError,
            "an ODI image holds one element of each kind",
        ));
    }
    if left.item().type_name() == "image" {
        return Err(element_refusal("swap", &left));
    }
    finish_unchanged(file, limits)
}

fn finish_unchanged(file: &OdfTree, limits: &SemanticLimits) -> Result<Commit> {
    let _ = limits;
    Ok(Commit {
        before: file.clone(),
        after: file.clone(),
        patch: Patch::Unchanged,
        operations: 1,
        snapshot_bytes: 0,
    })
}
