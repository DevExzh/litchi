//! Unified writes over the iWork leaves' focused transactions.
//!
//! Every verb resolves its targets against one immutable handle's tree,
//! hands the flavor the nodes (with the leaf positions the tree recorded)
//! and the unified properties, and the flavor runs its leaf transactions in
//! order, each over the package the previous one published. The last
//! package becomes a new handle whose tree is built on its first read; an
//! exact no-op returns the source handle.

#[cfg(any(feature = "keynote", feature = "pages", feature = "numbers"))]
use litchi_core::Position;
use litchi_core::selector::Path;

use super::{Flavor, Iwork, Nav, Node, Owner};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::path::PathArg;
#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
use crate::props::PropValue;
use crate::props::{IntoProps, Props};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
use crate::value::{Scalar, ValueType};

/// One leaf transaction's reversible proof.
#[allow(
    clippy::large_enum_variant,
    reason = "a write holds its proofs once; boxing would add an allocation per commit"
)]
#[allow(
    clippy::enum_variant_names,
    reason = "in a one-format build the remaining variants share that format's prefix"
)]
pub(crate) enum LeafPatch {
    #[cfg(feature = "pages")]
    Pages(super::pages::Patch),
    #[cfg(feature = "keynote")]
    Keynote(super::keynote::Patch),
    #[cfg(feature = "numbers")]
    NumbersCells(litchi_numbers::table::cells::Patch),
    /// One of the other Numbers transactions (record 0297).
    #[cfg(feature = "numbers")]
    Numbers(super::numbers::Patch),
}

impl LeafPatch {
    fn is_noop(&self) -> bool {
        match self {
            #[cfg(feature = "pages")]
            Self::Pages(patch) => patch.is_noop(),
            #[cfg(feature = "keynote")]
            Self::Keynote(patch) => patch.is_noop(),
            #[cfg(feature = "numbers")]
            Self::NumbersCells(patch) => patch.is_noop(),
            #[cfg(feature = "numbers")]
            Self::Numbers(patch) => patch.is_noop(),
        }
    }
}

/// The leaf proof of one write: the leaf transactions it ran, in order.
pub(crate) enum Patch {
    /// The write changed nothing.
    Unchanged,
    /// The transactions, each over the package the previous one published.
    Steps(Vec<LeafPatch>),
}

impl Patch {
    pub(crate) fn is_noop(&self) -> bool {
        match self {
            Self::Unchanged => true,
            Self::Steps(steps) => steps.iter().all(LeafPatch::is_noop),
        }
    }
}

/// The leaf proof and immutable endpoints of one write.
pub(crate) struct Commit {
    pub(crate) before: Iwork,
    pub(crate) after: Iwork,
    pub(crate) patch: Patch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

/// A refusal the caller's request caused, before the format is known.
pub(crate) fn refusal(code: Code, message: impl Into<std::borrow::Cow<'static, str>>) -> Error {
    Error::new(code, message)
}

/// A leaf transaction failure: the request's fault unless the leaf names
/// the file or a resource ceiling.
pub(crate) fn leaf_refusal(error: super::LeafError, format: litchi_core::FileFormat) -> Error {
    super::leaf_error(error, format)
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

/// The running state of one write: the package the next leaf transaction
/// starts from and the proofs so far.
pub(crate) struct Chain<P> {
    pub(crate) package: P,
    pub(crate) patches: Vec<LeafPatch>,
    pub(crate) changed: bool,
}

impl<P> Chain<P> {
    pub(crate) const fn new(package: P) -> Self {
        Self {
            package,
            patches: Vec::new(),
            changed: false,
        }
    }

    /// Record one committed transaction.
    pub(crate) fn push(&mut self, package: P, patch: LeafPatch, changed: bool) {
        self.package = package;
        self.patches.push(patch);
        self.changed |= changed;
    }
}

/// Counts bytes without keeping them.
struct Count(usize);

impl std::io::Write for Count {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.0 = self.0.saturating_add(buffer.len());
        Ok(buffer.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

/// Adopt the last committed leaf package, charging the retained endpoint.
pub(crate) fn finish(
    file: &Iwork,
    changed: bool,
    owner: Owner,
    patches: Vec<LeafPatch>,
    operations: usize,
    limits: &SemanticLimits,
) -> Result<Commit> {
    if !changed {
        return Ok(Commit {
            before: file.clone(),
            after: file.clone(),
            patch: Patch::Unchanged,
            operations,
            snapshot_bytes: 0,
        });
    }
    let mut count = Count(0);
    owner.write_to(&mut count)?;
    let snapshot_bytes = count.0;
    bound(
        LimitKind::SnapshotBytes,
        snapshot_bytes,
        limits.max_snapshot_bytes(),
    )?;
    let after = Iwork::from_owner(owner, file.origin().cloned());
    Ok(Commit {
        before: file.clone(),
        after,
        patch: Patch::Steps(patches),
        operations,
        snapshot_bytes,
    })
}

/// The nodes a set or remove targets: one path, or every match of a query.
pub(crate) fn resolve<'f, F: Flavor>(
    file: &'f Iwork,
    target: &TargetArg<'_>,
    maximum: usize,
) -> Result<Vec<Node<'f>>> {
    file.checked_tree()?;
    let mut nodes = Vec::new();
    if let Some(path) = target.path() {
        nodes.push(crate::path::resolve::get(&Nav::<F>::new(file), path)?);
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(Nav::<F>::new(file), query.clone())?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(refusal(
                Code::InvalidSelector,
                "the mutation selector contains an unsupported predicate",
            ));
        }
        for node in walk {
            nodes.push(node?);
            bound(LimitKind::ExpandedTargets, nodes.len(), maximum)?;
        }
        if nodes.is_empty() {
            return Err(refusal(
                Code::ZeroMatches,
                "the mutation selector matched nothing",
            ));
        }
    } else {
        return Err(refusal(
            Code::InvalidArgument,
            "the mutation target is empty",
        ));
    }
    bound(LimitKind::ExpandedTargets, nodes.len(), maximum)?;
    Ok(nodes)
}

/// Resolve one element by path.
#[cfg(any(feature = "keynote", feature = "numbers", feature = "pages"))]
pub(crate) fn node_at<'f, F: Flavor>(file: &'f Iwork, path: &PathArg<'_>) -> Result<Node<'f>> {
    file.checked_tree()?;
    crate::path::resolve::get(&Nav::<F>::new(file), path)
}

/// The value of `key` coerced to text; `null` is `None`.
#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
pub(crate) fn text(value: &PropValue<'_>, key: &'static str) -> Result<Option<String>> {
    value
        .with_coerced(ValueType::Text, |scalar| {
            Ok(scalar.map(|scalar| match scalar {
                Scalar::Text(text) => text.to_owned(),
                other => other.to_string(),
            }))
        })
        .map_err(|error| error.with_property(key))
}

/// The character properties of a `font.*` write (record 0309).
#[cfg(feature = "numbers")]
pub(crate) use litchi_numbers::cell::text_style::CharacterProperties as FontProperties;
#[cfg(all(feature = "pages", not(feature = "numbers")))]
pub(crate) use litchi_pages::section::text_style::CharacterProperties as FontProperties;

/// Fold one `font.*` key into `properties`; `false` for other keys.
#[cfg(any(feature = "pages", feature = "numbers"))]
pub(crate) fn font_property(
    properties: &mut FontProperties,
    key: &str,
    value: &PropValue<'_>,
    color_type: ValueType,
) -> Result<bool> {
    let cleared = |key: &str| {
        refusal(Code::InvalidValue, "a font property cannot be cleared")
            .with_property(key.to_owned())
    };
    let lower = key.to_ascii_lowercase();
    match lower.as_str() {
        "font.bold" => properties.bold = Some(boolean(value, "font.bold")?),
        "font.italic" => properties.italic = Some(boolean(value, "font.italic")?),
        "font.underline" => {
            properties.underline = Some(boolean(value, "font.underline")?);
        },
        "font.strike" | "font.strikethrough" => {
            properties.strikethrough = Some(boolean(value, "font.strike")?);
        },
        "font.name" => {
            properties.font_name = Some(text(value, "font.name")?.ok_or_else(|| cleared(key))?);
        },
        "font.size" => {
            let points = value
                .with_coerced(
                    ValueType::Length(crate::value::LengthSpec::FONT_SIZE),
                    |scalar| {
                        Ok(scalar.and_then(|scalar| match scalar {
                            Scalar::Length(length) => Some(length.points()),
                            Scalar::Number(value) => Some(value),
                            #[allow(clippy::cast_precision_loss, reason = "point sizes are small")]
                            Scalar::Int(value) => Some(value as f64),
                            _ => None,
                        }))
                    },
                )
                .map_err(|error| error.with_property("font.size"))?
                .ok_or_else(|| cleared(key))?;
            #[allow(
                clippy::cast_possible_truncation,
                reason = "iWork stores font sizes as 32-bit floats"
            )]
            let points = points as f32;
            if !points.is_finite() || points <= 0.0 {
                return Err(refusal(Code::InvalidValue, "a font size must be positive")
                    .with_property("font.size".to_owned()));
            }
            properties.size = Some(points);
        },
        "font.color" => {
            let color = value
                .with_coerced(color_type, |scalar| {
                    Ok(scalar.and_then(|scalar| match scalar {
                        Scalar::Color(color) if color.transforms().is_empty() => {
                            match color.base() {
                                crate::value::Base::Rgb(rgba) => {
                                    Some([rgba.red(), rgba.green(), rgba.blue()])
                                },
                                _ => None,
                            }
                        },
                        _ => None,
                    }))
                })
                .map_err(|error| error.with_property("font.color"))?
                .ok_or_else(|| {
                    refusal(Code::InvalidValue, "a font color must be an RGB color")
                        .with_property("font.color".to_owned())
                })?;
            properties.color = Some(color);
        },
        _ => return Ok(false),
    }
    Ok(true)
}

/// A table-cell value of a `text` or `value` key: text, a finite number,
/// a boolean, or `None` (clear) for `null` (record 0309).
#[cfg(any(feature = "keynote", feature = "pages"))]
pub(crate) enum CellScalar {
    Text(String),
    Number(f64),
    Boolean(bool),
    Empty,
}

#[cfg(any(feature = "keynote", feature = "pages"))]
pub(crate) fn cell_scalar(key: &str, value: &PropValue<'_>) -> Result<CellScalar> {
    if key.eq_ignore_ascii_case("text") {
        return Ok(text(value, "text")?.map_or(CellScalar::Empty, CellScalar::Text));
    }
    value
        .with_coerced(ValueType::Scalar, |scalar| {
            Ok(match scalar {
                None => CellScalar::Empty,
                Some(Scalar::Text(text)) => CellScalar::Text(text.to_owned()),
                Some(Scalar::Bool(flag)) => CellScalar::Boolean(flag),
                #[allow(clippy::cast_precision_loss, reason = "table cells store doubles")]
                Some(Scalar::Int(number)) => CellScalar::Number(number as f64),
                Some(Scalar::Number(number)) if number.is_finite() => CellScalar::Number(number),
                Some(other) => {
                    return Err(refusal(
                        Code::InvalidValue,
                        format!("a table cell cannot hold `{other}`"),
                    ));
                },
            })
        })
        .map_err(|error| error.with_property("value"))
}

/// The size, name and seeded text of a new Keynote or Pages table (record
/// 0329), from the OfficeCLI table keys: `rows`, `cols` (or `columns`),
/// `name`, and `data` (rows separated by `;`, cells by `,`). A missing size
/// is taken from `data`.
#[cfg(any(feature = "keynote", feature = "pages"))]
#[derive(Debug, Default)]
pub(crate) struct TableRequest {
    rows: Option<u32>,
    columns: Option<u32>,
    pub(crate) name: Option<String>,
    data: Vec<Vec<String>>,
}

#[cfg(any(feature = "keynote", feature = "pages"))]
impl TableRequest {
    /// Fold one key into the request; `false` for keys it does not own.
    pub(crate) fn key(&mut self, key: &str, value: &PropValue<'_>) -> Result<bool> {
        let size = |key: &'static str| -> Result<u32> {
            value
                .with_coerced(ValueType::Int, |scalar| {
                    Ok(scalar.and_then(|scalar| match scalar {
                        Scalar::Int(number) => u32::try_from(number).ok(),
                        _ => None,
                    }))
                })
                .map_err(|error| error.with_property(key))?
                .filter(|count| *count >= 1)
                .ok_or_else(|| {
                    refusal(Code::OutOfRange, "table dimensions are positive")
                        .with_property(key.to_owned())
                })
        };
        match key {
            "rows" => self.rows = Some(size("rows")?),
            "cols" | "columns" => self.columns = Some(size("cols")?),
            "name" => self.name = text(value, "name")?,
            "data" => {
                self.data = text(value, "data")?
                    .unwrap_or_default()
                    .split(';')
                    .map(|row| row.split(',').map(str::to_owned).collect())
                    .collect();
            },
            _ => return Ok(false),
        }
        Ok(true)
    }

    /// The rows, columns and non-empty seeded cells (zero-based).
    pub(crate) fn shape(&self) -> Result<(u32, u32, Vec<(u32, u32, String)>)> {
        let missing = |key: &'static str| {
            refusal(
                Code::MissingProperty,
                "a new table needs `rows` and `cols`, or `data`",
            )
            .with_property(key.to_owned())
        };
        let data_rows = u32::try_from(self.data.len()).ok().filter(|rows| *rows > 0);
        let data_columns = self
            .data
            .iter()
            .map(Vec::len)
            .max()
            .and_then(|columns| u32::try_from(columns).ok())
            .filter(|columns| *columns > 0);
        let rows = self.rows.or(data_rows).ok_or_else(|| missing("rows"))?;
        let columns = self
            .columns
            .or(data_columns)
            .ok_or_else(|| missing("cols"))?;
        let mut cells = Vec::new();
        for (row, values) in self.data.iter().enumerate() {
            for (column, value) in values.iter().enumerate() {
                if value.is_empty() {
                    continue;
                }
                let (Ok(row), Ok(column)) = (u32::try_from(row), u32::try_from(column)) else {
                    return Err(refusal(Code::OutOfRange, "table data is outside the table"));
                };
                if row >= rows || column >= columns {
                    return Err(refusal(Code::OutOfRange, "table data is outside the table")
                        .with_property("data".to_owned()));
                }
                cells.push((row, column, value.clone()));
            }
        }
        Ok((rows, columns, cells))
    }
}

/// The value of `key` coerced to a boolean; `null` is refused.
#[cfg(any(feature = "keynote", feature = "numbers", feature = "pages"))]
pub(crate) fn boolean(value: &PropValue<'_>, key: &'static str) -> Result<bool> {
    value
        .with_coerced(ValueType::Bool, |scalar| {
            Ok(scalar.and_then(|scalar| match scalar {
                Scalar::Bool(value) => Some(value),
                _ => None,
            }))
        })
        .map_err(|error| error.with_property(key))?
        .ok_or_else(|| {
            refusal(Code::InvalidValue, "the property cannot be cleared")
                .with_property(key.to_owned())
        })
}

/// The positive point count carried by a unified chart-font-size write.
///
/// The focused chart-font leaves all use the same finite, positive
/// point-size value. Keeping this coercion here makes the Pages,
/// Keynote and Numbers dispatch arms share the same invalid/null boundary;
/// no format is allowed to silently clamp or reinterpret a size.
#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
pub(crate) fn chart_font_points(value: &PropValue<'_>) -> Result<f32> {
    let points = value
        .with_coerced(
            ValueType::Length(crate::value::LengthSpec::FONT_SIZE),
            |scalar| {
                Ok(scalar.and_then(|scalar| match scalar {
                    Scalar::Length(length) => Some(length.points()),
                    Scalar::Number(value) => Some(value),
                    #[allow(clippy::cast_precision_loss, reason = "point sizes are small")]
                    Scalar::Int(value) => Some(value as f64),
                    _ => None,
                }))
            },
        )
        .map_err(|error| error.with_property("font.size"))?
        .ok_or_else(|| {
            refusal(Code::InvalidValue, "a chart font size must be positive")
                .with_property("font.size".to_owned())
        })?;
    #[allow(
        clippy::cast_possible_truncation,
        reason = "iWork stores chart font sizes as 32-bit floats"
    )]
    let points = points as f32;
    if !points.is_finite() || points <= 0.0 {
        return Err(refusal(
            Code::InvalidValue,
            "a chart font size must be positive and finite",
        )
        .with_property("font.size".to_owned()));
    }
    Ok(points)
}

/// The refusal of a key the element does not write.
pub(crate) fn unwritable(element: &'static str, key: &str) -> Error {
    refusal(
        Code::UnsupportedProperty,
        format!("`{key}` is not a writable property of `{element}`"),
    )
    .with_property(key.to_owned())
    .with_suggestion("list the writable properties with capabilities()")
}

/// The refusal of a read-only property, or an unknown key.
pub(crate) fn key_refusal(node: &Node<'_>, key: &str) -> Error {
    let item = node.item();
    if item
        .props()
        .iter()
        .any(|prop| prop.name.eq_ignore_ascii_case(key))
    {
        refusal(
            Code::UnsupportedCapability,
            format!("`{key}` of `{}` is read-only", item.type_name()),
        )
        .with_property(key.to_owned())
    } else {
        unwritable(item.type_name(), key)
    }
}

/// The refusal of a verb on an element the flavor does not write.
pub(crate) fn element_refusal(verb: &'static str, node: &Node<'_>) -> Error {
    refusal(
        Code::UnsupportedCapability,
        format!(
            "`{verb}` is not supported on `{}` elements",
            node.item().type_name()
        ),
    )
}

/// The zero-based destination of a move among `count` siblings, for an
/// element now at `from`; `anchor` resolves a `before`/`after` path to its
/// sibling position.
#[cfg(any(feature = "keynote", feature = "pages", feature = "numbers"))]
pub(crate) fn destination(
    from: usize,
    count: usize,
    placement: &crate::Placement<'_>,
    anchor: impl Fn(&PathArg<'_>) -> Result<usize>,
) -> Result<usize> {
    use crate::Placement;
    let last = count.saturating_sub(1);
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
        return Err(refusal(
            Code::OutOfRange,
            "the move position is outside the element's siblings",
        ));
    }
    Ok(to)
}

/// Set properties on every target.
pub(crate) fn set<'p, 'v>(
    file: &Iwork,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let props = Props::from_props(props)?;
    if props.is_empty() {
        return Err(refusal(
            Code::InvalidArgument,
            "a set needs at least one property",
        ));
    }
    match file.owner() {
        #[cfg(feature = "pages")]
        Owner::Pages(package) => super::pages::set(file, package, &target, &props, limits),
        #[cfg(feature = "keynote")]
        Owner::Keynote(package) => super::keynote::set(file, package, &target, &props, limits),
        #[cfg(feature = "numbers")]
        Owner::Numbers(package) => super::numbers::set(file, package, &target, &props, limits),
    }
}

/// Add one element (Pages paragraphs; Numbers merged ranges, record 0297).
pub(crate) fn add<'p, 'n>(
    file: &Iwork,
    parent: PathArg<'p>,
    node: NewNode<'n>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    match file.owner() {
        #[cfg(feature = "pages")]
        Owner::Pages(package) => super::pages::add(file, package, &parent, &node, limits),
        #[cfg(feature = "keynote")]
        Owner::Keynote(package) => super::keynote::add(file, package, &parent, &node, limits),
        #[cfg(feature = "numbers")]
        Owner::Numbers(package) => super::numbers::add(file, package, &parent, &node, limits),
        #[allow(
            unreachable_patterns,
            reason = "every iWork leaf adds elements; the arm covers a build without them"
        )]
        _ => {
            let _ = (parent, node, limits);
            Err(refusal(
                Code::UnsupportedCapability,
                "this build cannot add elements to the format",
            ))
        },
    }
}

/// Remove every target.
pub(crate) fn remove<'p>(
    file: &Iwork,
    target: TargetArg<'p>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    match file.owner() {
        #[cfg(feature = "pages")]
        Owner::Pages(package) => super::pages::remove(file, package, &target, limits),
        #[cfg(feature = "keynote")]
        Owner::Keynote(package) => super::keynote::remove(file, package, &target, limits),
        #[cfg(feature = "numbers")]
        Owner::Numbers(package) => super::numbers::remove(file, package, &target, limits),
        #[allow(unreachable_patterns, reason = "every iWork leaf removes elements")]
        _ => {
            let _ = (target, limits);
            let _ = file;
            Err(refusal(
                Code::UnsupportedCapability,
                "this iWork element cannot be removed",
            ))
        },
    }
}

/// Move one element (Keynote slides).
pub(crate) fn move_to<'p, 'q>(
    file: &Iwork,
    path: PathArg<'p>,
    placement: crate::Placement<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    match file.owner() {
        #[cfg(feature = "pages")]
        Owner::Pages(package) => super::pages::move_to(file, package, &path, &placement, limits),
        #[cfg(feature = "keynote")]
        Owner::Keynote(package) => {
            super::keynote::move_to(file, package, &path, &placement, limits)
        },
        #[cfg(feature = "numbers")]
        Owner::Numbers(package) => {
            super::numbers::move_to(file, package, &path, &placement, limits)
        },
        #[allow(unreachable_patterns, reason = "every iWork leaf orders elements")]
        _ => {
            let _ = (path, placement, limits);
            Err(refusal(
                Code::UnsupportedCapability,
                "this iWork element cannot be moved",
            ))
        },
    }
}

/// Exchange two elements (Pages paragraphs; Keynote slides, paragraphs and
/// drawables; Numbers sheets and tables, record 0347).
pub(crate) fn swap<'p, 'q>(
    file: &Iwork,
    left: PathArg<'p>,
    right: PathArg<'q>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    match file.owner() {
        #[cfg(feature = "pages")]
        Owner::Pages(package) => super::pages::swap(file, package, &left, &right, limits),
        #[cfg(feature = "keynote")]
        Owner::Keynote(package) => super::keynote::swap(file, package, &left, &right, limits),
        #[cfg(feature = "numbers")]
        Owner::Numbers(package) => super::numbers::swap(file, package, &left, &right, limits),
    }
}
