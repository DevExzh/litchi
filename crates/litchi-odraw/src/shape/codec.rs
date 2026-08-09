//! `OfficeArt` shape wire traversal and typed decoding.

use crate::prop::{Id, Props};
use crate::{Container, Error, Limit, Limits, Record, RecordKind, Result};

use super::model::{Bounds, Flags, Kind, Shape};
use super::validation::{
    Budget, Role, coordinate, detect, next_depth, scan_meta, validate_atom,
    validate_container_header, validate_container_record, validate_meta,
};

impl<'data> TryFrom<Record<'data>> for Shape<'data> {
    type Error = Error;

    /// Builds one typed shape from an `SpContainer` or `SpgrContainer`.
    fn try_from(record: Record<'data>) -> Result<Self> {
        let kind = record.kind();
        let container = Container::try_new(record)?;
        let mut budget = Budget::new(Limits::default());
        budget.visit()?;
        match kind {
            RecordKind::SpContainer => {
                build_shape(container, Vec::new(), Role::Standalone, &mut budget)
            },
            RecordKind::SpgrContainer => build_group(container, 0, Role::Standalone, &mut budget),
            RecordKind::DggContainer
            | RecordKind::BStoreContainer
            | RecordKind::DgContainer
            | RecordKind::SolverContainer
            | RecordKind::Dgg
            | RecordKind::Bse
            | RecordKind::Dg
            | RecordKind::Spgr
            | RecordKind::Sp
            | RecordKind::Opt
            | RecordKind::ClientTextbox
            | RecordKind::ChildAnchor
            | RecordKind::ClientAnchor
            | RecordKind::ClientData
            | RecordKind::ConnectorRule
            | RecordKind::AlignRule
            | RecordKind::ArcRule
            | RecordKind::ClientRule
            | RecordKind::CalloutRule
            | RecordKind::BlipEmf
            | RecordKind::BlipWmf
            | RecordKind::BlipPict
            | RecordKind::BlipJpeg
            | RecordKind::BlipPng
            | RecordKind::BlipDib
            | RecordKind::BlipTiff
            | RecordKind::ColorMru
            | RecordKind::SplitMenuColors
            | RecordKind::SecondaryOpt
            | RecordKind::TertiaryOpt
            | RecordKind::Unknown(_) => Err(Error::MalformedShape {
                reason: "record is not a shape container",
            }),
        }
    }
}

impl Bounds {
    /// Decodes one exact `[MS-ODRAW]` `FSPGR` record without copying its
    /// payload. The four fixed-width coordinates are decoded by value, while
    /// the record and any neighboring unknown records remain borrowed by the
    /// containing shape.
    ///
    /// # Errors
    ///
    /// Returns `Error::MalformedShape` if `record` is not a version-1,
    /// instance-0 `Spgr` atom with a 16-byte payload or a coordinate extends
    /// past the payload, or `Error::ArithmeticOverflow` if the coordinate
    /// offset arithmetic cannot be represented.
    pub fn from_record(record: &Record<'_>) -> Result<Self> {
        validate_atom(record, RecordKind::Spgr, 1, Some(0), 16)?;
        Ok(Self::new(
            coordinate(record.data(), 0)?,
            coordinate(record.data(), 4)?,
            coordinate(record.data(), 8)?,
            coordinate(record.data(), 12)?,
        ))
    }
}

/// Parses the user-visible shapes in one `OfficeArt` drawing.
///
/// # Errors
///
/// Returns an error under the same conditions as [`parse_with`] with the
/// default [`Limits`].
pub fn parse(data: &[u8]) -> Result<Vec<Shape<'_>>> {
    parse_with(data, Limits::default())
}

/// Parses user-visible shapes within explicit depth and record ceilings.
///
/// # Errors
///
/// Returns `Error::InvalidLimit` if `limits.max_depth` exceeds the
/// implementation's safe maximum, `Error::TruncatedHeader` or
/// `Error::TruncatedPayload` if the wire data ends mid-record,
/// `Error::TrailingData` if the top-level record does not consume all of
/// `data`, `Error::NotContainer` if the root record is not a container,
/// `Error::LimitExceeded` if a depth or record ceiling is reached, or
/// `Error::MalformedShape` if the root is not a drawing or shape container
/// or any nested record is structurally invalid.
pub fn parse_with(data: &[u8], limits: Limits) -> Result<Vec<Shape<'_>>> {
    const MAX_SAFE_DEPTH: u16 = 64;

    if limits.max_depth > MAX_SAFE_DEPTH {
        return Err(Error::InvalidLimit {
            limit: Limit::Depth,
            maximum: u32::from(MAX_SAFE_DEPTH),
        });
    }
    if data.is_empty() {
        return Ok(Vec::new());
    }

    let (record, consumed) = Record::parse(data, 0)?;
    if consumed != data.len() {
        return Err(Error::TrailingData { offset: consumed });
    }
    let root = Container::try_new(record)?;
    let mut budget = Budget::new(limits);
    budget.visit()?;
    budget.depth(0)?;

    match root.record().kind() {
        RecordKind::SpgrContainer => root_group(&root, 0, &mut budget),
        RecordKind::DgContainer => drawing(&root, 0, &mut budget),
        RecordKind::SpContainer => Ok(vec![build_shape(
            root,
            Vec::new(),
            Role::Standalone,
            &mut budget,
        )?]),
        RecordKind::DggContainer
        | RecordKind::BStoreContainer
        | RecordKind::SolverContainer
        | RecordKind::Dgg
        | RecordKind::Bse
        | RecordKind::Dg
        | RecordKind::Spgr
        | RecordKind::Sp
        | RecordKind::Opt
        | RecordKind::ClientTextbox
        | RecordKind::ChildAnchor
        | RecordKind::ClientAnchor
        | RecordKind::ClientData
        | RecordKind::ConnectorRule
        | RecordKind::AlignRule
        | RecordKind::ArcRule
        | RecordKind::ClientRule
        | RecordKind::CalloutRule
        | RecordKind::BlipEmf
        | RecordKind::BlipWmf
        | RecordKind::BlipPict
        | RecordKind::BlipJpeg
        | RecordKind::BlipPng
        | RecordKind::BlipDib
        | RecordKind::BlipTiff
        | RecordKind::ColorMru
        | RecordKind::SplitMenuColors
        | RecordKind::SecondaryOpt
        | RecordKind::TertiaryOpt
        | RecordKind::Unknown(_) => Err(Error::MalformedShape {
            reason: "root is not a drawing or shape container",
        }),
    }
}

fn drawing<'data>(
    container: &Container<'data>,
    depth: u16,
    budget: &mut Budget,
) -> Result<Vec<Shape<'data>>> {
    validate_container_header(container, RecordKind::DgContainer)?;
    budget.depth(depth)?;

    let mut dg = false;
    let mut root_group_seen = false;
    let mut shapes = Vec::new();
    for child_result in container.children() {
        let child = child_result?;
        budget.visit()?;
        match child.kind() {
            RecordKind::Dg => {
                if dg {
                    return Err(Error::MalformedShape {
                        reason: "drawing contains more than one Dg atom",
                    });
                }
                validate_atom(&child, RecordKind::Dg, 0, None, 8)?;
                dg = true;
            },
            RecordKind::SpContainer => {
                let shape =
                    build_shape(Container::try_new(child)?, Vec::new(), Role::Root, budget)?;
                if !shape.flags().contains(Flags::BACKGROUND) {
                    shapes.push(shape);
                }
            },
            RecordKind::SpgrContainer => {
                if root_group_seen {
                    return Err(Error::MalformedShape {
                        reason: "drawing contains more than one root shape group",
                    });
                }
                root_group_seen = true;
                shapes.extend(root_group(
                    &Container::try_new(child)?,
                    next_depth(depth)?,
                    budget,
                )?);
            },
            RecordKind::SolverContainer => {
                validate_container_record(&child, RecordKind::SolverContainer)?;
            },
            RecordKind::Unknown(_) => {},
            RecordKind::DggContainer
            | RecordKind::BStoreContainer
            | RecordKind::DgContainer
            | RecordKind::Dgg
            | RecordKind::Bse
            | RecordKind::Spgr
            | RecordKind::Sp
            | RecordKind::Opt
            | RecordKind::ClientTextbox
            | RecordKind::ChildAnchor
            | RecordKind::ClientAnchor
            | RecordKind::ClientData
            | RecordKind::ConnectorRule
            | RecordKind::AlignRule
            | RecordKind::ArcRule
            | RecordKind::ClientRule
            | RecordKind::CalloutRule
            | RecordKind::BlipEmf
            | RecordKind::BlipWmf
            | RecordKind::BlipPict
            | RecordKind::BlipJpeg
            | RecordKind::BlipPng
            | RecordKind::BlipDib
            | RecordKind::BlipTiff
            | RecordKind::ColorMru
            | RecordKind::SplitMenuColors
            | RecordKind::SecondaryOpt
            | RecordKind::TertiaryOpt => {
                return Err(Error::MalformedShape {
                    reason: "drawing contains an invalid direct child record",
                });
            },
        }
    }
    if !dg {
        return Err(Error::MalformedShape {
            reason: "drawing has no Dg atom",
        });
    }
    Ok(shapes)
}

fn root_group<'data>(
    container: &Container<'data>,
    depth: u16,
    budget: &mut Budget,
) -> Result<Vec<Shape<'data>>> {
    validate_container_header(container, RecordKind::SpgrContainer)?;
    budget.depth(depth)?;
    let mut children = container.children();
    let header = children.next().ok_or(Error::MalformedShape {
        reason: "shape-group container has no group header",
    })??;
    budget.visit()?;
    if header.kind() != RecordKind::SpContainer {
        return Err(Error::MalformedShape {
            reason: "shape-group container does not start with SpContainer",
        });
    }
    let header_container = Container::try_new(header)?;
    let meta = scan_meta(&header_container, budget)?;
    validate_meta(&meta, true, Role::Patriarch)?;

    let mut shapes = Vec::new();
    for child_result in children {
        let child = child_result?;
        budget.visit()?;
        match child.kind() {
            RecordKind::SpContainer => {
                let shape =
                    build_shape(Container::try_new(child)?, Vec::new(), Role::Root, budget)?;
                if !shape.flags().contains(Flags::BACKGROUND) {
                    shapes.push(shape);
                }
            },
            RecordKind::SpgrContainer => shapes.push(build_group(
                Container::try_new(child)?,
                next_depth(depth)?,
                Role::Root,
                budget,
            )?),
            RecordKind::DggContainer
            | RecordKind::BStoreContainer
            | RecordKind::DgContainer
            | RecordKind::SolverContainer
            | RecordKind::Dgg
            | RecordKind::Bse
            | RecordKind::Dg
            | RecordKind::Spgr
            | RecordKind::Sp
            | RecordKind::Opt
            | RecordKind::ClientTextbox
            | RecordKind::ChildAnchor
            | RecordKind::ClientAnchor
            | RecordKind::ClientData
            | RecordKind::ConnectorRule
            | RecordKind::AlignRule
            | RecordKind::ArcRule
            | RecordKind::ClientRule
            | RecordKind::CalloutRule
            | RecordKind::BlipEmf
            | RecordKind::BlipWmf
            | RecordKind::BlipPict
            | RecordKind::BlipJpeg
            | RecordKind::BlipPng
            | RecordKind::BlipDib
            | RecordKind::BlipTiff
            | RecordKind::ColorMru
            | RecordKind::SplitMenuColors
            | RecordKind::SecondaryOpt
            | RecordKind::TertiaryOpt
            | RecordKind::Unknown(_) => {
                return Err(Error::MalformedShape {
                    reason: "shape-group container has a non-shape child",
                });
            },
        }
    }
    Ok(shapes)
}

fn build_group<'data>(
    container: Container<'data>,
    depth: u16,
    role: Role,
    budget: &mut Budget,
) -> Result<Shape<'data>> {
    validate_container_header(&container, RecordKind::SpgrContainer)?;
    budget.depth(depth)?;
    let mut records = container.children();
    let header = records.next().ok_or(Error::MalformedShape {
        reason: "shape-group container has no group header",
    })??;
    budget.visit()?;
    if header.kind() != RecordKind::SpContainer {
        return Err(Error::MalformedShape {
            reason: "shape-group container does not start with SpContainer",
        });
    }
    let meta = Container::try_new(header)?;

    let mut children = Vec::new();
    for child_result in records {
        let child = child_result?;
        budget.visit()?;
        match child.kind() {
            RecordKind::SpContainer => {
                children.push(build_shape(
                    Container::try_new(child)?,
                    Vec::new(),
                    Role::Member,
                    budget,
                )?);
            },
            RecordKind::SpgrContainer => children.push(build_group(
                Container::try_new(child)?,
                next_depth(depth)?,
                Role::Member,
                budget,
            )?),
            RecordKind::DggContainer
            | RecordKind::BStoreContainer
            | RecordKind::DgContainer
            | RecordKind::SolverContainer
            | RecordKind::Dgg
            | RecordKind::Bse
            | RecordKind::Dg
            | RecordKind::Spgr
            | RecordKind::Sp
            | RecordKind::Opt
            | RecordKind::ClientTextbox
            | RecordKind::ChildAnchor
            | RecordKind::ClientAnchor
            | RecordKind::ClientData
            | RecordKind::ConnectorRule
            | RecordKind::AlignRule
            | RecordKind::ArcRule
            | RecordKind::ClientRule
            | RecordKind::CalloutRule
            | RecordKind::BlipEmf
            | RecordKind::BlipWmf
            | RecordKind::BlipPict
            | RecordKind::BlipJpeg
            | RecordKind::BlipPng
            | RecordKind::BlipDib
            | RecordKind::BlipTiff
            | RecordKind::ColorMru
            | RecordKind::SplitMenuColors
            | RecordKind::SecondaryOpt
            | RecordKind::TertiaryOpt
            | RecordKind::Unknown(_) => {
                return Err(Error::MalformedShape {
                    reason: "shape-group container has a non-shape child",
                });
            },
        }
    }

    build(container, meta, children, true, role, budget)
}

fn build_shape<'data>(
    container: Container<'data>,
    children: Vec<Shape<'data>>,
    role: Role,
    budget: &mut Budget,
) -> Result<Shape<'data>> {
    build(container.clone(), container, children, false, role, budget)
}

fn build<'data>(
    container: Container<'data>,
    meta: Container<'data>,
    children: Vec<Shape<'data>>,
    group: bool,
    role: Role,
    budget: &mut Budget,
) -> Result<Shape<'data>> {
    let mut records = scan_meta(&meta, budget)?;
    let (id, native_kind, flags, anchor) = validate_meta(&records, group, role)?;
    let group_bounds = records.spgr.as_ref().map(Bounds::from_record).transpose()?;
    let props = records.primary.take().unwrap_or_else(Props::new);
    let kind = if group {
        if records
            .tertiary
            .as_ref()
            .and_then(|tertiary| tertiary.get_int(Id::GroupTableProperties))
            .is_some_and(|value| value & 1 != 0)
        {
            Kind::Table
        } else {
            Kind::Group
        }
    } else {
        detect(native_kind, &props)
    };

    Ok(Shape::from_parts(
        kind,
        id,
        native_kind,
        flags,
        props,
        anchor,
        group_bounds,
        children,
        container,
        meta,
        records.client_data,
        records.textbox,
        records.client_anchor,
    ))
}
