//! Streaming builder that assembles an [`XlsxCustomGeometry`] from the
//! `a:custGeom` subtree of a worksheet drawing part.
//!
//! The shape inventory parser in [`crate::xlsx::shapes`] drives the builder:
//! it routes every DrawingML start/empty event under `a:custGeom` through
//! [`CustomGeometryBuilder::open`], every matching close through
//! [`CustomGeometryBuilder::close`], and finalizes the geometry with
//! [`CustomGeometryBuilder::finish`] when the `a:custGeom` element ends.
//! Unknown elements are skipped inertly; structurally invalid geometry
//! (missing required attributes, wrong point counts, a missing `a:pathLst`)
//! is an error, mirroring the anchor parser's strictness.

use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;

use litchi_ooxml_common::xml::unqualified_attribute_value;
use crate::error::{OoxmlError, Result};

use super::{
    MAX_ADJUST_HANDLES, MAX_CONNECTION_SITES, MAX_GEOMETRY_GUIDES, MAX_GEOMETRY_PATHS,
    MAX_PATH_COMMANDS, XlsxAdjustHandle, XlsxAdjustValue, XlsxConnectionSite, XlsxCustomGeometry,
    XlsxGeometryGuide, XlsxGeometryPath, XlsxGeometryPoint, XlsxGeometryRectangle,
    XlsxPathCommand, XlsxPathFillMode, XlsxPolarAdjustHandle, XlsxXyAdjustHandle,
};

/// One element of the `a:custGeom` subtree, used as parser context so close
/// events can be routed back to the builder.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum GeometryElement {
    /// `a:custGeom`.
    CustomGeometry,
    /// `a:avLst`.
    AdjustValueList,
    /// `a:gdLst`.
    GuideList,
    /// `a:gd`.
    Guide,
    /// `a:ahLst`.
    AdjustHandleList,
    /// `a:ahXY`.
    XyAdjustHandle,
    /// `a:ahPolar`.
    PolarAdjustHandle,
    /// `a:cxnLst`.
    ConnectionSiteList,
    /// `a:cxn`.
    ConnectionSite,
    /// `a:pos`.
    Position,
    /// `a:rect`.
    TextRectangle,
    /// `a:pathLst`.
    PathList,
    /// `a:path`.
    Path,
    /// `a:moveTo`.
    MoveTo,
    /// `a:lnTo`.
    LineTo,
    /// `a:arcTo`.
    ArcTo,
    /// `a:quadBezTo`.
    QuadraticBezierTo,
    /// `a:cubicBezTo`.
    CubicBezierTo,
    /// `a:close`.
    Close,
    /// `a:pt`.
    Point,
}

/// A point-consuming path command under construction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(clippy::enum_variant_names)] // variants mirror DrawingML path command names (moveTo/lnTo/...)
enum PendingCommandKind {
    MoveTo,
    LineTo,
    QuadraticBezierTo,
    CubicBezierTo,
}

impl PendingCommandKind {
    /// The exact `a:pt` count the command requires.
    fn point_count(self) -> usize {
        match self {
            Self::MoveTo | Self::LineTo => 1,
            Self::QuadraticBezierTo => 2,
            Self::CubicBezierTo => 3,
        }
    }

    fn description(self) -> &'static str {
        match self {
            Self::MoveTo => "moveTo",
            Self::LineTo => "lnTo",
            Self::QuadraticBezierTo => "quadBezTo",
            Self::CubicBezierTo => "cubicBezTo",
        }
    }
}

#[derive(Debug)]
struct PendingCommand {
    kind: PendingCommandKind,
    points: Vec<XlsxGeometryPoint>,
}

#[derive(Debug)]
enum PendingSite {
    Xy(XlsxXyAdjustHandle),
    Polar(XlsxPolarAdjustHandle),
    Connection(XlsxConnectionSite),
}

/// Streaming assembler for one `a:custGeom` subtree.
#[derive(Debug, Default)]
pub(crate) struct CustomGeometryBuilder {
    geometry: XlsxCustomGeometry,
    saw_path_list: bool,
    pending_site: Option<(PendingSite, bool)>,
    pending_path: Option<XlsxGeometryPath>,
    pending_command: Option<PendingCommand>,
}

impl CustomGeometryBuilder {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    /// Handle a start or empty event whose parent is `parent`; returns the
    /// element context for known children and `None` for elements to skip.
    pub(crate) fn open(
        &mut self,
        parent: GeometryElement,
        local: &[u8],
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<Option<GeometryElement>> {
        use GeometryElement as El;
        let context = match (parent, local) {
            (El::CustomGeometry, b"avLst") => El::AdjustValueList,
            (El::CustomGeometry, b"gdLst") => El::GuideList,
            (El::CustomGeometry, b"ahLst") => El::AdjustHandleList,
            (El::CustomGeometry, b"cxnLst") => El::ConnectionSiteList,
            (El::CustomGeometry, b"rect") => {
                self.open_text_rectangle(element, decoder)?;
                El::TextRectangle
            },
            (El::CustomGeometry, b"pathLst") => {
                self.saw_path_list = true;
                El::PathList
            },
            (El::AdjustValueList | El::GuideList, b"gd") => {
                self.open_guide(parent, element, decoder)?;
                El::Guide
            },
            (El::AdjustHandleList, b"ahXY") => {
                self.open_xy_handle(element, decoder)?;
                El::XyAdjustHandle
            },
            (El::AdjustHandleList, b"ahPolar") => {
                self.open_polar_handle(element, decoder)?;
                El::PolarAdjustHandle
            },
            (El::ConnectionSiteList, b"cxn") => {
                self.open_connection_site(element, decoder)?;
                El::ConnectionSite
            },
            (El::XyAdjustHandle | El::PolarAdjustHandle | El::ConnectionSite, b"pos") => {
                self.open_position(element, decoder)?;
                El::Position
            },
            (El::PathList, b"path") => {
                self.open_path(element, decoder)?;
                El::Path
            },
            (El::Path, b"moveTo") => self.open_command(PendingCommandKind::MoveTo)?,
            (El::Path, b"lnTo") => self.open_command(PendingCommandKind::LineTo)?,
            (El::Path, b"quadBezTo") => self.open_command(PendingCommandKind::QuadraticBezierTo)?,
            (El::Path, b"cubicBezTo") => self.open_command(PendingCommandKind::CubicBezierTo)?,
            (El::Path, b"arcTo") => {
                self.open_arc(element, decoder)?;
                El::ArcTo
            },
            (El::Path, b"close") => {
                self.push_command(XlsxPathCommand::Close)?;
                El::Close
            },
            (
                El::MoveTo | El::LineTo | El::QuadraticBezierTo | El::CubicBezierTo,
                b"pt",
            ) => {
                self.open_command_point(element, decoder)?;
                El::Point
            },
            _ => return Ok(None),
        };
        Ok(Some(context))
    }

    /// Handle the close event of a known geometry element.
    pub(crate) fn close(&mut self, element: GeometryElement) -> Result<()> {
        match element {
            GeometryElement::XyAdjustHandle
            | GeometryElement::PolarAdjustHandle
            | GeometryElement::ConnectionSite => self.close_site(),
            GeometryElement::Path => self.close_path(),
            GeometryElement::MoveTo
            | GeometryElement::LineTo
            | GeometryElement::QuadraticBezierTo
            | GeometryElement::CubicBezierTo => self.close_command(),
            _ => Ok(()),
        }
    }

    /// Finalize the geometry when `a:custGeom` closes.
    pub(crate) fn finish(self) -> Result<XlsxCustomGeometry> {
        if !self.saw_path_list {
            return Err(invalid("custom geometry is missing its path list"));
        }
        Ok(self.geometry)
    }

    fn open_text_rectangle(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let rectangle = XlsxGeometryRectangle {
            left: required_value(element, b"l", decoder, "text rectangle left edge")?,
            top: required_value(element, b"t", decoder, "text rectangle top edge")?,
            right: required_value(element, b"r", decoder, "text rectangle right edge")?,
            bottom: required_value(element, b"b", decoder, "text rectangle bottom edge")?,
        };
        if self.geometry.text_rectangle.replace(rectangle).is_some() {
            return Err(invalid("custom geometry has duplicate text rectangles"));
        }
        Ok(())
    }

    fn open_guide(
        &mut self,
        list: GeometryElement,
        element: &BytesStart<'_>,
        decoder: Decoder,
    ) -> Result<()> {
        let name = unqualified_attribute_value(element, b"name", decoder)?
            .ok_or_else(|| invalid("geometry guide is missing its name"))?;
        let formula = unqualified_attribute_value(element, b"fmla", decoder)?
            .ok_or_else(|| invalid("geometry guide is missing its formula"))?
            .parse()?;
        let target = if list == GeometryElement::AdjustValueList {
            &mut self.geometry.adjust_values
        } else {
            &mut self.geometry.guides
        };
        if target.len() >= MAX_GEOMETRY_GUIDES {
            return Err(limit("guide count"));
        }
        target.push(XlsxGeometryGuide { name, formula });
        Ok(())
    }

    fn open_xy_handle(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let handle = XlsxXyAdjustHandle {
            horizontal_guide: unqualified_attribute_value(element, b"gdRefX", decoder)?,
            minimum_x: optional_value(element, b"minX", decoder)?,
            maximum_x: optional_value(element, b"maxX", decoder)?,
            vertical_guide: unqualified_attribute_value(element, b"gdRefY", decoder)?,
            minimum_y: optional_value(element, b"minY", decoder)?,
            maximum_y: optional_value(element, b"maxY", decoder)?,
            position: XlsxGeometryPoint::default(),
        };
        self.open_site(PendingSite::Xy(handle))
    }

    fn open_polar_handle(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let handle = XlsxPolarAdjustHandle {
            radius_guide: unqualified_attribute_value(element, b"gdRefR", decoder)?,
            minimum_radius: optional_value(element, b"minR", decoder)?,
            maximum_radius: optional_value(element, b"maxR", decoder)?,
            angle_guide: unqualified_attribute_value(element, b"gdRefAng", decoder)?,
            minimum_angle: optional_value(element, b"minAng", decoder)?,
            maximum_angle: optional_value(element, b"maxAng", decoder)?,
            position: XlsxGeometryPoint::default(),
        };
        self.open_site(PendingSite::Polar(handle))
    }

    fn open_connection_site(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let site = XlsxConnectionSite {
            angle: required_value(element, b"ang", decoder, "connection site angle")?,
            position: XlsxGeometryPoint::default(),
        };
        self.open_site(PendingSite::Connection(site))
    }

    fn open_site(&mut self, site: PendingSite) -> Result<()> {
        let at_limit = match &site {
            PendingSite::Xy(_) | PendingSite::Polar(_) => {
                self.geometry.adjust_handles.len() >= MAX_ADJUST_HANDLES
            },
            PendingSite::Connection(_) => {
                self.geometry.connection_sites.len() >= MAX_CONNECTION_SITES
            },
        };
        if at_limit {
            return Err(limit("handle or connection site count"));
        }
        if self.pending_site.replace((site, false)).is_some() {
            return Err(invalid("nested custom geometry handles or connection sites"));
        }
        Ok(())
    }

    fn open_position(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let position = point_from_attributes(element, decoder)?;
        let (site, has_position) = self
            .pending_site
            .as_mut()
            .ok_or_else(|| invalid("geometry position outside a handle or connection site"))?;
        if *has_position {
            return Err(invalid("custom geometry site has duplicate positions"));
        }
        *has_position = true;
        match site {
            PendingSite::Xy(handle) => handle.position = position,
            PendingSite::Polar(handle) => handle.position = position,
            PendingSite::Connection(connection) => connection.position = position,
        }
        Ok(())
    }

    fn close_site(&mut self) -> Result<()> {
        let (site, has_position) = self
            .pending_site
            .take()
            .ok_or_else(|| invalid("mismatched custom geometry site close"))?;
        if !has_position {
            return Err(invalid("custom geometry site is missing its position"));
        }
        match site {
            PendingSite::Xy(handle) => {
                self.geometry.adjust_handles.push(XlsxAdjustHandle::Xy(handle));
            },
            PendingSite::Polar(handle) => {
                self.geometry.adjust_handles.push(XlsxAdjustHandle::Polar(handle));
            },
            PendingSite::Connection(connection) => {
                self.geometry.connection_sites.push(connection);
            },
        }
        Ok(())
    }

    fn open_path(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        if self.geometry.paths.len() >= MAX_GEOMETRY_PATHS {
            return Err(limit("path count"));
        }
        let defaults = XlsxGeometryPath::default();
        let path = XlsxGeometryPath {
            width: optional_number(element, b"w", decoder, "geometry path width")?
                .unwrap_or(defaults.width),
            height: optional_number(element, b"h", decoder, "geometry path height")?
                .unwrap_or(defaults.height),
            fill_mode: unqualified_attribute_value(element, b"fill", decoder)?
                .as_deref()
                .and_then(XlsxPathFillMode::from_token)
                .unwrap_or(defaults.fill_mode),
            stroked: unqualified_attribute_value(element, b"stroke", decoder)?
                .map_or(defaults.stroked, |value| is_on(&value)),
            extrusion_allowed: unqualified_attribute_value(element, b"extrusionOk", decoder)?
                .map_or(defaults.extrusion_allowed, |value| is_on(&value)),
            commands: Vec::new(),
        };
        if self.pending_path.replace(path).is_some() {
            return Err(invalid("nested custom geometry paths"));
        }
        Ok(())
    }

    fn close_path(&mut self) -> Result<()> {
        let path = self
            .pending_path
            .take()
            .ok_or_else(|| invalid("mismatched custom geometry path close"))?;
        self.geometry.paths.push(path);
        Ok(())
    }

    fn open_command(&mut self, kind: PendingCommandKind) -> Result<GeometryElement> {
        if self.pending_command.replace(PendingCommand {
            kind,
            points: Vec::new(),
        }).is_some()
        {
            return Err(invalid("nested custom geometry path commands"));
        }
        Ok(match kind {
            PendingCommandKind::MoveTo => GeometryElement::MoveTo,
            PendingCommandKind::LineTo => GeometryElement::LineTo,
            PendingCommandKind::QuadraticBezierTo => GeometryElement::QuadraticBezierTo,
            PendingCommandKind::CubicBezierTo => GeometryElement::CubicBezierTo,
        })
    }

    fn open_command_point(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let point = point_from_attributes(element, decoder)?;
        let command = self
            .pending_command
            .as_mut()
            .ok_or_else(|| invalid("geometry point outside a path command"))?;
        if command.points.len() >= command.kind.point_count() {
            return Err(invalid(format!(
                "geometry {} command has too many points",
                command.kind.description()
            )));
        }
        command.points.push(point);
        Ok(())
    }

    fn close_command(&mut self) -> Result<()> {
        let command = self
            .pending_command
            .take()
            .ok_or_else(|| invalid("mismatched custom geometry command close"))?;
        if command.points.len() != command.kind.point_count() {
            return Err(invalid(format!(
                "geometry {} command has the wrong point count",
                command.kind.description()
            )));
        }
        let mut points = command.points.into_iter();
        let mut next = || {
            points
                .next()
                .ok_or_else(|| invalid("geometry command point count mismatch"))
        };
        let command = match command.kind {
            PendingCommandKind::MoveTo => XlsxPathCommand::MoveTo(next()?),
            PendingCommandKind::LineTo => XlsxPathCommand::LineTo(next()?),
            PendingCommandKind::QuadraticBezierTo => XlsxPathCommand::QuadraticBezierTo {
                control: next()?,
                end: next()?,
            },
            PendingCommandKind::CubicBezierTo => XlsxPathCommand::CubicBezierTo {
                control1: next()?,
                control2: next()?,
                end: next()?,
            },
        };
        self.push_command(command)
    }

    fn open_arc(&mut self, element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        let command = XlsxPathCommand::ArcTo {
            width_radius: required_value(element, b"wR", decoder, "arc width radius")?,
            height_radius: required_value(element, b"hR", decoder, "arc height radius")?,
            start_angle: required_value(element, b"stAng", decoder, "arc start angle")?,
            swing_angle: required_value(element, b"swAng", decoder, "arc swing angle")?,
        };
        self.push_command(command)
    }

    fn push_command(&mut self, command: XlsxPathCommand) -> Result<()> {
        let path = self
            .pending_path
            .as_mut()
            .ok_or_else(|| invalid("geometry path command outside a path"))?;
        if path.commands.len() >= MAX_PATH_COMMANDS {
            return Err(limit("path command count"));
        }
        path.commands.push(command);
        Ok(())
    }
}

fn point_from_attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<XlsxGeometryPoint> {
    Ok(XlsxGeometryPoint {
        x: required_value(element, b"x", decoder, "geometry point x coordinate")?,
        y: required_value(element, b"y", decoder, "geometry point y coordinate")?,
    })
}

fn required_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<XlsxAdjustValue> {
    optional_value(element, name, decoder)?
        .ok_or_else(|| invalid(format!("{description} attribute is missing")))
}

fn optional_value(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<XlsxAdjustValue>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| value.parse())
        .transpose()
}

fn optional_number(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<i64>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse()
                .map_err(|_| invalid(format!("invalid {description} '{value}'")))
        })
        .transpose()
}

/// OOXML boolean attribute values: `1`, `true`, and `on` are truthy.
fn is_on(value: &str) -> bool {
    matches!(value, "1" | "true" | "on")
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(name: &str) -> OoxmlError {
    invalid(format!("custom geometry {name} limit exceeded"))
}
