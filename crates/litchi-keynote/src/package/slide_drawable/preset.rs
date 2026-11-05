//! The preset a shape's path source draws (record 0372; a port of the
//! host's `shapes::path::shape_preset`, read by `slide_shape_preset` and
//! `slide_shapes`): a rectangle or ellipse Bézier path (the leaf's or
//! Keynote's native ellipse), a rounded rectangle or regular polygon scalar
//! path, or an arrow or star point path. Any other path is no preset.

use super::{Error, Preset, lift};
use litchi_iwa_common::WireLimits;

/// The preset of one `TSWP.ShapeInfoArchive` payload (`None`: a path no
/// preset draws), on the shared `litchi_iwa_style::shape_path` reader
/// (record 0413b).
pub(super) fn preset(shape_info: &[u8], wire: WireLimits) -> Result<Option<Preset>, Error> {
    use litchi_iwa_common::shape::path::Preset as Shared;
    let shared =
        litchi_iwa_style::shape_path::preset(shape_info, wire).map_err(|_| Error::InvalidSource)?;
    Ok(shared.map(|preset| match preset {
        Shared::Rectangle => Preset::Rectangle,
        Shared::RoundedRectangle { corner_radius } => Preset::RoundedRectangle {
            radius: corner_radius.points(),
        },
        Shared::Ellipse => Preset::Ellipse,
        Shared::LeftArrow => Preset::LeftArrow,
        Shared::RightArrow => Preset::RightArrow,
        Shared::DoubleArrow => Preset::DoubleArrow,
        Shared::RegularPolygon { sides } => Preset::RegularPolygon { sides },
        Shared::Star {
            points,
            inner_radius,
        } => Preset::Star {
            points,
            inner_radius,
        },
    }))
}

impl super::Package {
    /// Read the preset one shape's path draws (host `slide_shape_preset`):
    /// `None` for a path no [`Preset`] draws, such as a line, a custom
    /// Bézier path or a callout.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`super::Refusal::Unsupported`]
    /// for drawables other than shapes and lines.
    pub fn slide_shape_preset<'s>(
        &self,
        slide: impl Into<crate::SlideSelector<'s>>,
        drawable: impl Into<crate::DrawableSelector>,
    ) -> Result<Option<Preset>, Error> {
        let selected = super::select_slide(self, slide.into())?;
        let inventory =
            super::read_inventory(self, selected.identifier, super::read_limits(self)?)?;
        let drawable = drawable.into();
        let item = inventory
            .items
            .get(drawable.as_index())
            .ok_or(Error::Refused(super::Refusal::DrawablePositionNotFound {
                position: drawable.as_position(),
            }))?;
        if !matches!(item.kind, super::Kind::Shape | super::Kind::Line) {
            return Err(Error::Refused(super::Refusal::Unsupported {
                reason: super::Unsupported::Kind { kind: item.kind },
            }));
        }
        let wire = self.wire_limits().map_err(|_| Error::InvalidSource)?;
        let workspace = super::Workspace::new(self).map_err(lift)?;
        let object = workspace.require(item.identifier).map_err(lift)?;
        let payload = &object
            .messages
            .get(item.message_index)
            .ok_or(Error::InvalidSource)?
            .data;
        preset(payload, wire)
    }
}

/// The position of the selected shape or line among the slide's owned
/// `TSWP.ShapeInfoArchive` drawables (text boxes included), the selector of
/// the shared caption seam.
pub(in crate::package) fn shape_info_position(
    package: &super::Package,
    slide: crate::SlideSelector<'_>,
    drawable: crate::DrawableSelector,
) -> Result<litchi_core::Position, Error> {
    const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
    let selected = super::select_slide(package, slide)?;
    let inventory =
        super::read_inventory(package, selected.identifier, super::read_limits(package)?)?;
    let item = inventory
        .items
        .get(drawable.as_index())
        .ok_or(Error::Refused(super::Refusal::DrawablePositionNotFound {
            position: drawable.as_position(),
        }))?;
    if !matches!(item.kind, super::Kind::Shape | super::Kind::Line) {
        return Err(Error::Refused(super::Refusal::Unsupported {
            reason: super::Unsupported::Kind { kind: item.kind },
        }));
    }
    Ok(litchi_core::Position::new(
        inventory.items[..drawable.as_index()]
            .iter()
            .filter(|item| item.message_type == SHAPE_INFO_MESSAGE_TYPE)
            .count(),
    ))
}
