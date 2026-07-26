//! General drawing shapes anchored in a sheet's `table:shapes` container.
//!
//! Spreadsheet sheets share the drawing model used by presentations and
//! drawings ([`crate::Shape`]), extended with the spreadsheet anchoring
//! attributes of ODF 1.3 §19.627–§19.633: `table:end-cell-address`,
//! `table:end-x`, `table:end-y`, and `table:table-background`. Picture
//! frames and embedded-object frames remain owned by the dedicated sheet
//! image and embedded-object models and are never represented here.

use super::sheet_image::{validate_length, validate_text};
use crate::odp::{DrawingAttribute, DrawingAttributeNamespace, PresentationBuilder};
use litchi_core::{Error, Result, ShapeType};

/// Safety limit for shapes stored in one sheet's `table:shapes` container.
pub(crate) const MAX_SHAPES_PER_SHEET: usize = 65_536;

const END_CELL_ADDRESS: &str = "end-cell-address";
const END_X: &str = "end-x";
const END_Y: &str = "end-y";
const TABLE_BACKGROUND: &str = "table-background";

/// Spreadsheet anchoring attributes of a sheet-level drawing shape.
///
/// All values are inert metadata: setting an end cell does not move or
/// resize the shape, and addresses are not resolved against sheet names.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SheetShapeAnchor {
    end_cell_address: Option<String>,
    end_x: Option<String>,
    end_y: Option<String>,
    table_background: Option<bool>,
}

impl SheetShapeAnchor {
    /// Create an empty anchor without end-cell attributes.
    pub fn new() -> Self {
        Self::default()
    }

    /// Return the `table:end-cell-address` cell address, when present.
    pub fn end_cell_address(&self) -> Option<&str> {
        self.end_cell_address.as_deref()
    }

    /// Set or remove the `table:end-cell-address` cell address.
    pub fn set_end_cell_address(&mut self, value: Option<String>) -> Result<()> {
        if let Some(value) = &value {
            validate_text(value, "table:end-cell-address", false)?;
        }
        self.end_cell_address = value;
        Ok(())
    }

    /// Return the `table:end-x` offset within the end cell, when present.
    pub fn end_x(&self) -> Option<&str> {
        self.end_x.as_deref()
    }

    /// Set or remove the `table:end-x` offset within the end cell.
    pub fn set_end_x(&mut self, value: Option<String>) -> Result<()> {
        if let Some(value) = &value {
            validate_length(value, "table:end-x", false)?;
        }
        self.end_x = value;
        Ok(())
    }

    /// Return the `table:end-y` offset within the end cell, when present.
    pub fn end_y(&self) -> Option<&str> {
        self.end_y.as_deref()
    }

    /// Set or remove the `table:end-y` offset within the end cell.
    pub fn set_end_y(&mut self, value: Option<String>) -> Result<()> {
        if let Some(value) = &value {
            validate_length(value, "table:end-y", false)?;
        }
        self.end_y = value;
        Ok(())
    }

    /// Return whether the shape is layered below the cell layer.
    pub fn table_background(&self) -> Option<bool> {
        self.table_background
    }

    /// Set or remove the `table:table-background` layering flag.
    pub fn set_table_background(&mut self, value: Option<bool>) {
        self.table_background = value;
    }

    fn is_empty(&self) -> bool {
        self.end_cell_address.is_none()
            && self.end_x.is_none()
            && self.end_y.is_none()
            && self.table_background.is_none()
    }
}

/// A general drawing shape anchored at sheet level in `table:shapes`.
///
/// Wraps the shared drawing-shape model with spreadsheet anchoring. The
/// shape tree is validated on construction and again on serialization.
#[derive(Debug, Clone)]
pub struct SheetShape {
    shape: crate::Shape,
    anchor: SheetShapeAnchor,
}

impl SheetShape {
    /// Wrap a drawing shape without spreadsheet anchoring attributes.
    pub fn new(shape: crate::Shape) -> Result<Self> {
        Self::with_anchor(shape, SheetShapeAnchor::new())
    }

    /// Wrap a drawing shape with spreadsheet anchoring attributes.
    pub fn with_anchor(shape: crate::Shape, anchor: SheetShapeAnchor) -> Result<Self> {
        let sheet_shape = Self { shape, anchor };
        validate_sheet_shape(&sheet_shape)?;
        Ok(sheet_shape)
    }

    /// Return the wrapped drawing shape.
    pub fn shape(&self) -> &crate::Shape {
        &self.shape
    }

    /// Return the wrapped drawing shape mutably.
    ///
    /// Edits are re-validated when the sheet is serialized.
    pub fn shape_mut(&mut self) -> &mut crate::Shape {
        &mut self.shape
    }

    /// Return the spreadsheet anchoring attributes.
    pub fn anchor(&self) -> &SheetShapeAnchor {
        &self.anchor
    }

    /// Return the spreadsheet anchoring attributes mutably.
    pub fn anchor_mut(&mut self) -> &mut SheetShapeAnchor {
        &mut self.anchor
    }

    /// Unwrap into the shared drawing-shape model, discarding anchoring.
    pub fn into_shape(self) -> crate::Shape {
        self.shape
    }
}

fn reserved_anchor_name(local_name: &str) -> bool {
    matches!(
        local_name,
        END_CELL_ADDRESS | END_X | END_Y | TABLE_BACKGROUND
    )
}

/// Validate one sheet shape, including its nested group children.
pub(crate) fn validate_sheet_shape(sheet_shape: &SheetShape) -> Result<()> {
    for attribute in sheet_shape.shape.drawing_attributes() {
        if attribute.namespace() == DrawingAttributeNamespace::Table
            && reserved_anchor_name(attribute.local_name())
        {
            return Err(Error::InvalidFormat(format!(
                "sheet shape attribute 'table:{}' must use the typed SheetShapeAnchor",
                attribute.local_name()
            )));
        }
    }
    validate_shape_tree(&sheet_shape.shape, 0)
}

fn validate_shape_tree(shape: &crate::Shape, depth: usize) -> Result<()> {
    if depth > 64 {
        return Err(Error::InvalidFormat(
            "sheet shape groups exceed 64 levels".to_string(),
        ));
    }
    if shape.drawing_kind().is_none() {
        return Err(Error::InvalidFormat(
            "sheet shapes require an exact ODF drawing element kind".to_string(),
        ));
    }
    match shape.shape_type() {
        ShapeType::TextBox
        | ShapeType::AutoShape
        | ShapeType::Line
        | ShapeType::Connector
        | ShapeType::Group => {},
        ShapeType::Picture => {
            return Err(Error::InvalidFormat(
                "sheet picture frames must use the dedicated sheet-image APIs".to_string(),
            ));
        },
        ShapeType::GraphicFrame | ShapeType::Table => {
            return Err(Error::InvalidFormat(
                "sheet object frames must use the dedicated embedded-object APIs".to_string(),
            ));
        },
        other => {
            return Err(Error::InvalidFormat(format!(
                "spreadsheet sheets cannot contain {other} drawing shapes"
            )));
        },
    }
    if shape.presentation_class().is_some()
        || shape.presentation_placeholder.is_some()
        || shape.presentation_user_transformed.is_some()
    {
        return Err(Error::InvalidFormat(
            "sheet shapes cannot carry presentation placeholder attributes".to_string(),
        ));
    }
    if shape.media().is_some() || shape.image_href().is_some() {
        return Err(Error::InvalidFormat(
            "sheet shapes cannot embed pictures or media plugins".to_string(),
        ));
    }
    for child in shape.children() {
        validate_shape_tree(child, depth + 1)?;
    }
    Ok(())
}

/// Convert a parsed `table:shapes` drawing shape into the sheet model.
///
/// Picture, embedded-object, and presentation-placeholder frames are owned
/// by their dedicated models and yield `None`. Anchoring attributes move
/// from the exact attribute list into the typed [`SheetShapeAnchor`].
pub(crate) fn sheet_shape_from_parsed(mut shape: crate::Shape) -> Result<Option<SheetShape>> {
    if matches!(
        shape.shape_type(),
        ShapeType::Picture | ShapeType::GraphicFrame | ShapeType::Table | ShapeType::Placeholder
    ) {
        return Ok(None);
    }
    let mut anchor = SheetShapeAnchor::new();
    let mut remaining = Vec::with_capacity(shape.drawing_attributes.len());
    for attribute in shape.drawing_attributes.drain(..) {
        if attribute.namespace() == DrawingAttributeNamespace::Table {
            match attribute.local_name() {
                END_CELL_ADDRESS => {
                    anchor.set_end_cell_address(Some(attribute.value().to_string()))?;
                    continue;
                },
                END_X => {
                    anchor.set_end_x(Some(attribute.value().to_string()))?;
                    continue;
                },
                END_Y => {
                    anchor.set_end_y(Some(attribute.value().to_string()))?;
                    continue;
                },
                TABLE_BACKGROUND => {
                    anchor.set_table_background(Some(parse_boolean(
                        attribute.value(),
                        "table:table-background",
                    )?));
                    continue;
                },
                _ => {},
            }
        }
        remaining.push(attribute);
    }
    shape.drawing_attributes = remaining;
    SheetShape::with_anchor(shape, anchor).map(Some)
}

fn parse_boolean(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        other => Err(Error::InvalidFormat(format!(
            "invalid {name} value '{other}'"
        ))),
    }
}

/// Validate a sheet's full shape list against the safety limits.
pub(crate) fn validate_sheet_shapes(shapes: &[SheetShape]) -> Result<()> {
    if shapes.len() > MAX_SHAPES_PER_SHEET {
        return Err(Error::InvalidFormat(format!(
            "sheet exceeds {MAX_SHAPES_PER_SHEET} drawing shapes"
        )));
    }
    shapes.iter().try_for_each(validate_sheet_shape)
}

/// Whether any shape in the sheet lists uses 3D drawing content.
pub(crate) fn sheet_shapes_use_3d(shapes: &[SheetShape]) -> bool {
    fn shape_uses_3d(shape: &crate::Shape) -> bool {
        shape
            .drawing_kind()
            .is_some_and(|kind| kind.is_three_dimensional())
            || shape
                .drawing_attributes()
                .iter()
                .any(|attribute| attribute.namespace() == DrawingAttributeNamespace::Dr3d)
            || shape.children().iter().any(shape_uses_3d)
    }
    shapes.iter().any(|shape| shape_uses_3d(&shape.shape))
}

/// Whether any shape carries inert `office:event-listeners` bindings.
pub(crate) fn sheet_shapes_have_event_listeners(shapes: &[SheetShape]) -> bool {
    fn shape_has_listeners(shape: &crate::Shape) -> bool {
        !shape.event_listeners().is_empty() || shape.children().iter().any(shape_has_listeners)
    }
    shapes.iter().any(|shape| shape_has_listeners(&shape.shape))
}

/// Write a sheet's `table:shapes` container with images and general shapes.
pub(crate) fn write_table_shapes(
    out: &mut String,
    images: &[crate::OdfImage],
    shapes: &[SheetShape],
) -> Result<()> {
    if images.is_empty() && shapes.is_empty() {
        return Ok(());
    }
    validate_sheet_shapes(shapes)?;
    out.push_str("<table:shapes>");
    super::sheet_image::write_sheet_images_content(out, images)?;
    for (index, sheet_shape) in shapes.iter().enumerate() {
        let mut shape = sheet_shape.shape.clone();
        let anchor = &sheet_shape.anchor;
        if !anchor.is_empty() {
            let mut push = |local_name: &str, value: String| -> Result<()> {
                shape.drawing_attributes.push(DrawingAttribute::new(
                    DrawingAttributeNamespace::Table,
                    local_name,
                    value,
                )?);
                Ok(())
            };
            if let Some(value) = anchor.end_cell_address() {
                push(END_CELL_ADDRESS, value.to_string())?;
            }
            if let Some(value) = anchor.end_x() {
                push(END_X, value.to_string())?;
            }
            if let Some(value) = anchor.end_y() {
                push(END_Y, value.to_string())?;
            }
            if let Some(value) = anchor.table_background() {
                push(TABLE_BACKGROUND, value.to_string())?;
            }
        }
        out.push_str(&PresentationBuilder::generate_shape_xml(&shape, index)?);
    }
    out.push_str("</table:shapes>");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::odp::DrawingShapeKind;

    fn rectangle() -> crate::Shape {
        crate::Shape {
            shape_type: ShapeType::AutoShape,
            drawing_kind: Some(DrawingShapeKind::Rectangle),
            name: Some("Box".to_string()),
            x: Some("1cm".to_string()),
            y: Some("2cm".to_string()),
            width: Some("3cm".to_string()),
            height: Some("4cm".to_string()),
            ..crate::Shape::new()
        }
    }

    #[test]
    fn writes_anchor_attributes_on_the_shape_element() {
        let mut anchor = SheetShapeAnchor::new();
        anchor
            .set_end_cell_address(Some("Sheet1.C5".to_string()))
            .unwrap();
        anchor.set_end_x(Some("0.5cm".to_string())).unwrap();
        anchor.set_end_y(Some("0.25cm".to_string())).unwrap();
        anchor.set_table_background(Some(false));
        let shape = SheetShape::with_anchor(rectangle(), anchor).unwrap();

        let mut xml = String::new();
        write_table_shapes(&mut xml, &[], std::slice::from_ref(&shape)).unwrap();
        assert!(xml.starts_with("<table:shapes><draw:rect"));
        assert!(xml.contains(r#"table:end-cell-address="Sheet1.C5""#));
        assert!(xml.contains(r#"table:end-x="0.5cm""#));
        assert!(xml.contains(r#"table:end-y="0.25cm""#));
        assert!(xml.contains(r#"table:table-background="false""#));
        assert!(xml.ends_with("</table:shapes>"));
    }

    #[test]
    fn rejects_untyped_anchor_attributes_and_missing_kinds() {
        let mut raw = rectangle();
        raw.drawing_attributes.push(
            DrawingAttribute::new(
                DrawingAttributeNamespace::Table,
                END_CELL_ADDRESS,
                "Sheet1.A1",
            )
            .unwrap(),
        );
        assert!(SheetShape::new(raw).is_err());

        let mut kindless = rectangle();
        kindless.drawing_kind = None;
        assert!(SheetShape::new(kindless).is_err());
    }

    #[test]
    fn rejects_picture_and_presentation_payloads() {
        let picture = rectangle().with_image_href("Pictures/a.png");
        assert!(SheetShape::new(picture).is_err());

        let mut placeholder = rectangle();
        placeholder.presentation_class = Some("object".to_string());
        assert!(SheetShape::new(placeholder).is_err());
    }

    #[test]
    fn parsed_conversion_extracts_typed_anchors() {
        let mut raw = rectangle();
        raw.drawing_attributes.extend([
            DrawingAttribute::new(
                DrawingAttributeNamespace::Table,
                END_CELL_ADDRESS,
                "Sheet1.B2",
            )
            .unwrap(),
            DrawingAttribute::new(DrawingAttributeNamespace::Table, TABLE_BACKGROUND, "true")
                .unwrap(),
            DrawingAttribute::new(DrawingAttributeNamespace::Drawing, "corner-radius", "0.1cm")
                .unwrap(),
        ]);
        let converted = sheet_shape_from_parsed(raw).unwrap().unwrap();
        assert_eq!(converted.anchor().end_cell_address(), Some("Sheet1.B2"));
        assert_eq!(converted.anchor().table_background(), Some(true));
        assert_eq!(converted.shape().drawing_attributes().len(), 1);

        let image_frame = crate::Shape {
            shape_type: ShapeType::Picture,
            drawing_kind: Some(DrawingShapeKind::Frame),
            ..crate::Shape::new()
        };
        assert!(sheet_shape_from_parsed(image_frame).unwrap().is_none());
    }

    #[test]
    fn anchor_setters_validate_value_spaces() {
        let mut anchor = SheetShapeAnchor::new();
        assert!(anchor.set_end_cell_address(Some(String::new())).is_err());
        assert!(anchor.set_end_x(Some("wide".to_string())).is_err());
        assert!(anchor.set_end_y(Some("1.5km".to_string())).is_err());
        assert!(anchor.set_end_x(Some("-0.5cm".to_string())).is_ok());
        assert!(parse_boolean("TRUE", "table:table-background").is_err());
    }
}
