//! Shape extraction for Excel workbooks.
//!
//! Excel workbooks store drawing objects in MsoDrawing and MsoDrawingGroup records.
//! This module projects host-neutral OfficeArt shapes and Excel TXO text into
//! an XLS-specific facade.

use litchi_biff::Records;
use litchi_odraw::{
    Record, RecordKind,
    shape::{self, Kind, Shape as DrawShape},
};
use std::collections::VecDeque;

const MSO_DRAWING: u16 = 0x00EC;
const TXO: u16 = 0x01B6;
const CONTINUE: u16 = 0x003C;

/// Shape information extracted from an Excel workbook.
#[derive(Debug, Clone)]
pub struct Shape {
    /// Shape type (rectangle, ellipse, line, etc.)
    pub shape_type: Kind,
    /// Shape ID
    pub shape_id: u32,
    /// Text content extracted from the shape (if any)
    pub text: Option<String>,
    /// Whether this is a group shape
    pub is_group: bool,
    /// Child shapes (for group shapes)
    pub children: Vec<Shape>,
}

impl Shape {
    /// Project a host-neutral OfficeArt shape and its following XLS TXO text.
    fn from_odraw(
        shape: &DrawShape<'_>,
        texts: &mut VecDeque<Option<String>>,
    ) -> std::io::Result<Self> {
        let text = if let Some(textbox) = shape.textbox() {
            // [MS-XLS] stores shape text in the following TXO/CONTINUE BIFF
            // records. Its OfficeArtClientTextbox is an empty boundary atom.
            if !textbox.is_empty() {
                return Err(invalid_data(
                    "XLS OfficeArtClientTextbox payload must be empty",
                ));
            }
            texts
                .pop_front()
                .ok_or_else(|| invalid_data("XLS ClientTextbox has no following TXO record"))?
        } else {
            None
        };
        let children = shape
            .children()
            .iter()
            .map(|child| Self::from_odraw(child, texts))
            .collect::<std::io::Result<_>>()?;

        Ok(Self {
            shape_type: shape.kind(),
            shape_id: shape.id(),
            text,
            is_group: matches!(shape.kind(), Kind::Group | Kind::Table),
            children,
        })
    }
}

fn invalid_data(error: impl ToString) -> std::io::Error {
    std::io::Error::new(std::io::ErrorKind::InvalidData, error.to_string())
}

#[derive(Default)]
struct WorkbookDrawing {
    officeart: Vec<u8>,
    texts: VecDeque<Option<String>>,
}

impl WorkbookDrawing {
    fn parse(data: &[u8]) -> std::io::Result<Self> {
        let mut drawing = Self::default();
        let mut pending: Option<(PendingTxo, bool)> = None;
        let mut drawing_run = false;
        let mut external_txo = false;
        let records = Records::new(data);

        for record in records {
            let record = record.map_err(invalid_data)?;
            let kind = record.kind().get();
            if pending.is_some() && kind != CONTINUE {
                return Err(invalid_data(
                    "XLS TXO text is not followed by the required CONTINUE record",
                ));
            }
            if external_txo && kind != TXO {
                return Err(invalid_data(
                    "standalone XLS ClientTextbox is not followed by a TXO record",
                ));
            }
            if kind != CONTINUE {
                drawing_run = false;
            }

            match kind {
                MSO_DRAWING => {
                    let textbox = standalone_textbox(record.payload())?;
                    if textbox && !officeart_needs_continuation(&drawing.officeart)? {
                        external_txo = true;
                    } else {
                        drawing.officeart.extend_from_slice(record.payload());
                        drawing_run = true;
                    }
                },
                TXO => {
                    if pending.is_some() {
                        return Err(invalid_data(
                            "XLS TXO text is incomplete before the next TXO record",
                        ));
                    }
                    let next = PendingTxo::new(record.payload())?;
                    let attach = !std::mem::take(&mut external_txo);
                    if next.is_complete() {
                        if attach {
                            drawing.texts.push_back(next.finish()?);
                        }
                    } else {
                        pending = Some((next, attach));
                    }
                },
                CONTINUE if pending.is_some() => {
                    let complete = pending
                        .as_mut()
                        .ok_or_else(|| invalid_data("missing XLS TXO state"))?
                        .0
                        .feed(record.payload())?;
                    if complete {
                        let (text, attach) = pending
                            .take()
                            .ok_or_else(|| invalid_data("missing completed XLS TXO state"))?;
                        if attach {
                            drawing.texts.push_back(text.finish()?);
                        }
                    }
                },
                CONTINUE if drawing_run => {
                    drawing.officeart.extend_from_slice(record.payload());
                },
                _ => {},
            }
        }
        if pending.is_some() {
            return Err(invalid_data("truncated XLS TXO text"));
        }
        if external_txo {
            return Err(invalid_data(
                "standalone XLS ClientTextbox has no TXO record",
            ));
        }
        Ok(drawing)
    }

    fn project(mut self) -> std::io::Result<Vec<Shape>> {
        let mut result = Vec::new();
        let mut offset = 0usize;
        while offset < self.officeart.len() {
            let (root, size) = Record::parse(&self.officeart, offset).map_err(invalid_data)?;
            if root.kind() != RecordKind::DgContainer {
                return Err(invalid_data(
                    "XLS MSODRAWING sequence does not start with a DgContainer",
                ));
            }
            let end = offset
                .checked_add(size)
                .ok_or_else(|| invalid_data("XLS OfficeArt drawing extent overflow"))?;
            let shapes = shape::parse(&self.officeart[offset..end]).map_err(invalid_data)?;
            for shape in &shapes {
                result.push(Shape::from_odraw(shape, &mut self.texts)?);
            }
            offset = end;
        }
        if !self.texts.is_empty() {
            return Err(invalid_data(
                "XLS workbook contains TXO text without a ClientTextbox shape",
            ));
        }
        Ok(result)
    }
}

fn standalone_textbox(data: &[u8]) -> std::io::Result<bool> {
    if data.len() < 8 {
        return Ok(false);
    }
    let Ok((record, size)) = Record::parse(data, 0) else {
        return Ok(false);
    };
    if size != data.len() || record.kind() != RecordKind::ClientTextbox {
        return Ok(false);
    }
    if !record.is_empty() {
        return Err(invalid_data(
            "standalone XLS ClientTextbox payload must be empty",
        ));
    }
    Ok(true)
}

/// Whether the final top-level OfficeArt record still needs bytes from a later
/// MsoDrawing BIFF record.
///
/// Excel is allowed to interleave an OBJ record between a shape's ClientData
/// and its final ClientTextbox. In that sequence, the second MsoDrawing record
/// consists solely of the ClientTextbox atom, but it remains part of the open
/// SpContainer. A standalone ClientTextbox after a complete drawing belongs to
/// a comment/control instead and must not be appended to the drawing tree.
fn officeart_needs_continuation(data: &[u8]) -> std::io::Result<bool> {
    let mut offset = 0usize;
    while offset < data.len() {
        let Some(header) = data.get(offset..offset + 8) else {
            return Ok(true);
        };
        let payload_len = usize::try_from(u32::from_le_bytes(
            header[4..8]
                .try_into()
                .map_err(|_| invalid_data("truncated OfficeArt record header"))?,
        ))
        .map_err(|_| invalid_data("OfficeArt record length exceeds usize"))?;
        let end = offset
            .checked_add(8)
            .and_then(|value| value.checked_add(payload_len))
            .ok_or_else(|| invalid_data("OfficeArt record extent overflows"))?;
        if end > data.len() {
            return Ok(true);
        }
        offset = end;
    }
    Ok(false)
}

struct PendingTxo {
    character_count: usize,
    run_byte_count: usize,
    code_units: Vec<u16>,
    run_bytes: usize,
}

impl PendingTxo {
    fn new(data: &[u8]) -> std::io::Result<Self> {
        let character_count = usize::from(read_u16(data, 10, "TXO character count")?);
        let run_byte_count = usize::from(read_u16(data, 12, "TXO run byte count")?);
        let formula_size = usize::from(read_u16(data, 16, "TXO formula size")?);
        let expected = 18usize
            .checked_add(formula_size)
            .ok_or_else(|| invalid_data("XLS TXO payload length overflow"))?;
        if data.len() != expected {
            return Err(invalid_data(format!(
                "XLS TXO payload length is {}, expected {expected}",
                data.len()
            )));
        }

        Ok(Self {
            character_count,
            run_byte_count,
            code_units: Vec::with_capacity(character_count),
            run_bytes: 0,
        })
    }

    fn feed(&mut self, data: &[u8]) -> std::io::Result<bool> {
        if self.code_units.len() < self.character_count {
            let (&flags, bytes) = data
                .split_first()
                .ok_or_else(|| invalid_data("empty XLS TXO text continuation"))?;
            if flags & !1 != 0 {
                return Err(invalid_data("XLS TXO text has reserved encoding flags"));
            }
            let width = if flags & 1 == 0 { 1 } else { 2 };
            let remaining = self.character_count - self.code_units.len();
            let available = bytes.len() / width;
            if width == 2 && remaining > available && !bytes.len().is_multiple_of(2) {
                return Err(invalid_data(
                    "XLS TXO UTF-16 continuation has an odd byte length",
                ));
            }
            let take = remaining.min(available);
            if take == 0 {
                return Err(invalid_data("XLS TXO text continuation has no characters"));
            }
            let character_bytes = take
                .checked_mul(width)
                .ok_or_else(|| invalid_data("XLS TXO character extent overflow"))?;
            if width == 1 {
                self.code_units
                    .extend(bytes[..character_bytes].iter().map(|&byte| u16::from(byte)));
            } else {
                self.code_units.extend(
                    bytes[..character_bytes]
                        .chunks_exact(2)
                        .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
                );
            }
            self.add_runs(&bytes[character_bytes..])?;
        } else {
            self.add_runs(data)?;
        }
        Ok(self.is_complete())
    }

    fn add_runs(&mut self, bytes: &[u8]) -> std::io::Result<()> {
        self.run_bytes = self
            .run_bytes
            .checked_add(bytes.len())
            .ok_or_else(|| invalid_data("XLS TXO run length overflow"))?;
        if self.run_bytes > self.run_byte_count {
            return Err(invalid_data(
                "XLS TXO formatting continuation exceeds its declared length",
            ));
        }
        Ok(())
    }

    fn is_complete(&self) -> bool {
        self.code_units.len() == self.character_count && self.run_bytes == self.run_byte_count
    }

    fn finish(self) -> std::io::Result<Option<String>> {
        if !self.is_complete() {
            return Err(invalid_data("incomplete XLS TXO text"));
        }
        if self.code_units.is_empty() {
            return Ok(None);
        }
        String::from_utf16(&self.code_units)
            .map(Some)
            .map_err(|_| invalid_data("XLS TXO text contains invalid UTF-16"))
    }
}

fn read_u16(data: &[u8], offset: usize, context: &str) -> std::io::Result<u16> {
    let end = offset
        .checked_add(2)
        .ok_or_else(|| invalid_data(format!("XLS {context} offset overflow")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| invalid_data(format!("truncated XLS {context}")))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn walk<'shape>(shapes: &'shape [Shape], out: &mut Vec<&'shape Shape>) {
    for shape in shapes {
        out.push(shape);
        walk(&shape.children, out);
    }
}

/// Extract shapes from Excel workbook MsoDrawing records.
///
/// # Arguments
///
/// * `workbook_data` - The raw workbook stream data
///
/// # Returns
///
/// A vector of shapes found in the workbook.
pub fn extract_shapes_from_workbook(workbook_data: &[u8]) -> std::io::Result<Vec<Shape>> {
    WorkbookDrawing::parse(workbook_data)?.project()
}

/// Extract text from all shapes in an Excel workbook.
///
/// # Arguments
///
/// * `workbook_data` - The raw workbook stream data
///
/// # Returns
///
/// A string containing all text extracted from shapes.
pub fn extract_shape_text_from_workbook(workbook_data: &[u8]) -> std::io::Result<String> {
    let shapes = extract_shapes_from_workbook(workbook_data)?;
    let mut flat = Vec::new();
    walk(&shapes, &mut flat);
    Ok(flat
        .into_iter()
        .filter_map(|shape| shape.text.as_deref())
        .collect::<Vec<_>>()
        .join("\n"))
}

/// Count shapes in an Excel workbook.
///
/// # Arguments
///
/// * `workbook_data` - The raw workbook stream data
///
/// # Returns
///
/// The number of shapes found.
pub fn count_shapes_in_workbook(workbook_data: &[u8]) -> std::io::Result<usize> {
    let shapes = extract_shapes_from_workbook(workbook_data)?;
    let mut flat = Vec::new();
    walk(&shapes, &mut flat);
    Ok(flat.len())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn biff(kind: u16, data: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(data.len() + 4);
        bytes.extend_from_slice(&kind.to_le_bytes());
        bytes.extend_from_slice(&(data.len() as u16).to_le_bytes());
        bytes.extend_from_slice(data);
        bytes
    }

    fn create_test_xls_shape(
        shape_type: Kind,
        shape_id: u32,
        text: Option<String>,
        is_group: bool,
        children: Vec<Shape>,
    ) -> Shape {
        Shape {
            shape_type,
            shape_id,
            text,
            is_group,
            children,
        }
    }

    #[test]
    fn test_xls_shape_creation() {
        let shape = create_test_xls_shape(
            Kind::Rectangle,
            100,
            Some("Shape text".to_string()),
            false,
            vec![],
        );

        assert_eq!(shape.shape_id, 100);
        assert_eq!(shape.text, Some("Shape text".to_string()));
        assert!(!shape.is_group);
        assert!(shape.children.is_empty());
    }

    #[test]
    fn test_xls_shape_group() {
        let child = create_test_xls_shape(Kind::Rectangle, 101, None, false, vec![]);

        let parent = create_test_xls_shape(Kind::Group, 100, None, true, vec![child]);

        assert!(parent.is_group);
        assert_eq!(parent.children.len(), 1);
        assert_eq!(parent.children[0].shape_id, 101);
    }

    #[test]
    fn test_xls_shape_clone() {
        let shape = create_test_xls_shape(
            Kind::Ellipse,
            50,
            Some("Clonable".to_string()),
            false,
            vec![],
        );
        let cloned = shape.clone();

        assert_eq!(cloned.shape_id, shape.shape_id);
        assert_eq!(cloned.text, shape.text);
        assert_eq!(cloned.is_group, shape.is_group);
    }

    #[test]
    fn test_xls_shape_debug() {
        let shape = create_test_xls_shape(
            Kind::Rectangle,
            1,
            Some("Debug test".to_string()),
            false,
            vec![],
        );
        let debug_str = format!("{:?}", shape);

        assert!(debug_str.contains("Shape"));
    }

    #[test]
    fn test_nested_groups() {
        let inner_child =
            create_test_xls_shape(Kind::Rectangle, 3, Some("Inner".to_string()), false, vec![]);

        let middle = create_test_xls_shape(Kind::Group, 2, None, true, vec![inner_child]);

        let outer = create_test_xls_shape(Kind::Group, 1, None, true, vec![middle]);

        assert!(outer.is_group);
        assert_eq!(outer.children.len(), 1);
        assert!(outer.children[0].is_group);
        assert_eq!(outer.children[0].children.len(), 1);
        assert_eq!(outer.children[0].children[0].shape_id, 3);
    }

    #[test]
    fn test_xls_shape_variants() {
        let shape_types = [
            Kind::Rectangle,
            Kind::Ellipse,
            Kind::Line,
            Kind::Group,
            Kind::Picture,
            Kind::TextBox,
            Kind::Polygon,
            Kind::AutoShape,
            Kind::Connector,
            Kind::Unknown,
        ];

        for (i, shape_type) in shape_types.iter().enumerate() {
            let shape = create_test_xls_shape(*shape_type, i as u32, None, false, vec![]);
            assert_eq!(shape.shape_id, i as u32);
            assert_eq!(shape.shape_type, *shape_type);
        }
    }

    #[test]
    fn test_xls_shape_empty_text() {
        let shape = create_test_xls_shape(Kind::TextBox, 1, None, false, vec![]);
        assert!(shape.text.is_none());
    }

    #[test]
    fn test_xls_shape_unicode_text() {
        let shape = create_test_xls_shape(
            Kind::TextBox,
            1,
            Some("Unicode: 你好世界 🎉".to_string()),
            false,
            vec![],
        );
        assert_eq!(shape.text.unwrap(), "Unicode: 你好世界 🎉");
    }

    #[test]
    fn test_extract_shapes_empty_data() {
        let result = extract_shapes_from_workbook(b"");
        assert!(result.is_ok());
        assert!(result.unwrap().is_empty());
    }

    #[test]
    fn test_extract_shape_text_empty_data() {
        let result = extract_shape_text_from_workbook(b"");
        assert!(result.is_ok());
        assert_eq!(result.unwrap(), "");
    }

    #[test]
    fn test_count_shapes_empty_data() {
        let count = count_shapes_in_workbook(b"").unwrap();
        assert_eq!(count, 0);
    }

    #[test]
    fn test_count_shapes_invalid_data() {
        // Random data that isn't valid BIFF
        let data = vec![0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07];
        assert!(count_shapes_in_workbook(&data).is_err());
    }

    #[test]
    fn reconstructs_mso_drawing_continue_fragments() {
        use litchi_odraw::write;

        let mut children = Vec::new();
        write::dg(&mut children, 0, 0).unwrap();
        let mut officeart = Vec::new();
        write::container(&mut officeart, 0, write::Container::Dg, &children).unwrap();
        let mut workbook = biff(MSO_DRAWING, &officeart[..5]);
        workbook.extend_from_slice(&biff(CONTINUE, &officeart[5..]));

        let drawing = WorkbookDrawing::parse(&workbook).unwrap();
        assert_eq!(drawing.officeart, officeart);
        assert!(drawing.project().unwrap().is_empty());
    }

    #[test]
    fn test_multiple_children() {
        let children: Vec<Shape> = (1..=5)
            .map(|i| {
                create_test_xls_shape(
                    Kind::Rectangle,
                    i,
                    Some(format!("Child {}", i)),
                    false,
                    vec![],
                )
            })
            .collect();

        let parent = create_test_xls_shape(Kind::Group, 0, None, true, children);

        assert_eq!(parent.children.len(), 5);
        for (i, child) in parent.children.iter().enumerate() {
            assert_eq!(child.shape_id, (i + 1) as u32);
            assert_eq!(child.text, Some(format!("Child {}", i + 1)));
        }
    }

    #[test]
    fn test_deeply_nested_groups() {
        let level4 = create_test_xls_shape(Kind::Rectangle, 4, None, false, vec![]);
        let level3 = create_test_xls_shape(Kind::Group, 3, None, true, vec![level4]);
        let level2 = create_test_xls_shape(Kind::Group, 2, None, true, vec![level3]);
        let level1 = create_test_xls_shape(Kind::Group, 1, None, true, vec![level2]);

        assert!(level1.is_group);
        assert!(level1.children[0].is_group);
        assert!(level1.children[0].children[0].is_group);
        assert!(!level1.children[0].children[0].children[0].is_group);
        assert_eq!(level1.children[0].children[0].children[0].shape_id, 4);
    }
}
