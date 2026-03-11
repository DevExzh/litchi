//! Tests for common module types and utilities.

#[cfg(test)]
mod error_tests {
    use crate::common::error::Error;

    #[test]
    fn test_error_display() {
        let err = Error::Io(std::io::Error::new(
            std::io::ErrorKind::NotFound,
            "file not found",
        ));
        assert!(err.to_string().contains("file not found"));

        let err = Error::ParseError("parse error".to_string());
        assert!(err.to_string().contains("parse error"));

        let err = Error::InvalidFormat("test format".to_string());
        assert!(err.to_string().contains("test format"));

        let err = Error::NotOfficeFile;
        assert!(err.to_string().contains("Not a valid Office file"));

        let err = Error::CorruptedFile("corrupted".to_string());
        assert!(err.to_string().contains("corrupted"));

        let err = Error::ComponentNotFound("component".to_string());
        assert!(err.to_string().contains("component"));

        let err = Error::XmlError("xml error".to_string());
        assert!(err.to_string().contains("xml error"));

        let err = Error::ZipError("zip error".to_string());
        assert!(err.to_string().contains("zip error"));

        let err = Error::InvalidContentType {
            expected: "expected".to_string(),
            got: "got".to_string(),
        };
        assert!(err.to_string().contains("expected"));
        assert!(err.to_string().contains("got"));

        let err = Error::Unsupported("unsupported feature".to_string());
        assert!(err.to_string().contains("unsupported feature"));

        let err = Error::FeatureDisabled("test_feature".to_string());
        assert!(err.to_string().contains("test_feature"));
        assert!(err.to_string().contains("disabled"));

        let err = Error::Other("other error".to_string());
        assert!(err.to_string().contains("other error"));
    }

    #[test]
    fn test_error_from_io() {
        let io_err = std::io::Error::new(std::io::ErrorKind::NotFound, "test");
        let err: Error = io_err.into();
        assert!(matches!(err, Error::Io(_)));
    }

    #[test]
    fn test_error_from_binary_error() {
        use crate::common::binary::BinaryError;
        let bin_err = BinaryError::InsufficientData {
            expected: 10,
            available: 5,
        };
        let err: Error = bin_err.into();
        assert!(matches!(err, Error::ParseError(_)));
        assert!(err.to_string().contains("Binary"));
    }
}

#[cfg(test)]
mod length_tests {
    use crate::common::Length;

    #[test]
    fn test_length_from_emus() {
        let length = Length::from_emus(914400);
        assert_eq!(length.emus(), 914400);
        assert!((length.inches() - 1.0).abs() < 0.001);
    }

    #[test]
    fn test_length_from_inches() {
        let length = Length::from_inches(2.0);
        assert_eq!(length.emus(), 2 * 914400);
        assert!((length.inches() - 2.0).abs() < 0.001);
    }

    #[test]
    fn test_length_from_cm() {
        let length = Length::from_cm(2.54);
        assert!((length.cm() - 2.54).abs() < 0.001);
        assert!((length.inches() - 1.0).abs() < 0.001);
    }

    #[test]
    fn test_length_points() {
        let length = Length::from_inches(1.0);
        assert!((length.points() - 72.0).abs() < 0.001);

        let length = Length::from_emus(914400 / 2);
        assert!((length.points() - 36.0).abs() < 0.001);
    }

    #[test]
    fn test_length_display() {
        let length = Length::from_inches(1.5);
        assert_eq!(format!("{}", length), "1.50\"");
    }

    #[test]
    fn test_length_clone_copy() {
        let length = Length::from_emus(1000);
        let length2 = length;
        assert_eq!(length.emus(), length2.emus());
        // length is still usable after copy
        assert_eq!(length.emus(), 1000);
    }

    #[test]
    fn test_length_equality() {
        let l1 = Length::from_emus(914400);
        let l2 = Length::from_emus(914400);
        let l3 = Length::from_emus(1000);
        assert_eq!(l1, l2);
        assert_ne!(l1, l3);
    }
}

#[cfg(test)]
mod vertical_position_tests {
    use crate::common::VerticalPosition;

    #[test]
    fn test_vertical_position_default() {
        let pos: VerticalPosition = Default::default();
        assert!(pos.is_normal());
        assert!(!pos.is_superscript());
        assert!(!pos.is_subscript());
    }

    #[test]
    fn test_vertical_position_variants() {
        let normal = VerticalPosition::Normal;
        assert!(normal.is_normal());
        assert!(!normal.is_superscript());
        assert!(!normal.is_subscript());

        let superscript = VerticalPosition::Superscript;
        assert!(!superscript.is_normal());
        assert!(superscript.is_superscript());
        assert!(!superscript.is_subscript());

        let subscript = VerticalPosition::Subscript;
        assert!(!subscript.is_normal());
        assert!(!subscript.is_superscript());
        assert!(subscript.is_subscript());
    }

    #[test]
    fn test_vertical_position_clone_copy() {
        let pos = VerticalPosition::Superscript;
        let pos2 = pos;
        assert!(pos2.is_superscript());
        // pos is still usable after copy
        assert!(pos.is_superscript());
    }

    #[test]
    fn test_vertical_position_equality() {
        assert_eq!(VerticalPosition::Normal, VerticalPosition::Normal);
        assert_eq!(VerticalPosition::Superscript, VerticalPosition::Superscript);
        assert_eq!(VerticalPosition::Subscript, VerticalPosition::Subscript);
        assert_ne!(VerticalPosition::Normal, VerticalPosition::Superscript);
        assert_ne!(VerticalPosition::Superscript, VerticalPosition::Subscript);
    }
}

#[cfg(test)]
mod bom_tests {
    use crate::common::bom::{
        BomKind, UTF8_BOM, UTF16_BE_BOM, UTF16_LE_BOM, UTF32_BE_BOM, UTF32_LE_BOM, strip_bom,
        write_bom,
    };
    use std::io::Cursor;

    #[test]
    fn test_utf8_bom() {
        assert_eq!(&UTF8_BOM[..], &[0xEF, 0xBB, 0xBF]);
        assert_eq!(BomKind::Utf8.as_bytes(), &UTF8_BOM);
        assert_eq!(BomKind::Utf8.len(), 3);
    }

    #[test]
    fn test_utf16_le_bom() {
        assert_eq!(&UTF16_LE_BOM[..], &[0xFF, 0xFE]);
        assert_eq!(BomKind::Utf16Le.as_bytes(), &UTF16_LE_BOM);
        assert_eq!(BomKind::Utf16Le.len(), 2);
    }

    #[test]
    fn test_utf16_be_bom() {
        assert_eq!(&UTF16_BE_BOM[..], &[0xFE, 0xFF]);
        assert_eq!(BomKind::Utf16Be.as_bytes(), &UTF16_BE_BOM);
        assert_eq!(BomKind::Utf16Be.len(), 2);
    }

    #[test]
    fn test_utf32_le_bom() {
        assert_eq!(&UTF32_LE_BOM[..], &[0xFF, 0xFE, 0x00, 0x00]);
        assert_eq!(BomKind::Utf32Le.as_bytes(), &UTF32_LE_BOM);
        assert_eq!(BomKind::Utf32Le.len(), 4);
    }

    #[test]
    fn test_utf32_be_bom() {
        assert_eq!(&UTF32_BE_BOM[..], &[0x00, 0x00, 0xFE, 0xFF]);
        assert_eq!(BomKind::Utf32Be.as_bytes(), &UTF32_BE_BOM);
        assert_eq!(BomKind::Utf32Be.len(), 4);
    }

    #[test]
    fn test_strip_bom_utf8() {
        let data = vec![0xEF, 0xBB, 0xBF, b'h', b'i'];
        let mut cursor = Cursor::new(data);
        let result = strip_bom(&mut cursor).unwrap();
        assert!(result.is_some());
        let (kind, len) = result.unwrap();
        assert_eq!(kind, BomKind::Utf8);
        assert_eq!(len, 3);
        // Check cursor position is after BOM
        assert_eq!(cursor.position(), 3);
    }

    #[test]
    fn test_strip_bom_utf16_le() {
        let data = vec![0xFF, 0xFE, 0x41, 0x00];
        let mut cursor = Cursor::new(data);
        let result = strip_bom(&mut cursor).unwrap();
        assert!(result.is_some());
        let (kind, len) = result.unwrap();
        assert_eq!(kind, BomKind::Utf16Le);
        assert_eq!(len, 2);
        assert_eq!(cursor.position(), 2);
    }

    #[test]
    fn test_strip_bom_utf16_be() {
        let data = vec![0xFE, 0xFF, 0x00, 0x41];
        let mut cursor = Cursor::new(data);
        let result = strip_bom(&mut cursor).unwrap();
        assert!(result.is_some());
        let (kind, len) = result.unwrap();
        assert_eq!(kind, BomKind::Utf16Be);
        assert_eq!(len, 2);
    }

    #[test]
    fn test_strip_bom_none() {
        let data = vec![b'h', b'i'];
        let mut cursor = Cursor::new(data);
        let result = strip_bom(&mut cursor).unwrap();
        assert!(result.is_none());
        // Cursor should be back at start
        assert_eq!(cursor.position(), 0);
    }

    #[test]
    fn test_strip_bom_empty() {
        let data: Vec<u8> = vec![];
        let mut cursor = Cursor::new(data);
        let result = strip_bom(&mut cursor).unwrap();
        assert!(result.is_none());
    }

    #[test]
    fn test_write_bom() {
        let mut buf = Vec::new();
        write_bom(&mut buf, BomKind::Utf8).unwrap();
        assert_eq!(buf, UTF8_BOM);

        let mut buf = Vec::new();
        write_bom(&mut buf, BomKind::Utf16Le).unwrap();
        assert_eq!(buf, UTF16_LE_BOM);

        let mut buf = Vec::new();
        write_bom(&mut buf, BomKind::Utf16Be).unwrap();
        assert_eq!(buf, UTF16_BE_BOM);

        let mut buf = Vec::new();
        write_bom(&mut buf, BomKind::Utf32Le).unwrap();
        assert_eq!(buf, UTF32_LE_BOM);

        let mut buf = Vec::new();
        write_bom(&mut buf, BomKind::Utf32Be).unwrap();
        assert_eq!(buf, UTF32_BE_BOM);
    }

    #[test]
    fn test_bom_kind_equality() {
        assert_eq!(BomKind::Utf8, BomKind::Utf8);
        assert_eq!(BomKind::Utf16Le, BomKind::Utf16Le);
        assert_ne!(BomKind::Utf8, BomKind::Utf16Le);
        assert_ne!(BomKind::Utf16Be, BomKind::Utf16Le);
    }
}

#[cfg(test)]
mod shape_type_tests {
    use crate::common::shapes::ShapeType;
    use std::fmt::Write;

    #[test]
    fn test_shape_type_variants() {
        // Test that all variants can be created and compared
        assert_eq!(ShapeType::TextBox, ShapeType::TextBox);
        assert_eq!(ShapeType::Placeholder, ShapeType::Placeholder);
        assert_eq!(ShapeType::AutoShape, ShapeType::AutoShape);
        assert_eq!(ShapeType::Picture, ShapeType::Picture);
        assert_eq!(ShapeType::Group, ShapeType::Group);
        assert_eq!(ShapeType::Line, ShapeType::Line);
        assert_eq!(ShapeType::Connector, ShapeType::Connector);
        assert_eq!(ShapeType::Table, ShapeType::Table);
        assert_eq!(ShapeType::GraphicFrame, ShapeType::GraphicFrame);
        assert_eq!(ShapeType::Unknown, ShapeType::Unknown);
    }

    #[test]
    fn test_shape_type_inequality() {
        assert_ne!(ShapeType::TextBox, ShapeType::Picture);
        assert_ne!(ShapeType::Line, ShapeType::Connector);
        assert_ne!(ShapeType::Unknown, ShapeType::AutoShape);
    }

    #[test]
    fn test_shape_type_display() {
        let mut buf = String::new();
        write!(&mut buf, "{}", ShapeType::TextBox).unwrap();
        assert_eq!(buf, "TextBox");

        buf.clear();
        write!(&mut buf, "{}", ShapeType::Picture).unwrap();
        assert_eq!(buf, "Picture");

        buf.clear();
        write!(&mut buf, "{}", ShapeType::Unknown).unwrap();
        assert_eq!(buf, "Unknown");
    }

    #[test]
    fn test_shape_type_clone_copy() {
        let shape = ShapeType::Picture;
        let shape2 = shape;
        assert_eq!(shape, shape2);
        // shape is still usable after copy
        assert_eq!(shape, ShapeType::Picture);
    }
}

#[cfg(test)]
mod placeholder_type_tests {
    use crate::common::shapes::PlaceholderType;
    use std::fmt::Write;

    #[test]
    fn test_placeholder_type_variants() {
        assert_eq!(PlaceholderType::Title, PlaceholderType::Title);
        assert_eq!(PlaceholderType::Body, PlaceholderType::Body);
        assert_eq!(
            PlaceholderType::CenteredTitle,
            PlaceholderType::CenteredTitle
        );
        assert_eq!(PlaceholderType::Subtitle, PlaceholderType::Subtitle);
        assert_eq!(PlaceholderType::Date, PlaceholderType::Date);
        assert_eq!(PlaceholderType::SlideNumber, PlaceholderType::SlideNumber);
        assert_eq!(PlaceholderType::Footer, PlaceholderType::Footer);
        assert_eq!(PlaceholderType::Header, PlaceholderType::Header);
        assert_eq!(PlaceholderType::Object, PlaceholderType::Object);
        assert_eq!(PlaceholderType::Chart, PlaceholderType::Chart);
        assert_eq!(PlaceholderType::Table, PlaceholderType::Table);
        assert_eq!(PlaceholderType::ClipArt, PlaceholderType::ClipArt);
        assert_eq!(PlaceholderType::Diagram, PlaceholderType::Diagram);
        assert_eq!(PlaceholderType::Media, PlaceholderType::Media);
        assert_eq!(PlaceholderType::Picture, PlaceholderType::Picture);
        assert_eq!(PlaceholderType::Unknown, PlaceholderType::Unknown);
    }

    #[test]
    fn test_placeholder_type_inequality() {
        assert_ne!(PlaceholderType::Title, PlaceholderType::Body);
        assert_ne!(PlaceholderType::Chart, PlaceholderType::Table);
        assert_ne!(PlaceholderType::Picture, PlaceholderType::ClipArt);
    }

    #[test]
    fn test_placeholder_type_display() {
        let mut buf = String::new();
        write!(&mut buf, "{}", PlaceholderType::Title).unwrap();
        assert_eq!(buf, "Title");

        buf.clear();
        write!(&mut buf, "{}", PlaceholderType::Body).unwrap();
        assert_eq!(buf, "Body");

        buf.clear();
        write!(&mut buf, "{}", PlaceholderType::Unknown).unwrap();
        assert_eq!(buf, "Unknown");
    }
}

#[cfg(test)]
mod xml_slice_tests {
    use crate::common::xml_slice::{XmlArenaBuilder, XmlSlice};
    use std::sync::Arc;

    #[test]
    fn test_xml_slice_new() {
        let arena = Arc::new(vec![b'h', b'e', b'l', b'l', b'o']);
        let slice = XmlSlice::new(arena, 0, 5);
        assert_eq!(slice.as_bytes(), b"hello");
        assert_eq!(slice.len(), 5);
        assert!(!slice.is_empty());
        assert_eq!(slice.start(), 0);
    }

    #[test]
    fn test_xml_slice_sub_slice() {
        let arena = Arc::new(vec![
            b'h', b'e', b'l', b'l', b'o', b' ', b'w', b'o', b'r', b'l', b'd',
        ]);
        let slice = XmlSlice::new(arena, 6, 5);
        assert_eq!(slice.as_bytes(), b"world");
        assert_eq!(slice.len(), 5);
        assert_eq!(slice.start(), 6);
    }

    #[test]
    fn test_xml_slice_empty() {
        let arena = Arc::new(vec![]);
        let slice = XmlSlice::new(arena, 0, 0);
        assert!(slice.is_empty());
        assert_eq!(slice.len(), 0);
        assert!(slice.as_bytes().is_empty());
    }

    #[test]
    fn test_xml_slice_arc() {
        let arena = Arc::new(vec![b't', b'e', b's', b't']);
        let slice = XmlSlice::new(arena.clone(), 0, 4);
        let arc = slice.arc();
        assert_eq!(arc.len(), 4);
    }

    #[test]
    fn test_xml_slice_clone() {
        let arena = Arc::new(vec![b'd', b'a', b't', b'a']);
        let slice = XmlSlice::new(arena, 0, 4);
        let slice2 = slice.clone();
        assert_eq!(slice.as_bytes(), slice2.as_bytes());
    }

    #[test]
    fn test_xml_arena_builder() {
        let mut builder = XmlArenaBuilder::with_capacity(100, 10);

        // First element
        let start1 = builder.start_element();
        builder.buffer_mut().extend_from_slice(b"hello");
        let idx1 = builder.finish_element(start1);

        // Second element
        let start2 = builder.start_element();
        builder.buffer_mut().extend_from_slice(b"world");
        let idx2 = builder.finish_element(start2);

        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
        assert_eq!(builder.element_count(), 2);

        let (arena, positions) = builder.build();
        assert_eq!(arena.len(), 10); // "hello" + "world"
        assert_eq!(positions.len(), 2);
        assert_eq!(positions[0], (0, 5)); // "hello" at offset 0, len 5
        assert_eq!(positions[1], (5, 5)); // "world" at offset 5, len 5
    }

    #[test]
    fn test_xml_arena_builder_empty() {
        let builder = XmlArenaBuilder::with_capacity(10, 5);
        assert_eq!(builder.element_count(), 0);
        let (arena, positions) = builder.build();
        assert!(arena.is_empty());
        assert!(positions.is_empty());
    }
}

#[cfg(test)]
mod detection_utils_tests {
    use crate::common::detection::utils::{OLE2_SIGNATURE, ZIP_SIGNATURE, find_in_buffer};

    #[test]
    fn test_ole2_signature() {
        assert_eq!(
            OLE2_SIGNATURE,
            &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
        );
    }

    #[test]
    fn test_zip_signature() {
        assert_eq!(ZIP_SIGNATURE, &[0x50, 0x4B, 0x03, 0x04]);
    }

    #[test]
    fn test_find_in_buffer_found() {
        let buffer = vec![
            0x00, 0x01, 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0xFF,
        ];
        assert!(find_in_buffer(&buffer, OLE2_SIGNATURE));

        let buffer = vec![0x00, 0x50, 0x4B, 0x03, 0x04, 0x00];
        assert!(find_in_buffer(&buffer, ZIP_SIGNATURE));
    }

    #[test]
    fn test_find_in_buffer_not_found() {
        let buffer = vec![0x00, 0x01, 0x02, 0x03, 0x04, 0x05];
        assert!(!find_in_buffer(&buffer, OLE2_SIGNATURE));
        assert!(!find_in_buffer(&buffer, ZIP_SIGNATURE));
    }

    #[test]
    fn test_find_in_buffer_at_start() {
        let mut buffer = OLE2_SIGNATURE.to_vec();
        buffer.extend_from_slice(&[0xFF, 0xFF]);
        assert!(find_in_buffer(&buffer, OLE2_SIGNATURE));
    }

    #[test]
    fn test_find_in_buffer_at_end() {
        let mut buffer = vec![0x00, 0x00];
        buffer.extend_from_slice(OLE2_SIGNATURE);
        assert!(find_in_buffer(&buffer, OLE2_SIGNATURE));
    }

    #[test]
    fn test_find_in_buffer_too_short() {
        let buffer = vec![0xD0, 0xCF];
        assert!(!find_in_buffer(&buffer, OLE2_SIGNATURE));
    }

    #[test]
    fn test_find_in_buffer_empty() {
        let buffer: Vec<u8> = vec![];
        assert!(!find_in_buffer(&buffer, OLE2_SIGNATURE));
        assert!(!find_in_buffer(&buffer, ZIP_SIGNATURE));
    }
}
