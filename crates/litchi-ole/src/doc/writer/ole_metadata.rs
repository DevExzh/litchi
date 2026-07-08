//! OLE metadata stream generation for DOC files
//!
//! Microsoft Word requires `\x01CompObj` and `\x01Ole` streams to recognize
//! the file as a valid Word document embedded object.
//!
//! Based on MS-OLEDS specification.

/// Generate the `\x01CompObj` stream
///
/// This stream contains OLE2 embedded object metadata including:
/// - Object CLSID
/// - User type string (display name)
/// - Clipboard format name
/// - ProgID (programmatic identifier)
///
/// # Returns
///
/// Vector of bytes containing the CompObj stream data
pub fn generate_compobj_stream() -> Vec<u8> {
    let mut data = Vec::new();

    // Version (4 bytes): 0x0001FFFE
    data.extend_from_slice(&[0x01, 0x00, 0xFE, 0xFF]);

    // Reserved (4 bytes)
    data.extend_from_slice(&[0x03, 0x0A, 0x00, 0x00]);

    // Reserved (4 bytes)
    data.extend_from_slice(&[0xFF, 0xFF, 0xFF, 0xFF]);

    // CLSID for Word.Document.8: {00020906-0000-0000-C000-000000000046}
    data.extend_from_slice(&[
        0x06, 0x09, 0x02, 0x00, // Data1
        0x00, 0x00, // Data2
        0x00, 0x00, // Data3
        0xC0, 0x00, // Data4[0-1]
        0x00, 0x00, 0x00, 0x00, 0x00, 0x46, // Data4[2-7]
    ]);

    // AnsiUserType string: "Microsoft Word Document" (null-terminated)
    let user_type = b"Microsoft Word Document\0";
    data.extend_from_slice(&(user_type.len() as u32).to_le_bytes());
    data.extend_from_slice(user_type);

    // Clipboard format name: "MSWordDoc" (null-terminated)
    let clip_format = b"MSWordDoc\0";
    data.extend_from_slice(&(clip_format.len() as u32).to_le_bytes());
    data.extend_from_slice(clip_format);

    // ProgID: "Word.Document.8" (null-terminated)
    let prog_id = b"Word.Document.8\0";
    data.extend_from_slice(&(prog_id.len() as u32).to_le_bytes());
    data.extend_from_slice(prog_id);

    // CRITICAL: Unicode marker (REQUIRED by Microsoft Word!)
    // If present, indicates Unicode versions of strings follow
    // Magic value: 0x71B239F4
    data.extend_from_slice(&0x71B239F4u32.to_le_bytes());

    // Unicode strings (empty but present to match Word format)
    // UnicodeUserType (4 bytes = length 0)
    data.extend_from_slice(&0u32.to_le_bytes());

    // UnicodeClipFormat (4 bytes = length 0)
    data.extend_from_slice(&0u32.to_le_bytes());

    // UnicodeProgID (4 bytes = length 0)
    data.extend_from_slice(&0u32.to_le_bytes());

    data
}

/// Generate the `\x01Ole` stream
///
/// This stream contains OLE version information.
///
/// # Returns
///
/// Vector of bytes containing the Ole stream data (20 bytes)
pub fn generate_ole_stream() -> Vec<u8> {
    let mut data = vec![0u8; 20];

    // OLE version: 0x02000001 (OLE 2.1)
    data[0..4].copy_from_slice(&[0x01, 0x00, 0x00, 0x02]);

    // Reserved (16 bytes): all zeros

    data
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_compobj_generation() {
        let compobj = generate_compobj_stream();

        // Check version marker
        assert_eq!(&compobj[0..4], &[0x01, 0x00, 0xFE, 0xFF]);

        // Check CLSID is present
        assert_eq!(
            &compobj[12..28],
            &[
                0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x46
            ]
        );

        // Should be at least 90 bytes
        assert!(compobj.len() >= 90);
    }

    #[test]
    fn test_ole_generation() {
        let ole = generate_ole_stream();

        // Check version
        assert_eq!(&ole[0..4], &[0x01, 0x00, 0x00, 0x02]);

        // Should be exactly 20 bytes
        assert_eq!(ole.len(), 20);
    }
}
