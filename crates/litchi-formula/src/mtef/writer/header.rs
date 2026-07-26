//! OLE equation header and MTEF header emission
//!
//! An equation embedded in a legacy OLE document (the `Equation Native` stream
//! of an `Equation.3` object) starts with a 28-byte header describing the
//! clipboard format and the payload length, followed by the MTEF header proper:
//!
//! ```text
//! OLE header   cb_hdr:u16  version:u32  format:u16  size:u32  reserved:[u32; 4]
//! MTEF header  version:u8  platform:u8  product:u8  prod_ver:u8  prod_sub:u8
//!              application_key:cstr  inline:u8
//! ```
//!
//! The MTEF header is written without the optional `"(\x04mt"` signature, which
//! matches what MathType and LibreOffice store in `Equation Native`.

use super::error::MtefWriteError;
use super::records::write_u8;
use crate::mtef::constants::*;

/// Reserve the space taken by the OLE header so it can be patched afterwards
///
/// Returns the offset the header starts at, for [`patch_ole_header`].
pub(super) fn reserve_ole_header(out: &mut Vec<u8>) -> usize {
    let offset = out.len();
    out.resize(offset + OLE_HEADER_LEN, 0);
    offset
}

/// Fill in the OLE header reserved by [`reserve_ole_header`]
///
/// `payload_len` is the number of bytes that follow the header, i.e. the MTEF
/// header plus the record stream.
pub(super) fn patch_ole_header(
    out: &mut [u8],
    offset: usize,
    payload_len: usize,
) -> Result<(), MtefWriteError> {
    let size =
        u32::try_from(payload_len).map_err(|_| MtefWriteError::OutputTooLarge(payload_len))?;

    let Some(header) = out.get_mut(offset..offset + OLE_HEADER_LEN) else {
        return Err(MtefWriteError::OutputTooLarge(payload_len));
    };

    // Written by hand rather than through a struct: the on-disk layout is
    // packed, whereas the equivalent Rust struct would be padded.
    header[0..2].copy_from_slice(&OLE_HEADER_CB_HDR.to_le_bytes());
    header[2..6].copy_from_slice(&OLE_HEADER_VERSION.to_le_bytes());
    header[6..8].copy_from_slice(&OLE_CLIPBOARD_FORMAT.to_le_bytes());
    header[8..12].copy_from_slice(&size.to_le_bytes());
    // The four reserved words stay zero; MathType ignores them on read.
    Ok(())
}

/// Write the MTEF 5 header
///
/// The application key is written as an empty C string, which is what MathType
/// emits for equations that are not owned by a specific host application.
pub(super) fn write_mtef_header(out: &mut Vec<u8>, inline: bool) {
    write_u8(out, MTEF_VERSION_5);
    write_u8(out, PLATFORM_WINDOWS);
    write_u8(out, PRODUCT_MATHTYPE);
    write_u8(out, PRODUCT_VERSION);
    write_u8(out, PRODUCT_SUB_VERSION);
    write_u8(out, 0); // empty, NUL-terminated application key
    write_u8(
        out,
        if inline {
            EQUATION_INLINE
        } else {
            EQUATION_DISPLAY
        },
    );
}
