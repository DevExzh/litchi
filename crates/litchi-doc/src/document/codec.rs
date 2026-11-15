//! Word binary stream codecs and bounded extraction helpers.

use super::model::Document;
#[cfg(feature = "equation")]
use crate::package::Error as PackageError;
use crate::package::Result;
use crate::parts::fib::FileInformationBlock;
use crate::parts::piece_table::PieceTable;
use litchi_cfb::OleFile;
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::sync::Arc;

impl Document {
    /// Decode the MTEF equations embedded in `ObjectPool`.
    ///
    /// Returns the LaTeX of each decoded equation keyed by its storage name
    /// (`_<id>`), and the storage ids of the equations whose data could not
    /// be decoded, in ascending order.
    #[cfg(feature = "equation")]
    pub(super) fn read_equations<R: Read + Seek>(
        ole: &mut OleFile<R>,
    ) -> Result<(HashMap<String, Arc<str>>, Vec<u32>)> {
        let streams = crate::mtef_extractor::equation_native_streams(ole).map_err(|e| {
            PackageError::InvalidFormat(format!("Failed to list embedded equations: {e}"))
        })?;
        Ok(Self::decode_equations(&streams))
    }

    /// Without the `equation` feature, embedded equations are not read.
    #[cfg(not(feature = "equation"))]
    pub(super) fn read_equations<R: Read + Seek>(
        _ole: &mut OleFile<R>,
    ) -> Result<(HashMap<String, Arc<str>>, Vec<u32>)> {
        Ok((HashMap::new(), Vec::new()))
    }

    /// Decode each `Equation Native` stream to LaTeX; see
    /// [`Self::read_equations`].
    #[cfg(feature = "equation")]
    pub(super) fn decode_equations(
        streams: &HashMap<String, Vec<u8>>,
    ) -> (HashMap<String, Arc<str>>, Vec<u32>) {
        let mut equations = HashMap::with_capacity(streams.len());
        let mut undecodable = Vec::new();
        for (storage, data) in streams {
            match litchi_equation::mtef_to_latex(data, &litchi_equation::Limits::DEFAULT) {
                Ok(latex) => {
                    equations.insert(storage.clone(), Arc::<str>::from(latex));
                },
                Err(_) => undecodable.push(crate::mtef_extractor::storage_id(storage)),
            }
        }
        undecodable.sort_unstable();
        (equations, undecodable)
    }

    pub(super) fn table_slice<'a>(
        fib: &FileInformationBlock,
        table_stream: &'a [u8],
        pointer_index: usize,
    ) -> Option<&'a [u8]> {
        let (offset, length) = fib.get_table_pointer(pointer_index)?;
        let start = usize::try_from(offset).ok()?;
        let length = usize::try_from(length).ok()?;
        if length == 0 {
            return None;
        }
        table_stream.get(start..start.checked_add(length)?)
    }

    /// Parse the CLX once for both property bin tables.
    pub(super) fn parse_piece_table(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Option<PieceTable> {
        Self::table_slice(fib, table_stream, 33).and_then(PieceTable::parse)
    }

    /// Parse the Main Document shape position table (`PlcfSpaMom`), if present.
    ///
    /// A malformed table yields no anchors rather than failing the document;
    /// floating shapes simply lose their positioning information.
    pub(super) fn parse_shape_anchors(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Vec<crate::parts::spa::ShapeAnchor> {
        Self::table_slice(fib, table_stream, crate::parts::spa::FIB_INDEX_PLC_SPA_MOM)
            .and_then(|data| crate::parts::spa::parse_plcf_spa(data).ok())
            .unwrap_or_default()
    }

    /// Parse the Header Document shape position table (`PlcfSpaHdr`), if present.
    pub(super) fn parse_header_shape_anchors(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Vec<crate::parts::spa::ShapeAnchor> {
        Self::table_slice(fib, table_stream, crate::parts::spa::FIB_INDEX_PLC_SPA_HDR)
            .and_then(|data| crate::parts::spa::parse_plcf_spa(data).ok())
            .unwrap_or_default()
    }

    /// Parse a textbox story position table (`PlcftxbxTxt` / `PlcfHdrtxbxTxt`),
    /// if present. A malformed table yields no entries rather than failing
    /// the document.
    pub(super) fn parse_textbox_entries(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        pointer_index: usize,
    ) -> Vec<crate::parts::textbox::TextBoxEntry> {
        Self::table_slice(fib, table_stream, pointer_index)
            .and_then(|data| crate::parts::textbox::parse_plcf_txbx_txt(data).ok())
            .unwrap_or_default()
    }
}
