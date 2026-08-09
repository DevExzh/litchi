//! Word binary stream codecs and bounded extraction helpers.

use super::model::Document;
#[cfg(feature = "formula")]
use crate::mtef_extractor::MtefExtractor;
#[cfg(feature = "formula")]
use crate::package::Error as PackageError;
use crate::package::Result;
use crate::parts::fib::FileInformationBlock;
use crate::parts::piece_table::PieceTable;
use litchi_cfb::OleFile;
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::sync::Arc;

impl Document {
    /// Extract MTEF data from OLE streams during document initialization
    ///
    /// This method extracts embedded equation objects from the `ObjectPool` directory.
    /// Each embedded equation is stored as a separate OLE object within `ObjectPool`.
    #[cfg(feature = "formula")]
    pub(super) fn extract_mtef_data<R: Read + Seek>(
        ole: &mut OleFile<R>,
    ) -> Result<HashMap<String, Vec<u8>>> {
        // Extract all MTEF formulas from ObjectPool (the primary location for embedded equations)
        let mtef_data = MtefExtractor::extract_all_mtef_from_objectpool(ole).map_err(|e| {
            PackageError::InvalidFormat(format!("Failed to extract MTEF data: {e}"))
        })?;

        // Also try direct stream names for compatibility with older formats
        let mut all_mtef = mtef_data;
        let direct_stream_names = ["Equation Native", "MSWordEquation", "Equation.3"];

        for stream_name in &direct_stream_names {
            if let Ok(Some(data)) = MtefExtractor::extract_mtef_from_stream(ole, &[stream_name]) {
                all_mtef.insert(stream_name.to_string(), data);
            }
        }

        Ok(all_mtef)
    }

    /// Extract MTEF data fallback (when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    pub(super) fn extract_mtef_data<R: Read + Seek>(
        _ole: &mut OleFile<R>,
    ) -> Result<HashMap<String, Vec<u8>>> {
        Ok(HashMap::new())
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

    /// Parse each MTEF stream in a scoped arena and retain an owned rendering.
    ///
    /// `MathNode` is intentionally arena-borrowing. Converting before the local
    /// arena is dropped keeps `Document` an ordinary owning type with no dependent
    /// fields, leaked allocations, or extended lifetimes.
    #[cfg(feature = "formula")]
    pub(super) fn parse_all_mtef_data(
        mtef_data: &HashMap<String, Vec<u8>>,
    ) -> Result<HashMap<String, Arc<str>>> {
        let mut parsed_mtef = HashMap::new();

        for (stream_name, data) in mtef_data {
            let formula = litchi_formula::Formula::new();
            let mut parser = litchi_formula::MtefParser::new(formula.arena(), data);

            if parser.is_valid() {
                match parser.parse() {
                    Ok(nodes) if !nodes.is_empty() => {
                        let mut converter = litchi_formula::LatexConverter::new();
                        let rendered = converter.convert_nodes(&nodes).map_err(|error| {
                            PackageError::InvalidFormat(format!(
                                "Failed to render MTEF formula {stream_name}: {error}"
                            ))
                        })?;
                        parsed_mtef.insert(stream_name.clone(), Arc::<str>::from(rendered));
                    },
                    Ok(_) => {},
                    Err(e) => {
                        parsed_mtef.insert(
                            stream_name.clone(),
                            Arc::<str>::from(format!("[Formula parsing error: {e}]")),
                        );
                    },
                }
            } else {
                parsed_mtef.insert(
                    stream_name.clone(),
                    Arc::<str>::from(format!("[Invalid MTEF format ({} bytes)]", data.len())),
                );
            }
        }

        Ok(parsed_mtef)
    }

    /// Parse all extracted MTEF data fallback (when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    pub(super) fn parse_all_mtef_data(
        _mtef_data: &HashMap<String, Vec<u8>>,
    ) -> Result<HashMap<String, Arc<str>>> {
        Ok(HashMap::new())
    }

    /// Check if text indicates a potential MTEF formula
    pub(super) fn is_potential_mtef_formula(text: &str) -> bool {
        let text = text.trim();

        // Common indicators of MathType equations in text
        text.contains("MathType")
            || text.contains("MTExtra")
            || text.contains('\\')
            || text.contains('{')
            || text.contains('}')
            || (text.len() > 10 && (text.contains('^') || text.contains('_')))
    }

    /// Parse MTEF data for a given text pattern
    #[cfg(feature = "formula")]
    pub(super) fn parse_mtef_for_text(&self, _text: &str) -> Option<Arc<str>> {
        // For now, try to find any parsed MTEF data
        // In a more sophisticated implementation, we'd match specific text patterns
        // to specific MTEF streams

        for parsed_ast in self.parsed_mtef.values() {
            if !parsed_ast.is_empty() {
                return Some(Arc::clone(parsed_ast));
            }
        }

        None
    }

    /// Parse MTEF data for a given text pattern (fallback when formula feature is disabled)
    #[cfg(not(feature = "formula"))]
    pub(super) fn parse_mtef_for_text(&self, _text: &str) -> Option<Arc<str>> {
        None
    }
}
