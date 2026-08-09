/// PPT record types
pub mod record_type {
    pub const DOCUMENT: u16 = 1000;
    pub const DOCUMENT_ATOM: u16 = 1001;
    pub const END_DOCUMENT: u16 = 1002; // RT_EndDocument (POI compatible)
    pub const SLIDE: u16 = 1006;
    pub const SLIDE_ATOM: u16 = 1007;
    pub const NOTES: u16 = 1008;
    pub const NOTES_ATOM: u16 = 1009;
    pub const ENVIRONMENT: u16 = 1010;
    pub const SLIDE_PERSIST_ATOM: u16 = 1011;
    pub const MAIN_MASTER: u16 = 1016;
    pub const SSSLIDEINFO_ATOM: u16 = 1017;
    pub const SLIDE_VIEW_INFO: u16 = 1018;
    pub const GUIDE_ATOM: u16 = 1019;
    pub const VIEW_INFO: u16 = 1020;
    pub const VIEW_INFO_ATOM: u16 = 1021;
    pub const SLIDE_VIEW_INFO_ATOM: u16 = 1022;
    pub const SHEET_PROPERTIES: u16 = 1044;
    pub const VBA_INFO: u16 = 1023;
    pub const VBA_INFO_ATOM: u16 = 1024;
    pub const PP_DRAWING_GROUP: u16 = 1035;
    pub const PP_DRAWING: u16 = 1036;
    pub const FONT_COLLECTION: u16 = 2005;
    pub const FONT_COLLECTION_10: u16 = 2006;
    pub const FONT_ENTITY_ATOM: u16 = 4023;
    pub const FONT_EMBED_FLAGS_10_ATOM: u16 = 0x32C8;
    pub const COLOR_SCHEME_ATOM: u16 = 2032;
    pub const TX_MASTER_STYLE_ATOM: u16 = 4003; // TxMasterStyleAtom
    pub const TX_CF_STYLE_ATOM: u16 = 4004; // TxCFStyleAtom
    pub const TX_PF_STYLE_ATOM: u16 = 4005; // TxPFStyleAtom
    pub const TX_SI_STYLE_ATOM: u16 = 4009; // TxSIStyleAtom
    pub const SR_KINSOKU: u16 = 4040; // SrKinsoku
    pub const SR_KINSOKU_ATOM: u16 = 4050; // SrKinsokuAtom
    pub const HEADERS_FOOTERS: u16 = 4057; // HeadersFooters container
    pub const HEADERS_FOOTERS_ATOM: u16 = 4058; // HeadersFootersAtom
    pub const DOC_INFO_LIST: u16 = 2000; // List container
    pub const SLIDE_LIST_WITH_TEXT: u16 = 4080;
    pub const TEXT_CHARS_ATOM: u16 = 4000;
    pub const TEXT_BYTES_ATOM: u16 = 4008;
    pub const PROG_TAGS: u16 = 5000;
    pub const PROG_BINARY_TAG: u16 = 5002;
    pub const BINARY_TAG_DATA: u16 = 5003;
    pub const CSTRING: u16 = 4026;
    pub const TEXT_HEADER_ATOM: u16 = 3999;
    pub const STYLE_TEXT_PROP_ATOM: u16 = 4001;
    // Escher types (payloads of PPDrawing/PPDrawingGroup)
    pub const DRAWING: u16 = 0xF008;
    pub const DRAWING_GROUP: u16 = 0xF006;
    pub const DG_CONTAINER: u16 = 0xF002;
    pub const SPGR_CONTAINER: u16 = 0xF003;
    pub const SP_CONTAINER: u16 = 0xF004;
    pub const PERSIST_PTR_HOLDER: u16 = 6001; // PersistDirectoryAtom (full)
    pub const PERSIST_PTR_INCREMENTAL_BLOCK: u16 = 6002; // PersistPtrIncrementalBlock (incremental)
    pub const USER_EDIT_ATOM: u16 = 4085;
    pub const INTERACTIVE_INFO: u16 = 4082; // InteractiveInfo container
    pub const INTERACTIVE_INFO_ATOM: u16 = 4083; // InteractiveInfoAtom
    // Comment records (PPT 2000+)
    pub const COMMENT2000: u16 = 12000; // EPP_Comment10 container
    pub const COMMENT2000_ATOM: u16 = 12001; // EPP_CommentAtom10
    // Named/Custom show records
    pub const NAMED_SHOWS: u16 = 1040; // EPP_NamedShows
    pub const NAMED_SHOW: u16 = 1041; // EPP_NamedShow container
    pub const NAMED_SHOW_SLIDES: u16 = 1042; // EPP_NamedShowSlides atom
}

use std::io::Write;

/// Error type for PPT operations
pub type Error = std::io::Error;

/// PPT record header
#[derive(Debug, Clone)]
pub struct RecordHeader {
    /// Record version (4 bits)
    pub version: u8,
    /// Record instance (12 bits)
    pub instance: u16,
    /// Record type
    pub record_type: u16,
    /// Record length (data only, not including header)
    pub length: u32,
}

impl RecordHeader {
    /// Create a new record header
    #[must_use]
    pub fn new(version: u8, instance: u16, record_type: u16, length: u32) -> Self {
        Self {
            version: version & 0x0F,
            instance: instance & 0x0FFF,
            record_type,
            length,
        }
    }

    /// Write the header to a writer (8 bytes)
    ///
    /// # Errors
    ///
    /// Returns an error if serialization fails or the underlying writer reports an error.
    pub fn write<W: Write>(&self, writer: &mut W) -> Result<(), Error> {
        // Combine version and instance into first 2 bytes
        let ver_inst = u16::from(self.version) | ((self.instance & 0x0FFF) << 4);
        writer.write_all(&ver_inst.to_le_bytes())?;

        // Write type (2 bytes)
        writer.write_all(&self.record_type.to_le_bytes())?;

        // Write length (4 bytes)
        writer.write_all(&self.length.to_le_bytes())?;

        Ok(())
    }

    /// Total size including header
    #[must_use]
    pub fn total_size(&self) -> u32 {
        8 + self.length
    }
}

/// PPT record builder
pub struct RecordBuilder {
    header: RecordHeader,
    data: Vec<u8>,
}

impl RecordBuilder {
    /// Create a new record builder
    #[must_use]
    pub fn new(version: u8, instance: u16, record_type: u16) -> Self {
        Self {
            header: RecordHeader::new(version, instance, record_type, 0),
            data: Vec::new(),
        }
    }

    /// Write data to the record
    #[allow(
        clippy::cast_possible_truncation,
        reason = "record payloads are assembled in memory and bounded far below the u32 length field of the on-disk format"
    )]
    pub fn write_data(&mut self, data: &[u8]) {
        self.data.extend_from_slice(data);
        self.header.length = self.data.len() as u32;
    }

    /// Write a child record
    #[allow(
        clippy::cast_possible_truncation,
        reason = "record payloads are assembled in memory and bounded far below the u32 length field of the on-disk format"
    )]
    pub fn write_child(&mut self, child: &[u8]) {
        self.data.extend_from_slice(child);
        self.header.length = self.data.len() as u32;
    }

    /// Build the complete record (header + data)
    ///
    /// # Errors
    ///
    /// Returns an error if serialization fails or the underlying writer reports an error.
    pub fn build(&self) -> Result<Vec<u8>, Error> {
        let mut record = Vec::new();
        self.header.write(&mut record)?;
        record.extend_from_slice(&self.data);
        Ok(record)
    }

    /// Get the current length
    #[must_use]
    pub fn len(&self) -> u32 {
        self.header.total_size()
    }

    /// Check if record is empty
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.data.is_empty()
    }
}
