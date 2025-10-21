/// Fields table parser for Word binary format.
///
/// Based on Apache POI's FieldsTables and Field implementations.
/// Fields in Word documents mark special content like page numbers, dates,
/// and most importantly for us: embedded equations (EMBED Equation.DSMT4).
use super::super::package::Result;
use super::fib::FileInformationBlock;
use crate::ole::plcf::PlcfParser;

/// Field types from Word specification
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum FieldType {
    /// Embedded object field (type 58, 0x3A)
    EmbeddedObject = 58,
    /// Hyperlink field
    Hyperlink = 88,
    /// Page reference
    PageRef = 37,
    /// Other/unknown field type
    Other(u8),
}

impl From<u8> for FieldType {
    fn from(value: u8) -> Self {
        match value {
            58 => FieldType::EmbeddedObject,
            88 => FieldType::Hyperlink,
            37 => FieldType::PageRef,
            other => FieldType::Other(other),
        }
    }
}

/// Field boundary markers
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldBoundary {
    /// Field begin marker (0x13)
    Begin = 0x13,
    /// Field separator marker (0x14)
    Separator = 0x14,
    /// Field end marker (0x15)
    End = 0x15,
}

/// A field descriptor structure (FLD - 2 bytes in MTEF format)
#[derive(Debug, Clone)]
pub struct FieldDescriptor {
    pub boundary_type: u8, // FIELD_BEGIN_MARK, FIELD_SEPARATOR_MARK, or FIELD_END_MARK
    pub field_type: u8,    // Field type (e.g., 58 for EMBED)
    pub flags: u8,         // Various field flags
}

impl FieldDescriptor {
    /// Parse a field descriptor from 2 bytes
    pub fn from_bytes(bytes: &[u8]) -> Option<Self> {
        if bytes.len() < 2 {
            return None;
        }

        // FLD structure (2 bytes):
        // Byte 0: ch (boundary type in bits 0-4, flags in bits 5-7)
        // Byte 1: flt (field type)
        let byte0 = bytes[0];
        let byte1 = bytes[1];

        let boundary_type = byte0 & 0x1F; // Lower 5 bits
        let flags = (byte0 >> 5) & 0x07; // Upper 3 bits
        let field_type = byte1;

        Some(Self {
            boundary_type,
            field_type,
            flags,
        })
    }

    /// Check if this is a field begin marker
    pub fn is_begin(&self) -> bool {
        self.boundary_type == 0x13 // FIELD_BEGIN_MARK
    }

    /// Check if this is a field separator marker
    pub fn is_separator(&self) -> bool {
        self.boundary_type == 0x14 // FIELD_SEPARATOR_MARK
    }

    /// Check if this is a field end marker
    pub fn is_end(&self) -> bool {
        self.boundary_type == 0x15 // FIELD_END_MARK
    }
}

/// A complete field structure with begin, optional separator, and end markers
#[derive(Debug, Clone)]
pub struct Field {
    /// Character position where field starts (begin marker)
    pub start_cp: u32,
    /// Character position of separator (if present)
    pub separator_cp: Option<u32>,
    /// Character position where field ends (end marker)
    pub end_cp: u32,
    /// Field type
    pub field_type: FieldType,
    /// Whether field has a separator
    pub has_separator: bool,
}

impl Field {
    /// Get the character range for the field code (between begin and separator/end)
    pub fn code_range(&self) -> (u32, u32) {
        let end = self.separator_cp.unwrap_or(self.end_cp);
        (self.start_cp + 1, end)
    }

    /// Get the character range for the field result (between separator and end)
    pub fn result_range(&self) -> Option<(u32, u32)> {
        self.separator_cp.map(|sep| (sep + 1, self.end_cp))
    }

    /// Check if this field is an embedded object field
    pub fn is_embedded_object(&self) -> bool {
        self.field_type == FieldType::EmbeddedObject
    }
}

/// Fields table parser
pub struct FieldsTable {
    /// Parsed fields from the main document
    main_document_fields: Vec<Field>,
}

impl FieldsTable {
    /// Parse fields table from table stream
    ///
    /// Based on Apache POI's FieldsTables constructor.
    /// Fields are stored in PLCF (Plex of Character Positions and Properties) structures.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        // Get the PLCF for main document fields (PLCFFLDMOM)
        // FIB offset 266 (fcPlcffldMom) and 270 (lcbPlcffldMom)
        let main_fields = if let Some((offset, length)) = fib.get_table_pointer(11) {
            if length > 0 && (offset as usize) < table_stream.len() {
                let fields_data = &table_stream[offset as usize..];
                let fields_len = length.min((table_stream.len() - offset as usize) as u32) as usize;
                if fields_len >= 4 {
                    // Parse PLCF with 2-byte field descriptors (FLD structure)
                    Self::parse_fields_plcf(&fields_data[..fields_len])
                } else {
                    Vec::new()
                }
            } else {
                Vec::new()
            }
        } else {
            Vec::new()
        };

        Ok(Self {
            main_document_fields: main_fields,
        })
    }

    /// Parse a PLCF structure containing field descriptors
    ///
    /// The PLCF format is:
    /// - Array of (n+1) CPs (character positions) - 4 bytes each
    /// - Array of n FLD structures - 2 bytes each
    fn parse_fields_plcf(data: &[u8]) -> Vec<Field> {
        if data.len() < 6 {
            // Minimum: 2 CPs (8 bytes) + 1 FLD (2 bytes) = 10 bytes
            // But we need at least 3 CPs for a complete field
            return Vec::new();
        }

        // Parse as PLCF with 2-byte properties (FLD structures)
        let plcf = PlcfParser::parse(data, 2);
        if plcf.is_none() {
            return Vec::new();
        }

        let plcf = plcf.unwrap();
        let mut fields = Vec::new();

        // Build fields from field markers
        // Each field consists of: BEGIN marker, optional SEPARATOR marker, END marker
        let mut i = 0;
        while i < plcf.count() {
            if let Some((cp, cp_end)) = plcf.range(i)
                && let Some(fld_data) = plcf.property(i)
                && let Some(descriptor) = FieldDescriptor::from_bytes(fld_data)
                && descriptor.is_begin()
            {
                // Found field begin - look for separator and end
                let start_cp = cp;
                let field_type = FieldType::from(descriptor.field_type);
                let mut separator_cp = None;
                let mut end_cp = cp_end;
                let mut has_separator = false;

                // Scan forward for separator and end markers
                let mut j = i + 1;
                while j < plcf.count() {
                    if let Some((sep_cp, _)) = plcf.range(j)
                        && let Some(next_fld) = plcf.property(j)
                        && let Some(next_desc) = FieldDescriptor::from_bytes(next_fld)
                    {
                        if next_desc.is_separator() {
                            separator_cp = Some(sep_cp);
                            has_separator = true;
                        } else if next_desc.is_end() {
                            end_cp = sep_cp;
                            i = j; // Move past this field
                            break;
                        }
                    }
                    j += 1;
                }

                // Create field
                fields.push(Field {
                    start_cp,
                    separator_cp,
                    end_cp,
                    field_type,
                    has_separator,
                });
            }
            i += 1;
        }

        fields
    }

    /// Get all fields in the main document
    pub fn main_document_fields(&self) -> &[Field] {
        &self.main_document_fields
    }

    /// Find a field at a specific character position
    pub fn find_field_at_position(&self, cp: u32) -> Option<&Field> {
        self.main_document_fields
            .iter()
            .find(|f| f.start_cp <= cp && cp <= f.end_cp)
    }

    /// Get all embedded object fields
    pub fn get_embedded_object_fields(&self) -> Vec<&Field> {
        self.main_document_fields
            .iter()
            .filter(|f| f.is_embedded_object())
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_field_descriptor_parsing() {
        // Test FIELD_BEGIN_MARK with type 58 (embedded object)
        let bytes = [0x13, 58];
        let desc = FieldDescriptor::from_bytes(&bytes).unwrap();
        assert!(desc.is_begin());
        assert_eq!(desc.field_type, 58);

        // Test FIELD_SEPARATOR_MARK
        let bytes = [0x14, 0];
        let desc = FieldDescriptor::from_bytes(&bytes).unwrap();
        assert!(desc.is_separator());

        // Test FIELD_END_MARK
        let bytes = [0x15, 0];
        let desc = FieldDescriptor::from_bytes(&bytes).unwrap();
        assert!(desc.is_end());
    }

    #[test]
    fn test_field_type_conversion() {
        assert_eq!(FieldType::from(58), FieldType::EmbeddedObject);
        assert_eq!(FieldType::from(88), FieldType::Hyperlink);

        match FieldType::from(99) {
            FieldType::Other(99) => {},
            _ => panic!("Expected Other(99)"),
        }
    }
}
