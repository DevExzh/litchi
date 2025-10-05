/// Magic bytes that should be at the beginning of every OLE file
pub const MAGIC: &[u8; 8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";

/// Minimal size of an empty OLE file with 512-byte sectors (1536 bytes)
pub const MINIMAL_OLEFILE_SIZE: usize = 1536;

/// Size of a directory entry in bytes
pub const DIRENTRY_SIZE: usize = 128;

/// Default sector size for version 3 (512 bytes)
pub const SECTOR_SIZE_V3: usize = 512;

/// Default sector size for version 4 (4096 bytes)
pub const SECTOR_SIZE_V4: usize = 4096;

// Sector IDs (from AAF specifications)
/// Maximum regular sector ID
pub const MAXREGSECT: u32 = 0xFFFFFFFA; // -6
/// Denotes a DIFAT sector in a FAT
pub const DIFSECT: u32 = 0xFFFFFFFC; // -4
/// Denotes a FAT sector in a FAT
pub const FATSECT: u32 = 0xFFFFFFFD; // -3
/// End of a virtual stream chain
pub const ENDOFCHAIN: u32 = 0xFFFFFFFE; // -2
/// Unallocated sector
pub const FREESECT: u32 = 0xFFFFFFFF; // -1

// Directory Entry IDs (from AAF specifications)
/// Maximum directory entry ID
pub const MAXREGSID: u32 = 0xFFFFFFFA; // -6
/// Unallocated directory entry
pub const NOSTREAM: u32 = 0xFFFFFFFF; // -1

// Object types in storage (from AAF specifications)
/// Empty directory entry
pub const STGTY_EMPTY: u8 = 0;
/// Element is a storage object
pub const STGTY_STORAGE: u8 = 1;
/// Element is a stream object
pub const STGTY_STREAM: u8 = 2;
/// Element is an ILockBytes object
pub const STGTY_LOCKBYTES: u8 = 3;
/// Element is an IPropertyStorage object
pub const STGTY_PROPERTY: u8 = 4;
/// Element is a root storage
pub const STGTY_ROOT: u8 = 5;

/// Unknown size for a stream (used when size is not known in advance)
pub const UNKNOWN_SIZE: u32 = 0x7FFFFFFF;

// Property types
pub const VT_EMPTY: u16 = 0;
pub const VT_NULL: u16 = 1;
pub const VT_I2: u16 = 2;
pub const VT_I4: u16 = 3;
pub const VT_R4: u16 = 4;
pub const VT_R8: u16 = 5;
pub const VT_CY: u16 = 6;
pub const VT_DATE: u16 = 7;
pub const VT_BSTR: u16 = 8;
pub const VT_DISPATCH: u16 = 9;
pub const VT_ERROR: u16 = 10;
pub const VT_BOOL: u16 = 11;
pub const VT_VARIANT: u16 = 12;
pub const VT_UNKNOWN: u16 = 13;
pub const VT_DECIMAL: u16 = 14;
pub const VT_I1: u16 = 16;
pub const VT_UI1: u16 = 17;
pub const VT_UI2: u16 = 18;
pub const VT_UI4: u16 = 19;
pub const VT_I8: u16 = 20;
pub const VT_UI8: u16 = 21;
pub const VT_INT: u16 = 22;
pub const VT_UINT: u16 = 23;
pub const VT_VOID: u16 = 24;
pub const VT_HRESULT: u16 = 25;
pub const VT_PTR: u16 = 26;
pub const VT_SAFEARRAY: u16 = 27;
pub const VT_CARRAY: u16 = 28;
pub const VT_USERDEFINED: u16 = 29;
pub const VT_LPSTR: u16 = 30;
pub const VT_LPWSTR: u16 = 31;
pub const VT_FILETIME: u16 = 64;
pub const VT_BLOB: u16 = 65;
pub const VT_STREAM: u16 = 66;
pub const VT_STORAGE: u16 = 67;
pub const VT_STREAMED_OBJECT: u16 = 68;
pub const VT_STORED_OBJECT: u16 = 69;
pub const VT_BLOB_OBJECT: u16 = 70;
pub const VT_CF: u16 = 71;
pub const VT_CLSID: u16 = 72;
pub const VT_VECTOR: u16 = 0x1000;

/// Common document type: Microsoft Word
pub const WORD_CLSID: &str = "00020900-0000-0000-C000-000000000046";

