/// ZIP creator system constants used in version_made_by field
pub(crate) const CREATOR_UNIX: u16 = 3;
pub(crate) const CREATOR_MACOS: u16 = 19;
pub(crate) const CREATOR_NTFS: u16 = 11;
pub(crate) const CREATOR_VFAT: u16 = 14;
pub(crate) const CREATOR_FAT: u16 = 0;

/// File mode information for a given zip file entry.
///
/// This represents Unix-style file permissions and type information.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EntryMode(u32);

impl EntryMode {
    /// Creates a new Mode from a raw mode value.
    #[must_use]
    pub(crate) const fn new(value: u32) -> Self {
        Self(value)
    }

    /// Returns the raw mode value
    #[must_use]
    pub const fn value(&self) -> u32 {
        self.0
    }

    /// Returns true if this is a symbolic link.
    #[must_use]
    pub const fn is_symlink(&self) -> bool {
        self.0 & S_IFMT == S_IFLNK
    }

    /// Returns the Unix permission bits (e.g., 0o755).
    #[must_use]
    pub const fn permissions(&self) -> u32 {
        self.0 & 0o777
    }
}

/// Unix file type and permission constants
const S_IFMT: u32 = 0o170000; // File type mask
const S_IFSOCK: u32 = 0o140000; // Socket
const S_IFLNK: u32 = 0o120000; // Symbolic link
const S_IFREG: u32 = 0o100000; // Regular file
const S_IFBLK: u32 = 0o060000; // Block device
const S_IFDIR: u32 = 0o040000; // Directory
const S_IFCHR: u32 = 0o020000; // Character device
const S_IFIFO: u32 = 0o010000; // FIFO
const S_ISUID: u32 = 0o004000; // Set user ID
const S_ISGID: u32 = 0o002000; // Set group ID
const S_ISVTX: u32 = 0o001000; // Sticky bit

/// MSDOS file attribute constants
const MSDOS_DIR: u32 = 0x10;
const MSDOS_READONLY: u32 = 0x01;

/// Converts Unix mode to file mode
pub(crate) fn unix_mode_to_file_mode(m: u32) -> u32 {
    let mut mode = m & 0o777; // Basic permissions

    // Set file type bits based on Unix mode
    match m & S_IFMT {
        S_IFBLK => mode |= S_IFBLK,
        S_IFCHR => mode |= S_IFCHR,
        S_IFDIR => mode |= S_IFDIR,
        S_IFIFO => mode |= S_IFIFO,
        S_IFLNK => mode |= S_IFLNK,
        S_IFSOCK => mode |= S_IFSOCK,
        _ => mode |= S_IFREG, // Default to regular file
    }

    // Set special permission bits
    if m & S_ISGID != 0 {
        mode |= S_ISGID;
    }
    if m & S_ISUID != 0 {
        mode |= S_ISUID;
    }
    if m & S_ISVTX != 0 {
        mode |= S_ISVTX;
    }

    mode
}

/// Converts MSDOS attributes to file mode, following Go's zip reader logic
pub(crate) fn msdos_mode_to_file_mode(m: u32) -> u32 {
    if m & MSDOS_DIR != 0 {
        S_IFDIR | 0o777
    } else if m & MSDOS_READONLY != 0 {
        S_IFREG | 0o444
    } else {
        S_IFREG | 0o666
    }
}
