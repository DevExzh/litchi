//! Explicit source profiles used by the DOCX replayable-tail benchmark.
//!
//! The benchmark owns the input capability in this module.  An owned source
//! keeps one caller-provided `Arc<[u8]>`; a file source keeps an explicitly
//! opened [`litchi_core::FileSource`] descriptor plus its prepared
//! length/digest.  A measured iteration clones that pinned descriptor and
//! does not read the file into memory or reopen a mutable pathname.  File
//! population, descriptor opening, and the first fingerprint read belong to
//! the caller's setup phase; [`SourceFileCapability::verify_fingerprint`] is
//! the explicit, potentially expensive preflight when a caller needs to
//! revalidate that setup before starting a sample.
//!
//! Short-read and latency profiles are harness models.  They wrap an explicit
//! positional source, cap every delegated request to a finite non-zero range,
//! and optionally sleep for per-request latency/overhead and returned-byte
//! transfer time.  They never resolve paths, URLs, or network resources.

use std::{
    fmt, io,
    num::{NonZeroU64, NonZeroUsize},
    path::Path,
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
    thread,
    time::Duration,
};

use litchi_core::{FileSource, ReadAt, SourceVersion};
use sha2::{Digest, Sha256};

/// Fixed buffer size used for setup-time file fingerprinting.
pub const FINGERPRINT_BUFFER_BYTES: usize = 64 * 1024;
/// Largest caller-request range accepted by a simulated input profile.
pub const MAX_PROFILE_RANGE_BYTES: usize = 64 * 1024 * 1024;
/// Maximum fixed-plus-transfer service time accepted for one bounded request.
pub const MAX_PROFILE_SERVICE_DELAY: Duration = Duration::from_secs(60);

static NEXT_OWNED_SOURCE_ID: AtomicU64 = AtomicU64::new(1 << 63);

/// Input shape selected by a benchmark case.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum InputMode {
    /// One caller-owned immutable byte vector with normal positional reads.
    Owned,
    /// A caller-selected local file opened through `FileSource`.
    File,
    /// An explicit positional source with bounded short reads and no delay.
    ShortRead,
    /// An explicit positional source with bounded reads and service delays.
    Latency,
}

/// Physical storage selected for an input capability.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum InputStorageKind {
    /// Caller-owned immutable bytes.
    Owned,
    /// A caller-selected local file descriptor.
    File,
}

impl InputStorageKind {
    /// Returns the stable report spelling.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Owned => "owned",
            Self::File => "file",
        }
    }
}

impl fmt::Display for InputStorageKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

impl InputMode {
    /// Parses the stable command/report spelling for one input mode.
    pub fn parse(value: &str) -> Result<Self, InputProfileError> {
        match value {
            "owned" => Ok(Self::Owned),
            "file" => Ok(Self::File),
            "short-read" | "short_read" => Ok(Self::ShortRead),
            "latency" => Ok(Self::Latency),
            _ => Err(InputProfileError::InvalidMode(value.to_owned())),
        }
    }

    /// Returns the stable command/report spelling.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Owned => "owned",
            Self::File => "file",
            Self::ShortRead => "short-read",
            Self::Latency => "latency",
        }
    }
}

impl fmt::Display for InputMode {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Validation failures for an [`InputProfile`].
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum InputProfileError {
    /// The command/report mode spelling was not recognized.
    InvalidMode(String),
    /// A short-read or latency profile requires a positive maximum range.
    InvalidMaxRange,
    /// A configured transfer rate must be positive.
    InvalidBandwidth,
    /// The configured range exceeds the finite harness ceiling.
    RangeTooLarge,
    /// One bounded request could sleep longer than the harness ceiling.
    ServiceDelayTooLarge,
    /// A setting was supplied for a mode whose contract does not use it.
    UnsupportedSetting {
        mode: InputMode,
        setting: &'static str,
    },
}

impl fmt::Display for InputProfileError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidMode(value) => {
                write!(formatter, "invalid input mode {value:?}")
            },
            Self::InvalidMaxRange => {
                formatter.write_str("input max range must be a finite non-zero value")
            },
            Self::InvalidBandwidth => {
                formatter.write_str("input bandwidth must be a finite non-zero value")
            },
            Self::RangeTooLarge => write!(
                formatter,
                "input max range exceeds the {}-byte harness ceiling",
                MAX_PROFILE_RANGE_BYTES
            ),
            Self::ServiceDelayTooLarge => write!(
                formatter,
                "input service delay exceeds {:?} per-request harness ceiling",
                MAX_PROFILE_SERVICE_DELAY
            ),
            Self::UnsupportedSetting { mode, setting } => {
                write!(
                    formatter,
                    "input setting {setting} is not valid for {mode} mode"
                )
            },
        }
    }
}

impl std::error::Error for InputProfileError {}

/// Validated configuration for one source adapter.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct InputProfile {
    mode: InputMode,
    max_range_bytes: Option<NonZeroUsize>,
    per_request_latency: Duration,
    request_overhead: Duration,
    bytes_per_second: Option<NonZeroU64>,
}

impl InputProfile {
    /// Builds a profile from CLI-sized values.
    ///
    /// `ShortRead` and `Latency` require `max_range_bytes`.  `Owned` and
    /// `File` reject range and delay settings so that a report cannot label a
    /// delayed or bounded source as the ordinary input case.  A bandwidth of
    /// zero is rejected rather than silently becoming an unpaced profile.
    pub fn try_new(
        mode: InputMode,
        max_range_bytes: Option<usize>,
        per_request_latency: Duration,
        request_overhead: Duration,
        bytes_per_second: Option<u64>,
    ) -> Result<Self, InputProfileError> {
        let max_range_bytes = match max_range_bytes {
            Some(value) => {
                if value > MAX_PROFILE_RANGE_BYTES {
                    return Err(InputProfileError::RangeTooLarge);
                }
                Some(NonZeroUsize::new(value).ok_or(InputProfileError::InvalidMaxRange)?)
            },
            None => None,
        };
        let bytes_per_second = match bytes_per_second {
            Some(value) => Some(NonZeroU64::new(value).ok_or(InputProfileError::InvalidBandwidth)?),
            None => None,
        };

        match mode {
            InputMode::Owned | InputMode::File => {
                if max_range_bytes.is_some() {
                    return Err(InputProfileError::UnsupportedSetting {
                        mode,
                        setting: "max_range_bytes",
                    });
                }
                if !per_request_latency.is_zero() {
                    return Err(InputProfileError::UnsupportedSetting {
                        mode,
                        setting: "per_request_latency",
                    });
                }
                if !request_overhead.is_zero() {
                    return Err(InputProfileError::UnsupportedSetting {
                        mode,
                        setting: "request_overhead",
                    });
                }
                if bytes_per_second.is_some() {
                    return Err(InputProfileError::UnsupportedSetting {
                        mode,
                        setting: "bytes_per_second",
                    });
                }
            },
            InputMode::ShortRead => {
                if max_range_bytes.is_none() {
                    return Err(InputProfileError::InvalidMaxRange);
                }
                if !per_request_latency.is_zero() {
                    return Err(InputProfileError::UnsupportedSetting {
                        mode,
                        setting: "per_request_latency",
                    });
                }
                if !request_overhead.is_zero() {
                    return Err(InputProfileError::UnsupportedSetting {
                        mode,
                        setting: "request_overhead",
                    });
                }
                if bytes_per_second.is_some() {
                    return Err(InputProfileError::UnsupportedSetting {
                        mode,
                        setting: "bytes_per_second",
                    });
                }
            },
            InputMode::Latency => {
                let max_range_bytes = max_range_bytes.ok_or(InputProfileError::InvalidMaxRange)?;
                validate_service_delay(
                    max_range_bytes,
                    per_request_latency,
                    request_overhead,
                    bytes_per_second,
                )?;
            },
        }

        Ok(Self {
            mode,
            max_range_bytes,
            per_request_latency,
            request_overhead,
            bytes_per_second,
        })
    }

    /// Creates the ordinary caller-owned input profile.
    #[must_use]
    pub const fn owned() -> Self {
        Self {
            mode: InputMode::Owned,
            max_range_bytes: None,
            per_request_latency: Duration::ZERO,
            request_overhead: Duration::ZERO,
            bytes_per_second: None,
        }
    }

    /// Creates the ordinary local-file input profile.
    #[must_use]
    pub const fn file() -> Self {
        Self {
            mode: InputMode::File,
            max_range_bytes: None,
            per_request_latency: Duration::ZERO,
            request_overhead: Duration::ZERO,
            bytes_per_second: None,
        }
    }

    /// Creates a bounded short-read profile.
    #[cfg(test)]
    pub fn short_read(max_range_bytes: usize) -> Result<Self, InputProfileError> {
        Self::try_new(
            InputMode::ShortRead,
            Some(max_range_bytes),
            Duration::ZERO,
            Duration::ZERO,
            None,
        )
    }

    /// Creates a bounded latency/bandwidth profile.
    pub fn latency(
        max_range_bytes: usize,
        per_request_latency: Duration,
        request_overhead: Duration,
        bytes_per_second: Option<u64>,
    ) -> Result<Self, InputProfileError> {
        Self::try_new(
            InputMode::Latency,
            Some(max_range_bytes),
            per_request_latency,
            request_overhead,
            bytes_per_second,
        )
    }

    /// Returns the selected input mode.
    #[must_use]
    pub const fn mode(self) -> InputMode {
        self.mode
    }

    /// Opens one source without adding read-counter instrumentation.
    ///
    /// `Owned` and `File` therefore return their positional source directly;
    /// bounded profiles retain only their requested range/delay adapter.
    pub fn open(self, capability: &InputSourceCapability) -> io::Result<Arc<dyn ReadAt>> {
        let inner = self.open_inner(capability)?;
        if matches!(self.mode, InputMode::Owned | InputMode::File) {
            return Ok(inner);
        }
        Ok(Arc::new(ProfileReadAt {
            inner,
            profile: self,
        }))
    }

    fn open_inner(self, capability: &InputSourceCapability) -> io::Result<Arc<dyn ReadAt>> {
        match (self.mode, capability.is_file()) {
            (InputMode::Owned, true) => {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "owned input profile requires caller-owned bytes",
                ));
            },
            (InputMode::File, false) => {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "file input profile requires a prepared file capability",
                ));
            },
            _ => {},
        }
        let inner = capability.open_for_timing()?;
        let expected_length = capability.fingerprint().len();
        let observed_length = inner.len()?;
        if observed_length != expected_length {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                format!(
                    "input source length changed: expected {expected_length}, observed {observed_length}"
                ),
            ));
        }
        Ok(inner)
    }
}

/// Length and SHA-256 identity prepared for one source snapshot.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct SourceFingerprint {
    len: u64,
    sha256: [u8; 32],
}

impl SourceFingerprint {
    fn from_bytes(bytes: &[u8]) -> Self {
        let digest = Sha256::digest(bytes);
        let mut sha256 = [0_u8; 32];
        sha256.copy_from_slice(&digest);
        Self {
            len: bytes.len() as u64,
            sha256,
        }
    }

    fn from_parts(len: u64, sha256: [u8; 32]) -> Self {
        Self { len, sha256 }
    }

    /// Returns the exact source length used by the prepared identity.
    #[must_use]
    pub const fn len(self) -> u64 {
        self.len
    }

    /// Returns the digest in stable lowercase hexadecimal form.
    #[must_use]
    pub fn sha256_hex(self) -> String {
        hex_digest(&self.sha256)
    }
}

/// A prepared caller-owned local-file capability.
///
/// The path is explicit and remains outside report serialization.  The
/// length/digest are setup-time evidence, and the retained descriptor pins the
/// prepared file object.  `open_for_timing` clones that descriptor for the
/// iteration, while `verify_fingerprint` provides an explicit setup-time
/// revalidation when the descriptor's bytes may have changed.
#[derive(Clone, Debug)]
pub struct SourceFileCapability {
    fingerprint: SourceFingerprint,
    source: FileSource,
    version: SourceVersion,
}

impl SourceFileCapability {
    /// Opens and fingerprints a regular file without retaining its bytes.
    pub fn prepare(path: impl AsRef<Path>) -> io::Result<Self> {
        let source = FileSource::open(path)?;
        let version_before = source.version()?;
        let fingerprint = fingerprint_read_at(&source)?;
        let final_length = source.len()?;
        let version_after = source.version()?;
        if final_length != fingerprint.len() || version_after != version_before {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "file changed while its input fingerprint was prepared",
            ));
        }
        Ok(Self {
            fingerprint,
            source,
            version: version_after,
        })
    }

    /// Returns the setup-time source identity.
    #[must_use]
    pub const fn fingerprint(&self) -> SourceFingerprint {
        self.fingerprint
    }

    /// Hashes the pinned file to revalidate the prepared identity.
    pub fn verify_fingerprint(&self) -> io::Result<()> {
        let source = self.source.clone();
        let version_before = source.version()?;
        if version_before != self.version {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "prepared file source version changed",
            ));
        }
        let current = fingerprint_read_at(&source)?;
        let version_after = source.version()?;
        if version_after != version_before || current != self.fingerprint {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "file input fingerprint changed after preparation",
            ));
        }
        Ok(())
    }

    fn open_for_timing(&self) -> io::Result<Arc<dyn ReadAt>> {
        let source = self.source.clone();
        let version_before = source.version()?;
        if version_before != self.version {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "prepared file source version changed before timed open",
            ));
        }
        let observed_length = source.len()?;
        if observed_length != self.fingerprint.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                format!(
                    "file input length changed: expected {}, observed {observed_length}",
                    self.fingerprint.len()
                ),
            ));
        }
        let version_after = source.version()?;
        if version_after != version_before {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "prepared file source changed during timed open",
            ));
        }
        Ok(Arc::new(source))
    }
}

/// Explicit source capability supplied to an [`InputProfile`].
#[derive(Clone, Debug)]
pub enum InputSourceCapability {
    /// Caller-owned immutable bytes.  Cloning this capability clones only the
    /// `Arc`; opening a profile creates a fresh positional identity around it.
    Owned {
        bytes: Arc<[u8]>,
        fingerprint: SourceFingerprint,
    },
    /// A setup-prepared local-file capability.
    File(SourceFileCapability),
}

impl InputSourceCapability {
    /// Creates an owned capability without copying the supplied byte owner.
    #[must_use]
    pub fn owned(bytes: Arc<[u8]>) -> Self {
        let fingerprint = SourceFingerprint::from_bytes(&bytes);
        Self::Owned { bytes, fingerprint }
    }

    /// Moves a byte vector into the caller-owned capability.
    #[must_use]
    #[cfg(test)]
    pub fn from_bytes(bytes: Vec<u8>) -> Self {
        Self::owned(Arc::from(bytes))
    }

    /// Prepares a file capability outside a measured sample.
    pub fn prepare_file(path: impl AsRef<Path>) -> io::Result<Self> {
        Ok(Self::File(SourceFileCapability::prepare(path)?))
    }

    /// Wraps an already prepared file capability.
    #[must_use]
    #[cfg(test)]
    pub fn from_file_capability(capability: SourceFileCapability) -> Self {
        Self::File(capability)
    }

    /// Returns the setup-time length/digest identity.
    #[must_use]
    pub fn fingerprint(&self) -> SourceFingerprint {
        match self {
            Self::Owned { fingerprint, .. } => *fingerprint,
            Self::File(capability) => capability.fingerprint(),
        }
    }

    /// Returns the base storage kind to report alongside an adapter mode.
    #[must_use]
    pub const fn storage_kind(&self) -> InputStorageKind {
        match self {
            Self::Owned { .. } => InputStorageKind::Owned,
            Self::File(_) => InputStorageKind::File,
        }
    }

    /// Returns whether this capability is backed by a prepared file.
    #[must_use]
    pub const fn is_file(&self) -> bool {
        matches!(self, Self::File(_))
    }

    /// Revalidates a prepared file, or succeeds immediately for owned bytes.
    pub fn verify_fingerprint(&self) -> io::Result<()> {
        match self {
            Self::Owned { .. } => Ok(()),
            Self::File(capability) => capability.verify_fingerprint(),
        }
    }

    fn open_for_timing(&self) -> io::Result<Arc<dyn ReadAt>> {
        match self {
            Self::Owned { bytes, .. } => Ok(Arc::new(OwnedBytesSource {
                bytes: Arc::clone(bytes),
                version: fresh_owned_version(),
            })),
            Self::File(capability) => capability.open_for_timing(),
        }
    }
}

#[derive(Debug)]
struct OwnedBytesSource {
    bytes: Arc<[u8]>,
    version: SourceVersion,
}

impl ReadAt for OwnedBytesSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("owned input length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let Ok(start) = usize::try_from(offset) else {
            return Ok(0);
        };
        let Some(input) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

struct ProfileReadAt {
    inner: Arc<dyn ReadAt>,
    profile: InputProfile,
}

impl ReadAt for ProfileReadAt {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        sleep_if_nonzero(self.profile.per_request_latency);
        let delegated_len = self
            .profile
            .max_range_bytes
            .map_or(output.len(), |maximum| output.len().min(maximum.get()));
        let returned = self.inner.read_at(offset, &mut output[..delegated_len])?;
        if returned > delegated_len {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "input source returned more bytes than the bounded request",
            ));
        }

        sleep_if_nonzero(self.profile.request_overhead);
        if returned != 0
            && let Some(rate) = self.profile.bytes_per_second
        {
            let delay = transfer_delay(returned, rate)?;
            sleep_if_nonzero(delay);
        }

        Ok(returned)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

fn fresh_owned_version() -> SourceVersion {
    SourceVersion::new(NEXT_OWNED_SOURCE_ID.fetch_add(1, Ordering::Relaxed), 0)
}

fn sleep_if_nonzero(duration: Duration) {
    if !duration.is_zero() {
        thread::sleep(duration);
    }
}

fn transfer_delay(bytes: usize, rate: NonZeroU64) -> io::Result<Duration> {
    let bytes = u64::try_from(bytes)
        .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "input transfer overflow"))?;
    let nanos = (u128::from(bytes) * 1_000_000_000).div_ceil(u128::from(rate.get()));
    let nanos = u64::try_from(nanos).map_err(|_| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "input transfer delay does not fit u64 nanoseconds",
        )
    })?;
    Ok(Duration::from_nanos(nanos))
}

fn validate_service_delay(
    max_range_bytes: NonZeroUsize,
    per_request_latency: Duration,
    request_overhead: Duration,
    bytes_per_second: Option<NonZeroU64>,
) -> Result<(), InputProfileError> {
    let fixed = per_request_latency
        .checked_add(request_overhead)
        .ok_or(InputProfileError::ServiceDelayTooLarge)?;
    if fixed > MAX_PROFILE_SERVICE_DELAY {
        return Err(InputProfileError::ServiceDelayTooLarge);
    }
    if let Some(rate) = bytes_per_second {
        let transfer = transfer_delay(max_range_bytes.get(), rate)
            .map_err(|_| InputProfileError::ServiceDelayTooLarge)?;
        if transfer > MAX_PROFILE_SERVICE_DELAY
            || fixed
                .checked_add(transfer)
                .is_none_or(|total| total > MAX_PROFILE_SERVICE_DELAY)
        {
            return Err(InputProfileError::ServiceDelayTooLarge);
        }
    }
    Ok(())
}

fn fingerprint_read_at(source: &dyn ReadAt) -> io::Result<SourceFingerprint> {
    let length = source.len()?;
    let mut hasher = Sha256::new();
    let mut buffer = vec![0_u8; FINGERPRINT_BUFFER_BYTES];
    let mut offset = 0_u64;
    while offset < length {
        let remaining = length - offset;
        let request = usize::try_from(remaining)
            .unwrap_or(FINGERPRINT_BUFFER_BYTES)
            .min(FINGERPRINT_BUFFER_BYTES);
        let read = source.read_at(offset, &mut buffer[..request])?;
        if read == 0 {
            return Err(io::Error::new(
                io::ErrorKind::UnexpectedEof,
                "source ended while its fingerprint was being prepared",
            ));
        }
        if read > request {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "source returned more bytes than its fingerprint request",
            ));
        }
        hasher.update(&buffer[..read]);
        offset = offset
            .checked_add(u64::try_from(read).map_err(|_| {
                io::Error::new(io::ErrorKind::InvalidInput, "fingerprint offset overflow")
            })?)
            .ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidInput, "fingerprint offset overflow")
            })?;
    }
    let digest = hasher.finalize();
    let mut sha256 = [0_u8; 32];
    sha256.copy_from_slice(&digest);
    Ok(SourceFingerprint::from_parts(length, sha256))
}

fn hex_digest(bytes: &[u8; 32]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let mut output = String::with_capacity(bytes.len() * 2);
    for &byte in bytes {
        output.push(char::from(HEX[usize::from(byte >> 4)]));
        output.push(char::from(HEX[usize::from(byte & 0x0f)]));
    }
    output
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "focused adapter tests use direct assertions"
    )]

    use std::{
        fs::OpenOptions,
        io::Write,
        path::PathBuf,
        sync::atomic::{AtomicU64, Ordering},
    };

    use super::*;

    static NEXT_TEST_FILE: AtomicU64 = AtomicU64::new(1);

    struct TestFile(PathBuf);

    impl Drop for TestFile {
        fn drop(&mut self) {
            let _ = std::fs::remove_file(&self.0);
        }
    }

    fn test_file(bytes: &[u8]) -> TestFile {
        let path = std::env::temp_dir().join(format!(
            "litchi-docx-input-profile-{}-{}.bin",
            std::process::id(),
            NEXT_TEST_FILE.fetch_add(1, Ordering::Relaxed)
        ));
        let mut file = OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&path)
            .expect("unique test file");
        file.write_all(bytes).expect("write test file");
        file.sync_all().expect("sync test file");
        TestFile(path)
    }

    fn read_all(source: &dyn ReadAt) -> Vec<u8> {
        let length = usize::try_from(source.len().expect("source length")).expect("test length");
        let mut result = Vec::with_capacity(length);
        let mut offset = 0_u64;
        while result.len() < length {
            let mut buffer = [0_u8; 4];
            let read = source.read_at(offset, &mut buffer).expect("source read");
            assert!(read > 0, "source must make progress");
            result.extend_from_slice(&buffer[..read]);
            offset += u64::try_from(read).expect("test read fits u64");
        }
        result.truncate(length);
        result
    }

    #[test]
    fn owned_source_is_zero_copy_and_each_open_has_fresh_identity() {
        let bytes: Arc<[u8]> = Arc::from(b"source bytes".as_slice());
        let capability = InputSourceCapability::owned(Arc::clone(&bytes));
        let first = InputProfile::owned().open(&capability).expect("first open");
        let second = InputProfile::owned()
            .open(&capability)
            .expect("second open");
        assert_eq!(read_all(first.as_ref()), bytes.as_ref());
        assert_ne!(
            first.version().expect("first version"),
            second.version().expect("second version")
        );
        assert!(Arc::strong_count(&bytes) >= 4);
    }

    #[test]
    fn short_read_reconstructs_exact_bytes_with_bounded_requests() {
        let capability = InputSourceCapability::from_bytes(b"abcdefghij".to_vec());
        let opened = InputProfile::short_read(2)
            .expect("valid range")
            .open(&capability)
            .expect("open");
        assert_eq!(read_all(opened.as_ref()), b"abcdefghij");
    }

    #[test]
    fn file_capability_pins_identity_across_path_replacement() {
        let file = test_file(b"file source");
        let capability = SourceFileCapability::prepare(&file.0).expect("prepare file");
        assert_eq!(capability.fingerprint().len(), 11);
        capability.verify_fingerprint().expect("initial identity");
        let opened = InputProfile::file()
            .open(&InputSourceCapability::from_file_capability(
                capability.clone(),
            ))
            .expect("open file");
        assert_eq!(read_all(opened.as_ref()), b"file source");

        let replacement = file.0.with_extension("replacement");
        std::fs::write(&replacement, b"other data!").expect("write replacement");
        std::fs::rename(&replacement, &file.0).expect("replace pathname");
        capability
            .verify_fingerprint()
            .expect("pinned descriptor must retain the prepared identity");
        let reopened = InputProfile::file()
            .open(&InputSourceCapability::from_file_capability(capability))
            .expect("open pinned file");
        assert_eq!(read_all(reopened.as_ref()), b"file source");
    }

    #[test]
    fn in_place_file_mutation_is_rejected_by_the_prepared_version_fence() {
        let file = test_file(b"stable bytes");
        let capability = SourceFileCapability::prepare(&file.0).expect("prepare file");
        std::fs::write(&file.0, b"mutated bytes").expect("mutate file in place");
        let error = capability
            .verify_fingerprint()
            .expect_err("in-place mutation must fail setup verification");
        assert_eq!(error.kind(), io::ErrorKind::InvalidData);
    }

    #[test]
    fn invalid_range_bandwidth_and_mode_settings_fail_before_open() {
        assert_eq!(
            InputMode::parse("short-read").expect("mode"),
            InputMode::ShortRead
        );
        assert!(InputProfile::short_read(0).is_err());
        assert!(InputProfile::latency(0, Duration::ZERO, Duration::ZERO, None).is_err());
        assert!(
            InputProfile::try_new(
                InputMode::Latency,
                Some(4),
                Duration::ZERO,
                Duration::ZERO,
                Some(0),
            )
            .is_err()
        );
        assert!(
            InputProfile::try_new(
                InputMode::Owned,
                Some(4),
                Duration::ZERO,
                Duration::ZERO,
                None,
            )
            .is_err()
        );
        assert!(
            InputProfile::try_new(
                InputMode::File,
                None,
                Duration::from_nanos(1),
                Duration::ZERO,
                None,
            )
            .is_err()
        );
        assert!(InputProfile::short_read(MAX_PROFILE_RANGE_BYTES + 1).is_err());
        assert!(InputProfile::latency(1, Duration::from_secs(61), Duration::ZERO, None,).is_err());
        assert!(
            InputProfile::latency(
                MAX_PROFILE_RANGE_BYTES,
                Duration::ZERO,
                Duration::ZERO,
                Some(1),
            )
            .is_err()
        );
    }

    #[test]
    fn latency_profile_preserves_fingerprint_and_returned_bytes() {
        let capability = InputSourceCapability::from_bytes(b"latency".to_vec());
        let opened = InputProfile::latency(3, Duration::ZERO, Duration::ZERO, Some(u64::MAX))
            .expect("valid latency profile")
            .open(&capability)
            .expect("open");
        assert_eq!(read_all(opened.as_ref()), b"latency");
    }
}
