//! Archive-free media asset discovery: the materialized `Data/*` members of
//! an iWork package as a bounded, deterministically ordered catalog.
//!
//! A format owner feeds the package's member names and sizes to
//! [`Catalog::from_members`]; the catalog keeps the `Data/` members, checks
//! basename uniqueness and the [`Limits`], and answers the read queries the
//! migration host's `MediaManager` answered (record 0375).

use std::collections::HashMap;
use std::fmt;
use std::path::{Path, PathBuf};

use super::Type;

const DATA_PREFIX: &str = "Data/";
const HARD_MAX_ASSETS: usize = 100_000;
const HARD_MAX_ASSET_BYTES: u64 = 512 * 1024 * 1024;
const HARD_MAX_TOTAL_BYTES: u64 = 2 * 1024 * 1024 * 1024;

/// The resource a [`Error::Limit`] names.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Discovered assets.
    Assets,
    /// Bytes of one asset.
    AssetBytes,
    /// Aggregate bytes of all assets.
    TotalBytes,
}

/// Media discovery failures.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A limit was 0 or above its hard ceiling.
    #[error("media {kind:?} limit is zero or above its hard ceiling")]
    InvalidLimit {
        /// The resource.
        kind: LimitKind,
    },
    /// Two members share a basename, so basename lookup would be ambiguous.
    #[error("media basenames are ambiguous: {first} and {second}")]
    AmbiguousFileName {
        /// The member found first.
        first: String,
        /// The member found second.
        second: String,
    },
    /// A resource limit was exceeded.
    #[error("media {kind:?} limit exceeded: {observed} > {maximum}")]
    Limit {
        /// The resource.
        kind: LimitKind,
        /// Amount required.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
}

/// Resource ceilings of media discovery and extraction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_assets: usize,
    max_asset_bytes: u64,
    max_total_bytes: u64,
}

impl Limits {
    /// Hard ceiling of the asset count.
    pub const HARD_MAX_ASSETS: usize = HARD_MAX_ASSETS;
    /// Hard ceiling of one asset's bytes.
    pub const HARD_MAX_ASSET_BYTES: u64 = HARD_MAX_ASSET_BYTES;
    /// Hard ceiling of the aggregate bytes.
    pub const HARD_MAX_TOTAL_BYTES: u64 = HARD_MAX_TOTAL_BYTES;

    /// Construct a checked profile no wider than the hard ceilings.
    ///
    /// # Errors
    ///
    /// [`Error::InvalidLimit`] for a zero value or one above its ceiling
    /// (checked in the order assets, asset bytes, total bytes).
    pub const fn new(
        max_assets: usize,
        max_asset_bytes: u64,
        max_total_bytes: u64,
    ) -> Result<Self, Error> {
        if max_assets == 0 || max_assets > HARD_MAX_ASSETS {
            return Err(Error::InvalidLimit {
                kind: LimitKind::Assets,
            });
        }
        if max_asset_bytes == 0 || max_asset_bytes > HARD_MAX_ASSET_BYTES {
            return Err(Error::InvalidLimit {
                kind: LimitKind::AssetBytes,
            });
        }
        if max_total_bytes == 0 || max_total_bytes > HARD_MAX_TOTAL_BYTES {
            return Err(Error::InvalidLimit {
                kind: LimitKind::TotalBytes,
            });
        }
        Ok(Self {
            max_assets,
            max_asset_bytes,
            max_total_bytes,
        })
    }

    /// Maximum asset count.
    #[must_use]
    pub const fn max_assets(self) -> usize {
        self.max_assets
    }

    /// Maximum bytes of one asset.
    #[must_use]
    pub const fn max_asset_bytes(self) -> u64 {
        self.max_asset_bytes
    }

    /// Maximum aggregate bytes.
    #[must_use]
    pub const fn max_total_bytes(self) -> u64 {
        self.max_total_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_assets: HARD_MAX_ASSETS,
            max_asset_bytes: HARD_MAX_ASSET_BYTES,
            max_total_bytes: HARD_MAX_TOTAL_BYTES,
        }
    }
}

/// One materialized media member.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Asset {
    path: PathBuf,
    media_type: Type,
    size: u64,
    file_name: String,
}

impl Asset {
    /// Describe a member: its type comes from the extension and its
    /// basename (`"unknown"` when the path has none) is the lookup key.
    #[must_use]
    pub fn new(path: PathBuf, size: u64) -> Self {
        let file_name = path
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("unknown")
            .to_owned();
        let media_type = path
            .extension()
            .and_then(|extension| extension.to_str())
            .map_or(Type::Unknown, Type::from_extension);
        Self {
            path,
            media_type,
            size,
            file_name,
        }
    }

    /// The member path relative to the package root.
    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }

    /// The type inferred from the extension.
    #[must_use]
    pub const fn media_type(&self) -> Type {
        self.media_type
    }

    /// The member's byte length.
    #[must_use]
    pub const fn size(&self) -> u64 {
        self.size
    }

    /// The basename.
    #[must_use]
    pub fn file_name(&self) -> &str {
        &self.file_name
    }

    /// Whether the asset is an image.
    #[must_use]
    pub fn is_image(&self) -> bool {
        self.media_type == Type::Image
    }

    /// Whether the asset is a video.
    #[must_use]
    pub fn is_video(&self) -> bool {
        self.media_type == Type::Video
    }

    /// Whether the asset is audio.
    #[must_use]
    pub fn is_audio(&self) -> bool {
        self.media_type == Type::Audio
    }
}

/// Counts and total size of a catalog's assets.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Stats {
    total_count: usize,
    total_size: u64,
    image_count: usize,
    video_count: usize,
    audio_count: usize,
    pdf_count: usize,
    unknown_count: usize,
}

impl Stats {
    /// Count `assets` (sizes saturate).
    pub fn from_assets<'a>(assets: impl IntoIterator<Item = &'a Asset>) -> Self {
        let mut stats = Self::default();
        for asset in assets {
            stats.total_count += 1;
            stats.total_size = stats.total_size.saturating_add(asset.size);
            match asset.media_type {
                Type::Image => stats.image_count += 1,
                Type::Video => stats.video_count += 1,
                Type::Audio => stats.audio_count += 1,
                Type::Pdf => stats.pdf_count += 1,
                Type::Unknown => stats.unknown_count += 1,
            }
        }
        stats
    }

    /// Assets counted.
    #[must_use]
    pub const fn total_count(&self) -> usize {
        self.total_count
    }

    /// Aggregate bytes.
    #[must_use]
    pub const fn total_size(&self) -> u64 {
        self.total_size
    }

    /// Images.
    #[must_use]
    pub const fn image_count(&self) -> usize {
        self.image_count
    }

    /// Videos.
    #[must_use]
    pub const fn video_count(&self) -> usize {
        self.video_count
    }

    /// Audio files.
    #[must_use]
    pub const fn audio_count(&self) -> usize {
        self.audio_count
    }

    /// PDF documents.
    #[must_use]
    pub const fn pdf_count(&self) -> usize {
        self.pdf_count
    }

    /// Unclassified files.
    #[must_use]
    pub const fn unknown_count(&self) -> usize {
        self.unknown_count
    }

    /// The aggregate size with a binary unit, two decimals (`"1.50 KB"`).
    #[must_use]
    pub fn total_size_human(&self) -> String {
        format_bytes(self.total_size)
    }

    /// One-line summary (`"3 files (1.50 KB) - 2 images, …"`).
    #[must_use]
    pub fn summary(&self) -> String {
        format!(
            "{} files ({}) - {} images, {} videos, {} audio, {} PDFs",
            self.total_count,
            self.total_size_human(),
            self.image_count,
            self.video_count,
            self.audio_count,
            self.pdf_count
        )
    }
}

/// Format a byte count with a binary unit and two decimals.
#[must_use]
pub fn format_bytes(bytes: u64) -> String {
    const UNITS: &[&str] = &["B", "KB", "MB", "GB", "TB"];
    #[allow(clippy::cast_precision_loss, reason = "a display approximation")]
    let mut size = bytes as f64;
    let mut unit = 0;
    while size >= 1024.0 && unit < UNITS.len() - 1 {
        size /= 1024.0;
        unit += 1;
    }
    format!("{size:.2} {}", UNITS[unit])
}

/// The materialized assets of one package, ordered by path then basename,
/// with unique basenames.
#[derive(Clone, PartialEq, Eq)]
pub struct Catalog {
    ordered: Box<[Asset]>,
    by_file_name: HashMap<String, usize>,
}

impl fmt::Debug for Catalog {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_list().entries(self.ordered.iter()).finish()
    }
}

impl Catalog {
    /// Discover the `Data/` members among `(name, size)` package members.
    ///
    /// # Errors
    ///
    /// [`Error::AmbiguousFileName`] for a repeated basename, then
    /// [`Error::Limit`] for the asset count, one asset's bytes and the
    /// aggregate bytes, checked per member in that order.
    pub fn from_members<'a>(
        members: impl IntoIterator<Item = (&'a str, u64)>,
        limits: Limits,
    ) -> Result<Self, Error> {
        let mut assets: Vec<Asset> = Vec::new();
        let mut by_file_name: HashMap<String, usize> = HashMap::new();
        let mut total: u64 = 0;
        for (name, size) in members {
            if !name.starts_with(DATA_PREFIX) {
                continue;
            }
            let asset = Asset::new(PathBuf::from(name), size);
            if let Some(&previous) = by_file_name.get(&asset.file_name) {
                return Err(Error::AmbiguousFileName {
                    first: assets[previous].path.display().to_string(),
                    second: asset.path.display().to_string(),
                });
            }
            if assets.len() >= limits.max_assets {
                return Err(Error::Limit {
                    kind: LimitKind::Assets,
                    observed: assets.len() as u64 + 1,
                    maximum: limits.max_assets as u64,
                });
            }
            if asset.size > limits.max_asset_bytes {
                return Err(Error::Limit {
                    kind: LimitKind::AssetBytes,
                    observed: asset.size,
                    maximum: limits.max_asset_bytes,
                });
            }
            let next = total.saturating_add(asset.size);
            if next > limits.max_total_bytes || total.checked_add(asset.size).is_none() {
                return Err(Error::Limit {
                    kind: LimitKind::TotalBytes,
                    observed: next,
                    maximum: limits.max_total_bytes,
                });
            }
            total = next;
            by_file_name.insert(asset.file_name.clone(), assets.len());
            assets.push(asset);
        }
        assets.sort_unstable_by(|left, right| {
            left.path
                .as_os_str()
                .cmp(right.path.as_os_str())
                .then_with(|| left.file_name.cmp(&right.file_name))
        });
        let by_file_name = assets
            .iter()
            .enumerate()
            .map(|(index, asset)| (asset.file_name.clone(), index))
            .collect();
        Ok(Self {
            ordered: assets.into_boxed_slice(),
            by_file_name,
        })
    }

    /// Every asset, in order.
    #[must_use]
    pub fn assets(&self) -> &[Asset] {
        &self.ordered
    }

    /// The asset with basename `file_name`.
    #[must_use]
    pub fn get(&self, file_name: &str) -> Option<&Asset> {
        self.by_file_name
            .get(file_name)
            .and_then(|&index| self.ordered.get(index))
    }

    /// The assets of one type, in order.
    pub fn by_type(&self, media_type: Type) -> impl Iterator<Item = &Asset> + '_ {
        self.ordered
            .iter()
            .filter(move |asset| asset.media_type == media_type)
    }

    /// Number of assets.
    #[must_use]
    pub fn len(&self) -> usize {
        self.ordered.len()
    }

    /// Whether the catalog is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.ordered.is_empty()
    }

    /// Counts and total size.
    #[must_use]
    pub fn stats(&self) -> Stats {
        Stats::from_assets(self.ordered.iter())
    }
}

#[cfg(test)]
mod tests {
    use std::path::PathBuf;

    use super::{Asset, Catalog, Error, LimitKind, Limits, Stats, Type, format_bytes};

    #[test]
    fn assets_classify_by_extension() {
        let asset = Asset::new(PathBuf::from("Data/photo-12.PNG"), 1536);
        assert_eq!(asset.file_name(), "photo-12.PNG");
        assert!(asset.is_image() && !asset.is_video() && !asset.is_audio());
        assert_eq!(
            Asset::new(PathBuf::from("Data/clip.mov"), 1).media_type(),
            Type::Video
        );
        assert_eq!(Asset::new(PathBuf::from(""), 0).file_name(), "unknown");
    }

    #[test]
    fn catalogs_order_filter_and_bound() -> Result<(), Error> {
        let members = [
            ("Index/Document.iwa", 10),
            ("Data/b.m4a", 2048),
            ("Data/a.png", 1024),
            ("preview.jpg", 5),
        ];
        let catalog = Catalog::from_members(members, Limits::default())?;
        let names: Vec<_> = catalog.assets().iter().map(Asset::file_name).collect();
        assert_eq!(names, ["a.png", "b.m4a"]);
        assert_eq!(catalog.get("b.m4a").map(Asset::size), Some(2048));
        assert_eq!(catalog.by_type(Type::Audio).count(), 1);
        let stats = catalog.stats();
        assert_eq!(stats, Stats::from_assets(catalog.assets()));
        assert_eq!(
            stats.summary(),
            "2 files (3.00 KB) - 1 images, 0 videos, 1 audio, 0 PDFs"
        );
        assert_eq!(format_bytes(512), "512.00 B");

        assert!(matches!(
            Catalog::from_members([("Data/x/a.png", 1), ("Data/a.png", 1)], Limits::default()),
            Err(Error::AmbiguousFileName { .. })
        ));
        let one = Limits::new(1, 10, 10)?;
        assert!(matches!(
            Catalog::from_members([("Data/a.png", 1), ("Data/b.png", 1)], one),
            Err(Error::Limit {
                kind: LimitKind::Assets,
                ..
            })
        ));
        assert!(matches!(
            Catalog::from_members([("Data/a.png", 11)], one),
            Err(Error::Limit {
                kind: LimitKind::AssetBytes,
                ..
            })
        ));
        assert_eq!(
            Limits::new(0, 1, 1),
            Err(Error::InvalidLimit {
                kind: LimitKind::Assets
            })
        );
        assert_eq!(
            Limits::new(1, Limits::HARD_MAX_ASSET_BYTES + 1, 1),
            Err(Error::InvalidLimit {
                kind: LimitKind::AssetBytes
            })
        );
        Ok(())
    }
}
