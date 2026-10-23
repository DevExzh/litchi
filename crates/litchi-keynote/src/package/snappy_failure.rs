//! Content-free classification of retained Snappy encoder failures.
//!
//! Keynote writers encode rewritten IWA members with the package's retained
//! `SnappyLimits`. A caller-selected compressed-entry or decoded-stream ceiling
//! is therefore an ordinary typed limit refusal at the owner boundary, never a
//! malformed source. Each owner maps these categories onto its own public
//! `LimitKind`.

/// One Snappy encoder failure, reduced to the owner-visible categories.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SnappyFailure {
    /// A compressed frame or the compressed stream exceeded its ceiling.
    CompressedBytes { observed: u64, maximum: u64 },
    /// A decoded chunk or the decoded stream exceeded its ceiling.
    DecodedBytes { observed: u64, maximum: u64 },
    /// The frame count or another bounded encoder resource exceeded its
    /// ceiling.
    Work { observed: u64, maximum: u64 },
    /// A checked allocation of `requested` bytes failed.
    Allocation { requested: usize },
    /// The encoder rejected its input for a non-resource reason.
    Invalid,
}

impl From<litchi_iwa_core::Error> for SnappyFailure {
    fn from(error: litchi_iwa_core::Error) -> Self {
        use litchi_iwa_core::LimitKind as K;
        match error {
            litchi_iwa_core::Error::Limit {
                kind,
                observed,
                maximum,
            } => {
                let observed = u64::try_from(observed).unwrap_or(u64::MAX);
                let maximum = u64::try_from(maximum).unwrap_or(u64::MAX);
                match kind {
                    K::SnappyCompressedChunkBytes | K::SnappyCompressedStreamBytes => {
                        Self::CompressedBytes { observed, maximum }
                    },
                    K::SnappyChunkBytes | K::SnappyStreamBytes => {
                        Self::DecodedBytes { observed, maximum }
                    },
                    _ => Self::Work { observed, maximum },
                }
            },
            litchi_iwa_core::Error::Allocation { requested, .. } => Self::Allocation { requested },
            _ => Self::Invalid,
        }
    }
}

#[cfg(test)]
mod tests {
    use litchi_iwa_core::{SnappyLimits, SnappyStream};

    use super::SnappyFailure;

    #[test]
    fn retained_encoder_ceilings_stay_typed() {
        let data = vec![7_u8; 4_096];
        let exact = SnappyStream::compress(&data, SnappyLimits::default())
            .expect("default profile")
            .len();
        let tight = SnappyLimits::default()
            .with_compressed_stream_bytes(exact - 1)
            .expect("tight compressed ceiling");
        let error = SnappyStream::compress(&data, tight).expect_err("one byte under");
        assert_eq!(
            SnappyFailure::from(error),
            SnappyFailure::CompressedBytes {
                observed: exact as u64,
                maximum: (exact - 1) as u64,
            }
        );
        let exact_profile = SnappyLimits::default()
            .with_compressed_stream_bytes(exact)
            .expect("exact compressed ceiling");
        assert_eq!(
            SnappyStream::compress(&data, exact_profile)
                .expect("exact ceiling")
                .len(),
            exact
        );

        let decoded = SnappyLimits::new(1_024, 2_048).expect("decoded ceiling");
        let error = SnappyStream::maximum_compressed_len(data.len(), decoded)
            .expect_err("decoded stream over its ceiling");
        assert_eq!(
            SnappyFailure::from(error),
            SnappyFailure::DecodedBytes {
                observed: 4_096,
                maximum: 2_048,
            }
        );
    }
}
