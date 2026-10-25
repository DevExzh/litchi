#![cfg(feature = "rtf")]

use litchi_ops::{Code, File, OpenOptions, open_bytes_with, open_with};
use litchi_rtf::SourceLimits;
use litchi_rtf::leniency::Leniency;

const PLAIN: &[u8] = br"{\rtf1\ansi Policy retained\par Second paragraph}";

fn options(limits: SourceLimits) -> OpenOptions {
    OpenOptions::new().rtf_options(litchi_rtf::OpenOptions::new().source_limits(limits))
}

#[test]
fn source_policy_applies_to_bytes_and_paths() {
    let denied = options(
        SourceLimits::builder()
            .max_input_bytes(PLAIN.len() - 1)
            .unwrap()
            .build()
            .unwrap(),
    );
    assert_eq!(
        open_bytes_with(PLAIN, &denied).unwrap_err().code(),
        Code::LimitExceeded
    );
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("policy.rtf");
    std::fs::write(&path, PLAIN).unwrap();
    assert_eq!(
        open_with(&path, &denied).unwrap_err().code(),
        Code::LimitExceeded
    );
    let admitted = options(
        SourceLimits::builder()
            .max_input_bytes(PLAIN.len())
            .unwrap()
            .build()
            .unwrap(),
    );
    for file in [
        open_bytes_with(PLAIN, &admitted).unwrap(),
        open_with(&path, &admitted).unwrap(),
    ] {
        let File::Rtf(rtf) = file else {
            panic!("expected RTF")
        };
        assert!(rtf.text().unwrap().contains("Policy retained"));
    }
    assert_eq!(
        open_bytes_with(PLAIN, &admitted.max_input_bytes(1))
            .unwrap_err()
            .code(),
        Code::LimitExceeded
    );
}

#[test]
fn compressed_expansion_uses_native_profile() {
    let denied = options(
        SourceLimits::builder()
            .max_decompressed_bytes(PLAIN.len() - 1)
            .unwrap()
            .build()
            .unwrap(),
    );
    for lzfu in [false, true] {
        let bytes = litchi_rtf::transport::compress(
            PLAIN,
            lzfu,
            litchi_rtf::transport::CompressionLimits::default(),
        )
        .unwrap();
        assert_eq!(
            open_bytes_with(bytes, &denied).unwrap_err().code(),
            Code::LimitExceeded
        );
    }
}

#[test]
fn repair_requires_explicit_native_policy() {
    let input = br"{\rtf1\ansi unclosed";
    assert_eq!(
        open_bytes_with(input.as_slice(), &OpenOptions::new())
            .unwrap_err()
            .code(),
        Code::CorruptFile
    );
    let options = OpenOptions::new()
        .rtf_options(litchi_rtf::OpenOptions::new().leniency(Leniency::TolerateUnclosedGroups));
    let File::Rtf(rtf) = open_bytes_with(input.as_slice(), &options).unwrap() else {
        panic!("expected RTF")
    };
    assert!(rtf.text().unwrap().contains("unclosed"));
}

#[test]
fn invalid_native_profile_maps_to_invalid_argument() {
    let native = SourceLimits::builder().max_events(usize::MAX).unwrap_err();
    let error = litchi_ops::Error::from(native);
    assert_eq!(error.code(), Code::InvalidArgument);
}

#[test]
fn zero_opaque_nodes_is_a_denial() {
    let options = options(
        SourceLimits::builder()
            .max_opaque_nodes(0)
            .unwrap()
            .build()
            .unwrap(),
    );
    assert_eq!(
        open_bytes_with(br"{\rtf1\future1 x}".as_slice(), &options)
            .unwrap_err()
            .code(),
        Code::LimitExceeded
    );
}
