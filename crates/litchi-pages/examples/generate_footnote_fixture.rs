//! Generate the exact synthetic body-footnote fixture used by the lifecycle
//! integration tests.  Keeping the builder included from the test source
//! prevents the qualification workload from silently drifting from the
//! tested graph and metadata sidecars.

#[allow(dead_code)]
mod body_footnote {
    include!("../tests/body_footnote.rs");

    pub(super) fn generated() -> Vec<u8> {
        build_fixture()
            .expect("synthetic body-footnote fixture")
            .bytes
    }
}

fn main() {
    let mut args = std::env::args().skip(1);
    let output = args
        .next()
        .filter(|value| value == "--output")
        .and_then(|_| args.next())
        .map(std::path::PathBuf::from)
        .unwrap_or_else(|| panic!("usage: generate_footnote_fixture --output PATH"));
    assert!(args.next().is_none(), "unexpected generator argument");
    std::fs::write(output, body_footnote::generated())
        .expect("write generated body-footnote fixture");
}
