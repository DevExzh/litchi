#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Focused unit coverage for mapping and nested-record rewriting.

mod limits;
mod mapping;
mod rewrite;

fn ppt_record(version: u16, kind: u16, data: &[u8]) -> Vec<u8> {
    let mut output = version.to_le_bytes().to_vec();
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(&u32::try_from(data.len()).unwrap().to_le_bytes());
    output.extend_from_slice(data);
    output
}
