//! Focused coverage for field models, codecs, and package integration.

fn plcf(cps: &[u32], descriptors: &[[u8; 2]]) -> Vec<u8> {
    assert_eq!(cps.len(), descriptors.len() + 1);
    let mut data = Vec::new();
    for cp in cps {
        data.extend_from_slice(&cp.to_le_bytes());
    }
    for descriptor in descriptors {
        data.extend_from_slice(descriptor);
    }
    data
}

mod codec;
mod model;
mod parse;
mod roundtrip;
