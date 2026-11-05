use std::{fmt::Debug, fmt::Display, hint::black_box, sync::OnceLock};

use litchi_keynote::{OpenOptions, Package, SemanticLimits, SourceLimits};

pub const MAX_INPUT_BYTES: u64 = 1024 * 1024;
pub const MAX_PACKAGE_BYTES: u64 = 8 * 1024 * 1024;
pub const MAX_ENTRIES: usize = 256;
pub const MAX_ENTRY_BYTES: u64 = 2 * 1024 * 1024;
pub const MAX_EXPANDED_BYTES: u64 = 8 * 1024 * 1024;
pub const MAX_IWA_STREAM_BYTES: usize = 2 * 1024 * 1024;
pub const MAX_OBJECTS: usize = 16 * 1024;
pub const MAX_SLIDES: usize = 512;
pub const MAX_REFERENCES: usize = 32 * 1024;
pub const MAX_TEXT_STORAGES: usize = 8 * 1024;
pub const MAX_TEXT_FRAGMENTS: usize = 32 * 1024;
pub const MAX_TEXT_BYTES: usize = 2 * 1024 * 1024;

pub fn read_options() -> OpenOptions {
    static OPTIONS: OnceLock<OpenOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = SourceLimits::builder()
            .max_input_bytes(MAX_PACKAGE_BYTES)
            .and_then(|limits| limits.max_entries(MAX_ENTRIES))
            .and_then(|limits| limits.max_entry_bytes(MAX_ENTRY_BYTES))
            .and_then(|limits| limits.max_aggregate_bytes(MAX_EXPANDED_BYTES))
            .and_then(|limits| limits.max_component_bytes(MAX_IWA_STREAM_BYTES))
            .and_then(|limits| limits.build())
            .unwrap_or_else(|error| unreachable!("valid fuzz archive limits: {error}"));
        let semantic = SemanticLimits::builder()
            .max_objects(MAX_OBJECTS)
            .and_then(|limits| limits.max_slides(MAX_SLIDES))
            .and_then(|limits| limits.max_references(MAX_REFERENCES))
            .and_then(|limits| limits.max_text_storages(MAX_TEXT_STORAGES))
            .and_then(|limits| limits.max_text_fragments(MAX_TEXT_FRAGMENTS))
            .and_then(|limits| limits.max_text_bytes(MAX_TEXT_BYTES))
            .and_then(|limits| limits.build())
            .unwrap_or_else(|error| unreachable!("valid fuzz semantic limits: {error}"));
        OpenOptions::new()
            .source_limits(archive)
            .semantic_limits(semantic)
    })
}

pub fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing Keynote fuzz package must succeed: {error}"));
    bytes
}

pub fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}

pub fn control(data: &[u8], offset: usize) -> u8 {
    data.get(offset).copied().unwrap_or_default()
}

pub fn finite_axis(data: &[u8], offset: usize, fallback: f32) -> f32 {
    let bits = u32::from(control(data, offset))
        | (u32::from(control(data, offset.saturating_add(1))) << 8)
        | (u32::from(control(data, offset.saturating_add(2))) << 16)
        | (u32::from(control(data, offset.saturating_add(3))) << 24);
    let value = f32::from_bits(bits);
    if value.is_finite() {
        let value = value.clamp(-1_000_000.0, 1_000_000.0);
        if value == 0.0 { 0.0 } else { value }
    } else {
        fallback
    }
}
