#![no_main]

#[path = "common/mod.rs"]
mod common;

use libfuzzer_sys::fuzz_target;
use litchi_imgconv::emfplus::objects::DecodeLimits;
use litchi_imgconv::emfplus::playback::PlaybackLimits;
use litchi_imgconv::emfplus::renderer::{EmfPlusSvgRenderer, RendererLimits};
use litchi_imgconv::emfplus::svg::SvgLimits;
use litchi_imgconv::emfplus::ParserLimits;
use litchi_imgconv::{InputFormat, Options, OutputFormat, convert_metafile};

fn renderer_limits() -> RendererLimits {
    RendererLimits {
        parser: ParserLimits {
            max_bytes: 1 << 20,
            max_records: 20_000,
            max_object_slots: 64,
        },
        playback: PlaybackLimits {
            max_points: 65_536,
            max_output: 20_000,
            max_depth: 32,
            max_objects: 64,
            max_bytes: 1 << 20,
        },
        objects: DecodeLimits {
            max_bytes: 1 << 20,
            max_points: 65_536,
            max_recursion: 32,
        },
        svg: SvgLimits {
            max_output_bytes: 1 << 20,
            max_elements: 20_000,
            max_definitions: 1_024,
            max_path_commands: 65_536,
            max_image_bytes: 1 << 20,
            max_diagnostics: 256,
        },
    }
}

fuzz_target!(|data: &[u8]| {
    if let Ok(mut renderer) = EmfPlusSvgRenderer::new(256.0, 256.0, renderer_limits()) {
        let _ = renderer.push_payload(data);
        let _ = renderer.finish();
    }

    let emf = common::wrap_emfplus(data);
    let _ = convert_metafile(&emf, InputFormat::Emf, OutputFormat::Svg, Options::default());
});
