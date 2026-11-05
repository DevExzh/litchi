//! Writes the packages of the macOS queue's Keynote image-label row from their
//! frozen inputs (`docs/unified-ops/0435-caption-native-style.md`).
//!
//! ```text
//! cargo run -p litchi-keynote --example keynote_mac_queue_caption_routes -- <queue-out> <output-directory>
//! ```
//!
//! Most inputs were built by the host record 0429 deleted, with random UUIDs, so
//! they cannot be rebuilt; `<queue-out>/inputs` holds the ones the queue's Mac
//! passes rest on. Each route takes the first image of the first slide, as the
//! deleted parity cases did. The output is deterministic.

use std::path::Path;

use litchi_keynote::slide::image::text::Text;
use litchi_keynote::{ImageSelector, Package, SlideSelector};

type Result<T = ()> = std::result::Result<T, Box<dyn std::error::Error>>;

/// Package, input (a frozen input's name, or a tracked fixture), label, and the
/// text to set (`None`: remove the label).
type Route = (&'static str, &'static str, Text, Option<&'static str>);

const ROUTES: &[Route] = &[
    (
        "P0290",
        "image-adjustments-native.key",
        Text::Caption,
        Some("Photo: litchi"),
    ),
    ("P0291", "0e41fab47772002a", Text::Caption, None),
    ("P0292", "cc0fa1ed3ac37e9a", Text::Title, None),
    (
        "P0293",
        "cbce7d08ad1807a9",
        Text::Title,
        Some("Abstract art"),
    ),
    (
        "P0294",
        "image-adjustments-native.key",
        Text::Title,
        Some("Abstract"),
    ),
    (
        "P0295",
        "62f66e75dc7df17c",
        Text::Caption,
        Some("Photo: litchi, 2026"),
    ),
];

fn open(queue_out: &Path, input: &str) -> Result<Package> {
    let path = if input.ends_with(".key") {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote")
            .join(input)
    } else {
        queue_out.join("inputs").join(format!("{input}.key"))
    };
    Ok(Package::open(path)?)
}

fn main() -> Result {
    let mut arguments = std::env::args_os().skip(1);
    let (Some(queue_out), Some(output)) = (arguments.next(), arguments.next()) else {
        return Err("usage: mac_queue_caption_routes <queue-out> <output-directory>".into());
    };
    let (queue_out, output) = (Path::new(&queue_out), Path::new(&output));
    std::fs::create_dir_all(output)?;
    for (name, input, text, after) in ROUTES {
        let package = open(queue_out, input)?;
        let edit = package.edit_slide_image_text(
            SlideSelector::index(0),
            ImageSelector::index(0),
            *text,
        )?;
        let edit = match after {
            Some(after) => edit.set(after)?,
            None => edit.clear()?,
        };
        let mut bytes = Vec::new();
        edit.commit()?.package().write_to(&mut bytes)?;
        std::fs::write(output.join(format!("{name}.key")), bytes)?;
    }
    Ok(())
}
