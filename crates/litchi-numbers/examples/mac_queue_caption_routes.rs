//! Writes the packages of the macOS queue's Numbers caption rows from their
//! frozen inputs (`docs/unified-ops/0435-caption-native-style.md`).
//!
//! ```text
//! cargo run -p litchi-numbers --example numbers_mac_queue_caption_routes -- <queue-out> <output-directory>
//! ```
//!
//! Most inputs were built by the host record 0429 deleted, with random UUIDs, so
//! they cannot be rebuilt; `<queue-out>/inputs` holds the ones the queue's Mac
//! passes rest on. Each route takes the first drawable of its kind on the first
//! sheet, as the deleted parity cases did. The output is deterministic.

use std::path::Path;

use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::caption::Label;
use litchi_numbers::sheet::drawable::{DrawableSelector, Kind};

type Result<T = ()> = std::result::Result<T, Box<dyn std::error::Error>>;

/// Package, input (a frozen input's name, or a tracked fixture), drawable,
/// label, and the text to set (`None`: remove the label).
type Route = (
    &'static str,
    &'static str,
    Kind,
    Label,
    Option<&'static str>,
);

const ROUTES: &[Route] = &[
    (
        "P1126",
        "fb202c1041968bfe",
        Kind::Movie,
        Label::Title,
        Some("Updated label"),
    ),
    (
        "P1127",
        "03d3cf290c1f72f1",
        Kind::Movie,
        Label::Title,
        Some("Zürich — 2026"),
    ),
    (
        "P1128",
        "55da573ae9e6e6f2",
        Kind::Shape,
        Label::Title,
        Some("Zürich — 2026"),
    ),
    (
        "P1129",
        "14bbef3b02cde0fe",
        Kind::Movie,
        Label::Caption,
        Some("Zürich — 2026"),
    ),
    (
        "P1130",
        "image-adjustments-native.numbers",
        Kind::Image,
        Label::Caption,
        Some("Zürich — 2026"),
    ),
    (
        "P1131",
        "656bfb2acca15fb6",
        Kind::Shape,
        Label::Caption,
        Some("Updated label"),
    ),
    ("P1132", "942418dd055d4747", Kind::Movie, Label::Title, None),
    (
        "P1133",
        "f810284723fbf72a",
        Kind::Chart,
        Label::Caption,
        None,
    ),
    (
        "P1134",
        "8d7d6a87c36f0d9b",
        Kind::Chart,
        Label::Caption,
        Some("Updated label"),
    ),
    (
        "P1135",
        "image-adjustments-native.numbers",
        Kind::Image,
        Label::Title,
        Some("Zürich — 2026"),
    ),
    (
        "P1136",
        "817da2eae0719db6",
        Kind::Image,
        Label::Caption,
        None,
    ),
    (
        "P1137",
        "a21219e6178bb6ee",
        Kind::Image,
        Label::Caption,
        Some("Updated label"),
    ),
    ("P1138", "bec2a7eefe50d8fc", Kind::Image, Label::Title, None),
    (
        "P1139",
        "c77b719ab02302eb",
        Kind::Shape,
        Label::Title,
        Some("Updated label"),
    ),
    (
        "P1140",
        "f094406184c5a649",
        Kind::Movie,
        Label::Caption,
        None,
    ),
    (
        "P1141",
        "08d37cfb9d9e8935",
        Kind::Movie,
        Label::Caption,
        Some("Updated label"),
    ),
    ("P1142", "df049f66fa7899b3", Kind::Shape, Label::Title, None),
    (
        "P1143",
        "d5fceb20f36b29f6",
        Kind::Image,
        Label::Title,
        Some("Updated label"),
    ),
    (
        "P1144",
        "24c2cad74f58a7e4",
        Kind::Shape,
        Label::Caption,
        None,
    ),
    (
        "P1145",
        "e1d6f75bfa687bf3",
        Kind::Shape,
        Label::Caption,
        Some("Zürich — 2026"),
    ),
    (
        "P1146",
        "5da0f02f24e6f0bb",
        Kind::Chart,
        Label::Caption,
        Some("Zürich — 2026"),
    ),
    (
        "P1155",
        "04f8d42895a0eb7c",
        Kind::Chart,
        Label::Caption,
        Some("Updated label"),
    ),
    (
        "W8-native-parent-new-caption",
        "chart-data-native.numbers",
        Kind::Chart,
        Label::Caption,
        Some("Native caption"),
    ),
];

fn open(queue_out: &Path, input: &str) -> Result<Package> {
    let path = if input.ends_with(".numbers") {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers")
            .join(input)
    } else {
        queue_out.join("inputs").join(format!("{input}.numbers"))
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
    for (name, input, kind, label, text) in ROUTES {
        let package = open(queue_out, input)?;
        let mut edit =
            package.edit_sheet_drawable_caption(0, DrawableSelector::of(*kind, 0), *label)?;
        match text {
            Some(text) => drop(edit.set(text)?),
            None => drop(edit.clear()?),
        }
        let mut bytes = Vec::new();
        edit.commit()?.package().write_to(&mut bytes)?;
        std::fs::write(output.join(format!("{name}.numbers")), bytes)?;
    }
    Ok(())
}
