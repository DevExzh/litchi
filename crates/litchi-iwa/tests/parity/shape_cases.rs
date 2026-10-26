//! Shared shape engine (wave W1-E, record 0375): moving a native straight
//! line through the host editors and through the engine —
//! `litchi_iwa_common::shape::line::Segment` and
//! `litchi_iwa_protos::shape_style_codec` — on a line the host added to a
//! native-saved fixture of each format.

use litchi_iwa::shapes::DrawablePoint;
use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_common::shape::line::Segment;
use litchi_iwa_protos::shape_style_codec::{self as codec, Geometry, LINE_FLAGS, Limits};

use crate::chart_style_cases::{Components, publish};
use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};

const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;

#[derive(Clone, Copy)]
enum Target {
    Keynote { slide: usize },
    Numbers { sheet: u64 },
    Pages,
}

const fn point(x: f32, y: f32) -> DrawablePoint {
    DrawablePoint { x, y }
}

fn input(format: Format) -> TestResult<(Fixture, Target, u64)> {
    let (start, end) = (point(100.0, 120.0), point(260.0, 200.0));
    Ok(match format {
        Format::Keynote => {
            let path = "iwork/keynote/chart-data-native.key";
            let mut editor = keynote::Draft::from_bytes(&fixture(path)?.bytes)?;
            let line = editor.add_slide_line(0, start, end)?;
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Keynote { slide: 0 },
                line.drawable_object_id,
            )
        },
        Format::Numbers => {
            let path = "iwork/numbers/chart-data-native.numbers";
            let mut editor = numbers::Draft::from_bytes(&fixture(path)?.bytes)?;
            let sheet = editor.sheets()?.first().ok_or("no sheet")?.id();
            let line = editor.add_sheet_line(sheet, start, end)?;
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Numbers { sheet },
                line.drawable_object_id,
            )
        },
        Format::Pages => {
            let path = "iwork/pages/chart-data-native.pages";
            let mut editor = pages::Draft::from_bytes(&fixture(path)?.bytes)?;
            let line = editor.add_body_line(0, start, end)?;
            (
                Fixture {
                    path,
                    bytes: editor.to_bytes()?,
                },
                Target::Pages,
                line.drawable_object_id,
            )
        },
    })
}

fn host_move(source: &[u8], target: Target, line: u64, segment: Segment) -> TestResult<Vec<u8>> {
    let (start, end) = (segment.start(), segment.end());
    Ok(match target {
        Target::Keynote { slide } => {
            let mut editor = keynote::Draft::from_bytes(source)?;
            editor.set_slide_line_segment(slide, line, start, end)?;
            editor.to_bytes()?
        },
        Target::Numbers { sheet } => {
            let mut editor = numbers::Draft::from_bytes(source)?;
            editor.set_sheet_line_segment(sheet, line, start, end)?;
            editor.to_bytes()?
        },
        Target::Pages => {
            let mut editor = pages::Draft::from_bytes(source)?;
            editor.set_body_line_segment(line, start, end)?;
            editor.to_bytes()?
        },
    })
}

fn host_read(bytes: &[u8], target: Target, line: u64) -> TestResult<String> {
    let segment = match target {
        Target::Keynote { slide } => {
            keynote::Draft::from_bytes(bytes)?.slide_line_segment(slide, line)?
        },
        Target::Numbers { sheet } => {
            numbers::Draft::from_bytes(bytes)?.sheet_line_segment(sheet, line)?
        },
        Target::Pages => pages::Draft::from_bytes(bytes)?.body_line_segment(line)?,
    };
    Ok(format!(
        "{:?} {:?} {:.3}",
        segment.start(),
        segment.end(),
        segment.length()
    ))
}

/// The engine read: the stored line and the segment it encodes.
fn engine_segment(bytes: &[u8], line: u64) -> TestResult<Segment> {
    let components = Components::read(bytes)?;
    let location = components.locate(line, SHAPE_INFO_MESSAGE_TYPE)?;
    let stored = codec::line(components.payload(line, location)?, Limits::DEFAULT)?
        .ok_or("not a straight line")?;
    let [x, y] = stored.geometry.position.ok_or("no position")?;
    let [width, height] = stored.geometry.size.ok_or("no size")?;
    let segment = Segment::from_geometry(
        litchi_iwa_common::shape::geometry::Point { x, y },
        litchi_iwa_common::shape::geometry::Size { width, height },
        stored.geometry.angle.ok_or("no angle")?,
        stored.natural_width,
    )?;
    Ok(if stored.horizontal_flip {
        segment.reversed()
    } else {
        segment
    })
}

fn engine_read(bytes: &[u8], line: u64) -> TestResult<String> {
    let segment = engine_segment(bytes, line)?;
    Ok(format!(
        "{:?} {:?} {:.3}",
        segment.start(),
        segment.end(),
        segment.length()
    ))
}

fn engine_move(source: &[u8], line: u64, segment: Segment) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let location = components.locate(line, SHAPE_INFO_MESSAGE_TYPE)?;
    let payload = components.payload(line, location)?;
    let stored = codec::line(payload, Limits::DEFAULT)?.ok_or("not a straight line")?;
    let directed = if stored.horizontal_flip {
        segment.reversed()
    } else {
        segment
    };
    let (position, size, angle) = directed.geometry();
    let geometry = Geometry {
        position: Some([position.x, position.y]),
        size: Some([size.width, size.height]),
        flags: Some(LINE_FLAGS),
        angle: Some(angle),
    };
    let rewritten = codec::rewrite_line(payload, geometry, segment.length(), Limits::DEFAULT)?;
    let data = rewritten.into_bytes().into_owned();
    publish(
        source,
        components,
        vec![(line, SHAPE_INFO_MESSAGE_TYPE, data)],
    )
}

fn line_case(name: &'static str, format: Format, tier: Tier, evidence: Evidence) -> TestResult {
    let (input, target, line) = input(format)?;
    let segment = Segment::new(point(40.0, 300.0), point(10.0, 90.0))?;
    Case::new(name, format, tier, evidence, &input)
        .host(|source| host_move(source, target, line, segment))
        .leaf(|source| Ok(LeafRun::output(engine_move(source, line, segment)?)))
        .read("host line segment", move |bytes| {
            host_read(bytes, target, line)
        })
        .read("engine line segment", move |bytes| engine_read(bytes, line))
        .run()?;
    Ok(())
}

#[test]
fn keynote_line_segment_parity() -> TestResult {
    line_case(
        "keynote_line_segment_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
    )
}

#[test]
fn numbers_line_segment_parity() -> TestResult {
    line_case(
        "numbers_line_segment_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
    )
}

#[test]
fn pages_line_segment_parity() -> TestResult {
    line_case(
        "pages_line_segment_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
    )
}
