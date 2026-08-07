//! End-to-end conformance checks for the public raw-metafile API.

use std::{
    fs,
    panic::{AssertUnwindSafe, catch_unwind},
    path::{Path, PathBuf},
};

use image::{GenericImageView, ImageFormat};
use litchi_core::error::Error;
use litchi_imgconv::{
    ConvertedFormat, InputFormat, Limits, Options, OutputFormat, convert_metafile,
};
use quick_xml::{Reader, events::Event};

const EMF_HEADER_SIZE: usize = 88;
const WMF_PLACEABLE_SIZE: usize = 22;
const WMF_HEADER_SIZE: usize = 18;
const WMF_EOF: u16 = 0;
const WMF_RECTANGLE: u16 = 0x041B;
const EMR_EOF: u32 = 14;
const EMR_SAVE_DC: u32 = 33;
const EMR_RECTANGLE: u32 = 43;
const EMR_BEGIN_PATH: u32 = 59;
const EMR_MOVE_TO_EX: u32 = 27;
const EMR_STRETCH_DIBITS: u32 = 81;

#[test]
fn emf_vector_converts_to_all_requested_formats() {
    let emf = emf_vector(40, 20);
    let svg = convert(
        &emf,
        InputFormat::Emf,
        OutputFormat::Svg,
        Options::default(),
    );
    assert_eq!(svg.format, ConvertedFormat::Svg);
    assert_eq!(svg.mime_type, "image/svg+xml");
    assert_eq!(svg.extension, "svg");
    let svg_text = utf8(&svg.bytes);
    assert_semantic_svg("synthetic EMF rectangle", svg_text);

    let auto = convert(
        &emf,
        InputFormat::Emf,
        OutputFormat::Auto,
        Options::default(),
    );
    assert_eq!(auto.format, ConvertedFormat::Svg);
    assert_eq!(auto.report.diagnostics.len(), 1);
    assert_eq!(auto.report.diagnostics[0].code, "auto-vector-first");

    let png = convert(
        &emf,
        InputFormat::Emf,
        OutputFormat::Png,
        Options::default().width(80),
    );
    assert_png(&png.bytes, 80, 40);

    let jpeg = convert(
        &emf,
        InputFormat::Emf,
        OutputFormat::Jpeg,
        Options::default().width(80).height(30),
    );
    assert_jpeg(&jpeg.bytes, 80, 30);
}

#[test]
fn placeable_wmf_vector_converts_to_all_requested_formats() {
    let wmf = wmf_vector(30, 20);
    let svg = convert(
        &wmf,
        InputFormat::Wmf,
        OutputFormat::Svg,
        Options::default(),
    );
    assert_eq!(svg.format, ConvertedFormat::Svg);
    let svg_text = utf8(&svg.bytes);
    assert_well_formed_svg(svg_text);
    assert!(svg_text.contains("<rect"));

    let auto = convert(
        &wmf,
        InputFormat::Wmf,
        OutputFormat::Auto,
        Options::default(),
    );
    assert_eq!(auto.format, ConvertedFormat::Svg);
    assert_eq!(auto.report.diagnostics[0].code, "auto-vector-first");

    let png = convert(
        &wmf,
        InputFormat::Wmf,
        OutputFormat::Png,
        Options::default().width(60),
    );
    assert_png(&png.bytes, 60, 40);

    let jpeg = convert(
        &wmf,
        InputFormat::Wmf,
        OutputFormat::Jpeg,
        Options::default().width(60).height(24),
    );
    assert_jpeg(&jpeg.bytes, 60, 24);
}

#[test]
fn auto_selects_png_for_bitmap_only_emf() {
    let emf = emf_file(
        16,
        12,
        &[emr(EMR_STRETCH_DIBITS, &stretch_dibits()), emr_eof()],
    );
    let converted = convert(
        &emf,
        InputFormat::Emf,
        OutputFormat::Auto,
        Options::default(),
    );
    assert_eq!(converted.format, ConvertedFormat::Png);
    assert_eq!(
        converted.report.diagnostics[0].code,
        "auto-raster-only-metafile"
    );
    assert_png(&converted.bytes, 16, 12);
}

#[test]
fn malformed_emf_lengths_and_missing_eof_are_rejected() {
    let mut malformed_header = emf_vector(10, 10);
    malformed_header[48..52].copy_from_slice(&u32::MAX.to_le_bytes());
    assert_err(
        &malformed_header,
        InputFormat::Emf,
        OutputFormat::Svg,
        Options::default(),
    );

    let malformed_record = emf_file(10, 10, &[emr(EMR_RECTANGLE, &rectl(0, 0, 10, 10))]);
    assert_err(
        &malformed_record,
        InputFormat::Emf,
        OutputFormat::Svg,
        Options::default(),
    );

    let mut overlong_record = emf_vector(10, 10);
    overlong_record[EMF_HEADER_SIZE + 4..EMF_HEADER_SIZE + 8]
        .copy_from_slice(&u32::MAX.to_le_bytes());
    assert_err(
        &overlong_record,
        InputFormat::Emf,
        OutputFormat::Svg,
        Options::default(),
    );
}

#[test]
fn malformed_wmf_lengths_missing_eof_and_checksum_are_rejected() {
    let mut bad_checksum = wmf_vector(10, 10);
    bad_checksum[20..22].copy_from_slice(&0_u16.to_le_bytes());
    assert_err(
        &bad_checksum,
        InputFormat::Wmf,
        OutputFormat::Svg,
        Options::default(),
    );

    let wmf_without_eof = wmf_file(
        10,
        10,
        &[wmf_record(WMF_RECTANGLE, &wmf_rect(0, 0, 10, 10))],
    );
    assert_err(
        &wmf_without_eof,
        InputFormat::Wmf,
        OutputFormat::Svg,
        Options::default(),
    );

    let mut overlong_record = wmf_vector(10, 10);
    let record_offset = WMF_PLACEABLE_SIZE + WMF_HEADER_SIZE;
    overlong_record[record_offset..record_offset + 4].copy_from_slice(&u32::MAX.to_le_bytes());
    assert_err(
        &overlong_record,
        InputFormat::Wmf,
        OutputFormat::Svg,
        Options::default(),
    );
}

#[test]
fn target_and_output_limits_are_enforced() {
    let emf = emf_vector(10, 10);
    assert_err(
        &emf,
        InputFormat::Emf,
        OutputFormat::Png,
        Options::default().width(0),
    );
    let target_limited = Limits {
        max_width: 9,
        ..Limits::default()
    };
    assert_err(
        &emf,
        InputFormat::Emf,
        OutputFormat::Png,
        Options::default().limits(target_limited),
    );
    let svg_limited = Limits {
        max_output_bytes: 1,
        ..Limits::default()
    };
    assert_err(
        &emf,
        InputFormat::Emf,
        OutputFormat::Svg,
        Options::default().limits(svg_limited),
    );
    let png_limited = Limits {
        max_output_bytes: 1,
        ..Limits::default()
    };
    assert_err(
        &emf,
        InputFormat::Emf,
        OutputFormat::Png,
        Options::default().limits(png_limited),
    );

    let record_limited = Limits {
        max_records: 2,
        ..Limits::default()
    };
    assert_err(
        &emf,
        InputFormat::Emf,
        OutputFormat::Svg,
        Options::default().limits(record_limited),
    );

    let nested = emf_file(
        10,
        10,
        &[emr(EMR_SAVE_DC, &[]), emr(EMR_SAVE_DC, &[]), emr_eof()],
    );
    let depth_limited = Limits {
        max_state_depth: 1,
        ..Limits::default()
    };
    assert_err(
        &nested,
        InputFormat::Emf,
        OutputFormat::Svg,
        Options::default().limits(depth_limited),
    );

    let mut path_records = vec![emr(EMR_BEGIN_PATH, &[])];
    for coordinate in [1_i32, 2] {
        let mut point = Vec::new();
        point.extend_from_slice(&coordinate.to_le_bytes());
        point.extend_from_slice(&coordinate.to_le_bytes());
        path_records.push(emr(EMR_MOVE_TO_EX, &point));
    }
    path_records.push(emr_eof());
    let path_limited = Limits {
        max_path_points: 1,
        ..Limits::default()
    };
    assert_err(
        &emf_file(10, 10, &path_records),
        InputFormat::Emf,
        OutputFormat::Svg,
        Options::default().limits(path_limited),
    );

    let wmf_record_limited = Limits {
        max_records: 1,
        ..Limits::default()
    };
    assert_err(
        &wmf_vector(10, 10),
        InputFormat::Wmf,
        OutputFormat::Svg,
        Options::default().limits(wmf_record_limited),
    );
}

#[test]
fn bundled_metafile_corpus_never_panics_or_returns_an_empty_svg() {
    let corpus = [
        ("emf/jack-sign.emf", InputFormat::Emf),
        ("emf/vector_image.emf", InputFormat::Emf),
        ("emf/wrench.emf", InputFormat::Emf),
        ("wmf/santa.wmf", InputFormat::Wmf),
    ];
    let mut successful = 0_usize;
    let mut unsupported = 0_usize;
    for (relative, input) in corpus {
        let bytes = corpus_file(relative);
        let result = catch_unwind(AssertUnwindSafe(|| {
            convert_metafile(&bytes, input, OutputFormat::Auto, Options::default())
        }));
        let converted = match result {
            Ok(Ok(converted)) => converted,
            Ok(Err(Error::Unsupported(_))) => {
                unsupported += 1;
                continue;
            },
            Ok(Err(error)) => panic!("{relative} returned an unexpected conversion error: {error}"),
            Err(_) => panic!("{relative} panicked during conversion"),
        };
        successful += 1;
        if relative == "emf/jack-sign.emf" {
            assert!(
                converted
                    .report
                    .diagnostics
                    .iter()
                    .any(|diagnostic| diagnostic.code == "noncanonical-emf-eof-size-last")
            );
        }
        assert!(
            !converted.bytes.is_empty(),
            "{relative} returned no output bytes"
        );
        match converted.format {
            ConvertedFormat::Svg => assert_semantic_svg(relative, utf8(&converted.bytes)),
            ConvertedFormat::Png => assert_raster(&converted.bytes, ImageFormat::Png),
            ConvertedFormat::Jpeg => assert_raster(&converted.bytes, ImageFormat::Jpeg),
        }
    }
    assert_eq!(successful, corpus.len());
    assert_eq!(unsupported, 0);
}

#[test]
#[allow(
    clippy::print_stdout,
    reason = "The corpus category totals are intentional test evidence."
)]
fn all_bundled_metafiles_are_bounded_and_panic_safe() {
    let paths = discover_metafiles();
    let discovered = paths.len();
    assert!(discovered >= 175, "metafile corpus unexpectedly shrank");
    let options = bounded_corpus_options();
    let mut counts = CorpusCounts::default();
    for (path, input) in paths {
        let bytes = match fs::read(&path) {
            Ok(bytes) => bytes,
            Err(error) => panic!("failed to read {}: {error}", path.display()),
        };
        let result = catch_unwind(AssertUnwindSafe(|| {
            convert_metafile(&bytes, input, OutputFormat::Auto, options)
        }));
        match result {
            Ok(Ok(converted)) => {
                assert!(
                    !converted.bytes.is_empty(),
                    "{} returned no output",
                    path.display()
                );
                assert_converted_output(&converted, &path);
                counts.record_success(input);
            },
            Ok(Err(Error::Unsupported(_))) => counts.unsupported += 1,
            Ok(Err(Error::ParseError(_))) => counts.parse_invalid += 1,
            Ok(Err(error)) => panic!("{} returned an unexpected error: {error}", path.display()),
            Err(_) => panic!("{} panicked during conversion", path.display()),
        }
    }
    println!(
        "metafile corpus: success={} unsupported={} parse-invalid={} total={}",
        counts.success,
        counts.unsupported,
        counts.parse_invalid,
        counts.total()
    );
    assert!(
        counts.success > 0,
        "the corpus had no successful conversions"
    );
    assert!(
        counts.emf_success > 0,
        "the corpus had no successful EMF conversion"
    );
    assert!(
        counts.wmf_success > 0,
        "the corpus had no successful WMF conversion"
    );
    assert_eq!(counts.total(), discovered);
}

fn convert(
    data: &[u8],
    input: InputFormat,
    output: OutputFormat,
    options: Options,
) -> litchi_imgconv::ConvertedImage {
    match convert_metafile(data, input, output, options) {
        Ok(converted) => converted,
        Err(error) => panic!("conversion unexpectedly failed: {error}"),
    }
}

fn assert_err(data: &[u8], input: InputFormat, output: OutputFormat, options: Options) {
    assert!(convert_metafile(data, input, output, options).is_err());
}

fn assert_converted_output(converted: &litchi_imgconv::ConvertedImage, path: &Path) {
    match converted.format {
        ConvertedFormat::Svg => assert_well_formed_svg(utf8(&converted.bytes)),
        ConvertedFormat::Png => assert_raster(&converted.bytes, ImageFormat::Png),
        ConvertedFormat::Jpeg => assert_raster(&converted.bytes, ImageFormat::Jpeg),
    }
    assert_eq!(
        converted.report.selected,
        converted.format,
        "{}",
        path.display()
    );
}

fn bounded_corpus_options() -> Options {
    let limits = Limits {
        max_encoded_bytes: 8 * 1024 * 1024,
        max_uncompressed_bytes: 8 * 1024 * 1024,
        max_width: 512,
        max_height: 512,
        max_pixels: 512 * 512,
        max_output_bytes: 1024 * 1024,
        max_records: 20_000,
        max_objects: 2_048,
        max_state_depth: 64,
        max_path_points: 100_000,
        max_svg_elements: 20_000,
    };
    Options::default().limits(limits)
}

#[derive(Default)]
struct CorpusCounts {
    success: usize,
    unsupported: usize,
    parse_invalid: usize,
    emf_success: usize,
    wmf_success: usize,
}

impl CorpusCounts {
    fn record_success(&mut self, input: InputFormat) {
        self.success += 1;
        match input {
            InputFormat::Emf => self.emf_success += 1,
            InputFormat::Wmf => self.wmf_success += 1,
        }
    }

    const fn total(&self) -> usize {
        self.success + self.unsupported + self.parse_invalid
    }
}

fn assert_png(bytes: &[u8], width: u32, height: u32) {
    assert!(bytes.starts_with(b"\x89PNG\r\n\x1a\n"));
    assert_dimensions(bytes, ImageFormat::Png, width, height);
}

fn assert_jpeg(bytes: &[u8], width: u32, height: u32) {
    assert!(bytes.starts_with(&[0xFF, 0xD8]));
    assert!(bytes.ends_with(&[0xFF, 0xD9]));
    assert_dimensions(bytes, ImageFormat::Jpeg, width, height);
}

fn assert_raster(bytes: &[u8], format: ImageFormat) {
    let image = match image::load_from_memory_with_format(bytes, format) {
        Ok(image) => image,
        Err(error) => panic!("encoded image could not be decoded: {error}"),
    };
    assert!(image.width() > 0);
    assert!(image.height() > 0);
}

fn assert_dimensions(bytes: &[u8], format: ImageFormat, width: u32, height: u32) {
    let image = match image::load_from_memory_with_format(bytes, format) {
        Ok(image) => image,
        Err(error) => panic!("encoded image could not be decoded: {error}"),
    };
    assert_eq!(image.dimensions(), (width, height));
}

fn utf8(bytes: &[u8]) -> &str {
    match std::str::from_utf8(bytes) {
        Ok(text) => text,
        Err(error) => panic!("SVG output was not UTF-8: {error}"),
    }
}

fn assert_semantic_svg(name: &str, svg: &str) {
    assert_well_formed_svg(svg);
    assert!(
        svg.contains("<path")
            || svg.contains("<rect")
            || svg.contains("<ellipse")
            || svg.contains("<line")
            || svg.contains("<polygon")
            || svg.contains("<polyline")
            || svg.contains("<text")
            || svg.contains("<image"),
        "{name} produced an SVG document with no semantic drawing elements"
    );
}

fn assert_well_formed_svg(svg: &str) {
    let mut reader = Reader::from_str(svg);
    let mut root_count = 0_usize;
    let mut depth = 0_usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if depth == 0 {
                    root_count += 1;
                    assert_eq!(start.name().as_ref(), b"svg");
                    assert!(start.attributes().flatten().any(|attribute| {
                        attribute.key.as_ref() == b"xmlns"
                            && attribute.value.as_ref() == b"http://www.w3.org/2000/svg"
                    }));
                }
                depth += 1;
            },
            Ok(Event::Empty(element)) if depth == 0 => {
                root_count += 1;
                assert_eq!(element.name().as_ref(), b"svg");
            },
            Ok(Event::Empty(_)) => {},
            Ok(Event::End(_)) => {
                depth = depth
                    .checked_sub(1)
                    .unwrap_or_else(|| panic!("unexpected XML close tag"))
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => panic!("invalid SVG XML: {error}"),
        }
    }
    assert_eq!(depth, 0);
    assert_eq!(root_count, 1);
}

fn emf_vector(width: i32, height: i32) -> Vec<u8> {
    emf_file(
        width,
        height,
        &[
            emr(EMR_RECTANGLE, &rectl(1, 1, width - 1, height - 1)),
            emr_eof(),
        ],
    )
}

fn emf_file(width: i32, height: i32, records: &[Vec<u8>]) -> Vec<u8> {
    let total = EMF_HEADER_SIZE + records.iter().map(Vec::len).sum::<usize>();
    let count =
        u32::try_from(records.len() + 1).unwrap_or_else(|_| panic!("record count overflow"));
    let total = u32::try_from(total).unwrap_or_else(|_| panic!("EMF file size overflow"));
    let mut data = vec![0_u8; EMF_HEADER_SIZE];
    write_u32(&mut data, 0, 1);
    write_u32(&mut data, 4, 88);
    write_i32(&mut data, 16, width);
    write_i32(&mut data, 20, height);
    write_u32(&mut data, 40, 0x464D_4520);
    write_u32(&mut data, 44, 0x0001_0000);
    write_u32(&mut data, 48, total);
    write_u32(&mut data, 52, count);
    write_i32(&mut data, 72, width);
    write_i32(&mut data, 76, height);
    write_i32(&mut data, 80, 1);
    write_i32(&mut data, 84, 1);
    for record in records {
        data.extend_from_slice(record);
    }
    data
}

fn emr(kind: u32, payload: &[u8]) -> Vec<u8> {
    let size = u32::try_from(payload.len() + 8).unwrap_or_else(|_| panic!("EMR size overflow"));
    let mut record = Vec::with_capacity(payload.len() + 8);
    record.extend_from_slice(&kind.to_le_bytes());
    record.extend_from_slice(&size.to_le_bytes());
    record.extend_from_slice(payload);
    record
}

fn emr_eof() -> Vec<u8> {
    let mut payload = vec![0_u8; 12];
    payload[8..12].copy_from_slice(&20_u32.to_le_bytes());
    emr(EMR_EOF, &payload)
}

fn rectl(left: i32, top: i32, right: i32, bottom: i32) -> Vec<u8> {
    let mut rect = Vec::with_capacity(16);
    for coordinate in [left, top, right, bottom] {
        rect.extend_from_slice(&coordinate.to_le_bytes());
    }
    rect
}

fn stretch_dibits() -> Vec<u8> {
    const FIXED_SIZE: usize = 72;
    let mut info = [0_u8; 40];
    write_u32(&mut info, 0, 40);
    write_i32(&mut info, 4, 1);
    write_i32(&mut info, 8, 1);
    info[12..14].copy_from_slice(&1_u16.to_le_bytes());
    info[14..16].copy_from_slice(&24_u16.to_le_bytes());
    write_u32(&mut info, 20, 4);
    let bits = [0_u8, 0, 255, 0];
    let mut payload = vec![0_u8; FIXED_SIZE + info.len() + bits.len()];
    write_i32(&mut payload, 16, 0);
    write_i32(&mut payload, 20, 0);
    write_i32(&mut payload, 24, 0);
    write_i32(&mut payload, 28, 0);
    write_i32(&mut payload, 32, 1);
    write_i32(&mut payload, 36, 1);
    write_u32(&mut payload, 40, as_u32(8 + FIXED_SIZE));
    write_u32(&mut payload, 44, as_u32(info.len()));
    write_u32(&mut payload, 48, as_u32(8 + FIXED_SIZE + info.len()));
    write_u32(&mut payload, 52, as_u32(bits.len()));
    write_u32(&mut payload, 60, 0x00CC_0020);
    write_i32(&mut payload, 64, 16);
    write_i32(&mut payload, 68, 12);
    payload[FIXED_SIZE..FIXED_SIZE + info.len()].copy_from_slice(&info);
    payload[FIXED_SIZE + info.len()..].copy_from_slice(&bits);
    payload
}

fn wmf_vector(width: i16, height: i16) -> Vec<u8> {
    wmf_file(
        width,
        height,
        &[
            wmf_record(WMF_RECTANGLE, &wmf_rect(1, 1, width - 1, height - 1)),
            wmf_record(WMF_EOF, &[]),
        ],
    )
}

fn wmf_file(width: i16, height: i16, records: &[Vec<u8>]) -> Vec<u8> {
    let records_size = records.iter().map(Vec::len).sum::<usize>();
    let metafile_size = WMF_HEADER_SIZE + records_size;
    let file_words =
        u32::try_from(metafile_size / 2).unwrap_or_else(|_| panic!("WMF file size overflow"));
    let max_record = records.iter().map(Vec::len).max().unwrap_or(0);
    let max_record =
        u32::try_from(max_record / 2).unwrap_or_else(|_| panic!("WMF record size overflow"));
    let mut data = placeable_header(0, 0, width, height);
    data.extend_from_slice(&1_u16.to_le_bytes());
    data.extend_from_slice(&9_u16.to_le_bytes());
    data.extend_from_slice(&0x0300_u16.to_le_bytes());
    data.extend_from_slice(&file_words.to_le_bytes());
    data.extend_from_slice(&0_u16.to_le_bytes());
    data.extend_from_slice(&max_record.to_le_bytes());
    data.extend_from_slice(&0_u16.to_le_bytes());
    for record in records {
        data.extend_from_slice(record);
    }
    data
}

fn placeable_header(left: i16, top: i16, right: i16, bottom: i16) -> Vec<u8> {
    let mut data = Vec::with_capacity(WMF_PLACEABLE_SIZE);
    data.extend_from_slice(&0x9AC6_CDD7_u32.to_le_bytes());
    data.extend_from_slice(&0_u16.to_le_bytes());
    for coordinate in [left, top, right, bottom] {
        data.extend_from_slice(&coordinate.to_le_bytes());
    }
    data.extend_from_slice(&1440_u16.to_le_bytes());
    data.extend_from_slice(&0_u32.to_le_bytes());
    let checksum = data.chunks_exact(2).fold(0_u16, |sum, word| {
        sum ^ u16::from_le_bytes([word[0], word[1]])
    });
    data.extend_from_slice(&checksum.to_le_bytes());
    data
}

fn wmf_record(function: u16, params: &[u8]) -> Vec<u8> {
    assert!(params.len().is_multiple_of(2));
    let words =
        u32::try_from((params.len() + 6) / 2).unwrap_or_else(|_| panic!("WMF words overflow"));
    let mut record = Vec::with_capacity(params.len() + 6);
    record.extend_from_slice(&words.to_le_bytes());
    record.extend_from_slice(&function.to_le_bytes());
    record.extend_from_slice(params);
    record
}

fn wmf_rect(left: i16, top: i16, right: i16, bottom: i16) -> Vec<u8> {
    let mut rect = Vec::with_capacity(8);
    for coordinate in [bottom, right, top, left] {
        rect.extend_from_slice(&coordinate.to_le_bytes());
    }
    rect
}

fn write_u32(bytes: &mut [u8], offset: usize, value: u32) {
    bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

fn as_u32(value: usize) -> u32 {
    u32::try_from(value).unwrap_or_else(|_| panic!("test fixture value does not fit u32"))
}

fn write_i32(bytes: &mut [u8], offset: usize, value: i32) {
    bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

fn corpus_file(relative: &str) -> Vec<u8> {
    let mut path = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    path.push("../../test-data/images");
    path.push(relative);
    match fs::read(&path) {
        Ok(bytes) => bytes,
        Err(error) => panic!("failed to read {}: {error}", path.display()),
    }
}

fn discover_metafiles() -> Vec<(PathBuf, InputFormat)> {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let mut paths = Vec::new();
    for relative in ["../../3rdparty", "../../test-data"] {
        collect_metafiles(&root.join(relative), &mut paths);
    }
    paths.sort_by(|left, right| left.0.cmp(&right.0));
    paths
}

fn collect_metafiles(directory: &Path, paths: &mut Vec<(PathBuf, InputFormat)>) {
    let entries = match fs::read_dir(directory) {
        Ok(entries) => entries,
        Err(error) => panic!("failed to read {}: {error}", directory.display()),
    };
    for entry in entries {
        let entry = match entry {
            Ok(entry) => entry,
            Err(error) => panic!("failed to enumerate {}: {error}", directory.display()),
        };
        let file_type = match entry.file_type() {
            Ok(file_type) => file_type,
            Err(error) => panic!("failed to inspect {}: {error}", entry.path().display()),
        };
        let path = entry.path();
        if file_type.is_dir() {
            collect_metafiles(&path, paths);
        } else if file_type.is_file()
            && let Some(input) = metafile_input_format(&path)
        {
            paths.push((path, input));
        }
    }
}

fn metafile_input_format(path: &Path) -> Option<InputFormat> {
    let extension = path.extension()?.to_str()?;
    if extension.eq_ignore_ascii_case("emf") {
        Some(InputFormat::Emf)
    } else if extension.eq_ignore_ascii_case("wmf") {
        Some(InputFormat::Wmf)
    } else {
        None
    }
}
