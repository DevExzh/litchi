use litchi_odf_common::core::PackageWriter;
use litchi_odg::Drawing;
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::{
    env,
    fs::File,
    hint::black_box,
    io::{BufWriter, Write},
    path::PathBuf,
    time::Instant,
};

const MIMETYPE: &str = "application/vnd.oasis.opendocument.graphics";

#[derive(Clone, Copy, Debug)]
struct CorpusSpec {
    name: &'static str,
    pages: usize,
    shapes_per_page: usize,
    rich: bool,
}

const CORPORA: [CorpusSpec; 4] = [
    CorpusSpec {
        name: "plain-small",
        pages: 4,
        shapes_per_page: 64,
        rich: false,
    },
    CorpusSpec {
        name: "plain-large",
        pages: 32,
        shapes_per_page: 256,
        rich: false,
    },
    CorpusSpec {
        name: "metadata-small",
        pages: 4,
        shapes_per_page: 32,
        rich: true,
    },
    CorpusSpec {
        name: "metadata-large",
        pages: 16,
        shapes_per_page: 128,
        rich: true,
    },
];

#[derive(Debug, Serialize)]
struct Report {
    schema: &'static str,
    probe: &'static str,
    revision: Option<String>,
    corpus: &'static str,
    rich_metadata: bool,
    pages: usize,
    shapes_per_page: usize,
    input_bytes: usize,
    input_sha256: String,
    binary_sha256: String,
    binary_bytes: usize,
    warmups: usize,
    samples: usize,
    elapsed_ns: Vec<u64>,
    statistics: Statistics,
    semantic_checksum: u64,
}

#[derive(Clone, Copy, Debug, Serialize)]
struct Statistics {
    min_ns: u64,
    mean_ns: f64,
    p50_ns: u64,
    p95_ns: u64,
    p99_ns: u64,
    max_ns: u64,
    throughput_input_bytes_per_second_p50: f64,
}

fn main() {
    let args = Args::parse();
    let spec = CORPORA
        .iter()
        .find(|candidate| candidate.name == args.corpus)
        .copied()
        .unwrap_or_else(|| panic!("unknown corpus {:?}", args.corpus));
    assert!(args.samples > 0, "samples must be positive");
    let package = package_for(spec);
    let input_sha256 = sha256(&package);

    let expected_checksum = measure_once(&package);
    for _ in 0..args.warmups {
        assert_eq!(measure_once(&package), expected_checksum);
    }

    let mut elapsed_ns = Vec::with_capacity(args.samples);
    for _ in 0..args.samples {
        let start = Instant::now();
        let checksum = measure_once(&package);
        assert_eq!(checksum, expected_checksum);
        elapsed_ns.push(start.elapsed().as_nanos() as u64);
    }
    let statistics = statistics(&elapsed_ns, package.len());
    let binary = env::current_exe().expect("current executable path");
    let binary_bytes = std::fs::read(&binary).expect("read current executable");
    let report = Report {
        schema: "litchi.odg.open-probe.v1",
        probe: "owned-bytes-open-and-semantic-traversal",
        revision: env::var("LITCHI_GIT_REV").ok(),
        corpus: spec.name,
        rich_metadata: spec.rich,
        pages: spec.pages,
        shapes_per_page: spec.shapes_per_page,
        input_bytes: package.len(),
        input_sha256,
        binary_sha256: sha256(&binary_bytes),
        binary_bytes: binary_bytes.len(),
        warmups: args.warmups,
        samples: args.samples,
        elapsed_ns,
        statistics,
        semantic_checksum: black_box(expected_checksum),
    };
    let output = args
        .output
        .unwrap_or_else(|| PathBuf::from(format!("odg-open-{}.json", spec.name)));
    let file = File::create(&output)
        .unwrap_or_else(|error| panic!("create {}: {error}", output.display()));
    let mut writer = BufWriter::new(file);
    serde_json::to_writer_pretty(&mut writer, &report)
        .unwrap_or_else(|error| panic!("write {}: {error}", output.display()));
    writer
        .flush()
        .unwrap_or_else(|error| panic!("flush {}: {error}", output.display()));
    println!(
        "{} input={}B p50={:.3}ms p95={:.3}ms p99={:.3}ms",
        spec.name,
        package.len(),
        statistics.p50_ns as f64 / 1_000_000.0,
        statistics.p95_ns as f64 / 1_000_000.0,
        statistics.p99_ns as f64 / 1_000_000.0
    );
}

fn measure_once(package: &[u8]) -> u64 {
    let drawing = Drawing::from_bytes(package.to_vec()).expect("generated ODG must open");
    let mut checksum = drawing.pages().len() as u64;
    for page in drawing.pages() {
        checksum = checksum.wrapping_add(page.name().map_or(0, str::len) as u64);
        checksum = checksum.wrapping_add(page.layers().len() as u64);
        checksum = checksum.wrapping_add(page.shapes().len() as u64);
        for shape in page.shapes() {
            checksum = checksum.wrapping_add(shape.name().map_or(0, str::len) as u64);
            checksum = checksum.wrapping_add(shape.text().len() as u64);
            checksum = checksum.wrapping_add(shape.title().map_or(0, str::len) as u64);
        }
    }
    black_box(checksum)
}

fn statistics(samples: &[u64], input_bytes: usize) -> Statistics {
    let mut sorted = samples.to_vec();
    sorted.sort_unstable();
    let sum = samples.iter().map(|value| *value as f64).sum::<f64>();
    let p50 = percentile(&sorted, 0.50);
    Statistics {
        min_ns: sorted[0],
        mean_ns: sum / samples.len() as f64,
        p50_ns: p50,
        p95_ns: percentile(&sorted, 0.95),
        p99_ns: percentile(&sorted, 0.99),
        max_ns: *sorted.last().unwrap(),
        throughput_input_bytes_per_second_p50: input_bytes as f64 * 1_000_000_000.0 / p50 as f64,
    }
}

fn percentile(sorted: &[u64], quantile: f64) -> u64 {
    let position = ((sorted.len() - 1) as f64 * quantile).round() as usize;
    sorted[position]
}

fn sha256(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    digest.iter().map(|byte| format!("{byte:02x}")).collect()
}

fn package_for(spec: CorpusSpec) -> Vec<u8> {
    let content = content_for(spec);
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIMETYPE).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    if spec.rich {
        writer
            .add_file("media/transition.wav", b"deterministic-inert-sound")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn content_for(spec: CorpusSpec) -> String {
    let mut content = String::with_capacity(
        512 + spec.pages * (128 + spec.shapes_per_page * if spec.rich { 780 } else { 250 }),
    );
    if spec.rich {
        content.push_str(r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="automatic" presentation:transition-style="fade-from-left" presentation:transition-speed="fast" smil:type="fade" smil:subtype="crossfade" smil:direction="forward" smil:fadeColor="#010203" presentation:duration="PT2S"><presentation:sound xlink:type="simple" xlink:href="media/transition.wav" xlink:actuate="onRequest" xlink:show="replace" presentation:play-full="true" xml:id="sound1"/></style:drawing-page-properties></style:style></office:automatic-styles><office:body><office:drawing>"##);
    } else {
        content.push_str(r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3"><office:body><office:drawing>"##);
    }
    for page in 0..spec.pages {
        if spec.rich {
            content.push_str(&format!(
                r#"<draw:page draw:name="Page {page}" draw:style-name="dp1">"#
            ));
            content.push_str(r#"<dr3d:scene draw:name="Scene"><dr3d:light dr3d:direction="(0 0 1)"/><dr3d:cube/></dr3d:scene><draw:custom-shape draw:name="Custom"><draw:enhanced-geometry draw:type="rectangle" svg:viewBox="0 0 21600 21600"><draw:equation draw:name="f0" draw:formula="width/2"/><draw:handle draw:handle-position="$0 0"/></draw:enhanced-geometry></draw:custom-shape>"#);
        } else {
            content.push_str(&format!(r#"<draw:page draw:name="Page {page}">"#));
        }
        for shape in 0..spec.shapes_per_page {
            if spec.rich && shape % 8 == 0 {
                content.push_str(&format!(r#"<draw:frame draw:name="Frame-{page}-{shape}" svg:width="2cm" svg:height="1cm"><draw:image xlink:type="simple" xlink:href="media/picture-{page}-{shape}.png"/><draw:glue-point draw:id="0" svg:x="1cm" svg:y="2cm" draw:escape-direction="auto"/><draw:image-map><draw:area-rectangle svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="3cm" xlink:type="simple" xlink:href="https://example.org/target" xlink:show="replace" office:name="area-{page}-{shape}"/><draw:area-polygon svg:x="0cm" svg:y="0cm" svg:width="2cm" svg:height="3cm" svg:viewBox="0 0 100 100" draw:points="0,0 100,0 50,100" draw:nohref="nohref"/></draw:image-map><draw:contour-polygon draw:recreate-on-edit="true" svg:width="2cm" svg:height="3cm" svg:viewBox="0 0 100 100" draw:points="0,0 100,0 50,100"/></draw:frame>"#));
            } else {
                content.push_str(&format!(r#"<draw:rect draw:name="Shape-{page}-{shape}" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><text:p>text-{page}-{shape}</text:p></draw:rect>"#));
            }
        }
        content.push_str("</draw:page>");
    }
    if spec.rich {
        content.push_str("</office:drawing></office:body></office:document-content>");
    } else {
        content.push_str("</office:drawing></office:body></office:document-content>");
    }
    content
}

struct Args {
    corpus: String,
    warmups: usize,
    samples: usize,
    output: Option<PathBuf>,
}

impl Args {
    fn parse() -> Self {
        let mut corpus = None;
        let mut warmups = 10;
        let mut samples = 100;
        let mut output = None;
        let mut args = env::args().skip(1);
        while let Some(arg) = args.next() {
            match arg.as_str() {
                "--corpus" => corpus = args.next(),
                "--warmups" => warmups = args.next().unwrap().parse().unwrap(),
                "--samples" => samples = args.next().unwrap().parse().unwrap(),
                "--output" => output = args.next().map(PathBuf::from),
                "--help" | "-h" => {
                    println!("--corpus NAME --warmups N --samples N --output PATH");
                    std::process::exit(0);
                },
                other => panic!("unknown argument {other:?}"),
            }
        }
        Self {
            corpus: corpus.unwrap_or_else(|| "plain-small".to_owned()),
            warmups,
            samples,
            output,
        }
    }
}
