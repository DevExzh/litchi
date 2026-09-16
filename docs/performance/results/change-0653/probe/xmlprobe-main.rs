//! Scratch probe for change 0588 (extends the 0587 survey probe).
//!
//! Modes:
//!   eager  <xlsx> [addr]        public eager open + one cell
//!   source <xlsx> [addr]        source-backed open + one cell
//!   mce    <xml> [reps]         run the MCE codec `reps` times; print in/out sizes
//!   canon  <xml|->              run the codec once and print a namespace-resolved
//!                               canonical form of its output (the differential oracle),
//!                               or the refusal's Debug + Display identity
//!   canonzip <pkg> <out-dir>    (unused; kept out to avoid extra deps)

use std::borrow::Cow;
use std::env;
use std::hint::black_box;
use std::io::{Read, Write};

use litchi_ooxml_common::mce;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;

fn main() {
    let args: Vec<String> = env::args().collect();
    let mode = args.get(1).map(String::as_str).unwrap_or("eager");
    match mode {
        "eager" => {
            let path = args.get(2).expect("path");
            let addr = args.get(3).map(String::as_str).unwrap_or("H680");
            let wb = litchi_xlsx::Workbook::open(path).expect("open");
            let sheet = wb.sheets().next().expect("sheet");
            let view = sheet.cell(addr).expect("cell");
            black_box(view);
            println!("eager ok");
        },
        "source" => {
            let path = args.get(2).expect("path");
            let addr = args.get(3).map(String::as_str).unwrap_or("H680");
            let wb = litchi_xlsx::SourceBackedWorkbook::from_path(path).expect("open");
            let sheet = wb.sheets().next().expect("sheet");
            let view = sheet.cell(addr).expect("cell");
            black_box(view);
            println!("source ok");
        },
        "mce" => {
            let path = args.get(2).expect("path");
            let reps: usize = args.get(3).map_or(1, |r| r.parse().expect("reps"));
            let bytes = std::fs::read(path).expect("read xml");
            let mut last = 0usize;
            let mut borrowed = false;
            for _ in 0..reps {
                let out = mce::process_ooxml(black_box(&bytes)).expect("mce");
                borrowed = matches!(out, Cow::Borrowed(_));
                last = out.len();
                black_box(&out);
                drop(out);
            }
            println!(
                "mce reps={reps} in={} out={last} borrowed={borrowed}",
                bytes.len()
            );
        },
        "raw" => {
            // Byte-identity oracle: the exact processed bytes, their ownership,
            // the report counters, or the refusal's Debug and Display identity.
            let path = args.get(2).expect("path");
            let bytes = if path == "-" {
                let mut v = Vec::new();
                std::io::stdin().read_to_end(&mut v).expect("stdin");
                v
            } else {
                std::fs::read(path).expect("read xml")
            };
            match mce::process_markup_compatibility(
                &bytes,
                &mce::Capabilities::default(),
                &mce::Limits::default(),
            ) {
                Err(error) => println!("ERR\t{error:?}\t{error}"),
                Ok(output) => {
                    let borrowed = matches!(output.xml, Cow::Borrowed(_));
                    let mut hash = 0xcbf2_9ce4_8422_2325u64;
                    for byte in output.xml.as_ref() {
                        hash = (hash ^ u64::from(*byte)).wrapping_mul(0x0000_0100_0000_01b3);
                    }
                    println!(
                        "OK\t{}\t{borrowed}\t{hash:016x}\t{:?}",
                        output.xml.len(),
                        output.report
                    );
                },
            }
        },
        "canon" => {
            let path = args.get(2).expect("path");
            let bytes = if path == "-" {
                let mut v = Vec::new();
                std::io::stdin().read_to_end(&mut v).expect("stdin");
                v
            } else {
                std::fs::read(path).expect("read xml")
            };
            let stdout = std::io::stdout();
            let mut w = std::io::BufWriter::new(stdout.lock());
            match mce::process_markup_compatibility(
                &bytes,
                &mce::Capabilities::default(),
                &mce::Limits::default(),
            ) {
                Err(error) => {
                    writeln!(w, "ERR\t{error:?}\t{error}").expect("write");
                },
                Ok(output) => {
                    writeln!(w, "REPORT\t{:?}", output.report).expect("write");
                    canonicalize(output.xml.as_ref(), &mut w);
                },
            }
            w.flush().expect("flush");
        },
        "time" => {
            // Paired timing leg: warm up, then print one wall-clock nanosecond
            // sample per line for the requested public read.
            let which = args.get(2).map(String::as_str).unwrap_or("eager");
            let path = args.get(3).expect("path");
            let addr = args.get(4).map(String::as_str).unwrap_or("H680");
            let warmup: usize = args.get(5).map_or(3, |v| v.parse().expect("warmup"));
            let samples: usize = args.get(6).map_or(30, |v| v.parse().expect("samples"));
            let once = |which: &str| {
                if which == "source" {
                    let wb = litchi_xlsx::SourceBackedWorkbook::from_path(path).expect("open");
                    let sheet = wb.sheets().next().expect("sheet");
                    let view = sheet.cell(addr).expect("cell");
                    black_box(&view);
                } else {
                    let wb = litchi_xlsx::Workbook::open(path).expect("open");
                    let sheet = wb.sheets().next().expect("sheet");
                    let view = sheet.cell(addr).expect("cell");
                    black_box(&view);
                }
            };
            for _ in 0..warmup {
                once(which);
            }
            let stdout = std::io::stdout();
            let mut w = std::io::BufWriter::new(stdout.lock());
            for _ in 0..samples {
                let start = std::time::Instant::now();
                once(which);
                writeln!(w, "{}", start.elapsed().as_nanos()).expect("write");
            }
            w.flush().expect("flush");
        },
        _ => panic!("mode"),
    }
}

fn esc(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            _ => out.push(c),
        }
    }
    out
}

fn ns_of(r: &ResolveResult<'_>) -> String {
    match r {
        ResolveResult::Unbound => String::new(),
        ResolveResult::Bound(ns) => String::from_utf8_lossy(ns.as_ref()).into_owned(),
        ResolveResult::Unknown(prefix) => {
            format!("!UNKNOWN-PREFIX:{}", String::from_utf8_lossy(prefix))
        },
    }
}

/// Emit a namespace-resolved canonical form: resolved (URI, local) for every
/// element and attribute with unescaped values, plus text, CDATA and comments.
/// Attribute order is normalized; namespace declarations themselves are not
/// emitted (they are exactly what this change alters).
fn canonicalize(xml: &[u8], w: &mut impl Write) {
    let mut reader = quick_xml::NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buf = Vec::new();
    let mut text = String::new();
    let mut depth: usize = 0;
    loop {
        let event = reader.read_event_into(&mut buf);
        match event {
            Err(error) => {
                writeln!(w, "PARSEERR\t{error}").expect("write");
                return;
            },
            Ok(Event::Text(ref e)) => {
                match e.decode() {
                    Ok(decoded) => text.push_str(&decoded),
                    Err(error) => {
                        writeln!(w, "DECODEERR\t{error}").expect("write");
                        return;
                    },
                }
                buf.clear();
                continue;
            },
            Ok(Event::CData(ref e)) => {
                text.push_str(&String::from_utf8_lossy(e.as_ref()));
                buf.clear();
                continue;
            },
            Ok(Event::GeneralRef(ref e)) => {
                match e.resolve_char_ref() {
                    Ok(Some(c)) => text.push(c),
                    _ => {
                        let Ok(n) = e.decode() else {
                            writeln!(w, "DECODEERR\tgeneral-ref").expect("write");
                            return;
                        };
                        match n.as_ref() {
                            "amp" => text.push('&'),
                            "lt" => text.push('<'),
                            "gt" => text.push('>'),
                            "apos" => text.push('\''),
                            "quot" => text.push('"'),
                            other => {
                                writeln!(w, "ENTITY\t{other}").expect("write");
                            },
                        }
                    },
                }
                buf.clear();
                continue;
            },
            _ => {},
        }
        if !text.is_empty() {
            writeln!(w, "T\t{}", esc(&text)).expect("write");
            text.clear();
        }
        let ev = event.expect("event");
        let mut emit_start = |e: &quick_xml::events::BytesStart<'_>,
                              reader: &quick_xml::NsReader<&[u8]>,
                              depth: usize| {
            let (rr, local) = reader.resolver().resolve_element(e.name());
            writeln!(
                w,
                "S\t{}\t{}\t{}",
                depth,
                ns_of(&rr),
                String::from_utf8_lossy(local.as_ref())
            )
            .expect("write");
            let mut attrs: Vec<(String, String, String)> = Vec::new();
            for a in e.attributes().with_checks(true) {
                let Ok(a) = a else {
                    writeln!(w, "ATTRERR").expect("write");
                    return;
                };
                if a.key.as_ref() == b"xmlns" || a.key.as_ref().starts_with(b"xmlns:") {
                    continue;
                }
                let (rr, local) = reader.resolver().resolve_attribute(a.key);
                let Ok(value) = a.decoded_and_normalized_value(
                    quick_xml::XmlVersion::Explicit1_0,
                    reader.decoder(),
                ) else {
                    writeln!(w, "ATTRVALERR").expect("write");
                    return;
                };
                let value = value.into_owned();
                attrs.push((
                    ns_of(&rr),
                    String::from_utf8_lossy(local.as_ref()).into_owned(),
                    value,
                ));
            }
            attrs.sort();
            for (ns, local, value) in attrs {
                writeln!(w, "A\t{ns}\t{local}\t{}", esc(&value)).expect("write");
            }
        };
        match ev {
            Event::Start(ref e) => {
                emit_start(e, &reader, depth);
                depth += 1;
            },
            Event::Empty(ref e) => {
                emit_start(e, &reader, depth);
                writeln!(w, "E\t{depth}").expect("write");
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                writeln!(w, "E\t{depth}").expect("write");
            },
            Event::Comment(ref e) => {
                writeln!(w, "C\t{}", esc(&String::from_utf8_lossy(e.as_ref()))).expect("write");
            },
            Event::Decl(ref e) => {
                writeln!(w, "D\t{}", esc(&String::from_utf8_lossy(e.as_ref()))).expect("write");
            },
            Event::PI(ref e) => {
                writeln!(w, "P\t{}", esc(&String::from_utf8_lossy(e.as_ref()))).expect("write");
            },
            Event::DocType(ref e) => {
                writeln!(w, "DT\t{}", esc(&String::from_utf8_lossy(e.as_ref()))).expect("write");
            },
            Event::Eof => break,
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) => unreachable!(),
        }
        buf.clear();
    }
    writeln!(w, "EOF").expect("write");
}
