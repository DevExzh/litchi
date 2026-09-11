#![forbid(unsafe_code)]
use litchi_core::TextOutputOptions;
use litchi_odt::Document;
use sha2::{Digest, Sha256};
use soapberry_zip::office::StreamingArchiveWriter;
use std::{
    hint::black_box,
    io::{self, Write},
    time::Instant,
};
#[derive(Default)]
struct Sink {
    bytes: u64,
}
impl Write for Sink {
    fn write(&mut self, value: &[u8]) -> io::Result<usize> {
        self.bytes += value.len() as u64;
        Ok(value.len())
    }
    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
fn sha(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}
fn main() {
    let samples: usize = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "100".into())
        .parse()
        .unwrap();
    println!("kind,case,index,elapsed_ns,bytes,objects,archive_sha256,output_sha256");
    for (length, blocks) in [(49usize, 10000usize), (1024, 500), (65536, 8)] {
        for pattern in ["dense_crlf", "all_cr", "sparse_crlf"] {
            let mut raw = vec![b'x'; length];
            match pattern {
                "all_cr" => raw.fill(b'\r'),
                "dense_crlf" => {
                    for (i, b) in raw.iter_mut().enumerate() {
                        *b = if i % 2 == 0 { b'\r' } else { b'\n' };
                    }
                },
                _ => {
                    for i in (length.min(128) - 1..length.saturating_sub(1)).step_by(257) {
                        raw[i] = b'\r';
                        raw[i + 1] = b'\n';
                    }
                },
            }
            let text = std::str::from_utf8(&raw).unwrap();
            let normalized = text.replace("\r\n", "\n").replace('\r', "\n");
            let mut expected = String::new();
            let mut xml = String::from(
                "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.3\"><office:body><office:text>",
            );
            for i in 0..blocks {
                xml.push_str("<text:p>");
                xml.push_str(text);
                xml.push_str("</text:p>");
                if i > 0 {
                    expected.push('\n');
                }
                expected.push_str(&normalized);
            }
            xml.push_str("</office:text></office:body></office:document-content>");
            let mut writer = StreamingArchiveWriter::new();
            writer
                .write_stored("mimetype", b"application/vnd.oasis.opendocument.text")
                .unwrap();
            writer.write_stored("content.xml", xml.as_bytes()).unwrap();
            writer.write_stored("META-INF/manifest.xml", b"<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/></manifest:manifest>").unwrap();
            let archive = writer.finish_to_bytes().unwrap();
            let archive_sha = sha(&archive);
            let output_sha = sha(expected.as_bytes());
            let document = Document::from_bytes(archive).unwrap();
            let options =
                TextOutputOptions::new("\n", "\n\n", expected.len() as u64, blocks as u64);
            let mut exact = Vec::new();
            let checked = document.write_text_to(&mut exact, options).unwrap();
            assert_eq!(exact, expected.as_bytes());
            assert_eq!(checked.objects_written(), blocks as u64);
            for index in 0..samples + 10 {
                let mut sink = Sink::default();
                let tick = Instant::now();
                let result = black_box(&document).write_text_to(black_box(&mut sink), options);
                let elapsed = tick.elapsed().as_nanos();
                let report = result.unwrap();
                assert_eq!(report.bytes_written(), expected.len() as u64);
                assert_eq!(report.objects_written(), blocks as u64);
                assert_eq!(sink.bytes, expected.len() as u64);
                if index >= 10 {
                    println!(
                        "sample,{pattern}-{length},{},{elapsed},{},{},{archive_sha},{output_sha}",
                        index - 10,
                        sink.bytes,
                        report.objects_written()
                    );
                }
            }
            exact.clear();
            document.write_text_to(&mut exact, options).unwrap();
            assert_eq!(exact, expected.as_bytes());
        }
    }
}
