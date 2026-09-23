//! Probe for change 0752: cost of per-write CRC32 and raw-Deflate calls on the
//! exact DOCX streaming write pattern, and chunking independence of zlib-rs.
use flate2::{Compress, Compression, FlushCompress, Status};
use std::time::Instant;

const DOCUMENT_PREFIX: &[u8] = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>"#,
)
.as_bytes();
const DOCUMENT_SUFFIX: &[u8] = b"<w:sectPr/></w:body></w:document>";

fn escape(text: &str) -> Vec<u8> {
    let mut out = Vec::new();
    for c in text.chars() {
        match c {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '>' => out.extend_from_slice(b"&gt;"),
            _ => {
                let mut b = [0u8; 4];
                out.extend_from_slice(c.encode_utf8(&mut b).as_bytes());
            }
        }
    }
    out
}

/// Returns the concatenated payload and the write boundaries.
fn docx_pattern(paragraphs: usize) -> (Vec<u8>, Vec<(usize, usize)>) {
    let mut data = Vec::new();
    let mut writes = Vec::new();
    let mut push = |data: &mut Vec<u8>, bytes: &[u8]| {
        let start = data.len();
        data.extend_from_slice(bytes);
        writes.push((start, data.len()));
    };
    push(&mut data, DOCUMENT_PREFIX);
    for index in 0..paragraphs {
        push(&mut data, b"<w:p>");
        push(&mut data, b"<w:r><w:t xml:space=\"preserve\">");
        let text = format!("litchi-perf-docx-streaming-v1-{index:06}-caf\u{e9}-<&>");
        let escaped = escape(&text);
        for chunk in escaped.chunks(64) {
            push(&mut data, chunk);
        }
        push(&mut data, b"</w:t></w:r>");
        push(&mut data, b"</w:p>");
    }
    push(&mut data, DOCUMENT_SUFFIX);
    (data, writes)
}

fn crc_chunked(data: &[u8], writes: &[(usize, usize)]) -> u32 {
    let mut crc = 0u32;
    for &(s, e) in writes {
        let mut h = crc32fast::Hasher::new_with_initial(crc);
        h.update(&data[s..e]);
        crc = h.finalize();
    }
    crc
}

fn crc_staged(data: &[u8], writes: &[(usize, usize)], stage_cap: usize) -> u32 {
    let mut crc = 0u32;
    let mut stage: Vec<u8> = Vec::with_capacity(stage_cap);
    for &(s, e) in writes {
        let b = &data[s..e];
        if stage.len() + b.len() > stage_cap {
            let mut h = crc32fast::Hasher::new_with_initial(crc);
            h.update(&stage);
            crc = h.finalize();
            stage.clear();
        }
        if b.len() >= stage_cap {
            let mut h = crc32fast::Hasher::new_with_initial(crc);
            h.update(b);
            crc = h.finalize();
        } else {
            stage.extend_from_slice(b);
        }
    }
    let mut h = crc32fast::Hasher::new_with_initial(crc);
    h.update(&stage);
    h.finalize()
}

/// Mirrors ReusableDeflateState::write_to's one-codec-call-per-write loop
/// (the drain is replaced by appending pending output to `out`).
fn deflate_writes(data: &[u8], writes: &[(usize, usize)]) -> Vec<u8> {
    let mut c = Compress::new(Compression::default(), false);
    let mut buf = vec![0u8; 32 * 1024];
    let mut out = Vec::new();
    for &(s, e) in writes {
        let mut input = &data[s..e];
        while !input.is_empty() {
            let before_in = c.total_in();
            let before_out = c.total_out();
            let _ = c.compress(input, &mut buf, FlushCompress::None).unwrap();
            let consumed = (c.total_in() - before_in) as usize;
            let produced = (c.total_out() - before_out) as usize;
            out.extend_from_slice(&buf[..produced]);
            input = &input[consumed..];
        }
    }
    // `ZipDataWriter::finish` flushes (a sync flush, drained with empty
    // no-flush calls, as `ReusableDeflateState::flush_to` does) before
    // `finish_to` ends the stream.
    let before_out = c.total_out();
    let _ = c.compress(&[], &mut buf, FlushCompress::Sync).unwrap();
    let produced = (c.total_out() - before_out) as usize;
    out.extend_from_slice(&buf[..produced]);
    loop {
        let before_out = c.total_out();
        let _ = c.compress(&[], &mut buf, FlushCompress::None).unwrap();
        let produced = (c.total_out() - before_out) as usize;
        out.extend_from_slice(&buf[..produced]);
        if produced == 0 {
            break;
        }
    }
    loop {
        let before_out = c.total_out();
        let status = c.compress(&[], &mut buf, FlushCompress::Finish).unwrap();
        let produced = (c.total_out() - before_out) as usize;
        out.extend_from_slice(&buf[..produced]);
        if status == Status::StreamEnd {
            break;
        }
    }
    out
}

fn inflate(stream: &[u8]) -> Vec<u8> {
    use std::io::Read;
    let mut out = Vec::new();
    flate2::read::DeflateDecoder::new(stream).read_to_end(&mut out).unwrap();
    out
}

fn chunk_writes(total: usize, size: usize) -> Vec<(usize, usize)> {
    (0..total).step_by(size).map(|s| (s, (s + size).min(total))).collect()
}

fn time<F: FnMut() -> u64>(label: &str, reps: usize, mut f: F) {
    let mut samples = Vec::new();
    let mut sink = 0u64;
    for _ in 0..reps {
        let t = Instant::now();
        sink = sink.wrapping_add(f());
        samples.push(t.elapsed().as_nanos() as f64 / 1e6);
    }
    samples.sort_by(|a, b| a.partial_cmp(b).unwrap());
    println!("{label:48} p50 {:8.3} ms  min {:8.3} ms  (sink {sink})", samples[reps / 2], samples[0]);
}

fn main() {
    let paragraphs: usize = std::env::args().nth(1).map(|v| v.parse().unwrap()).unwrap_or(131_072);
    let (data, writes) = docx_pattern(paragraphs);
    println!("payload {} bytes, {} writes, mean write {:.1} bytes", data.len(), writes.len(), data.len() as f64 / writes.len() as f64);
    let reference = crc32fast::hash(&data);
    assert_eq!(crc_chunked(&data, &writes), reference);
    for cap in [256usize, 1024, 4096, 16384] {
        assert_eq!(crc_staged(&data, &writes, cap), reference);
    }
    // Chunking independence of the compressed stream.
    let base = deflate_writes(&data, &writes);
    assert!(inflate(&base) == data, "the writer's split must inflate to the payload");
    println!("deflate writer split: {} bytes, inflates to the payload: true", base.len());
    if let Some(path) = std::env::var_os("MODEL_STREAM_OUT") {
        std::fs::write(path, &base).unwrap();
    }
    for size in [1usize, 7, 64, 4096, 16 * 1024, 64 * 1024, data.len()] {
        let other = deflate_writes(&data, &chunk_writes(data.len(), size));
        let inflates = inflate(&other) == data;
        println!("deflate chunk {size:>9}: {} bytes, identical to write pattern: {}, inflates to the payload: {inflates}", other.len(), other == base);
    }
    let reps = 15;
    time("crc per write (current)", reps, || crc_chunked(&data, &writes) as u64);
    for cap in [256usize, 1024, 4096, 16384] {
        time(&format!("crc staged {cap}"), reps, || crc_staged(&data, &writes, cap) as u64);
    }
    time("crc one buffer", reps, || crc32fast::hash(&data) as u64);
    time("deflate per write (current)", reps, || deflate_writes(&data, &writes).len() as u64);
    for size in [4096usize, 16 * 1024, 64 * 1024] {
        let w = chunk_writes(data.len(), size);
        time(&format!("deflate chunks {size}"), reps, || deflate_writes(&data, &w).len() as u64);
    }
}
