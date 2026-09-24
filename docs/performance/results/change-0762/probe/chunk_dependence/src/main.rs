use flate2::{Compress, Compression, FlushCompress, Status};
fn xml_payload(paragraphs: usize) -> Vec<u8> {
    let mut output = Vec::new();
    for index in 0..paragraphs {
        output.extend_from_slice(format!("<w:p><w:r><w:t xml:space=\"preserve\">paragraph {index:05} caf\u{e9} &amp; &lt;tags&gt; {}</w:t></w:r></w:p>", "lorem ipsum ".repeat(index % 5)).as_bytes());
    }
    output
}
fn run(payload: &[u8], level: u32, pieces: &mut dyn FnMut(usize) -> usize) -> Vec<u8> {
    let mut c = Compress::new(Compression::new(level), false);
    let mut buf = vec![0u8; 32768];
    let mut out = Vec::new();
    let mut pos = 0; let mut n = 0;
    while pos < payload.len() {
        let end = (pos + pieces(n).max(1)).min(payload.len()); n += 1;
        let mut input = &payload[pos..end];
        while !input.is_empty() {
            let (bi, bo) = (c.total_in(), c.total_out());
            c.compress(input, &mut buf, FlushCompress::None).unwrap();
            let consumed = (c.total_in() - bi) as usize; let produced = (c.total_out() - bo) as usize;
            out.extend_from_slice(&buf[..produced]); input = &input[consumed..];
        }
        pos = end;
    }
    loop { let bo = c.total_out(); let st = c.compress(&[], &mut buf, FlushCompress::Finish).unwrap(); let produced = (c.total_out()-bo) as usize; out.extend_from_slice(&buf[..produced]); if st == Status::StreamEnd { break } }
    out
}
fn main() {
    let payload = xml_payload(1200);
    for level in [5u32, 6] {
        let whole = run(&payload, level, &mut |_| usize::MAX);
        let chunked = run(&payload, level, &mut |_| 16384);
        let mut s = 7u64;
        let tiny = run(&payload, level, &mut |_| { s ^= s << 13; s ^= s >> 7; s ^= s << 17; 1 + (s % 23) as usize });
        let mut s2 = 9u64;
        let mixed = run(&payload, level, &mut |_| { s2 ^= s2 << 13; s2 ^= s2 >> 7; s2 ^= s2 << 17; 1 + (s2 % 5000) as usize });
        println!("level {level}: whole {} chunked {} tiny {} mixed {} | whole==chunked {} whole==tiny {} whole==mixed {}", whole.len(), chunked.len(), tiny.len(), mixed.len(), whole == chunked, whole == tiny, whole == mixed);
    }
}
