//! Descriptive ZIP transport probe. Build instructions are retained in README.md.
use soapberry_zip::office::{
    ArchiveLimits, IndexedArchive, StreamingArchiveLimits, StreamingArchiveWriter,
};
use soapberry_zip::{CompressionMethod, FileReader};
use std::fs::File;
use std::io::{self, Read, Write};
use std::time::Instant;

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

struct Pattern<'a> {
    block: &'a [u8],
    remaining: u64,
    offset: usize,
}
impl Read for Pattern<'_> {
    fn read(&mut self, out: &mut [u8]) -> io::Result<usize> {
        let mut written = 0;
        let count = self.remaining.min(out.len() as u64) as usize;
        while written < count {
            let n = (count - written).min(self.block.len() - self.offset);
            out[written..written + n].copy_from_slice(&self.block[self.offset..self.offset + n]);
            written += n;
            self.offset = (self.offset + n) % self.block.len();
        }
        self.remaining -= count as u64;
        Ok(count)
    }
}

struct Observe<W> {
    inner: W,
    bytes: u64,
    writes: u64,
    crc: crc32fast::Hasher,
}
impl<W> Observe<W> {
    fn new(inner: W) -> Self {
        Self {
            inner,
            bytes: 0,
            writes: 0,
            crc: crc32fast::Hasher::new(),
        }
    }
}
impl<W: Write> Write for Observe<W> {
    fn write(&mut self, data: &[u8]) -> io::Result<usize> {
        let n = self.inner.write(data)?;
        self.bytes += n as u64;
        self.writes += 1;
        self.crc.update(&data[..n]);
        Ok(n)
    }
    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

fn block(kind: &str) -> Vec<u8> {
    let mut bytes = vec![0; 65536];
    match kind {
        "zeros" => (),
        "mixed" => {
            let mut state = 0x0315_0415_a987_1234u64;
            for b in &mut bytes {
                state ^= state << 13;
                state ^= state >> 7;
                state ^= state << 17;
                *b = state as u8;
            }
        },
        _ => panic!("unknown payload kind"),
    }
    bytes
}

fn emit<W: Write>(sink: W, mode: &str, size: u64, block: &[u8]) -> Result<Observe<W>> {
    let mut limits = StreamingArchiveLimits::default();
    limits.max_entry_size = limits.max_entry_size.max(size);
    limits.max_total_size = limits.max_total_size.max(size);
    let mut archive = StreamingArchiveWriter::with_writer_and_limits(Observe::new(sink), limits);
    let mut source = Pattern {
        block,
        remaining: size,
        offset: 0,
    };
    match mode {
        "borrowed" => archive.write_deflated_stream("payload.bin", source)?,
        "owned" => {
            let mut entry = archive.start_entry("payload.bin", CompressionMethod::Deflate)?;
            let mut buffer = [0u8; 65536];
            loop {
                let n = source.read(&mut buffer)?;
                if n == 0 {
                    break;
                }
                entry.write_all(&buffer[..n])?;
            }
            archive = entry.finish()?;
        },
        _ => return Err("unknown writer mode".into()),
    }
    Ok(archive.finish()?)
}

fn verify(path: &str) -> Result<serde_json::Value> {
    let file = File::open(path)?;
    let len = file.metadata()?.len();
    let limits = ArchiveLimits {
        max_entry_size: 8 << 30,
        max_total_size: 9 << 30,
        ..ArchiveLimits::default()
    };
    let archive = IndexedArchive::from_reader_with_limits(FileReader::from(file), len, limits)?;
    let mut entries = Vec::new();
    for name in archive.file_names() {
        let mut sink = Observe::new(io::sink());
        let metadata = archive.metadata(name)?;
        let decoded = archive.read_to(name, &mut sink)?;
        assert_eq!(decoded, metadata.uncompressed_size());
        assert_eq!(decoded, sink.bytes);
        entries.push(serde_json::json!({"name": name, "uncompressed_bytes": decoded, "compressed_bytes": metadata.compressed_size(), "crc32": sink.crc.finalize()}));
    }
    Ok(
        serde_json::json!({"archive_bytes":len,"archive_zip64":archive.archive_is_zip64(),"has_data_descriptors":archive.has_data_descriptor_entries(),"entries":entries}),
    )
}

fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().collect();
    let output = match args.get(1).map(String::as_str) {
        Some("verify") => verify(&args[2])?,
        Some("write") => {
            let size: u64 = args[3].parse()?;
            let input = block("zeros");
            let file = File::create(&args[4])?;
            let start = Instant::now();
            let output = emit(file, &args[2], size, &input)?;
            let elapsed = start.elapsed().as_nanos();
            serde_json::json!({"mode":args[2],"input_bytes":size,"input_buffer_bytes":65536,"elapsed_ns":elapsed,"output_bytes":output.bytes,"output_writes":output.writes,"output_crc32":output.crc.finalize()})
        }
        Some("guard") => {
            let size: u64 = args[4].parse()?;
            let samples: usize = args[5].parse()?;
            let warmup: usize = args[6].parse()?;
            let input = block(&args[3]);
            let mut times = Vec::with_capacity(samples);
            let mut oracle = None;
            for index in 0..samples+warmup {
                let start = Instant::now();
                let output = emit(io::sink(), &args[2], size, &input)?;
                let elapsed = start.elapsed().as_nanos() as u64;
                let current = (output.bytes,output.writes,output.crc.finalize());
                if let Some(expected) = oracle { assert_eq!(expected,current); } else { oracle=Some(current); }
                if index >= warmup { times.push(elapsed); }
            }
            let (bytes,writes,crc)=oracle.ok_or("at least one operation is required")?;
            serde_json::json!({"mode":args[2],"payload":args[3],"input_bytes":size,"warmup":warmup,"samples_ns":times,"output_bytes":bytes,"output_writes":writes,"output_crc32":crc})
        }
        _ => return Err("usage: probe verify PATH | write borrowed|owned SIZE PATH | guard borrowed|owned zeros|mixed SIZE SAMPLES WARMUP".into()),
    };
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}
