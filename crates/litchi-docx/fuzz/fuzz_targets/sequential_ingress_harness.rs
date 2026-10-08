use std::io::{self, Read};
use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
};
use litchi_docx::{OpenOptions, SourceCacheLimits, SourceLimits, SourceReadPolicy, source_backed};

struct Reader<'a> {
    bytes: &'a [u8],
    chunk: usize,
    interrupts: u8,
    overreport: bool,
    cancel: Option<CancellationSource>,
}

impl Read for Reader<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.interrupts != 0 {
            self.interrupts -= 1;
            return Err(io::ErrorKind::Interrupted.into());
        }
        if self.overreport {
            self.overreport = false;
            return Ok(output.len() + 1);
        }
        let length = self.chunk.min(output.len());
        let count = self.bytes.read(&mut output[..length])?;
        if let Some(cancel) = self.cancel.take() {
            cancel.cancel();
        }
        Ok(count)
    }
}

pub fn exercise(data: &[u8]) -> bool {
    let Some((&flags, bytes)) = data.split_first() else {
        return false;
    };
    if bytes.len() > 64 * 1024 {
        return false;
    }
    let budget = Budget::root(
        "fuzz-docx-ingress",
        Limits::new(8 << 20, 8 << 20, 1 << 20, 100_000, 64, 8 << 20),
    );
    let (cancel, token) = CancellationSource::pair();
    let execution = ExecutionContext::new(
        budget.clone(),
        token,
        ExecutionLimits::new(
            NonZeroUsize::MIN,
            NonZeroUsize::MIN,
            NonZeroU64::new(8 << 20).unwrap(),
            0,
        )
        .unwrap(),
    );
    let limits = SourceLimits::builder()
        .max_input_bytes(64 << 10)
        .unwrap()
        .max_members(64)
        .unwrap()
        .max_parts(64)
        .unwrap()
        .max_relationship_parts(64)
        .unwrap()
        .max_entry_bytes(128 << 10)
        .unwrap()
        .max_aggregate_bytes(256 << 10)
        .unwrap()
        .max_part_bytes(128 << 10)
        .unwrap()
        .max_total_part_bytes(256 << 10)
        .unwrap()
        .max_content_types_bytes(16 << 10)
        .unwrap()
        .max_relationship_xml_bytes(16 << 10)
        .unwrap()
        .max_total_relationship_xml_bytes(64 << 10)
        .unwrap()
        .max_events(8192)
        .unwrap()
        .max_depth(32)
        .unwrap()
        .build()
        .unwrap();
    let options = OpenOptions::default()
        .source_limits(limits)
        .cache_limits(SourceCacheLimits::new(128 << 10, 4).unwrap())
        .source_read_policy(if flags & 1 == 0 {
            SourceReadPolicy::exact()
        } else {
            SourceReadPolicy::forward_start(4096).unwrap()
        })
        .execution_context(execution);
    let reader = Reader {
        bytes,
        chunk: 1 + usize::from(flags >> 4) * 64,
        interrupts: if flags & 2 == 0 { 0 } else { 3 },
        overreport: flags & 4 != 0,
        cancel: (flags & 8 != 0).then_some(cancel),
    };
    let opened = match source_backed::Package::from_reader_with_options(reader, options) {
        Ok(package) => {
            let _ = package
                .document()
                .and_then(|document| document.extract_text());
            let clone = package.clone();
            let artifact = package.source_backed_package().unwrap().source_artifact();
            drop(package);
            drop((clone, artifact));
            true
        },
        Err(_) => false,
    };
    assert_eq!(budget.used(Resource::Memory), 0);
    opened
}
