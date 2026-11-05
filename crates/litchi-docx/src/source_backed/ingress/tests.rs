use super::*;
use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits};
use std::num::{NonZeroU64, NonZeroUsize};

const OWNER: usize = size_of::<Input>() + size_of::<Vec<u8>>() + 4 * size_of::<usize>();

fn config(
    memory: usize,
    input: u64,
    work: u64,
    maximum: u64,
) -> (Budget, CancellationSource, OpenOptions) {
    let budget = Budget::root(
        "sequential-input",
        Limits::new(memory as u64, input, 1_000, 1_000, 32, work),
    );
    let (cancel, token) = CancellationSource::pair();
    let execution = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory.max(1) as u64).unwrap(),
        0,
    )
    .unwrap();
    let limits = crate::SourceLimits::builder()
        .max_input_bytes(maximum)
        .unwrap()
        .build()
        .unwrap();
    (
        budget.clone(),
        cancel,
        OpenOptions::default()
            .source_limits(limits)
            .execution_context(ExecutionContext::new(budget, token, execution)),
    )
}

struct OneByte<'a>(&'a [u8]);
impl Read for OneByte<'_> {
    fn read(&mut self, out: &mut [u8]) -> io::Result<usize> {
        let length = out.len().min(1);
        self.0.read(&mut out[..length])
    }
}

fn assert_resource(error: Error, resource: Resource) {
    assert!(
        matches!(error, Error::Opc(OpcError::Execution(ExecutionError::ResourceLimit(limit))) if limit.resource == resource)
    );
}

#[test]
fn exact_growth_peak_and_retained_capacity_are_charged_until_last_source_drop() {
    // Three one-byte reads grow 0 -> 1 -> 2 -> 3 (source ceiling), with a
    // two-byte old allocation and three-byte replacement at the peak.
    let (budget, _, options) = config(OWNER + CHUNK + 5, 4, 100, 3);
    let source = from_reader(OneByte(b"abc"), &options).unwrap();
    assert_eq!(budget.used(Resource::Memory), (OWNER + 3) as u64);
    assert_eq!(budget.used(Resource::InputBytes), 3);
    assert_eq!(source.len().unwrap(), 3);
    let mut output = [0; 4];
    assert_eq!(source.read_at(1, &mut output).unwrap(), 2);
    assert_eq!(&output[..2], b"bc");
    assert_eq!(source.read_at(u64::MAX, &mut output).unwrap(), 0);
    let clone = source.clone();
    drop(source);
    assert_eq!(budget.used(Resource::Memory), (OWNER + 3) as u64);
    drop(clone);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _, options) = config(OWNER + CHUNK + 4, 4, 100, 3);
    assert_resource(
        from_reader(OneByte(b"abc"), &options).err().unwrap(),
        Resource::Memory,
    );
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::InputBytes), 3);
}

#[test]
fn source_limit_probe_commits_one_over_and_releases_every_memory_charge() {
    let (budget, _, options) = config(OWNER + CHUNK + 16, 5, 100, 3);
    assert!(matches!(
        from_reader(&b"abcd"[..], &options),
        Err(Error::Opc(OpcError::Limit {
            kind: SourceLimitKind::InputBytes,
            observed: 4,
            maximum: 3
        }))
    ));
    assert_eq!(budget.used(Resource::InputBytes), 4);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn adaptive_input_window_refuses_before_reading_past_available_input() {
    let (budget, _, options) = config(OWNER + CHUNK + 16, 2, 100, 4);
    let mut reader = &b"abc"[..];
    assert_resource(
        from_reader(&mut reader, &options).err().unwrap(),
        Resource::InputBytes,
    );
    assert_eq!(reader, b"c");
    assert_eq!(budget.used(Resource::InputBytes), 2);
    assert_eq!(
        budget.used(Resource::Work),
        3,
        "one actual read plus two copied bytes; refused admission is not a read"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
    // Even EOF requires one byte of preflight headroom; none is consumed if
    // the probe returns zero. Never silently read outside the caller's budget.
    let (budget, _, options) = config(OWNER + CHUNK + 16, 3, 100, 3);
    assert_resource(
        from_reader(&b"abc"[..], &options).err().unwrap(),
        Resource::InputBytes,
    );
    assert_eq!(budget.used(Resource::InputBytes), 3);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn interrupted_retries_are_bounded_by_work_and_release_input_reservations() {
    struct Interrupted(usize);
    impl Read for Interrupted {
        fn read(&mut self, _: &mut [u8]) -> io::Result<usize> {
            self.0 += 1;
            Err(io::ErrorKind::Interrupted.into())
        }
    }
    let (budget, _, options) = config(OWNER + CHUNK + 16, 1, 2, 3);
    let mut reader = Interrupted(0);
    assert_resource(
        from_reader(&mut reader, &options).err().unwrap(),
        Resource::Work,
    );
    assert_eq!(reader.0, 2);
    assert_eq!(budget.used(Resource::InputBytes), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn cancellation_during_a_read_commits_accepted_input_before_releasing_memory() {
    struct Cancel(CancellationSource);
    impl Read for Cancel {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            output[0] = b'x';
            self.0.cancel();
            Ok(1)
        }
    }
    let (budget, cancel, options) = config(OWNER + CHUNK + 16, 4, 100, 3);
    assert!(matches!(
        from_reader(Cancel(cancel), &options),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert_eq!(budget.used(Resource::InputBytes), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn overreporting_is_invalid_data_including_at_the_source_limit_probe() {
    struct Overreport {
        first: bool,
    }
    impl Read for Overreport {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            if self.first {
                self.first = false;
                output.fill(b'x');
                Ok(output.len())
            } else {
                Ok(output.len() + 1)
            }
        }
    }
    for first in [false, true] {
        let (budget, _, options) = config(OWNER + CHUNK + 16, 10, 100, 3);
        assert!(
            matches!(from_reader(Overreport { first }, &options), Err(Error::Io(error)) if error.kind() == io::ErrorKind::InvalidData)
        );
        assert_eq!(budget.used(Resource::InputBytes), if first { 3 } else { 0 });
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}

#[test]
fn failed_scratch_admission_does_not_touch_the_reader() {
    struct NeverRead;
    impl Read for NeverRead {
        fn read(&mut self, _: &mut [u8]) -> io::Result<usize> {
            panic!("read before workspace admission");
        }
    }
    let (budget, _, options) = config(OWNER + CHUNK - 1, 4, 100, 3);
    assert_resource(
        from_reader(NeverRead, &options).err().unwrap(),
        Resource::Memory,
    );
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::InputBytes), 0);
}

#[test]
fn io_failure_after_a_short_read_keeps_cumulative_input_but_no_retained_memory() {
    struct Failing(bool);
    impl Read for Failing {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            if self.0 {
                return Err(io::ErrorKind::BrokenPipe.into());
            }
            self.0 = true;
            output[0] = b'x';
            Ok(1)
        }
    }
    let (budget, _, options) = config(OWNER + CHUNK + 16, 4, 100, 3);
    assert!(
        matches!(from_reader(Failing(false), &options), Err(Error::Io(error)) if error.kind() == io::ErrorKind::BrokenPipe)
    );
    assert_eq!(budget.used(Resource::InputBytes), 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}
