//! Deterministic CFB regular-stream chain guards.
//!
//! This example is a small, public-API-only enabler for measuring the full
//! litchi_cfb::OleFile::open path while it validates a regular stream's FAT
//! chain. It writes one deterministic large stream with litchi_cfb::OleWriter,
//! locates the writer-produced FAT through the CFB header/DIFAT, and changes
//! one FAT entry for each malformed-chain case. The mutation is completed
//! before timing starts; each timed operation constructs an OleFile over Cursor<&[u8]>; result
//! checks and destruction happen after the timestamp.
//!
//! The valid case reopens the generated source and reads the stream before
//! timing. Malformed cases perform one untimed oracle open and retain the
//! exact OleError::CorruptedFile message. Every timed sample checks that same
//! message without formatting or allocating an error string.
//!
//! Run one case with the fixed release recipe used by the performance
//! records:
//!
//!     cargo run --release -p litchi-cfb --features write \
//!       --example perf_chain_guard -- \
//!       --case valid --size 128 --warmup 20 --samples 200 --json
//!
//! --case accepts valid, shortselfcycle, prefixcycle, latercycle, earlyend,
//! invalidmarker, invalidindex, or lateexcess; hyphens and underscores are
//! accepted as spelling separators. --size is the declared regular-stream
//! length in 512-byte sectors. --json PATH also writes the same bytes to PATH;
//! --json without a path prints JSON only.

#[cfg(feature = "write")]
mod benchmark {
    use std::fmt::{self, Write as _};
    use std::io::Cursor;
    use std::path::PathBuf;
    use std::time::Instant;

    use litchi_cfb::{OleError, OleFile, OleWriter};
    use sha2::{Digest, Sha256};

    type AnyResult<T> = Result<T, Box<dyn std::error::Error>>;

    const STREAM_NAME: &str = "GuardStream";
    const SECTOR_SIZE: usize = 512;
    const MAX_CHAIN_SECTORS: usize = 65_536;
    const MIN_CHAIN_SECTORS: usize = 8;
    const MAX_WARMUP: usize = 10_000;
    const MAX_SAMPLES: usize = 10_000;
    const HEADER_DIFAT_OFFSET: usize = 0x4C;
    const HEADER_DIFAT_ENTRIES: usize = 109;
    const NUM_FAT_SECTORS_OFFSET: usize = 0x2C;
    const SECTOR_SHIFT_OFFSET: usize = 0x1E;
    const FIRST_DIFAT_SECTOR_OFFSET: usize = 0x44;
    const NUM_DIFAT_SECTORS_OFFSET: usize = 0x48;
    const ENDOFCHAIN: u32 = 0xFFFF_FFFE;
    const FATSECT: u32 = 0xFFFF_FFFD;
    const MAXREGSECT: u32 = 0xFFFF_FFFA;

    #[derive(Clone, Copy, Debug, Eq, PartialEq)]
    enum GuardCase {
        Valid,
        ShortSelfCycle,
        PrefixCycle,
        LaterCycle,
        EarlyEnd,
        InvalidMarker,
        InvalidIndex,
        LateExcess,
    }

    impl GuardCase {
        fn parse(value: &str) -> AnyResult<Self> {
            let normalized = value
                .chars()
                .filter(|character| *character != '-' && *character != '_')
                .flat_map(char::to_lowercase)
                .collect::<String>();
            match normalized.as_str() {
                "valid" => Ok(Self::Valid),
                "shortselfcycle" => Ok(Self::ShortSelfCycle),
                "prefixcycle" => Ok(Self::PrefixCycle),
                "latercycle" => Ok(Self::LaterCycle),
                "earlyend" => Ok(Self::EarlyEnd),
                "invalidmarker" => Ok(Self::InvalidMarker),
                "invalidindex" => Ok(Self::InvalidIndex),
                "lateexcess" => Ok(Self::LateExcess),
                _ => Err(format!(
                    "unknown CFB chain-guard case '{value}'; expected valid, \
                     shortselfcycle, prefixcycle, latercycle, earlyend, \
                     invalidmarker, invalidindex, or lateexcess"
                )
                .into()),
            }
        }

        const fn name(self) -> &'static str {
            match self {
                Self::Valid => "valid",
                Self::ShortSelfCycle => "shortselfcycle",
                Self::PrefixCycle => "prefixcycle",
                Self::LaterCycle => "latercycle",
                Self::EarlyEnd => "earlyend",
                Self::InvalidMarker => "invalidmarker",
                Self::InvalidIndex => "invalidindex",
                Self::LateExcess => "lateexcess",
            }
        }

        const fn is_valid(self) -> bool {
            matches!(self, Self::Valid)
        }
    }

    #[derive(Debug)]
    struct Arguments {
        case: GuardCase,
        size: usize,
        warmup: usize,
        samples: usize,
        json_path: Option<PathBuf>,
    }

    fn usage() -> &'static str {
        "usage: perf_chain_guard --case <valid|shortselfcycle|prefixcycle|latercycle|earlyend|invalidmarker|invalidindex|lateexcess> --size <sectors> [--warmup N] [--samples N] [--json [PATH]]"
    }

    fn parse_count(value: &str, label: &str, maximum: usize) -> AnyResult<usize> {
        let count = value
            .parse::<usize>()
            .map_err(|source| format!("invalid {label} '{value}': {source}"))?;
        if count > maximum {
            return Err(format!("{label} {count} exceeds maximum {maximum}").into());
        }
        Ok(count)
    }

    fn parse_args() -> AnyResult<Arguments> {
        let mut case = GuardCase::Valid;
        let mut size = None;
        let mut warmup = 20;
        let mut samples = 200;
        let mut json_path = None;
        let mut arguments = std::env::args_os()
            .skip(1)
            .map(|argument| argument.to_string_lossy().into_owned())
            .peekable();

        while let Some(argument) = arguments.next() {
            match argument.as_str() {
                "--case" => {
                    let value = arguments.next().ok_or("missing value for --case")?;
                    case = GuardCase::parse(&value)?;
                },
                "--size" => {
                    let value = arguments.next().ok_or("missing value for --size")?;
                    size = Some(parse_count(&value, "size", MAX_CHAIN_SECTORS)?);
                },
                "--warmup" => {
                    let value = arguments.next().ok_or("missing value for --warmup")?;
                    warmup = parse_count(&value, "warmup", MAX_WARMUP)?;
                },
                "--samples" => {
                    let value = arguments.next().ok_or("missing value for --samples")?;
                    samples = parse_count(&value, "samples", MAX_SAMPLES)?;
                },
                "--json" => {
                    if arguments
                        .peek()
                        .is_some_and(|value| !value.starts_with('-'))
                    {
                        json_path =
                            Some(PathBuf::from(arguments.next().ok_or("missing JSON path")?));
                    }
                },
                "--help" | "-h" => return Err(usage().into()),
                other => return Err(format!("unknown argument '{other}'\n{}", usage()).into()),
            }
        }

        let size = size.ok_or_else(|| format!("missing --size\n{}", usage()))?;
        if size < MIN_CHAIN_SECTORS {
            return Err(format!(
                "size {size} is too small; at least {MIN_CHAIN_SECTORS} sectors are required"
            )
            .into());
        }
        if samples == 0 {
            return Err("samples must be positive".into());
        }
        Ok(Arguments {
            case,
            size,
            warmup,
            samples,
            json_path,
        })
    }

    #[derive(Debug)]
    struct FatLayout {
        sector_size: usize,
        fat_sector_ids: Vec<u32>,
        fat_entry_count: usize,
    }

    #[derive(Debug)]
    struct Fixture {
        source: Vec<u8>,
        source_sha256: String,
        case: GuardCase,
        declared_chain_sectors: usize,
        stream_start_sector: u32,
        stream_bytes: usize,
        fat_entry_count: usize,
        expected_error: Option<String>,
        observed_error: Option<String>,
        valid_stream_content_checked: bool,
        valid_reopen_checked: bool,
    }

    fn sha256_hex(bytes: &[u8]) -> AnyResult<String> {
        let digest = Sha256::digest(bytes);
        let mut output = String::with_capacity(64);
        for byte in digest {
            write!(&mut output, "{byte:02x}")?;
        }
        Ok(output)
    }

    fn read_u16(bytes: &[u8], offset: usize, name: &str) -> AnyResult<u16> {
        let end = offset
            .checked_add(2)
            .ok_or_else(|| format!("{name} offset overflows usize"))?;
        let value = bytes
            .get(offset..end)
            .ok_or_else(|| format!("{name} is truncated"))?;
        Ok(u16::from_le_bytes([value[0], value[1]]))
    }

    fn read_u32(bytes: &[u8], offset: usize, name: &str) -> AnyResult<u32> {
        let end = offset
            .checked_add(4)
            .ok_or_else(|| format!("{name} offset overflows usize"))?;
        let value = bytes
            .get(offset..end)
            .ok_or_else(|| format!("{name} is truncated"))?;
        Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
    }

    fn sector_offset(sector: u32, sector_size: usize) -> AnyResult<usize> {
        let sector_size = u64::try_from(sector_size)?;
        let offset = u64::from(sector)
            .checked_add(1)
            .and_then(|value| value.checked_mul(sector_size))
            .ok_or("CFB sector offset overflows u64")?;
        Ok(usize::try_from(offset)?)
    }

    fn inspect_fat_layout(source: &[u8]) -> AnyResult<FatLayout> {
        let sector_shift = read_u16(source, SECTOR_SHIFT_OFFSET, "sector shift")?;
        let sector_size: usize = match sector_shift {
            9 => 512,
            12 => 4096,
            other => return Err(format!("unsupported generated sector shift {other}").into()),
        };
        let fat_count = usize::try_from(read_u32(
            source,
            NUM_FAT_SECTORS_OFFSET,
            "FAT sector count",
        )?)?;
        if fat_count == 0 {
            return Err("generated CFB contains no FAT sectors".into());
        }

        let header_count = fat_count.min(HEADER_DIFAT_ENTRIES);
        let mut fat_sector_ids = Vec::with_capacity(fat_count);
        for index in 0..header_count {
            let offset = HEADER_DIFAT_OFFSET
                .checked_add(
                    index
                        .checked_mul(4)
                        .ok_or("header DIFAT offset overflows")?,
                )
                .ok_or("header DIFAT offset overflows")?;
            fat_sector_ids.push(read_u32(source, offset, "header FAT sector")?);
        }

        let difat_count = usize::try_from(read_u32(
            source,
            NUM_DIFAT_SECTORS_OFFSET,
            "DIFAT sector count",
        )?)?;
        let mut difat_sector = read_u32(source, FIRST_DIFAT_SECTOR_OFFSET, "DIFAT start sector")?;
        let entries_per_difat = sector_size
            .checked_div(4)
            .and_then(|entries| entries.checked_sub(1))
            .ok_or("DIFAT sector geometry is invalid")?;
        for difat_index in 0..difat_count {
            let offset = sector_offset(difat_sector, sector_size)?;
            for entry_index in 0..entries_per_difat {
                if fat_sector_ids.len() >= fat_count {
                    break;
                }
                let entry_offset = offset
                    .checked_add(entry_index.checked_mul(4).ok_or("DIFAT offset overflows")?)
                    .ok_or("DIFAT offset overflows")?;
                fat_sector_ids.push(read_u32(source, entry_offset, "DIFAT FAT sector")?);
            }
            let continuation_offset = offset
                .checked_add(sector_size)
                .and_then(|value| value.checked_sub(4))
                .ok_or("DIFAT continuation offset overflows")?;
            if difat_index + 1 < difat_count {
                difat_sector = read_u32(source, continuation_offset, "DIFAT continuation")?;
            }
        }
        if fat_sector_ids.len() != fat_count {
            return Err(format!(
                "generated CFB FAT list contains {} sectors, expected {fat_count}",
                fat_sector_ids.len()
            )
            .into());
        }

        let entries_per_sector = sector_size / 4;
        let fat_entry_count = fat_count
            .checked_mul(entries_per_sector)
            .ok_or("FAT entry count overflows usize")?;
        Ok(FatLayout {
            sector_size,
            fat_sector_ids,
            fat_entry_count,
        })
    }

    fn pattern_byte(index: usize) -> u8 {
        match u8::try_from(index % 251) {
            Ok(byte) => byte,
            Err(_) => unreachable!("modulo 251 always fits in u8"),
        }
    }

    fn pattern_matches(bytes: &[u8], expected_len: usize) -> bool {
        bytes.len() == expected_len
            && bytes
                .iter()
                .copied()
                .enumerate()
                .all(|(index, byte)| byte == pattern_byte(index))
    }

    fn build_valid_source(sector_count: usize) -> AnyResult<Vec<u8>> {
        let stream_bytes = sector_count
            .checked_mul(SECTOR_SIZE)
            .ok_or("stream byte count overflows usize")?;
        let mut payload = Vec::with_capacity(stream_bytes);
        for index in 0..stream_bytes {
            payload.push(pattern_byte(index));
        }
        let mut writer = OleWriter::new();
        writer.create_stream_owned(&[STREAM_NAME], payload)?;
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        Ok(output.into_inner())
    }

    fn stream_location(source: &[u8], expected_bytes: usize) -> AnyResult<(u32, bool)> {
        let mut file = OleFile::open(Cursor::new(source))?;
        let (start_sector, size, is_minifat) = {
            let entries = file.list_directory_entries(&[])?;
            let entry = entries
                .into_iter()
                .find(|entry| entry.name == STREAM_NAME)
                .ok_or("generated CFB has no guard stream")?;
            (entry.start_sector, entry.size, entry.is_minifat)
        };
        if is_minifat {
            return Err("generated guard stream unexpectedly uses MiniFAT".into());
        }
        if size != u64::try_from(expected_bytes)? {
            return Err(format!(
                "generated guard stream declares {size} bytes, expected {expected_bytes}"
            )
            .into());
        }
        let stream = file.open_stream(&[STREAM_NAME])?;
        if !pattern_matches(&stream, expected_bytes) {
            return Err("generated guard stream content failed its deterministic oracle".into());
        }
        drop(file);

        // A second public open is intentional: it proves the writer output is
        // still independently reopenable before any malformed mutation.
        let mut reopened = OleFile::open(Cursor::new(source))?;
        let second_stream = reopened.open_stream(&[STREAM_NAME])?;
        if !pattern_matches(&second_stream, expected_bytes) {
            return Err("normal guard-stream reopen failed its content oracle".into());
        }
        Ok((start_sector, true))
    }

    fn mutation(
        case: GuardCase,
        sector_count: usize,
        stream_start: u32,
        fat_entry_count: usize,
    ) -> AnyResult<(u32, u32, Option<String>)> {
        let add = |base: u32, offset: usize| -> AnyResult<u32> {
            Ok(base
                .checked_add(u32::try_from(offset)?)
                .ok_or("stream sector index overflows u32")?)
        };
        let target_offset = match case {
            GuardCase::Valid => return Ok((0, 0, None)),
            GuardCase::ShortSelfCycle => 0,
            GuardCase::PrefixCycle => sector_count / 4,
            GuardCase::LaterCycle => sector_count - 2,
            GuardCase::EarlyEnd | GuardCase::InvalidMarker | GuardCase::InvalidIndex => {
                sector_count / 2
            },
            GuardCase::LateExcess => sector_count - 1,
        };
        let target = add(stream_start, target_offset)?;
        let next = match case {
            GuardCase::ShortSelfCycle => target,
            GuardCase::PrefixCycle => add(stream_start, 1)?,
            GuardCase::LaterCycle => add(stream_start, sector_count / 2)?,
            GuardCase::EarlyEnd => ENDOFCHAIN,
            GuardCase::InvalidMarker => FATSECT,
            GuardCase::InvalidIndex => {
                let value = u32::try_from(fat_entry_count)?;
                if value >= MAXREGSECT {
                    return Err("FAT table is too large for an invalid-index guard".into());
                }
                value
            },
            GuardCase::LateExcess => stream_start,
            GuardCase::Valid => unreachable!("valid case returned above"),
        };
        let expected = match case {
            GuardCase::ShortSelfCycle => {
                format!("Cycle detected in regular stream chain at sector {target}")
            },
            GuardCase::PrefixCycle => format!(
                "Cycle detected in regular stream chain at sector {}",
                add(stream_start, 1)?
            ),
            GuardCase::LaterCycle => format!(
                "Cycle detected in regular stream chain at sector {}",
                add(stream_start, sector_count / 2)?
            ),
            GuardCase::EarlyEnd => {
                "regular stream chain ends before its declared length".to_owned()
            },
            GuardCase::InvalidMarker => {
                "Invalid sector marker 0xFFFFFFFD in regular stream chain".to_owned()
            },
            GuardCase::InvalidIndex => format!("Invalid sector index {next} in regular stream"),
            GuardCase::LateExcess => "regular stream chain exceeds its declared length".to_owned(),
            GuardCase::Valid => unreachable!("valid case returned above"),
        };
        Ok((target, next, Some(expected)))
    }

    fn patch_fat_entry(
        source: &mut [u8],
        layout: &FatLayout,
        sector: u32,
        next: u32,
    ) -> AnyResult<()> {
        let entry = usize::try_from(sector)?;
        if entry >= layout.fat_entry_count {
            return Err(format!("FAT entry {sector} is outside the generated FAT").into());
        }
        let entries_per_sector = layout.sector_size / 4;
        let fat_sector_index = entry / entries_per_sector;
        let within_sector = entry % entries_per_sector;
        let fat_sector = *layout
            .fat_sector_ids
            .get(fat_sector_index)
            .ok_or("FAT sector index is outside the generated FAT list")?;
        let offset = sector_offset(fat_sector, layout.sector_size)?
            .checked_add(
                within_sector
                    .checked_mul(4)
                    .ok_or("FAT entry offset overflows")?,
            )
            .ok_or("FAT entry offset overflows")?;
        let end = offset.checked_add(4).ok_or("FAT entry end overflows")?;
        let slot = source
            .get_mut(offset..end)
            .ok_or("FAT entry is outside the generated source")?;
        slot.copy_from_slice(&next.to_le_bytes());
        Ok(())
    }

    fn corruption_message(error: &OleError) -> Option<&str> {
        match error {
            OleError::CorruptedFile(message) => Some(message.as_str()),
            _ => None,
        }
    }

    fn build_fixture(case: GuardCase, sector_count: usize) -> AnyResult<Fixture> {
        let stream_bytes = sector_count
            .checked_mul(SECTOR_SIZE)
            .ok_or("stream byte count overflows usize")?;
        let mut source = build_valid_source(sector_count)?;
        if !litchi_cfb::is_ole_file(&source) {
            return Err("writer output does not have a CFB signature and minimum size".into());
        }
        let (stream_start_sector, valid_reopen_checked) = stream_location(&source, stream_bytes)?;
        let layout = inspect_fat_layout(&source)?;
        let (target, next, expected_error) = mutation(
            case,
            sector_count,
            stream_start_sector,
            layout.fat_entry_count,
        )?;
        if !case.is_valid() {
            patch_fat_entry(&mut source, &layout, target, next)?;
        }
        let source_sha256 = sha256_hex(&source)?;
        let observed_error = if let Some(expected) = expected_error.as_deref() {
            let error = match OleFile::open(Cursor::new(source.as_slice())) {
                Ok(_) => "malformed guard unexpectedly opened successfully".to_owned(),
                Err(error) => {
                    corruption_message(&error).map_or_else(|| error.to_string(), ToOwned::to_owned)
                },
            };
            if error != expected {
                return Err(format!(
                    "malformed guard oracle mismatch: expected '{expected}', observed '{error}'"
                )
                .into());
            }
            Some(error)
        } else {
            None
        };
        Ok(Fixture {
            source,
            source_sha256,
            case,
            declared_chain_sectors: sector_count,
            stream_start_sector,
            stream_bytes,
            fat_entry_count: layout.fat_entry_count,
            expected_error,
            observed_error,
            valid_stream_content_checked: case.is_valid(),
            valid_reopen_checked,
        })
    }

    fn timed_open(fixture: &Fixture) -> AnyResult<(u128, bool)> {
        let expected = fixture.expected_error.as_deref();
        let reader = Cursor::new(fixture.source.as_slice());
        let started = Instant::now();
        let result = OleFile::open(reader);
        let elapsed_ns = started.elapsed().as_nanos();
        let exact = match result {
            Ok(file) => {
                let observed_size = file.file_size();
                std::hint::black_box(observed_size);
                drop(file);
                expected.is_none()
            },
            Err(error) => {
                let exact = expected.is_some_and(|expected_message| {
                    corruption_message(&error) == Some(expected_message)
                });
                std::hint::black_box(exact);
                drop(error);
                exact
            },
        };
        if !exact {
            return Err("timed CFB chain guard produced a result different from its oracle".into());
        }
        Ok((elapsed_ns, exact))
    }

    fn append_json_string(output: &mut String, value: &str) -> fmt::Result {
        output.push('"');
        for byte in value.bytes() {
            match byte {
                b'"' => output.push_str("\\\""),
                b'\\' => output.push_str("\\\\"),
                b'\n' => output.push_str("\\n"),
                b'\r' => output.push_str("\\r"),
                b'\t' => output.push_str("\\t"),
                0x20..=0x7e => output.push(char::from(byte)),
                _ => write!(output, "\\u00{byte:02x}")?,
            }
        }
        output.push('"');
        Ok(())
    }

    fn append_optional_json_string(output: &mut String, value: Option<&str>) -> fmt::Result {
        match value {
            Some(value) => append_json_string(output, value),
            None => {
                output.push_str("null");
                Ok(())
            },
        }
        Ok(())
    }

    fn render_json(
        arguments: &Arguments,
        fixture: &Fixture,
        samples_ns: &[u128],
        exact_error: bool,
    ) -> AnyResult<String> {
        let mut output = String::new();
        output.push_str("{\n  \"schema\": \"litchi-cfb.perf-chain-guard.v1\",\n");
        write!(&mut output, "  \"case\": ")?;
        append_json_string(&mut output, fixture.case.name())?;
        output.push_str(",\n");
        write!(
            &mut output,
            "  \"declared_chain_sectors\": {},\n  \"sector_size\": {},\n  \"stream_start_sector\": {},\n  \"stream_bytes\": {},\n  \"fat_entry_count\": {},\n",
            fixture.declared_chain_sectors,
            SECTOR_SIZE,
            fixture.stream_start_sector,
            fixture.stream_bytes,
            fixture.fat_entry_count
        )?;
        output.push_str("  \"input_sha256\": ");
        append_json_string(&mut output, &fixture.source_sha256)?;
        output.push_str(",\n  \"expected_error\": ");
        append_optional_json_string(&mut output, fixture.expected_error.as_deref())?;
        output.push_str(",\n  \"observed_error\": ");
        append_optional_json_string(&mut output, fixture.observed_error.as_deref())?;
        output.push_str(",\n");
        write!(
            &mut output,
            "  \"warmup_iterations\": {},\n  \"sample_count\": {},\n  \"timed_operation\": \"OleFile::open(Cursor<&[u8]>)\",\n  \"cleanup_inside_timing\": false,\n  \"input_clone_outside_timing\": true,\n  \"allocation_counters\": {{\"available\": false, \"reason\": \"not instrumented by this public example\"}},\n",
            arguments.warmup,
            samples_ns.len()
        )?;
        output.push_str("  \"oracle\": {\n");
        write!(
            &mut output,
            "    \"expected_error_exact\": {},\n    \"valid_stream_content_checked\": {},\n    \"valid_reopen_checked\": {},\n    \"all_samples_match\": {}\n  }},\n",
            exact_error,
            fixture.valid_stream_content_checked,
            fixture.valid_reopen_checked,
            exact_error
        )?;
        output.push_str("  \"samples_ns\": [");
        for (index, sample) in samples_ns.iter().enumerate() {
            if index != 0 {
                output.push_str(", ");
            }
            write!(&mut output, "{sample}")?;
        }
        output.push_str("]\n}\n");
        Ok(output)
    }

    pub(super) fn run() -> AnyResult<()> {
        let arguments = parse_args()?;
        let fixture = build_fixture(arguments.case, arguments.size)?;
        let total_iterations = arguments
            .warmup
            .checked_add(arguments.samples)
            .ok_or("warmup plus samples overflows usize")?;
        let mut samples_ns = Vec::with_capacity(arguments.samples);
        let mut exact_error = true;
        for iteration in 0..total_iterations {
            let (elapsed_ns, exact) = timed_open(&fixture)?;
            exact_error &= exact;
            if iteration >= arguments.warmup {
                samples_ns.push(elapsed_ns);
            }
        }
        if samples_ns.len() != arguments.samples {
            return Err("timed CFB chain guard retained an unexpected sample count".into());
        }
        let json = render_json(&arguments, &fixture, &samples_ns, exact_error)?;
        if let Some(path) = arguments.json_path {
            std::fs::write(path, &json)?;
        }
        print!("{json}");
        Ok(())
    }
}

#[cfg(feature = "write")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    benchmark::run()
}

#[cfg(not(feature = "write"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    Err("perf_chain_guard requires the write feature; rerun with --features write".into())
}
