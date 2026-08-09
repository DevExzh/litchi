#![allow(
    clippy::pedantic,
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure extraction and exact fixture comparisons"
)]

use litchi_xlsb::calc::{self, Delta, Error, Mode, Opts, Props, Threads};
use litchi_xlsb::raw::{Error as RawError, Stage, Writer};

fn payload(id: u32, mode: u32, iters: u32, delta: f64, threads: i32, opts: u16) -> [u8; calc::LEN] {
    let mut bytes = [0_u8; calc::LEN];
    bytes[0..4].copy_from_slice(&id.to_le_bytes());
    bytes[4..8].copy_from_slice(&mode.to_le_bytes());
    bytes[8..12].copy_from_slice(&iters.to_le_bytes());
    bytes[12..20].copy_from_slice(&delta.to_le_bytes());
    bytes[20..24].copy_from_slice(&threads.to_le_bytes());
    bytes[24..26].copy_from_slice(&opts.to_le_bytes());
    bytes
}

#[test]
fn exact_payload_round_trip_preserves_every_field() {
    let opts = Opts::all();
    let bytes = payload(0xDEAD_BEEF, 2, u32::MAX, -1.25, 1024, opts.bits());

    let props = calc::read(&bytes).unwrap();
    assert_eq!(props.id(), 0xDEAD_BEEF);
    assert_eq!(props.mode(), Mode::AutoNoTables);
    assert_eq!(props.iters(), u32::MAX);
    assert_eq!(props.delta().get().to_bits(), (-1.25_f64).to_bits());
    assert_eq!(props.threads(), Threads::new(1024).unwrap());
    assert_eq!(props.opts(), opts);

    let mut written = Vec::new();
    calc::write(&props, &mut Writer::new(&mut written)).unwrap();
    assert_eq!(written, bytes);
}

#[test]
fn accepts_excel_12_single_byte_options_and_writes_canonical_form() {
    let full = payload(
        0x0001_DD63,
        1,
        100,
        0.001,
        1,
        (Opts::A1 | Opts::FULL_PRECISION | Opts::RECALC_ON_SAVE | Opts::MTR).bits(),
    );
    let legacy = &full[..calc::LEN - 1];

    let props = calc::read(legacy).unwrap();
    assert_eq!(props.id(), 0x0001_DD63);
    assert!(props.has(Opts::A1 | Opts::FULL_PRECISION | Opts::RECALC_ON_SAVE | Opts::MTR));
    assert!(!props.has(Opts::IGNORE_DEPS));

    let mut canonical = Vec::new();
    calc::write(&props, &mut Writer::new(&mut canonical)).unwrap();
    assert_eq!(canonical, full);
}

#[test]
fn rejects_payload_lengths_outside_supported_forms() {
    let full = payload(1, 1, 100, 0.001, 1, Opts::A1.bits());
    assert!(matches!(
        calc::read(&full[..calc::LEN - 2]),
        Err(Error::Wire(RawError::Truncated {
            stage: Stage::Value,
            ..
        }))
    ));

    let mut trailing = full.to_vec();
    trailing.push(0);
    assert!(matches!(
        calc::read(&trailing),
        Err(Error::Wire(RawError::Trailing {
            context: "BrtCalcProp",
            remaining: 1,
            ..
        }))
    ));
}

#[test]
fn rejects_unknown_modes_and_reserved_bits() {
    let unknown_mode = payload(1, 3, 100, 0.001, 1, 0);
    assert!(matches!(
        calc::read(&unknown_mode),
        Err(Error::Mode { value: 3 })
    ));

    let reserved = payload(1, 1, 100, 0.001, 1, 0x0200);
    assert!(matches!(
        calc::read(&reserved),
        Err(Error::ReservedOpts { bits: 0x0200 })
    ));

    let retained = Opts::from_bits_retain(0x8000);
    assert!(matches!(
        Props::new().with_opts(retained),
        Err(Error::ReservedOpts { bits: 0x8000 })
    ));
}

#[test]
fn xnum_delta_follows_ms_xlsb_section_2_5_172() {
    for value in [
        f64::INFINITY,
        f64::NEG_INFINITY,
        f64::NAN,
        f64::from_bits(1),
        f64::from_bits((1_u64 << 63) | 1),
        -0.0,
    ] {
        assert!(matches!(
            Delta::new(value),
            Err(Error::Delta { bits }) if bits == value.to_bits()
        ));
        let bytes = payload(1, 1, 100, value, 1, 0);
        assert!(matches!(
            calc::read(&bytes),
            Err(Error::Delta { bits }) if bits == value.to_bits()
        ));
    }

    for value in [0.0, f64::MIN_POSITIVE, -f64::MIN_POSITIVE, -42.5] {
        assert_eq!(Delta::new(value).unwrap().get().to_bits(), value.to_bits());
    }
}

#[test]
fn thread_count_follows_ms_xlsb_section_2_4_318() {
    for value in [1, 1024] {
        assert_eq!(Threads::new(value).unwrap().get(), value);
        assert_eq!(
            calc::read(&payload(1, 1, 100, 0.001, i32::from(value), 0))
                .unwrap()
                .threads()
                .get(),
            value
        );
    }

    for value in [0_i32, -1, 1025, i32::MAX] {
        assert!(matches!(
            calc::read(&payload(1, 1, 100, 0.001, value, 0)),
            Err(Error::Threads { value: found }) if found == value
        ));
    }
    assert!(matches!(Threads::new(0), Err(Error::Threads { value: 0 })));
    assert!(matches!(
        Threads::new(1025),
        Err(Error::Threads { value: 1025 })
    ));
}

#[test]
fn short_setters_and_builders_keep_state_valid() {
    let mut props = Props::new()
        .with_id(7)
        .with_mode(Mode::Manual)
        .with_iters(25)
        .with_delta(Delta::new(0.25).unwrap())
        .with_threads(Threads::new(8).unwrap())
        .with_opts(Opts::A1 | Opts::ITERATE)
        .unwrap();

    props
        .set_id(8)
        .set_mode(Mode::Auto)
        .set_iters(50)
        .set_delta(Delta::new(0.5).unwrap())
        .set_threads(Threads::new(16).unwrap());
    props.set_opt(Opts::MTR | Opts::USER_THREADS, true).unwrap();

    assert_eq!(props.id(), 8);
    assert_eq!(props.mode(), Mode::Auto);
    assert_eq!(props.iters(), 50);
    assert_eq!(props.delta().get(), 0.5);
    assert_eq!(props.threads().get(), 16);
    assert!(props.has(Opts::ITERATE | Opts::MTR | Opts::USER_THREADS));
}

#[test]
fn values_are_plain_send_sync_data() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Mode>();
    assert_send_sync::<Opts>();
    assert_send_sync::<Delta>();
    assert_send_sync::<Threads>();
    assert_send_sync::<Props>();
}
