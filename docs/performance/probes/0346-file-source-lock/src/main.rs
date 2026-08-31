use std::fs::{File, Metadata};
use std::hint::black_box;
use std::os::unix::fs::MetadataExt;
use std::sync::Mutex as StdMutex;
use std::time::Instant;

const BLOCKS: usize = 40;
const ITERATIONS: usize = 200_000;

#[derive(Clone, Copy, PartialEq, Eq)]
struct Fingerprint {
    device: u64,
    inode: u64,
    length: u64,
    modified_seconds: i64,
    modified_nanoseconds: i64,
}

impl Fingerprint {
    fn from_metadata(metadata: &Metadata) -> Self {
        Self {
            device: metadata.dev(),
            inode: metadata.ino(),
            length: metadata.len(),
            modified_seconds: metadata.mtime(),
            modified_nanoseconds: metadata.mtime_nsec(),
        }
    }
}

struct State {
    file: File,
    fingerprint: Fingerprint,
    revision: u64,
}

fn std_block(state: &StdMutex<State>) -> u128 {
    let started = Instant::now();
    for _ in 0..ITERATIONS {
        let mut state = state.lock().unwrap();
        let observed = Fingerprint::from_metadata(&state.file.metadata().unwrap());
        if observed != state.fingerprint {
            state.fingerprint = observed;
            state.revision += 1;
        }
        black_box(state.revision);
    }
    started.elapsed().as_nanos()
}

fn parking_block(state: &parking_lot::Mutex<State>) -> u128 {
    let started = Instant::now();
    for _ in 0..ITERATIONS {
        let mut state = state.lock();
        let observed = Fingerprint::from_metadata(&state.file.metadata().unwrap());
        if observed != state.fingerprint {
            state.fingerprint = observed;
            state.revision += 1;
        }
        black_box(state.revision);
    }
    started.elapsed().as_nanos()
}

fn main() {
    let path = std::env::args().nth(1).expect("input path");
    let std_file = File::open(&path).unwrap();
    let parking_file = File::open(&path).unwrap();
    let std_state = StdMutex::new(State {
        fingerprint: Fingerprint::from_metadata(&std_file.metadata().unwrap()),
        file: std_file,
        revision: 0,
    });
    let parking_state = parking_lot::Mutex::new(State {
        fingerprint: Fingerprint::from_metadata(&parking_file.metadata().unwrap()),
        file: parking_file,
        revision: 0,
    });

    black_box(std_block(&std_state));
    black_box(parking_block(&parking_state));

    println!("block\torder\timplementation\tns_per_call");
    for block in 0..BLOCKS {
        let order = block % 4;
        if order == 0 || order == 3 {
            let elapsed = std_block(&std_state);
            println!("{block}\t{order}\tstd\t{}", elapsed / ITERATIONS as u128);
            let elapsed = parking_block(&parking_state);
            println!(
                "{block}\t{order}\tparking_lot\t{}",
                elapsed / ITERATIONS as u128
            );
        } else {
            let elapsed = parking_block(&parking_state);
            println!(
                "{block}\t{order}\tparking_lot\t{}",
                elapsed / ITERATIONS as u128
            );
            let elapsed = std_block(&std_state);
            println!("{block}\t{order}\tstd\t{}", elapsed / ITERATIONS as u128);
        }
    }
}
