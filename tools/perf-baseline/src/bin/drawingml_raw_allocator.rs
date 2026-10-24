#![forbid(unsafe_code)]

#[path = "../drawingml_raw_probe.rs"]
mod probe;

fn main() -> std::io::Result<()> {
    probe::run(true)
}
