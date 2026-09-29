#![forbid(unsafe_code)]
mod codec;
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<_> = std::env::args_os().skip(1).collect();
    if args.len() != 1 {
        return Err("usage: codec-qualify NEW_OUTPUT_DIRECTORY".into());
    }
    codec::qualify(std::path::Path::new(&args[0]))
}
