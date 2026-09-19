use std::env;
use std::io;

use xls_index_retry_probe_0686::{parse_args, run, usage, write_report};

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let arguments = env::args().skip(1);
    match parse_args(arguments) {
        Ok(config) => write_report(&run(config)?)?,
        Err(error) if error == usage() => println!("{error}"),
        Err(error) => return Err(io::Error::new(io::ErrorKind::InvalidInput, error).into()),
    }
    Ok(())
}
