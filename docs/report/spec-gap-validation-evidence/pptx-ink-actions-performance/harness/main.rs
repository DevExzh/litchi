mod adapter;
mod support;

use std::env;
use std::process::ExitCode;

fn usage() {
    eprintln!(
        "usage: pptx-ink-actions-performance --list | --host-probe | \
         --matrix-correctness | --lane ID [--warmup N] [--samples N]"
    );
}

fn main() -> ExitCode {
    let mut args = env::args().skip(1);
    let Some(command) = args.next() else {
        usage();
        return ExitCode::from(2);
    };

    let result = match command.as_str() {
        "--help" | "-h" => {
            usage();
            return ExitCode::SUCCESS;
        },
        "--list" => {
            println!("{}", adapter::known_lanes().join("\n"));
            return ExitCode::SUCCESS;
        },
        "--matrix-correctness" => adapter::run_matrix(),
        "--host-probe" => adapter::host_probe(),
        "--lane" => {
            let Some(lane) = args.next() else {
                usage();
                return ExitCode::from(2);
            };
            let mut warmup = 2usize;
            let mut samples = 20usize;
            while let Some(flag) = args.next() {
                let Some(value) = args.next() else {
                    eprintln!("missing value for {flag}");
                    return ExitCode::from(2);
                };
                let parsed = match value.parse::<usize>() {
                    Ok(value) if value > 0 => value,
                    _ => {
                        eprintln!("invalid positive integer for {flag}: {value}");
                        return ExitCode::from(2);
                    },
                };
                match flag.as_str() {
                    "--warmup" => warmup = parsed,
                    "--samples" => samples = parsed,
                    _ => {
                        eprintln!("unknown flag: {flag}");
                        return ExitCode::from(2);
                    },
                }
            }
            adapter::run_lane(&lane, warmup, samples)
        },
        other => {
            eprintln!("unknown command: {other}");
            usage();
            return ExitCode::from(2);
        },
    };

    match result {
        Ok(output) => {
            println!("{output}");
            ExitCode::SUCCESS
        },
        Err(error) => {
            eprintln!("{error}");
            ExitCode::from(1)
        },
    }
}
