use serde_json::to_writer_pretty;
use std::env;
use std::io;
use std::path::PathBuf;
use xls_index_probe_0684::{
    CorpusConfig, ProbeMode, Route, RouteConfig, default_samples, default_warmups, run_corpus,
    run_route,
};

fn usage() -> &'static str {
    "usage:\n  xls-index-probe-0684 route --input FILE --route cold1|cold2|cold3|prepared|visit --mode owned|file|owned-native|file-native --worksheet N --row N --column N [--second-row N --second-column N] [--warmups N --samples N]\n  xls-index-probe-0684 corpus --root DIR [--mode owned|file|owned-native|file-native] [--sample-coordinates N] [--max-queries N]"
}

fn take_value(
    arguments: &mut impl Iterator<Item = String>,
    option: &str,
) -> Result<String, String> {
    arguments
        .next()
        .ok_or_else(|| format!("{option} requires a value"))
}

fn parse_usize(value: String, option: &str) -> Result<usize, String> {
    value
        .parse()
        .map_err(|_| format!("{option} requires a non-negative integer"))
}

fn parse_u32(value: String, option: &str) -> Result<u32, String> {
    value
        .parse()
        .map_err(|_| format!("{option} requires a u32"))
}

fn parse_route(arguments: impl IntoIterator<Item = String>) -> Result<RouteConfig, String> {
    let mut arguments = arguments.into_iter();
    let mut input = None;
    let mut route = None;
    let mut mode = ProbeMode::Owned;
    let mut worksheet = 0;
    let mut row = 0;
    let mut column = 0;
    let mut second_row = 1;
    let mut second_column = 0;
    let mut warmups = default_warmups();
    let mut samples = default_samples();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--input" => input = Some(PathBuf::from(take_value(&mut arguments, "--input")?)),
            "--route" => {
                route = Some(
                    take_value(&mut arguments, "--route")?
                        .parse::<Route>()
                        .map_err(|error| error.to_string())?,
                )
            },
            "--mode" => {
                mode = take_value(&mut arguments, "--mode")?
                    .parse::<ProbeMode>()
                    .map_err(|error| error.to_string())?;
            },
            "--worksheet" => {
                worksheet = parse_usize(take_value(&mut arguments, "--worksheet")?, "--worksheet")?
            },
            "--row" => row = parse_u32(take_value(&mut arguments, "--row")?, "--row")?,
            "--column" => column = parse_u32(take_value(&mut arguments, "--column")?, "--column")?,
            "--second-row" => {
                second_row = parse_u32(take_value(&mut arguments, "--second-row")?, "--second-row")?
            },
            "--second-column" => {
                second_column = parse_u32(
                    take_value(&mut arguments, "--second-column")?,
                    "--second-column",
                )?
            },
            "--warmups" => {
                warmups = parse_usize(take_value(&mut arguments, "--warmups")?, "--warmups")?
            },
            "--samples" => {
                samples = parse_usize(take_value(&mut arguments, "--samples")?, "--samples")?
            },
            "--help" | "-h" => return Err(usage().to_owned()),
            value => return Err(format!("unknown route option {value:?}\n{}", usage())),
        }
    }
    let input = input.ok_or_else(|| "--input FILE is required".to_owned())?;
    let route = route.ok_or_else(|| "--route is required".to_owned())?;
    if warmups == 0 || samples == 0 {
        return Err("--warmups and --samples must be nonzero".to_owned());
    }
    let mut config = RouteConfig::defaults(input, route, mode);
    config.worksheet = worksheet;
    config.row = row;
    config.column = column;
    config.second_row = second_row;
    config.second_column = second_column;
    config.warmups = warmups;
    config.samples = samples;
    Ok(config)
}

fn parse_corpus(arguments: impl IntoIterator<Item = String>) -> Result<CorpusConfig, String> {
    let mut arguments = arguments.into_iter();
    let mut config = CorpusConfig::default();
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--root" => config.root = PathBuf::from(take_value(&mut arguments, "--root")?),
            "--mode" => {
                config.mode = take_value(&mut arguments, "--mode")?
                    .parse::<ProbeMode>()
                    .map_err(|error| error.to_string())?;
            },
            "--sample-coordinates" => {
                config.sample_coordinates = parse_usize(
                    take_value(&mut arguments, "--sample-coordinates")?,
                    "--sample-coordinates",
                )?;
            },
            "--max-queries" => {
                config.max_queries = parse_usize(
                    take_value(&mut arguments, "--max-queries")?,
                    "--max-queries",
                )?;
            },
            "--help" | "-h" => return Err(usage().to_owned()),
            value => return Err(format!("unknown corpus option {value:?}\n{}", usage())),
        }
    }
    if config.sample_coordinates == 0 || config.max_queries == 0 {
        return Err("corpus bounds must be nonzero".to_owned());
    }
    Ok(config)
}

fn write_json<T: serde::Serialize>(value: &T) -> Result<(), io::Error> {
    let stdout = io::stdout();
    let mut output = stdout.lock();
    to_writer_pretty(&mut output, value).map_err(io::Error::other)?;
    use std::io::Write;
    writeln!(output)
}

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let mut arguments = env::args().skip(1);
    let command = arguments.next().ok_or_else(|| io::Error::other(usage()))?;
    match command.as_str() {
        "route" => write_json(&run_route(parse_route(arguments)?)?.report)?,
        "corpus" => write_json(&run_corpus(parse_corpus(arguments)?)?)?,
        "--help" | "-h" => println!("{}", usage()),
        other => {
            return Err(io::Error::other(format!("unknown command {other:?}\n{}", usage())).into());
        },
    }
    Ok(())
}
