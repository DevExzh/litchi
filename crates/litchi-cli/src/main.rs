//! The `litchi` binary; see the `litchi_cli` crate documentation.

#![forbid(unsafe_code)]

fn main() -> std::process::ExitCode {
    litchi_cli::main()
}
