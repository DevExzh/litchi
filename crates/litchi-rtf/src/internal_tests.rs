//! Regressions of the crate-private source-backed tail-append engine
//! (record 0359a).
//!
//! This suite drove `TailAppendSourceProof`, `TailAppendSourceEdit` and
//! `TailAppendSourcePublicationPlan` as an integration test. Record 0359a
//! makes them crate-private behind `source_backed::Package` and its
//! `tail_append` family, so the suite compiles here as a unit test of the
//! same code. The crate is named `litchi_rtf` inside it (`extern crate
//! self`), so its paths are unchanged.

#[path = "../tests/internal/source_tail_append.rs"]
mod source_tail_append;
