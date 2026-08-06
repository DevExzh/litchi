//! BIFF wire codecs for Obj records, control payloads, and record framing.

mod biff;
mod control;
mod obj;

pub(crate) use biff::{ranges, u32_at};
