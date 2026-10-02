//! Bounded `BrtCalcProp` wire codec.

use std::io::Write;

use crate::leniency::{ReservedFieldDefect, Tolerance};
use crate::raw::{Cursor, Writer, kind};

use super::{Error, Mode, Opts, Props, Result, Threads};

/// Read a canonical or historical Excel 12 `BrtCalcProp` payload with [`Cursor`].
///
/// Canonical records use the 26-byte layout in `[MS-XLSB]` section 2.4.318.
/// Early Excel 12 producers emitted the final option word as one byte; that
/// exact 25-byte form is accepted as a zero-extended option word. All other
/// lengths, mode values, reserved option bits, thread counts, and `Xnum`
/// values are rejected before a value is returned. [`write()`] always emits the
/// canonical 26-byte form.
///
/// This standalone function cannot report repairs, so it also rejects the
/// reserved option bits and trailing bytes an opened workbook repairs and
/// records; see [`crate::leniency`].
pub fn read(payload: &[u8]) -> Result<Props> {
    read_with(payload, Tolerance::refusing())
}

/// Read a `BrtCalcProp` payload, masking undefined option bits and ignoring
/// trailing payload bytes through `tolerance`.
///
/// Mode, iteration, delta, and thread-count domains stay strict.
pub(crate) fn read_with(payload: &[u8], mut tolerance: Tolerance<'_>) -> Result<Props> {
    let mut cursor = Cursor::new(payload, "BrtCalcProp");
    let id = cursor.read_u32()?;
    let mode = Mode::read(cursor.read_u32()?)?;
    let iters = cursor.read_u32()?;
    let delta = super::Delta::new(cursor.read_f64()?)?;
    let threads = Threads::from_wire(cursor.read_i32()?)?;
    let bits = if cursor.remaining() == 1 {
        u16::from(cursor.read_u8()?)
    } else {
        cursor.read_u16()?
    };
    // MS-XLSB 2.4.318: the reserved option bits MUST be ignored.
    let reserved = bits & !Opts::all().bits();
    if reserved != 0 {
        tolerance.tolerate(
            ReservedFieldDefect::CalcOptions,
            kind::CALC_PROP,
            0,
            u32::from(reserved),
            || Error::ReservedOpts { bits: reserved },
        )?;
    }
    let opts = Opts::from_bits_truncate(bits);
    let trailing = cursor.remaining();
    if trailing == 0 {
        cursor.finish()?;
    } else {
        tolerance.tolerate_trailing(kind::CALC_PROP, 0, trailing, || {
            Error::Wire(crate::raw::Error::Trailing {
                context: "BrtCalcProp",
                offset: cursor.position(),
                remaining: trailing,
            })
        })?;
    }

    Ok(Props::from_wire(id, mode, iters, delta, threads, opts))
}

/// Stream one exact `BrtCalcProp` payload through [`Writer`].
///
/// This writes exactly [`super::LEN`] bytes and does not allocate an intermediate
/// payload buffer.
pub fn write<W: Write>(props: &Props, writer: &mut Writer<W>) -> Result<()> {
    writer.write_u32(props.id())?;
    writer.write_u32(props.mode().wire())?;
    writer.write_u32(props.iters())?;
    writer.write_f64(props.delta().get())?;
    writer.write_i32(i32::from(props.threads().get()))?;
    writer.write_u16(props.opts().bits())?;
    Ok(())
}
