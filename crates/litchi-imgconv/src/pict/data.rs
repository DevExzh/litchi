//! PackBits decompression for PICT pixel rows.
//!
//! PackBits uses a signed control byte: `0..=127` copies the next
//! `control + 1` items literally, `-127..=-1` repeats the next item
//! `1 - control` times, and `-128` is ignored. Ordinary rows use one-byte
//! items; 16-bit pixel maps with `packType` 3 use two-byte items (*Inside
//! Macintosh: Imaging With QuickDraw*, "Pixel Maps" and `UnpackBits`).

/// How a packed row filled its unpacked buffer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Unpacked {
    /// Number of output bytes written.
    pub filled: usize,
    /// Whether the packed data described more bytes than the row holds.
    pub overrun: bool,
}

/// Decodes one PackBits row of one-byte items into `out`.
pub(super) fn unpack_bits(compressed: &[u8], out: &mut [u8]) -> Unpacked {
    unpack_items::<1>(compressed, out)
}

/// Decodes one PackBits row of two-byte items into `out`.
pub(super) fn unpack_words(compressed: &[u8], out: &mut [u8]) -> Unpacked {
    unpack_items::<2>(compressed, out)
}

fn unpack_items<const ITEM: usize>(compressed: &[u8], out: &mut [u8]) -> Unpacked {
    let mut input = 0usize;
    let mut output = 0usize;
    let mut overrun = false;
    while let Some(&control) = compressed.get(input) {
        input += 1;
        let control = control as i8;
        if control == -128 {
            continue;
        }
        if control < 0 {
            let Some(item) = compressed.get(input..input + ITEM) else {
                break;
            };
            input += ITEM;
            let count = usize::from(control.unsigned_abs()) + 1;
            for _ in 0..count {
                if output + ITEM > out.len() {
                    overrun = true;
                    break;
                }
                out[output..output + ITEM].copy_from_slice(item);
                output += ITEM;
            }
        } else {
            let bytes = (usize::from(control as u8) + 1) * ITEM;
            let available = compressed.len() - input;
            let take = bytes.min(available);
            let room = out.len() - output;
            let copy = take.min(room);
            out[output..output + copy].copy_from_slice(&compressed[input..input + copy]);
            output += copy;
            input += take;
            if take > room {
                overrun = true;
            }
        }
    }
    Unpacked {
        filled: output,
        overrun,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unpack_bits_literal_run() {
        let mut output = [0u8; 3];
        let result = unpack_bits(&[2, 0xAA, 0xBB, 0xCC], &mut output);
        assert_eq!(output, [0xAA, 0xBB, 0xCC]);
        assert_eq!(
            result,
            Unpacked {
                filled: 3,
                overrun: false
            }
        );
    }

    #[test]
    fn unpack_bits_repeated_run_and_noop() {
        let mut output = [0u8; 4];
        let result = unpack_bits(&[0x80, 0xFE, 0xDD, 0x00, 0x11], &mut output);
        assert_eq!(output, [0xDD, 0xDD, 0xDD, 0x11]);
        assert_eq!(result.filled, 4);
    }

    #[test]
    fn unpack_bits_reports_overrun_and_short_rows() {
        let mut output = [0u8; 2];
        let result = unpack_bits(&[5, 1, 2, 3, 4, 5, 6], &mut output);
        assert_eq!(output, [1, 2]);
        assert!(result.overrun);

        let mut short = [0u8; 4];
        let result = unpack_bits(&[0, 9], &mut short);
        assert_eq!(short, [9, 0, 0, 0]);
        assert_eq!(result.filled, 1);
        assert!(!result.overrun);
    }

    #[test]
    fn unpack_words_repeats_two_byte_items() {
        let mut output = [0u8; 8];
        let result = unpack_words(&[0xFE, 0x12, 0x34, 0x00, 0xAB, 0xCD], &mut output);
        assert_eq!(output, [0x12, 0x34, 0x12, 0x34, 0x12, 0x34, 0xAB, 0xCD]);
        assert_eq!(result.filled, 8);
    }
}
