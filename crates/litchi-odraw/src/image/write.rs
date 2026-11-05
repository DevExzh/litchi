//! Streaming `OfficeArt` BLIP copying and digest helpers.

use std::io::{self, Write};

use super::{Blip, Limits, Uid};
use crate::write::{self as record_write, Atom};

/// Copies a parsed BLIP while retaining its exact record type and instance.
///
/// # Errors
///
/// Returns an error if the record kind is not a known BLIP atom, or if the
/// writer fails.
pub fn copy<W: Write>(writer: &mut W, blip: &Blip<'_>, limits: Limits) -> io::Result<()> {
    let record = blip.record();
    bounded_blip_len(record.len(), limits)?;
    let atom = match record.raw_kind() {
        0xF01A => Atom::BlipEmf,
        0xF01B => Atom::BlipWmf,
        0xF01C => Atom::BlipPict,
        0xF01D => Atom::BlipJpeg,
        0xF01E => Atom::BlipPng,
        0xF01F => Atom::BlipDib,
        0xF029 => Atom::BlipTiff,
        0xF02A => Atom::BlipJpeg2,
        raw => Atom::unknown(raw, record.version())?,
    };
    record_write::atom_header(
        writer,
        record.instance(),
        atom,
        record.len(),
        limits.records,
    )?;
    writer.write_all(record.data())
}

pub(crate) fn bounded_record_len(
    body_len: u32,
    limits: Limits,
    label: &'static str,
) -> io::Result<u32> {
    if body_len > limits.records.max_record_bytes {
        return Err(invalid(match label {
            "BStore body" => "BStore body exceeds the configured record limit",
            _ => "OfficeArt record body exceeds the configured record limit",
        }));
    }
    Ok(body_len)
}

pub(crate) fn bounded_blip_len(body_len: u32, limits: Limits) -> io::Result<u32> {
    bounded_record_len(body_len, limits, "BLIP body")?;
    if body_len > limits.max_blip_bytes {
        return Err(invalid("BLIP body exceeds the configured image limit"));
    }
    Ok(body_len)
}

/// Computes the RFC1320 MD4 digest used for `OfficeArt` image UIDs.
#[must_use]
pub fn digest(data: &[u8]) -> Uid {
    let mut state = [0x6745_2301, 0xEFCD_AB89, 0x98BA_DCFE, 0x1032_5476];
    let (chunks, remainder) = data.as_chunks::<64>();
    for chunk in chunks {
        let mut block = [0; 64];
        block.copy_from_slice(chunk);
        compress(&mut state, &block);
    }

    let mut tail = [0; 128];
    tail[..remainder.len()].copy_from_slice(remainder);
    tail[remainder.len()] = 0x80;
    let padded = if remainder.len() < 56 { 64 } else { 128 };
    let bit_len = (data.len() as u64).wrapping_mul(8);
    tail[padded - 8..padded].copy_from_slice(&bit_len.to_le_bytes());
    for chunk in tail[..padded].as_chunks::<64>().0 {
        let mut block = [0; 64];
        block.copy_from_slice(chunk);
        compress(&mut state, &block);
    }

    let mut digest = [0; 16];
    for (bytes, word) in digest.as_chunks_mut::<4>().0.iter_mut().zip(state) {
        bytes.copy_from_slice(&word.to_le_bytes());
    }
    Uid::new(digest)
}

fn compress(state: &mut [u32; 4], block: &[u8; 64]) {
    let mut words = [0u32; 16];
    for (word, bytes) in words.iter_mut().zip(block.as_chunks::<4>().0) {
        *word = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
    }
    let [mut aa, mut bb, mut cc, mut dd] = *state;
    macro_rules! round {
        ($f:expr, $k:expr, [$($index:expr),*], [$($shift:expr),*]) => {{
            let indices = [$($index),*];
            let shifts = [$($shift),*];
            for (step, (&index, &shift)) in indices.iter().zip(&shifts).enumerate() {
                let value = match step & 3 {
                    0 => aa.wrapping_add($f(bb, cc, dd)).wrapping_add(words[index]).wrapping_add($k).rotate_left(shift),
                    1 => dd.wrapping_add($f(aa, bb, cc)).wrapping_add(words[index]).wrapping_add($k).rotate_left(shift),
                    2 => cc.wrapping_add($f(dd, aa, bb)).wrapping_add(words[index]).wrapping_add($k).rotate_left(shift),
                    _ => bb.wrapping_add($f(cc, dd, aa)).wrapping_add(words[index]).wrapping_add($k).rotate_left(shift),
                };
                match step & 3 { 0 => aa = value, 1 => dd = value, 2 => cc = value, _ => bb = value }
            }
        }};
    }
    round!(
        |left: u32, mid: u32, right: u32| (left & mid) | (!left & right),
        0,
        [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
        [3, 7, 11, 19, 3, 7, 11, 19, 3, 7, 11, 19, 3, 7, 11, 19]
    );
    round!(
        |left: u32, mid: u32, right: u32| (left & mid) | (left & right) | (mid & right),
        0x5A82_7999,
        [0, 4, 8, 12, 1, 5, 9, 13, 2, 6, 10, 14, 3, 7, 11, 15],
        [3, 5, 9, 13, 3, 5, 9, 13, 3, 5, 9, 13, 3, 5, 9, 13]
    );
    round!(
        |left: u32, mid: u32, right: u32| left ^ mid ^ right,
        0x6ED9_EBA1,
        [0, 8, 4, 12, 2, 10, 6, 14, 1, 9, 5, 13, 3, 11, 7, 15],
        [3, 9, 11, 15, 3, 9, 11, 15, 3, 9, 11, 15, 3, 9, 11, 15]
    );
    state[0] = state[0].wrapping_add(aa);
    state[1] = state[1].wrapping_add(bb);
    state[2] = state[2].wrapping_add(cc);
    state[3] = state[3].wrapping_add(dd);
}

pub(crate) fn invalid(error: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, error.into())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::super::builder::{BlipBuilder, StoreBuilder};
    use super::*;
    use crate::image::{Blip, Compression, Context, Delay, JpegFlavor, Kind, Point, Rect, Store};

    #[test]
    fn md4_matches_rfc_vectors() {
        assert_eq!(
            digest(b"").bytes(),
            [
                0x31, 0xD6, 0xCF, 0xE0, 0xD1, 0x6A, 0xE9, 0x31, 0xB7, 0x3C, 0x59, 0xD7, 0xE0, 0xC0,
                0x89, 0xC0,
            ]
        );
        assert_eq!(
            digest(b"abc").bytes(),
            [
                0xA4, 0x48, 0x01, 0x7A, 0xAF, 0x21, 0xD8, 0x52, 0x5F, 0xC1, 0x0A, 0xE8, 0x7A, 0xA6,
                0x72, 0x9D,
            ]
        );
    }

    #[test]
    fn metafile_writer_marks_raw_data_uncompressed() {
        let blip = BlipBuilder::meta(
            Kind::Wmf,
            b"wmf".as_slice(),
            Rect::default(),
            Point { x: 10, y: 20 },
        )
        .unwrap();
        let mut bytes = Vec::new();
        blip.write(&mut bytes, Limits::default()).unwrap();
        let Blip::Wmf(parsed) = Blip::parse(&bytes, Default::default()).unwrap() else {
            panic!("expected WMF")
        };
        assert_eq!(parsed.header().compression, Compression::None);
        assert_eq!(parsed.header().size, 3);
        assert_eq!(parsed.uids().effective(), digest(b"wmf"));
    }

    #[test]
    fn writes_two_uid_alternate_jpeg_losslessly() {
        let blip = BlipBuilder::bitmap(Kind::Jpeg, b"jpeg".as_slice())
            .unwrap()
            .jpeg_flavor(JpegFlavor::Alternate)
            .unwrap()
            .two_uids();
        let mut bytes = Vec::new();
        blip.write(&mut bytes, Limits::default()).unwrap();
        let Blip::Jpeg(parsed) = Blip::parse(&bytes, Default::default()).unwrap() else {
            panic!("expected JPEG")
        };
        assert_eq!(parsed.jpeg_flavor(), Some(JpegFlavor::Alternate));
        assert_eq!(parsed.uids().second(), Some(digest(b"jpeg")));
    }

    #[test]
    fn store_writer_resolves_embedded_and_delayed_payloads() {
        let mut builder = StoreBuilder::new();
        let embedded = builder
            .add_embedded(BlipBuilder::bitmap(Kind::Png, b"one".as_slice()).unwrap())
            .unwrap();
        let delayed = builder
            .add_delayed(BlipBuilder::bitmap(Kind::Png, Vec::from(&b"two"[..])).unwrap())
            .unwrap();
        let mut store_bytes = Vec::new();
        let mut delay_bytes = Vec::new();
        builder
            .write_store(&mut store_bytes, Limits::default())
            .unwrap();
        builder
            .write_delay(&mut delay_bytes, Limits::default())
            .unwrap();

        let store = Store::parse(&store_bytes, Default::default()).unwrap();
        assert!(matches!(
            store.resolve(embedded, Context::new()).unwrap(),
            Some(Blip::Png(_))
        ));
        assert!(matches!(
            store
                .resolve(
                    delayed,
                    Context::new().with_delay(Delay::new(&delay_bytes, Default::default()))
                )
                .unwrap(),
            Some(Blip::Png(_))
        ));
    }
}
