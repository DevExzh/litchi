//! Public picture-bullet construction and lossless serialization.

use litchi_ppt::records::Record;
use litchi_ppt::{PictureBullet, PictureBulletCollection, PictureBulletLimits, PictureBulletType};

const PNG_SIGNATURE: &[u8] = b"\x89PNG\r\n\x1a\n";

fn png_blip() -> Vec<u8> {
    let mut blip = Vec::new();
    blip.extend_from_slice(&(0x06e0u16 << 4).to_le_bytes());
    blip.extend_from_slice(&0xf01eu16.to_le_bytes());
    blip.extend_from_slice(&25u32.to_le_bytes());
    blip.extend_from_slice(&[0; 17]);
    blip.extend_from_slice(PNG_SIGNATURE);
    blip
}

#[test]
fn public_picture_bullet_construction_reads_and_serializes_losslessly() {
    let limits = PictureBulletLimits::default();
    let officeart_record = png_blip();
    let officeart_pointer = officeart_record.as_ptr();
    let expected_officeart_record = officeart_record.clone();
    let bullet = PictureBullet::new(3, PictureBulletType::Png, 0x7f, officeart_record, limits)
        .expect("checked picture-bullet construction");

    assert_eq!(bullet.officeart_record.as_ptr(), officeart_pointer);
    assert_eq!(
        bullet.blip().expect("read constructed BLIP").data(),
        PNG_SIGNATURE
    );
    let record = bullet.to_record().expect("serialize constructed bullet");
    assert_eq!(record.record_type_raw, 2041);
    assert_eq!(
        record.data.as_slice(),
        [
            &[PictureBulletType::Png as u8, 0x7f][..],
            expected_officeart_record.as_slice()
        ]
        .concat()
    );

    let collection = PictureBulletCollection {
        bullets: vec![bullet],
    };
    let bytes = collection.to_bytes().expect("serialize collection");
    let parsed_record = Record::parse(&bytes, 0, &litchi_ppt::SourceLimits::default())
        .expect("parse serialized collection")
        .0;
    let parsed = PictureBulletCollection::parse(&parsed_record, &limits)
        .expect("read serialized collection");
    let parsed_bullet = parsed
        .get(3)
        .expect("constructed bullet survives round trip");
    assert_eq!(parsed_bullet.officeart_record, expected_officeart_record);
    assert_eq!(
        parsed_bullet
            .blip()
            .expect("read round-tripped BLIP")
            .data(),
        PNG_SIGNATURE
    );
}

#[test]
fn public_picture_bullet_constructor_rejects_a_record_over_its_limit() {
    let limits = PictureBulletLimits {
        max_officeart_record_bytes: 24,
        ..PictureBulletLimits::default()
    };
    assert!(PictureBullet::new(3, PictureBulletType::Png, 0x7f, png_blip(), limits,).is_err());
}
