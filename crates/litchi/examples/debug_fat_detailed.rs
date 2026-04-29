fn main() {
    // Simulate FAT allocation
    let sector_size: usize = 512;
    let stream_size: usize = 4096;
    let num_sectors = stream_size.div_ceil(sector_size);

    println!("Stream size: {} bytes", stream_size);
    println!("Sector size: {} bytes", sector_size);
    println!("Number of sectors needed: {}", num_sectors);
    println!();

    // Simulate the loop
    let mut fat: Vec<u32> = Vec::new();
    let mut next_sector = 0u32;

    // Pre-allocate
    let new_size = (next_sector as usize + num_sectors).max(fat.len());
    println!(
        "Resizing FAT to {} entries (filled with FREESECT)",
        new_size
    );
    fat.resize(new_size, 0xFFFFFFFF); // FREESECT

    println!("\nAllocating chain:");
    let start_sector = next_sector;
    for i in 0..num_sectors {
        let current_sector = next_sector;
        next_sector += 1;

        let next_value = if i < num_sectors - 1 {
            current_sector + 1
        } else {
            0xFFFFFFFE // ENDOFCHAIN
        };

        println!(
            "  i={}, current_sector={}, next_value=0x{:08X}",
            i, current_sector, next_value
        );

        fat[current_sector as usize] = next_value;
    }

    println!(
        "\nFinal FAT (sectors {} to {}):",
        start_sector,
        next_sector - 1
    );
    for i in start_sector..next_sector {
        let value = fat[i as usize];
        let status = match value {
            0xFFFFFFFF => "FREE".to_string(),
            0xFFFFFFFE => "END".to_string(),
            _ => format!("-> {}", value),
        };
        println!("  Sector {}: {}", i, status);
    }
}
