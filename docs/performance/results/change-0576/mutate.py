import sys
p = "crates/litchi-xls/src/records.rs"
t = open(p, encoding="utf-8").read()
which = sys.argv[1]
M = {
 # M1: forget the pending high surrogate at a compressed continuation boundary.
 "M1": ("""        if self.pending_high && !chunk.is_empty() {
            self.pending_high = false;
            self.malformed = true;
        }""", """        let _ = chunk;"""),
 # M2: ignore a high surrogate still pending at the end of the string.
 "M2": ("        !self.malformed && !self.pending_high", "        !self.malformed"),
 # M3: report the refusal from the walk instead of re-walking through
 # `String::from_utf16`, using char::decode_utf16's wording.
 "M3": ("""        self.segment_index = restart.0;
        self.offset = restart.1;
        self.read_characters(count, high_byte).map(|_| ())""",
        """        let _ = restart;
        Err(Error::Encoding(
            "UTF-16 decoding error: unpaired surrogate found: d800".to_string(),
        ))"""),
 # M4: take the chunk fast exit even when a high surrogate is pending.
 "M4": ("        if !self.pending_high && !pairs.iter().any(|pair| pair[1] & 0xF8 == 0xD8) {",
        "        if !pairs.iter().any(|pair| pair[1] & 0xF8 == 0xD8) {"),
 # M5: swap which surrogate half opens a pair.
 "M5": ("""            match unit {
                0xD800..=0xDBFF => pending_high = true,
                0xDC00..=0xDFFF => malformed = true,
                _ => {},
            }""",
        """            match unit {
                0xDC00..=0xDFFF => pending_high = true,
                0xD800..=0xDBFF => malformed = true,
                _ => {},
            }"""),
 # M6: drop the re-examination of the unit that failed to complete a pair.
 "M6": ("""                malformed = true;
            }
            match unit {""",
        """                malformed = true;
                continue;
            }
            match unit {"""),
}
old, new = M[which]
assert t.count(old) == 1, (which, t.count(old))
open(p, "w", encoding="utf-8").write(t.replace(old, new, 1))
print(f"applied {which}")
