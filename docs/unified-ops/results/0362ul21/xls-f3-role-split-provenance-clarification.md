# XLS F3 role-split packet provenance clarification

The original paired packet remains byte-for-byte unchanged at
`sheet-native-readiness-472b-f3-role-split.json`, SHA-256
`51258c4b0e296af95cb675621f91a3f370cfbafb5c410880b1c7dc5bae205df5`.
After JSON decoding, its three shell command strings contain ordinary
`"$CPU"` quoting and no literal backslash-quote; the earlier concern came from
escaped presentation of JSON text, not packet drift.

The final corrected packet is the separate
`sheet-native-readiness-472b-f3-role-split-corrected.json`, SHA-256
`2a2b61ef7a5d0a13489bd4285ffe4b600afbd7ebed2223581229c625ea477e03`. It
adds structured argv/env launch facts and remains frozen. No source, driver,
historical endpoint, binary, target, or qualification evidence was changed.
