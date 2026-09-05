# Descriptive live-byte changes from sample entry

All table values are mean MiB above each sample’s entry checkpoint.
They are global allocator observations, not object-owned memory or RSS.

| Corpus | API | Repeat | Planned | Published | Plan dropped | Documents dropped | Sink dropped |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| plain | owned | R1 | 0.509020 | 0.590243 | 0.550676 | 0.122667 | 0.000000 |
| plain | source-backed | R1 | 0.339646 | 0.264227 | 0.261890 | 0.183219 | 0.000000 |
| plain | source-backed | R2 | 0.339646 | 0.264227 | 0.261890 | 0.183219 | 0.000000 |
| plain | owned | R2 | 0.509020 | 0.590243 | 0.550676 | 0.122667 | 0.000000 |
| media-rich | owned | R1 | 128.600019 | 160.719714 | 160.662287 | 64.149172 | 0.000000 |
| media-rich | source-backed | R1 | 118.399294 | 118.309136 | 102.293739 | 96.222899 | 0.000000 |
| media-rich | source-backed | R2 | 118.399294 | 118.309136 | 102.293739 | 96.222899 | 0.000000 |
| media-rich | owned | R2 | 128.600019 | 160.719714 | 160.662287 | 64.149172 | 0.000000 |
