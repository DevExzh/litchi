#define _GNU_SOURCE
#include <errno.h>
#include <inttypes.h>
#include <sched.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <unistd.h>

static volatile uint64_t sink;

static int parse_size(const char *s, size_t *out) {
    char *end = NULL;
    errno = 0;
    unsigned long long value = strtoull(s, &end, 0);
    if (errno || end == s || *end != '\0' || value == 0 || value > SIZE_MAX) return -1;
    *out = (size_t)value;
    return 0;
}

static int parse_u64(const char *s, uint64_t *out) {
    char *end = NULL;
    errno = 0;
    unsigned long long value = strtoull(s, &end, 0);
    if (errno || end == s || *end != '\0' || value == 0) return -1;
    *out = (uint64_t)value;
    return 0;
}

int main(int argc, char **argv) {
    if (argc != 4) {
        fprintf(stderr, "usage: %s CPU BYTES PASSES\n", argv[0]);
        return 2;
    }
    int cpu = atoi(argv[1]);
    size_t bytes;
    uint64_t passes;
    if (cpu < 0 || parse_size(argv[2], &bytes) || parse_u64(argv[3], &passes) || bytes < 64) {
        fprintf(stderr, "invalid arguments\n");
        return 2;
    }
    cpu_set_t set;
    CPU_ZERO(&set);
    CPU_SET(cpu, &set);
    if (sched_setaffinity(0, sizeof(set), &set) != 0) {
        fprintf(stderr, "sched_setaffinity(%d): %s\n", cpu, strerror(errno));
        return 1;
    }
    unsigned char *buf = NULL;
    if (posix_memalign((void **)&buf, 64, bytes) != 0 || !buf) {
        fprintf(stderr, "posix_memalign failed for %zu bytes\n", bytes);
        return 1;
    }
    for (size_t i = 0; i < bytes; i += 64) buf[i] = (unsigned char)(i >> 6);
    uint64_t lines = bytes / 64;
    for (uint64_t pass = 0; pass < passes; ++pass) {
        for (uint64_t line = 0; line < lines; ++line) {
            sink += buf[line * 64];
        }
    }
    printf("cpu=%d bytes=%zu passes=%" PRIu64 " lines=%" PRIu64 " sink=%" PRIu64 "\n",
           sched_getcpu(), bytes, passes, lines, sink);
    free(buf);
    return 0;
}
