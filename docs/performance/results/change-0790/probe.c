#define _GNU_SOURCE
#define _POSIX_C_SOURCE 200809L

#include <errno.h>
#include <inttypes.h>
#include <pthread.h>
#include <stdbool.h>
#include <stddef.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/mman.h>
#include <sys/resource.h>
#include <sys/types.h>
#include <unistd.h>

#define MAX_WORKERS 32U
#define LINE_CAPACITY 256U
#define PAGE_PATTERN_PERIOD 251U

enum touch_mode {
    TOUCH_MAIN = 0,
    TOUCH_WORKERS = 1,
};

struct options {
    uint64_t mib;
    unsigned workers;
    enum touch_mode touch;
    bool have_mib;
    bool have_workers;
    bool have_touch;
};

struct worker_arg {
    volatile unsigned char *mapping;
    size_t page_size;
    size_t first_page;
    size_t last_page;
    unsigned worker_index;
    bool touch_payload;
};

static const char usage_text[] =
    "usage: probe --mib {0,1,4,16,64} --workers {0,4,32} "
    "--touch {main,workers}\n";

static int write_all(int fd, const void *data, size_t length)
{
    const unsigned char *cursor = (const unsigned char *)data;
    size_t written = 0U;

    while (written < length) {
        ssize_t count = write(fd, cursor + written, length - written);
        if (count < 0) {
            if (errno == EINTR) {
                continue;
            }
            return -1;
        }
        if (count == 0) {
            errno = EIO;
            return -1;
        }
        written += (size_t)count;
    }
    return 0;
}

static void write_error(const char *message)
{
    (void)write_all(STDERR_FILENO, message, strlen(message));
}

static int parse_u64(const char *text, uint64_t *value)
{
    const unsigned char *cursor = (const unsigned char *)text;
    char *end = NULL;
    unsigned long long parsed;

    if (text == NULL || *text == '\0') {
        return -1;
    }
    while (*cursor != '\0') {
        if (*cursor < (unsigned char)'0' || *cursor > (unsigned char)'9') {
            return -1;
        }
        ++cursor;
    }

    errno = 0;
    parsed = strtoull(text, &end, 10);
    if (errno == ERANGE || end == text || *end != '\0') {
        return -1;
    }
    *value = (uint64_t)parsed;
    return 0;
}

static int parse_options(int argc, char **argv, struct options *options)
{
    int index;

    memset(options, 0, sizeof(*options));
    for (index = 1; index < argc; index += 2) {
        const char *option;
        const char *value;
        uint64_t parsed;

        if (index + 1 >= argc) {
            return -1;
        }
        option = argv[index];
        value = argv[index + 1];
        if (strcmp(option, "--mib") == 0) {
            if (options->have_mib || parse_u64(value, &parsed) != 0) {
                return -1;
            }
            if (parsed != 0U && parsed != 1U && parsed != 4U &&
                parsed != 16U && parsed != 64U) {
                return -1;
            }
            options->mib = parsed;
            options->have_mib = true;
        } else if (strcmp(option, "--workers") == 0) {
            if (options->have_workers || parse_u64(value, &parsed) != 0) {
                return -1;
            }
            if (parsed != 0U && parsed != 4U && parsed != 32U) {
                return -1;
            }
            options->workers = (unsigned)parsed;
            options->have_workers = true;
        } else if (strcmp(option, "--touch") == 0) {
            if (options->have_touch) {
                return -1;
            }
            if (strcmp(value, "main") == 0) {
                options->touch = TOUCH_MAIN;
            } else if (strcmp(value, "workers") == 0) {
                options->touch = TOUCH_WORKERS;
            } else {
                return -1;
            }
            options->have_touch = true;
        } else {
            return -1;
        }
    }

    if (!options->have_mib || !options->have_workers || !options->have_touch) {
        return -1;
    }
    if (options->touch == TOUCH_WORKERS && options->workers == 0U) {
        return -1;
    }
    return 0;
}

static unsigned char page_value(size_t page_index)
{
    return (unsigned char)((page_index % PAGE_PATTERN_PERIOD) + 1U);
}

static void touch_page_range(const struct worker_arg *argument)
{
    size_t page;

    for (page = argument->first_page; page < argument->last_page; ++page) {
        argument->mapping[page * argument->page_size] = page_value(page);
    }
}

static void touch_worker_stack(unsigned worker_index)
{
    volatile unsigned char stack_touch[64];
    unsigned char folded = 0U;
    size_t index;

    for (index = 0U; index < sizeof(stack_touch); ++index) {
        stack_touch[index] = (unsigned char)(index + worker_index + 1U);
    }
    for (index = 0U; index < sizeof(stack_touch); ++index) {
        folded ^= stack_touch[index];
    }
    if (folded == 0xffU) {
        stack_touch[0] ^= 1U;
    }
}

static void *worker_main(void *opaque)
{
    struct worker_arg *argument = (struct worker_arg *)opaque;

    touch_worker_stack(argument->worker_index);
    if (argument->touch_payload && argument->first_page < argument->last_page) {
        touch_page_range(argument);
    }
    return NULL;
}

static int run_workers(unsigned worker_count, volatile unsigned char *mapping,
                       size_t page_size, size_t page_count,
                       bool touch_payload)
{
    pthread_t threads[MAX_WORKERS];
    struct worker_arg arguments[MAX_WORKERS];
    unsigned created = 0U;
    unsigned index;
    size_t quotient;
    size_t remainder;
    int status = 0;

    if (worker_count == 0U) {
        return 0;
    }

    quotient = page_count / (size_t)worker_count;
    remainder = page_count % (size_t)worker_count;
    for (index = 0U; index < worker_count; ++index) {
        size_t extra_before = index < remainder ? (size_t)index : remainder;
        size_t first_page = ((size_t)index * quotient) + extra_before;
        size_t page_span = quotient + (index < remainder ? 1U : 0U);
        int result;

        arguments[index].mapping = mapping;
        arguments[index].page_size = page_size;
        arguments[index].first_page = first_page;
        arguments[index].last_page = first_page + page_span;
        arguments[index].worker_index = index;
        arguments[index].touch_payload = touch_payload;

        result = pthread_create(&threads[index], NULL, worker_main,
                                &arguments[index]);
        if (result != 0) {
            status = -1;
            break;
        }
        ++created;
    }

    for (index = 0U; index < created; ++index) {
        int result = pthread_join(threads[index], NULL);
        if (result != 0) {
            write_error("pthread_join failed; terminate without releasing worker storage\n");
            _exit(EXIT_FAILURE);
        }
    }
    return status;
}

static int verify_pages(volatile unsigned char *mapping, size_t page_size,
                        size_t page_count, uint64_t *checksum)
{
    uint64_t total = 0U;
    size_t page;

    if (page_count != 0U && mapping == NULL) {
        return -1;
    }
    for (page = 0U; page < page_count; ++page) {
        unsigned char observed = mapping[page * page_size];
        unsigned char expected = page_value(page);
        if (observed != expected) {
            return -1;
        }
        total += (uint64_t)observed;
    }
    *checksum = total;
    return 0;
}

static int read_ack(void)
{
    unsigned char ack[2];
    size_t received = 0U;

    while (received < sizeof(ack)) {
        ssize_t count = read(STDIN_FILENO, ack + received,
                             sizeof(ack) - received);
        if (count < 0) {
            if (errno == EINTR) {
                continue;
            }
            return -1;
        }
        if (count == 0) {
            errno = EPIPE;
            return -1;
        }
        received += (size_t)count;
    }
    if (ack[0] != (unsigned char)'+' || ack[1] != (unsigned char)'\n') {
        errno = EPROTO;
        return -1;
    }
    return 0;
}

static int emit_checkpoint(const char *phase, pid_t pid, uint64_t bytes,
                           uint64_t checksum)
{
    struct rusage before_ack;
    struct rusage after_ack;
    char line[LINE_CAPACITY];
    int length;

    if (getrusage(RUSAGE_SELF, &before_ack) != 0) {
        return -1;
    }
    length = snprintf(line, sizeof(line),
                      "RSS0789\t%s\t%d\t%ld\t%ld\t%ld\t%" PRIu64
                      "\t%" PRIu64 "\n",
                      phase, (int)pid, before_ack.ru_maxrss,
                      before_ack.ru_minflt, before_ack.ru_majflt, bytes,
                      checksum);
    if (length < 0 || (size_t)length >= sizeof(line)) {
        errno = EOVERFLOW;
        return -1;
    }
    if (write_all(STDOUT_FILENO, line, (size_t)length) != 0) {
        return -1;
    }
    if (read_ack() != 0) {
        return -1;
    }
    if (getrusage(RUSAGE_SELF, &after_ack) != 0) {
        return -1;
    }
    length = snprintf(line, sizeof(line), "ACK0789\t%s\t%d\t%ld\t%ld\t%ld\n",
                      phase, (int)pid, after_ack.ru_maxrss,
                      after_ack.ru_minflt, after_ack.ru_majflt);
    if (length < 0 || (size_t)length >= sizeof(line)) {
        errno = EOVERFLOW;
        return -1;
    }
    return write_all(STDOUT_FILENO, line, (size_t)length);
}

int main(int argc, char **argv)
{
    struct options options;
    long page_size_result;
    size_t page_size;
    size_t bytes;
    size_t page_count;
    uint64_t bytes_u64;
    volatile unsigned char *mapping = NULL;
    void *mapping_base = NULL;
    bool mapping_is_live = false;
    uint64_t checksum = 0U;
    pid_t pid = getpid();
    int status = EXIT_FAILURE;

    if (parse_options(argc, argv, &options) != 0) {
        write_error(usage_text);
        return EXIT_FAILURE;
    }

    page_size_result = sysconf(_SC_PAGESIZE);
    if (page_size_result <= 0) {
        write_error("sysconf(_SC_PAGESIZE) failed\n");
        return EXIT_FAILURE;
    }
    page_size = (size_t)page_size_result;
    bytes_u64 = options.mib * UINT64_C(1024) * UINT64_C(1024);
    if (bytes_u64 > (uint64_t)SIZE_MAX) {
        write_error("mapping size does not fit size_t\n");
        return EXIT_FAILURE;
    }
    bytes = (size_t)bytes_u64;
    page_count = bytes / page_size;
    if (bytes % page_size != 0U) {
        ++page_count;
    }

    if (emit_checkpoint("startup", pid, 0U, 0U) != 0) {
        write_error("startup checkpoint failed\n");
        goto cleanup;
    }

    if (bytes != 0U) {
        mapping_base = mmap(NULL, bytes, PROT_READ | PROT_WRITE,
                            MAP_PRIVATE | MAP_ANONYMOUS, -1, 0);
        if (mapping_base == MAP_FAILED) {
            mapping_base = NULL;
            write_error("mmap failed\n");
            goto cleanup;
        }
        mapping = (volatile unsigned char *)mapping_base;
        mapping_is_live = true;
    }

    if (emit_checkpoint("mapped", pid, bytes, 0U) != 0) {
        write_error("mapped checkpoint failed\n");
        goto cleanup;
    }

    if (options.touch == TOUCH_WORKERS) {
        if (run_workers(options.workers, mapping, page_size, page_count, true) !=
            0) {
            write_error("payload worker execution failed\n");
            goto cleanup;
        }
    } else if (page_count != 0U) {
        struct worker_arg main_argument;

        main_argument.mapping = mapping;
        main_argument.page_size = page_size;
        main_argument.first_page = 0U;
        main_argument.last_page = page_count;
        main_argument.worker_index = 0U;
        main_argument.touch_payload = true;
        touch_page_range(&main_argument);
    }

    if (verify_pages(mapping, page_size, page_count, &checksum) != 0) {
        write_error("page pattern verification failed\n");
        goto cleanup;
    }
    if (options.touch == TOUCH_MAIN) {
        if (run_workers(options.workers, mapping, page_size, page_count, false) !=
            0) {
            write_error("control worker execution failed\n");
            goto cleanup;
        }
    }
    if (emit_checkpoint("touched", pid, bytes, checksum) != 0) {
        write_error("touched checkpoint failed\n");
        goto cleanup;
    }
    if (emit_checkpoint("workers_joined", pid, bytes, checksum) != 0) {
        write_error("workers_joined checkpoint failed\n");
        goto cleanup;
    }

    if (mapping_is_live) {
        if (munmap(mapping_base, bytes) != 0) {
            write_error("munmap failed\n");
            goto cleanup;
        }
        mapping_base = NULL;
        mapping = NULL;
        mapping_is_live = false;
    }
    if (emit_checkpoint("unmapped", pid, 0U, checksum) != 0) {
        write_error("unmapped checkpoint failed\n");
        goto cleanup;
    }
    if (emit_checkpoint("final", pid, 0U, checksum) != 0) {
        write_error("final checkpoint failed\n");
        goto cleanup;
    }

    status = EXIT_SUCCESS;

cleanup:
    if (mapping_is_live) {
        (void)munmap(mapping_base, bytes);
    }
    return status;
}
