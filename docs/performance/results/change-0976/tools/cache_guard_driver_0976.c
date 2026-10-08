#define _GNU_SOURCE
#include <elfutils/libdwfl.h>
#include <errno.h>
#include <fcntl.h>
#include <inttypes.h>
#include <stdbool.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <unistd.h>

static const Dwfl_Callbacks callbacks = {
    .find_elf = dwfl_linux_proc_find_elf,
    .find_debuginfo = dwfl_standard_find_debuginfo,
};

static unsigned long long parse_u64(const char *text) {
    char *end = NULL;
    errno = 0;
    unsigned long long value = strtoull(text, &end, 0);
    if (errno != 0 || end == text || *end != '\0') {
        fprintf(stderr, "invalid address: %s\n", text);
        exit(2);
    }
    return value;
}

static Dwfl_Module *report(Dwfl *dwfl, const char *name, const char *path,
                           int fd, GElf_Addr base, bool add_p_vaddr) {
    Dwfl_Module *module = dwfl_report_elf(dwfl, name, path, fd, base, add_p_vaddr);
    if (module == NULL)
        fprintf(stderr, "report failed name=%s path=%s fd=%d base=0x%" PRIx64 " add=%d errno=%d dwfl=%s\n",
                name, path, fd, (uint64_t)base, add_p_vaddr, errno,
                dwfl_errmsg(-1));
    return module;
}

static Dwfl *begin(void) {
    Dwfl *dwfl = dwfl_begin(&callbacks);
    if (dwfl == NULL) {
        fprintf(stderr, "dwfl_begin failed: %s\n", dwfl_errmsg(-1));
        exit(3);
    }
    dwfl_report_begin(dwfl);
    return dwfl;
}

int main(int argc, char **argv) {
    if (argc != 4) {
        fprintf(stderr, "usage: %s TARGET NAME EXPECTED_HEX\n", argv[0]);
        return 2;
    }
    const char *path = argv[1];
    const char *name = argv[2];
    GElf_Addr expected = (GElf_Addr)parse_u64(argv[3]);
    GElf_Addr corrected = expected - 0x1000;

    /* One context proves initial rewrite, two exact reuses, and report_end. */
    Dwfl *dwfl = begin();
    Dwfl_Module *first = report(dwfl, name, path, -1, expected, false);
    Dwfl_Module *reuse_one = report(dwfl, name, path, -1, corrected, false);
    Dwfl_Module *reuse_two = report(dwfl, name, path, -1, corrected, false);
    if (first == NULL || reuse_one == NULL || reuse_two == NULL)
        return 4;
    if (dwfl_report_end(dwfl, NULL, NULL) < 0)
        fprintf(stderr, "dwfl_report_end returned failure\n");
    dwfl_end(dwfl);

    /* A new Dwfl after dwfl_end must not inherit the previous cache entry. */
    dwfl = begin();
    if (report(dwfl, name, path, -1, expected, false) == NULL)
        return 5;
    dwfl_end(dwfl);

    /* Actual negative argument calls use a fresh Dwfl so libdw cannot hide a
       rejected report behind a prior module; the final shim still sees the
       real call and must forward every one unchanged. */
    dwfl = begin();
    char wrong_name[512];
    if (snprintf(wrong_name, sizeof(wrong_name), "%s.wrong", name) >= (int)sizeof(wrong_name))
        return 6;
    (void)report(dwfl, wrong_name, path, -1, expected, false);
    dwfl_end(dwfl);

    dwfl = begin();
    (void)report(dwfl, name, "/tmp/litchi-0976-no-such-target", -1, expected, false);
    dwfl_end(dwfl);

    dwfl = begin();
    (void)report(dwfl, name, path, -1, expected + 0x2000, false);
    dwfl_end(dwfl);

    dwfl = begin();
    int fd = open(path, O_RDONLY | O_CLOEXEC);
    if (fd < 0) {
        perror("open target");
        return 7;
    }
    (void)report(dwfl, name, path, fd, expected, false);
    dwfl_end(dwfl);

    dwfl = begin();
    (void)report(dwfl, name, path, -1, expected, true);
    dwfl_end(dwfl);
    return 0;
}
