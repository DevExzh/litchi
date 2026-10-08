#define _GNU_SOURCE

#include <dlfcn.h>
#include <elfutils/libdwfl.h>
#include <ctype.h>
#include <errno.h>
#include <fcntl.h>
#include <inttypes.h>
#include <limits.h>
#include <pthread.h>
#include <stdbool.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/stat.h>
#include <unistd.h>

typedef Dwfl_Module *(*dwfl_report_elf_fn)(Dwfl *, const char *, const char *,
                                            int, GElf_Addr, bool);
typedef const char *(*dwfl_module_info_fn)(Dwfl_Module *, void ***, Dwarf_Addr *,
                                            Dwarf_Addr *, Dwarf_Addr *,
                                            Dwarf_Addr *, const char **,
                                            const char **);
typedef int (*dwfl_errno_fn)(void);
typedef Dwfl_Module *(*dwfl_addrmodule_fn)(Dwfl *, Dwarf_Addr);
typedef int (*dwfl_module_build_id_fn)(Dwfl_Module *, const unsigned char **,
                                       GElf_Addr *);
typedef Elf *(*dwfl_module_getelf_fn)(Dwfl_Module *, GElf_Addr *);
typedef void (*dwfl_report_begin_fn)(Dwfl *);
typedef int (*dwfl_report_end_fn)(Dwfl *, int (*)(Dwfl_Module *, void *,
                                                   const char *, Dwarf_Addr,
                                                   void *), void *);
typedef void (*dwfl_end_fn)(Dwfl *);

#define CACHE_CAP 8
#define LOW_PROBE_OFFSET ((GElf_Addr)0x1000)
/* The retained ELF layout's executable PT_LOAD starts at vaddr 0xe5000. */
#define EXEC_PROBE_OFFSET ((GElf_Addr)0xe5a00)
/* Exact retained target build ID; a merely well-formed caller value is not
 * sufficient to authorize the base rewrite. */
static const unsigned char REQUIRED_TARGET_BUILD_ID[20] = {
    0xe5, 0xc4, 0x57, 0x38, 0x4e, 0xee, 0xbd, 0x05,
    0x79, 0x32, 0xc4, 0x44, 0xef, 0xda, 0x77, 0x61,
    0x8f, 0x13, 0x73, 0x99,
};

typedef struct {
    Dwfl *dwfl;
    Dwfl_Module *module;
    GElf_Addr expected_base;
    GElf_Addr corrected_base;
    int fd;
    bool add_p_vaddr;
    dev_t st_dev;
    ino_t st_ino;
    off_t st_size;
    char name[PATH_MAX];
    char file_name[PATH_MAX];
    unsigned char build_id[64];
    size_t build_id_len;
    char target_sha256[65];
    Dwarf_Addr module_start;
    Dwarf_Addr module_end;
    Dwfl_Module *probe_low;
    Dwfl_Module *probe_exec;
    int range_end_known;
    int identity_verified;
    int valid;
} target_cache_entry;

static target_cache_entry target_cache[CACHE_CAP];
static pthread_mutex_t cache_mutex = PTHREAD_MUTEX_INITIALIZER;
static unsigned long long report_epoch;

static dwfl_report_elf_fn real_report;
static dwfl_module_info_fn real_info;
static dwfl_errno_fn real_errno;
static dwfl_addrmodule_fn real_addrmodule;
static dwfl_module_build_id_fn real_build_id;
static dwfl_module_getelf_fn real_getelf;
static dwfl_report_begin_fn real_begin;
static dwfl_report_begin_fn real_begin_add;
static dwfl_report_end_fn real_end;
static dwfl_end_fn real_end_session;
static int log_fd = -1;
static __thread int resolving;
/* dwfl_module_getelf may call back through dwfl_module_info. */
static __thread int identity_probe;

static void *resolve_symbol(const char *name, const char *version) {
    void *value = dlvsym(RTLD_NEXT, name, version);
    if (value == NULL)
        value = dlsym(RTLD_NEXT, name);
    return value;
}

#define RESOLVER(fn_type, fn_var, symbol, version)                         \
    static fn_type resolve_##fn_var(void) {                                \
        if (fn_var != NULL)                                                 \
            return fn_var;                                                  \
        if (resolving)                                                      \
            return NULL;                                                    \
        resolving = 1;                                                      \
        fn_var = (fn_type)resolve_symbol(symbol, version);                 \
        resolving = 0;                                                      \
        return fn_var;                                                       \
    }

RESOLVER(dwfl_report_elf_fn, real_report, "dwfl_report_elf", "ELFUTILS_0.156")
RESOLVER(dwfl_module_info_fn, real_info, "dwfl_module_info", "ELFUTILS_0.122")
RESOLVER(dwfl_errno_fn, real_errno, "dwfl_errno", "ELFUTILS_0.122")
RESOLVER(dwfl_addrmodule_fn, real_addrmodule, "dwfl_addrmodule", "ELFUTILS_0.122")
RESOLVER(dwfl_module_build_id_fn, real_build_id, "dwfl_module_build_id", "ELFUTILS_0.138")
RESOLVER(dwfl_module_getelf_fn, real_getelf, "dwfl_module_getelf", "ELFUTILS_0.122")
RESOLVER(dwfl_report_begin_fn, real_begin, "dwfl_report_begin", "ELFUTILS_0.122")
RESOLVER(dwfl_report_begin_fn, real_begin_add, "dwfl_report_begin_add", "ELFUTILS_0.127")
RESOLVER(dwfl_report_end_fn, real_end, "dwfl_report_end", "ELFUTILS_0.122")
RESOLVER(dwfl_end_fn, real_end_session, "dwfl_end", "ELFUTILS_0.122")

static const char *mode_name(void) {
    const char *mode = getenv("LITCHI_0973_SHIM_MODE");
    if (mode != NULL && (strcmp(mode, "noop") == 0 || strcmp(mode, "corrected") == 0))
        return mode;
    return "invalid";
}

static const char *configuration(const char *key) {
    const char *value = getenv(key);
    return value != NULL ? value : "";
}

static int hex_value(char c) {
    if (c >= '0' && c <= '9')
        return c - '0';
    if (c >= 'a' && c <= 'f')
        return c - 'a' + 10;
    if (c >= 'A' && c <= 'F')
        return c - 'A' + 10;
    return -1;
}

/* Reject signs and whitespace: strtoull otherwise accepts both. */
static bool parse_address(const char *key, GElf_Addr *out) {
    const char *text = configuration(key);
    if (*text == '\0' || *text == '-' || *text == '+' || isspace((unsigned char)*text))
        return false;
    char *end = NULL;
    errno = 0;
    unsigned long long value = strtoull(text, &end, 0);
    if (errno != 0 || end == text || *end != '\0' ||
        value > (unsigned long long)UINT64_MAX)
        return false;
    *out = (GElf_Addr)value;
    return true;
}

static bool parse_build_id(unsigned char *out, size_t *out_len) {
    const char *text = configuration("LITCHI_0976_TARGET_BUILD_ID");
    size_t length = strlen(text);
    /* The retained target has a 20-byte GNU build ID; require that exact form. */
    if (length != 40)
        return false;
    for (size_t i = 0; i < length / 2; ++i) {
        int high = hex_value(text[2 * i]);
        int low = hex_value(text[2 * i + 1]);
        if (high < 0 || low < 0)
            return false;
        out[i] = (unsigned char)((high << 4) | low);
    }
    *out_len = length / 2;
    return true;
}

static bool parse_sha256(char out[65]) {
    const char *text = configuration("LITCHI_0976_TARGET_SHA256");
    if (strlen(text) != 64)
        return false;
    for (size_t i = 0; i < 64; ++i) {
        if (hex_value(text[i]) < 0)
            return false;
        out[i] = (char)tolower((unsigned char)text[i]);
    }
    out[64] = '\0';
    return true;
}

/* LITCHI_0976_TARGET_SHA256 is an external custody binding.  The offline
 * runner hashes the immutable target before and after replay; this shim stores
 * the exact expected digest but deliberately does not shell out or claim to
 * hash the file inside libdw's report callback. */
static bool config_valid(GElf_Addr *corrected_base, GElf_Addr *perf_expected_start,
                         unsigned char *build_id, size_t *build_id_len,
                         char target_sha256[65]) {
    const char *target_path = configuration("LITCHI_0973_TARGET_PATH");
    const char *target_name = configuration("LITCHI_0973_TARGET_NAME");
    GElf_Addr delta = 0;
    if (*target_path == '\0' || *target_name == '\0' ||
        strlen(target_path) >= PATH_MAX || strlen(target_name) >= PATH_MAX ||
        !parse_build_id(build_id, build_id_len) ||
        *build_id_len != sizeof(REQUIRED_TARGET_BUILD_ID) ||
        memcmp(build_id, REQUIRED_TARGET_BUILD_ID, sizeof(REQUIRED_TARGET_BUILD_ID)) != 0 ||
        !parse_sha256(target_sha256) ||
        !parse_address("LITCHI_0973_CORRECTED_BASE", corrected_base) ||
        !parse_address("LITCHI_0973_PERF_EXPECTED_START", perf_expected_start) ||
        !parse_address("LITCHI_0976_EXPECTED_DELTA", &delta))
        return false;
    if (*corrected_base == 0 || *perf_expected_start == 0 ||
        *perf_expected_start <= *corrected_base || delta == 0 ||
        *perf_expected_start - *corrected_base != delta ||
        (*corrected_base & 0xfff) != 0 || (*perf_expected_start & 0xfff) != 0 ||
        (delta & 0xfff) != 0 ||
        *corrected_base > (GElf_Addr)-1 - EXEC_PROBE_OFFSET)
        return false;
    return true;
}

static int output_fd(void) {
    if (log_fd >= 0)
        return log_fd;
    const char *path = getenv("LITCHI_0973_SHIM_LOG");
    if (path == NULL || *path == '\0')
        return -1;
    int fd = open(path, O_WRONLY | O_CREAT | O_APPEND | O_CLOEXEC, 0644);
    if (fd >= 0)
        log_fd = fd;
    return fd;
}

static int caller_is_perf_from_pc(void *pc, const char **caller_file,
                                  const char **caller_symbol) {
    Dl_info info;
    memset(&info, 0, sizeof(info));
    int found = dladdr(pc, &info);
    if (caller_file != NULL)
        *caller_file = found && info.dli_fname != NULL ? info.dli_fname : "";
    if (caller_symbol != NULL)
        *caller_symbol = found && info.dli_sname != NULL ? info.dli_sname : "";
    return found && info.dli_fname != NULL &&
           strcmp(info.dli_fname, "/usr/bin/perf") == 0;
}

static target_cache_entry *lookup_locked(Dwfl *dwfl) {
    for (size_t i = 0; i < CACHE_CAP; ++i) {
        if (target_cache[i].valid && target_cache[i].dwfl == dwfl)
            return &target_cache[i];
    }
    return NULL;
}

static target_cache_entry *free_entry_locked(void) {
    for (size_t i = 0; i < CACHE_CAP; ++i) {
        if (!target_cache[i].valid)
            return &target_cache[i];
    }
    return NULL;
}

static void clear_cache_locked(Dwfl *dwfl) {
    for (size_t i = 0; i < CACHE_CAP; ++i) {
        if (target_cache[i].valid && target_cache[i].dwfl == dwfl)
            target_cache[i].valid = 0;
    }
}

static bool stat_identity(const char *path, dev_t *dev, ino_t *ino, off_t *size) {
    struct stat st;
    if (stat(path, &st) != 0)
        return false;
    *dev = st.st_dev;
    *ino = st.st_ino;
    *size = st.st_size;
    return true;
}

static bool module_build_id_matches(Dwfl_Module *module,
                                    const unsigned char *want, size_t want_len) {
    dwfl_module_build_id_fn get_id = resolve_real_build_id();
    if (get_id == NULL || module == NULL || want == NULL || want_len == 0)
        return false;
    const unsigned char *bits = NULL;
    GElf_Addr vaddr = 0;
    int length = get_id(module, &bits, &vaddr);
    return length == (int)want_len && bits != NULL &&
           memcmp(bits, want, want_len) == 0;
}

static bool cache_module_name_matches(const target_cache_entry *entry,
                                      const char *name,
                                      const char *file_name) {
    return entry != NULL && name != NULL && file_name != NULL &&
           strcmp(entry->name, name) == 0 &&
           strcmp(entry->file_name, file_name) == 0;
}

/*
 * Revalidate the live libdw module while the cache lock is held.  getelf is
 * deliberate: libdw documents that module_info/build_id may be incomplete
 * until the ELF has been opened.  The thread flag prevents a getelf callback
 * through our info interposer from taking cache_mutex recursively.
 */
static bool validate_cached_module_locked(target_cache_entry *entry,
                                          Dwfl *dwfl,
                                          const unsigned char *build_id,
                                          size_t build_id_len,
                                          const char target_sha256[65],
                                          Dwfl_Module **probe_low,
                                          Dwfl_Module **probe_exec) {
    if (entry == NULL || !entry->valid || entry->dwfl != dwfl ||
        entry->module == NULL || entry->fd != -1 || entry->add_p_vaddr ||
        build_id == NULL || build_id_len == 0 || target_sha256 == NULL ||
        strcmp(entry->target_sha256, target_sha256) != 0)
        return false;

    dev_t dev = 0;
    ino_t ino = 0;
    off_t size = 0;
    if (!stat_identity(entry->file_name, &dev, &ino, &size) ||
        dev != entry->st_dev || ino != entry->st_ino || size != entry->st_size)
        return false;

    dwfl_module_getelf_fn getelf = resolve_real_getelf();
    dwfl_module_info_fn get_info = resolve_real_info();
    dwfl_addrmodule_fn addrmodule = resolve_real_addrmodule();
    if (getelf == NULL || get_info == NULL || addrmodule == NULL)
        return false;

    GElf_Addr elf_bias = 0;
    Dwarf_Addr start = 0, end = 0, dwbias = 0, symbias = 0;
    const char *mainfile = NULL;
    const char *debugfile = NULL;
    identity_probe++;
    Elf *elf = getelf(entry->module, &elf_bias);
    const char *module_name = elf != NULL
        ? get_info(entry->module, NULL, &start, &end, &dwbias, &symbias,
                   &mainfile, &debugfile)
        : NULL;
    identity_probe--;
    if (elf == NULL || module_name == NULL ||
        strcmp(module_name, entry->name) != 0 || mainfile == NULL ||
        strcmp(mainfile, entry->file_name) != 0 || start != entry->corrected_base ||
        (end != 0 && end <= start) ||
        !module_build_id_matches(entry->module, build_id, build_id_len))
        return false;

    /* Two addresses cover both the first page and the retained executable PT_LOAD. */
    Dwfl_Module *low = addrmodule(dwfl, entry->corrected_base + LOW_PROBE_OFFSET);
    Dwfl_Module *exec = addrmodule(dwfl, entry->corrected_base + EXEC_PROBE_OFFSET);
    if (low != entry->module || exec != entry->module)
        return false;
    if (end != 0 &&
        (entry->corrected_base + LOW_PROBE_OFFSET < start ||
         entry->corrected_base + EXEC_PROBE_OFFSET >= end))
        return false;

    entry->module_start = start;
    entry->module_end = end;
    entry->probe_low = low;
    entry->probe_exec = exec;
    entry->range_end_known = end != 0;
    entry->identity_verified = 1;
    if (probe_low != NULL)
        *probe_low = low;
    if (probe_exec != NULL)
        *probe_exec = exec;
    return true;
}

/* Recheck the live module for each info call before changing its base.  An
 * empty *mainfile is treated as "not reported by this libdw query", never as
 * proof of identity; stat/build-ID and the stored getelf identity still have
 * to match. */
static bool validate_info_identity_locked(const target_cache_entry *entry,
                                         Dwfl *dwfl, const char *name,
                                         const char **mainfile,
                                         Dwarf_Addr *start, Dwarf_Addr *end,
                                         const unsigned char *build_id,
                                         size_t build_id_len) {
    if (entry == NULL || !entry->valid || !entry->identity_verified ||
        entry->dwfl != dwfl || entry->module == NULL || name == NULL ||
        strcmp(name, entry->name) != 0 || mainfile == NULL ||
        (entry->probe_low != entry->module || entry->probe_exec != entry->module) ||
        start == NULL || *start != entry->corrected_base ||
        (end != NULL && *end != 0 && *end <= *start))
        return false;
    if (*mainfile != NULL && **mainfile != '\0' &&
        strcmp(*mainfile, entry->file_name) != 0)
        return false;
    dev_t dev = 0;
    ino_t ino = 0;
    off_t size = 0;
    return stat_identity(entry->file_name, &dev, &ino, &size) &&
           dev == entry->st_dev && ino == entry->st_ino && size == entry->st_size &&
           module_build_id_matches(entry->module, build_id, build_id_len);
}

static bool exact_module_identity_locked(target_cache_entry *entry, Dwfl *dwfl,
                                         const char *name, const char *file_name,
                                         int fd, bool add_p_vaddr, GElf_Addr base,
                                         GElf_Addr expected, GElf_Addr corrected,
                                         int caller_perf, const char *caller_file,
                                         const unsigned char *build_id,
                                         size_t build_id_len,
                                         const char target_sha256[65],
                                         Dwfl_Module **probe_low,
                                         Dwfl_Module **probe_exec) {
    if (entry == NULL || !entry->valid || entry->dwfl != dwfl ||
        entry->fd != fd || entry->add_p_vaddr != add_p_vaddr || fd != -1 ||
        add_p_vaddr || entry->expected_base != expected ||
        entry->corrected_base != corrected ||
        (base != expected && base != corrected) || !caller_perf ||
        caller_file == NULL || strcmp(caller_file, "/usr/bin/perf") != 0 ||
        !cache_module_name_matches(entry, name, file_name))
        return false;
    bool valid = validate_cached_module_locked(entry, dwfl, build_id, build_id_len,
                                               target_sha256, probe_low,
                                               probe_exec);
    if (!valid) {
        /* Never let a failed reuse validation authorize later info adjustment. */
        entry->identity_verified = 0;
        entry->valid = 0;
    }
    return valid;
}

static bool prepare_cache_entry_locked(target_cache_entry *entry, Dwfl *dwfl,
                                       Dwfl_Module *module, const char *name,
                                       const char *file_name, int fd,
                                       bool add_p_vaddr, GElf_Addr expected,
                                       GElf_Addr corrected,
                                       const unsigned char *build_id,
                                       size_t build_id_len,
                                       const char target_sha256[65]) {
    if (entry == NULL || module == NULL || name == NULL || file_name == NULL ||
        fd != -1 || add_p_vaddr || build_id == NULL || build_id_len == 0 ||
        target_sha256 == NULL || lookup_locked(dwfl) != NULL)
        return false;
    if (strcmp(name, configuration("LITCHI_0973_TARGET_NAME")) != 0 ||
        strcmp(file_name, configuration("LITCHI_0973_TARGET_PATH")) != 0)
        return false;

    memset(entry, 0, sizeof(*entry));
    entry->dwfl = dwfl;
    entry->module = module;
    entry->expected_base = expected;
    entry->corrected_base = corrected;
    entry->fd = fd;
    entry->add_p_vaddr = add_p_vaddr;
    entry->build_id_len = build_id_len;
    memcpy(entry->build_id, build_id, build_id_len);
    memcpy(entry->target_sha256, target_sha256, 65);
    strncpy(entry->name, name, sizeof(entry->name) - 1);
    strncpy(entry->file_name, file_name, sizeof(entry->file_name) - 1);
    if (!stat_identity(file_name, &entry->st_dev, &entry->st_ino,
                       &entry->st_size))
        return false;
    entry->valid = 1;
    if (!validate_cached_module_locked(entry, dwfl, build_id, build_id_len,
                                       target_sha256, NULL, NULL)) {
        entry->valid = 0;
        return false;
    }
    return true;
}

static void log_report(const char *mode, const char *name, const char *file_name,
                       Dwfl *dwfl, int fd, GElf_Addr original,
                       GElf_Addr forwarded, bool add_p_vaddr, int path_match,
                       int name_match, int admission, int initial_rewrite,
                       int reused_existing, int identity_valid, int cache_stored,
                       int caller_perf, const char *caller_file,
                       const char *caller_symbol, void *caller_pc,
                       Dwfl_Module *result, Dwfl_Module *probe_low,
                       Dwfl_Module *probe_exec, int saved_errno, int dwfl_error,
                       int config_ok) {
    int out = output_fd();
    if (out < 0)
        return;
    dprintf(out,
            "{\"schema\":\"litchi.performance.0976.dwfl_report.final\","
            "\"mode\":\"%s\",\"epoch\":%llu,\"name\":\"%s\","
            "\"file_name\":\"%s\",\"dwfl\":\"%p\",\"fd\":%d,"
            "\"path_match\":%s,\"name_match\":%s,\"config_valid\":%s,"
            "\"admission\":%s,\"initial_rewrite\":%s,"
            "\"original_base\":\"0x%llx\",\"forwarded_base\":\"0x%llx\","
            "\"add_p_vaddr\":%s,\"reused_existing\":%s,"
            "\"identity_valid\":%s,\"cache_stored\":%s,"
            "\"caller_is_perf\":%s,\"caller_file\":\"%s\","
            "\"caller_symbol\":\"%s\",\"caller_pc\":\"%p\","
            "\"result_module\":\"%p\",\"probe_low\":\"%p\","
            "\"probe_exec\":\"%p\",\"errno_after\":%d,"
            "\"dwfl_errno_after\":%d}\n",
            mode, report_epoch, name != NULL ? name : "",
            file_name != NULL ? file_name : "", (void *)dwfl, fd,
            path_match ? "true" : "false", name_match ? "true" : "false",
            config_ok ? "true" : "false", admission ? "true" : "false",
            initial_rewrite ? "true" : "false",
            (unsigned long long)original, (unsigned long long)forwarded,
            add_p_vaddr ? "true" : "false", reused_existing ? "true" : "false",
            identity_valid ? "true" : "false", cache_stored ? "true" : "false",
            caller_perf ? "true" : "false", caller_file != NULL ? caller_file : "",
            caller_symbol != NULL ? caller_symbol : "", caller_pc,
            (void *)result, (void *)probe_low, (void *)probe_exec,
            saved_errno, dwfl_error);
}

static void log_info(const char *mode, const char *module_name,
                     const char *mainfile, const char *debugfile,
                     Dwarf_Addr before_start, Dwarf_Addr after_start,
                     Dwarf_Addr end, Dwarf_Addr dwbias, Dwarf_Addr symbias,
                     int target_match, int caller_perf, int cache_bound,
                     int current_path_valid, int current_range_valid,
                     const char *caller_file, const char *caller_symbol,
                     void *caller_pc, int adjusted, int config_ok) {
    int out = output_fd();
    if (out < 0)
        return;
    dprintf(out,
            "{\"schema\":\"litchi.performance.0976.dwfl_info.final\","
            "\"mode\":\"%s\",\"module_name\":\"%s\","
            "\"mainfile\":\"%s\",\"debugfile\":\"%s\","
            "\"before_start\":\"0x%llx\",\"after_start\":\"0x%llx\","
            "\"end\":\"0x%llx\",\"dwbias\":\"0x%llx\","
            "\"symbias\":\"0x%llx\",\"target_match\":%s,"
            "\"caller_is_perf\":%s,\"cache_bound\":%s,"
            "\"current_path_valid\":%s,\"current_range_valid\":%s,"
            "\"caller_file\":\"%s\",\"caller_symbol\":\"%s\","
            "\"caller_pc\":\"%p\",\"start_adjusted\":%s,"
            "\"config_valid\":%s}\n",
            mode, module_name != NULL ? module_name : "",
            mainfile != NULL ? mainfile : "", debugfile != NULL ? debugfile : "",
            (unsigned long long)before_start, (unsigned long long)after_start,
            (unsigned long long)end, (unsigned long long)dwbias,
            (unsigned long long)symbias, target_match ? "true" : "false",
            caller_perf ? "true" : "false", cache_bound ? "true" : "false",
            current_path_valid ? "true" : "false",
            current_range_valid ? "true" : "false",
            caller_file != NULL ? caller_file : "",
            caller_symbol != NULL ? caller_symbol : "", caller_pc,
            adjusted ? "true" : "false", config_ok ? "true" : "false");
}

void dwfl_report_begin(Dwfl *dwfl) {
    dwfl_report_begin_fn next = resolve_real_begin();
    pthread_mutex_lock(&cache_mutex);
    clear_cache_locked(dwfl);
    ++report_epoch;
    pthread_mutex_unlock(&cache_mutex);
    if (next != NULL)
        next(dwfl);
}

void dwfl_report_begin_add(Dwfl *dwfl) {
    dwfl_report_begin_fn next = resolve_real_begin_add();
    pthread_mutex_lock(&cache_mutex);
    clear_cache_locked(dwfl);
    ++report_epoch;
    pthread_mutex_unlock(&cache_mutex);
    if (next != NULL)
        next(dwfl);
}

int dwfl_report_end(Dwfl *dwfl, int (*removed)(Dwfl_Module *, void *, const char *,
                                               Dwarf_Addr, void *), void *arg) {
    dwfl_report_end_fn next = resolve_real_end();
    int result = next != NULL ? next(dwfl, removed, arg) : -1;
    pthread_mutex_lock(&cache_mutex);
    clear_cache_locked(dwfl);
    pthread_mutex_unlock(&cache_mutex);
    return result;
}

void dwfl_end(Dwfl *dwfl) {
    dwfl_end_fn next = resolve_real_end_session();
    pthread_mutex_lock(&cache_mutex);
    clear_cache_locked(dwfl);
    pthread_mutex_unlock(&cache_mutex);
    if (next != NULL)
        next(dwfl);
}

Dwfl_Module *dwfl_report_elf(Dwfl *dwfl, const char *name, const char *file_name,
                             int fd, GElf_Addr base, bool add_p_vaddr) {
    dwfl_report_elf_fn next = resolve_real_report();
    if (next == NULL)
        return NULL;

    int saved_errno = errno;
    GElf_Addr corrected_base = 0;
    GElf_Addr perf_expected_start = 0;
    unsigned char build_id[64] = {0};
    size_t build_id_len = 0;
    char target_sha256[65] = {0};
    int config_ok = config_valid(&corrected_base, &perf_expected_start,
                                 build_id, &build_id_len, target_sha256);
    const char *target_path = configuration("LITCHI_0973_TARGET_PATH");
    const char *target_name = configuration("LITCHI_0973_TARGET_NAME");
    int path_match = *target_path != '\0' && file_name != NULL &&
                     strcmp(file_name, target_path) == 0;
    int name_match = *target_name != '\0' && name != NULL &&
                     strcmp(name, target_name) == 0;
    const char *mode = mode_name();
    void *caller_pc = __builtin_return_address(0);
    const char *caller_file = "";
    const char *caller_symbol = "";
    int caller_perf = caller_is_perf_from_pc(caller_pc, &caller_file,
                                             &caller_symbol);
    int base_allowed = config_ok &&
                       (base == perf_expected_start || base == corrected_base);
    /* This is the only admission predicate. Every other caller delegates. */
    int admission = config_ok && strcmp(mode, "corrected") == 0 &&
                    path_match && name_match && caller_perf &&
                    caller_file != NULL && strcmp(caller_file, "/usr/bin/perf") == 0 &&
                    fd == -1 && !add_p_vaddr && base_allowed;
    int initial_rewrite = admission && base == perf_expected_start;
    GElf_Addr forwarded = admission ? corrected_base : base;
    Dwfl_Module *result = NULL;
    Dwfl_Module *probe_low = NULL;
    Dwfl_Module *probe_exec = NULL;
    int reused_existing = 0;
    int identity_valid = 0;
    int cache_stored = 0;
    int locked = admission;

    if (locked)
        pthread_mutex_lock(&cache_mutex);
    target_cache_entry *entry = locked ? lookup_locked(dwfl) : NULL;
    if (locked && exact_module_identity_locked(
            entry, dwfl, name, file_name, fd, add_p_vaddr, base,
            perf_expected_start, corrected_base, caller_perf, caller_file,
            build_id, build_id_len, target_sha256, &probe_low, &probe_exec)) {
        identity_valid = 1;
        reused_existing = 1;
        result = entry->module;
        log_report(mode, name, file_name, dwfl, fd, base, forwarded, add_p_vaddr,
                   path_match, name_match, admission, initial_rewrite,
                   reused_existing, identity_valid, cache_stored, caller_perf,
                   caller_file, caller_symbol, caller_pc, result, probe_low,
                   probe_exec, saved_errno, -1, config_ok);
        pthread_mutex_unlock(&cache_mutex);
        errno = saved_errno;
        return result;
    }

    result = next(dwfl, name, file_name, fd, forwarded, add_p_vaddr);
    saved_errno = errno;
    int error = -1;
    dwfl_errno_fn get_errno = resolve_real_errno();
    if (get_errno != NULL)
        error = get_errno();
    /* A corrected-base call never creates a cache entry.  It may reuse only. */
    if (initial_rewrite && result != NULL) {
        target_cache_entry *free_entry = free_entry_locked();
        if (free_entry != NULL && lookup_locked(dwfl) == NULL) {
            cache_stored = prepare_cache_entry_locked(
                free_entry, dwfl, result, name, file_name, fd, add_p_vaddr,
                perf_expected_start, corrected_base, build_id, build_id_len,
                target_sha256);
            identity_valid = cache_stored;
            if (cache_stored) {
                probe_low = free_entry->probe_low;
                probe_exec = free_entry->probe_exec;
            }
        }
    }
    log_report(mode, name, file_name, dwfl, fd, base, forwarded, add_p_vaddr,
               path_match, name_match, admission, initial_rewrite,
               reused_existing, identity_valid, cache_stored, caller_perf,
               caller_file, caller_symbol, caller_pc, result, probe_low,
               probe_exec, saved_errno, error, config_ok);
    if (locked)
        pthread_mutex_unlock(&cache_mutex);
    errno = saved_errno;
    return result;
}

const char *dwfl_module_info(Dwfl_Module *mod, void ***userdata,
                             Dwarf_Addr *start, Dwarf_Addr *end,
                             Dwarf_Addr *dwbias, Dwarf_Addr *symbias,
                             const char **mainfile, const char **debugfile) {
    dwfl_module_info_fn next = resolve_real_info();
    if (next == NULL)
        return NULL;
    const char *name = next(mod, userdata, start, end, dwbias, symbias,
                            mainfile, debugfile);
    int saved_errno = errno;
    /* getelf can reach here while the cache lock is held. */
    if (identity_probe) {
        errno = saved_errno;
        return name;
    }

    Dwarf_Addr before = start != NULL ? *start : 0;
    Dwarf_Addr after = before;
    const char *caller_file = "";
    const char *caller_symbol = "";
    void *caller_pc = __builtin_return_address(0);
    int direct_perf = caller_is_perf_from_pc(caller_pc, &caller_file,
                                             &caller_symbol);
    const char *target_name = configuration("LITCHI_0973_TARGET_NAME");
    const char *target_path = configuration("LITCHI_0973_TARGET_PATH");
    int target_match = *target_name != '\0' && name != NULL &&
                       strcmp(name, target_name) == 0;
    GElf_Addr corrected_base = 0;
    GElf_Addr perf_expected_start = 0;
    unsigned char build_id[64] = {0};
    size_t build_id_len = 0;
    char target_sha256[65] = {0};
    int config_ok = config_valid(&corrected_base, &perf_expected_start,
                                 build_id, &build_id_len, target_sha256);
    int current_path_valid = mainfile != NULL &&
        (*mainfile == NULL || **mainfile == '\0' ||
         strcmp(*mainfile, target_path) == 0);
    int current_range_valid = true;
    if (start != NULL && *start != 0) {
        current_range_valid = *start == corrected_base;
        if (end != NULL && *end != 0)
            current_range_valid = current_range_valid && *end > *start;
    }
    if (start != NULL && *start == 0 && end != NULL && *end != 0)
        current_range_valid = false;

    int cache_bound = 0;
    int info_candidate = config_ok && target_match && direct_perf &&
                         current_path_valid && current_range_valid &&
                         start != NULL && *start == corrected_base &&
                         strcmp(mode_name(), "corrected") == 0;
    pthread_mutex_lock(&cache_mutex);
    if (info_candidate) {
        for (size_t i = 0; i < CACHE_CAP; ++i) {
            target_cache_entry *entry = &target_cache[i];
            if (!entry->valid || entry->module != mod ||
                !cache_module_name_matches(entry, target_name, target_path) ||
                entry->fd != -1 || entry->add_p_vaddr ||
                entry->expected_base != perf_expected_start ||
                entry->corrected_base != corrected_base ||
                strcmp(entry->target_sha256, target_sha256) != 0)
                continue;
            if (validate_info_identity_locked(entry, entry->dwfl, name, mainfile,
                                              start, end, build_id, build_id_len)) {
                cache_bound = 1;
                break;
            }
            /* A current identity failure must revoke adjustment authority. */
            entry->identity_verified = 0;
            entry->valid = 0;
        }
    }
    if (info_candidate && cache_bound) {
        *start = perf_expected_start;
        after = *start;
    }
    int adjusted = after != before;
    pthread_mutex_unlock(&cache_mutex);
    if (start != NULL)
        after = *start;
    log_info(mode_name(), name, mainfile != NULL ? *mainfile : NULL,
             debugfile != NULL ? *debugfile : NULL, before, after,
             end != NULL ? *end : 0, dwbias != NULL ? *dwbias : 0,
             symbias != NULL ? *symbias : 0, target_match, direct_perf,
             cache_bound, current_path_valid, current_range_valid, caller_file,
             caller_symbol, caller_pc, adjusted, config_ok);
    errno = saved_errno;
    return name;
}

