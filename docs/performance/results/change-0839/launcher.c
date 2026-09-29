#define _GNU_SOURCE
#include <errno.h>
#include <fcntl.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/resource.h>
#include <sys/wait.h>
#include <unistd.h>

/* Archive-only diagnostic. The fresh ordinary-fork child discards signal
 * maxrss inherited by this launcher across the parent's posix_spawn exec. */
int main(int argc, char **argv)
{
    struct rusage start, child;
    int status;
    if (argc < 4 || strcmp(argv[1], "--usage") != 0 || argv[3][0] != '/') {
        fputs("usage: launcher --usage NEW_FILE /absolute/command [args...]\n", stderr);
        return 2;
    }
    int fd = open(argv[2], O_WRONLY | O_CREAT | O_EXCL | O_CLOEXEC, 0600);
    if (fd < 0) { perror("usage file"); return 2; }
    if (getrusage(RUSAGE_SELF, &start) != 0) { perror("getrusage"); close(fd); return 2; }
    pid_t pid = fork();
    if (pid < 0) { perror("fork"); close(fd); return 2; }
    if (pid == 0) {
        close(fd);
        execv(argv[3], &argv[3]);
        /* Avoid libc teardown of inherited parent state after failed exec. */
        _exit(127);
    }
    pid_t waited;
    do { waited = wait4(pid, &status, 0, &child); } while (waited < 0 && errno == EINTR);
    if (waited != pid) { perror("wait4"); close(fd); return 2; }
    int code = WIFEXITED(status) ? WEXITSTATUS(status) :
               WIFSIGNALED(status) ? 128 + WTERMSIG(status) : 2;
    int written = dprintf(fd,
        "{\"child_pid\":%d,\"wait_status\":%d,\"exit_code\":%d,"
        "\"launcher_start_maxrss_kib\":%ld,\"rusage\":{"
        "\"maxrss_kib\":%ld,\"minor_faults\":%ld,\"major_faults\":%ld}}\n",
        (int)pid, status, code, start.ru_maxrss, child.ru_maxrss,
        child.ru_minflt, child.ru_majflt);
    int closed = close(fd);
    if (written < 0 || closed != 0) { perror("usage write"); return 2; }
    return code;
}
