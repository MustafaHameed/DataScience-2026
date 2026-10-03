/* ch03_popf.c -- show that POPF fails silently in user mode.
   Build: gcc -O0 -o ch03_popf ch03_popf.c     Run: ./ch03_popf   */
#include <stdio.h>

/* --- 1. Read EFLAGS, so we can see what the processor really did. */
static unsigned long read_flags(void) {
    unsigned long f;
    __asm__ volatile ("pushfq ; popq %0" : "=r"(f));
    return f;
}

/* --- 2. Try to write EFLAGS, clearing the interrupt flag (bit 9). */
static void write_flags(unsigned long f) {
    __asm__ volatile ("pushq %0 ; popfq" : : "r"(f) : "cc");
}

int main(void) {
    unsigned long before = read_flags();
    printf("before : EFLAGS=0x%08lx  IF=%lu\n", before, (before >> 9) & 1);

    /* --- 3. Ask for interrupts off, and for the carry flag on.
       A kernel entering a critical section does exactly this. */
    unsigned long want = (before & ~(1UL << 9)) | 1UL;   /* IF=0, CF=1 */
    write_flags(want);

    /* --- 4. Look at what we actually got. */
    unsigned long after = read_flags();
    printf("wanted : EFLAGS=0x%08lx  IF=%lu  CF=%lu\n",
           want, (want >> 9) & 1, want & 1);
    printf("after  : EFLAGS=0x%08lx  IF=%lu  CF=%lu\n",
           after, (after >> 9) & 1, after & 1);

    /* --- 5. The verdict. No signal was raised; nothing faulted. */
    printf("\ncarry flag taken?     %s\n",
           (after & 1) ? "yes -- the instruction 'worked'" : "no");
    printf("interrupt flag taken? %s\n",
           (((after >> 9) & 1) == 0) ? "yes" :
           "NO -- written, ignored, and no trap. This is the bug.");
    return 0;
}
