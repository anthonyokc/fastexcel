// Forward routine registration from C to Rust so the linker keeps the static library.
void R_init_fastexcel_extendr(void *dll);

void R_init_fastexcel(void *dll) {
    R_init_fastexcel_extendr(dll);
}
