var c
var p
var ch

set_pgstat_mode 2
set_max_bandwidth 40
set_range ba 100u
set_autoranging ba 1n 100u

mux_config 4i 0x0000
cell_on

# channel 1
store_var ch 1i ea
mux_set_channel 1i
meas_loop_cv p c 0 -500m 500m 2m 100m
    pck_start
        pck_add ch
        pck_add p
        pck_add c
    pck_end
endloop

# channel 2
store_var ch 2i ea
mux_set_channel 2i
meas_loop_cv p c 0 -500m 500m 2m 100m
    pck_start
        pck_add ch
        pck_add p
        pck_add c
    pck_end
endloop

on_finished:
cell_off
