var c
var p
var h

set_pgstat_mode 2
set_max_bandwidth 40
set_range ba 100u
set_autoranging ba 1n 100u

mux_config 4i 0x0000i
cell_on

store_var h 1i ea
mux_set_channel 1i
meas_loop_cv p c 0 -500m 500m 2m 100m
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

on_finished:
cell_off

