var c
var p
var ch

set_pgstat_mode 2
set_max_bandwidth 400
set_range ba 100u
set_autoranging ba 10u 1m

mux_config 4i 0x0000

cell_on

# loop ch = 1..16
store_var ch 1i ea
loop ch <= 16i

    mux_set_channel ch
    # optional settle time
    wait 100m

    meas_loop_lsv p c -1 1 10m 1
        pck_start
            pck_add ch
            pck_add p
            pck_add c
        pck_end
    endloop

    add_var ch 1i
endloop

on_finished:
cell_off

