e
set_gpio_cfg 0x3FFi 1

var p
var c
var h

set_pgstat_chan 0
set_pgstat_mode 2
set_max_bandwidth 400
set_pot_range -1 1
set_cr 1m
set_autoranging 10u 1m

cell_on

# --- CH 1 ---
set_gpio 0x00i
store_var h 1i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 2 ---
set_gpio 0x11i
store_var h 2i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 3 ---
set_gpio 0x22i
store_var h 3i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 4 ---
set_gpio 0x33i
store_var h 4i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 5 ---
set_gpio 0x44i
store_var h 5i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 6 ---
set_gpio 0x55i
store_var h 6i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 7 ---
set_gpio 0x66i
store_var h 7i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 8 ---
set_gpio 0x77i
store_var h 8i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 9 ---
set_gpio 0x88i
store_var h 9i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 10 ---
set_gpio 0x99i
store_var h 10i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 11 ---
set_gpio 0xAAi
store_var h 11i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 12 ---
set_gpio 0xBBi
store_var h 12i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 13 ---
set_gpio 0xCCi
store_var h 13i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 14 ---
set_gpio 0xDDi
store_var h 14i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 15 ---
set_gpio 0xEEi
store_var h 15i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

# --- CH 16 ---
set_gpio 0xFFi
store_var h 16i ea
set_e -1000m
wait 100m
meas_loop_lsv p c -1 1 10m 1
    pck_start
        pck_add h
        pck_add p
        pck_add c
    pck_end
endloop

on_finished:
cell_off
