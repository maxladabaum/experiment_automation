e
var c
var p
var f
var r
var ch
set_pgstat_mode 2
set_max_bandwidth 1600
set_range_minmax da -520m 520m
set_range ba 5m
set_autoranging ba 100n 5m
mux_config 4i 0x0000
cell_on
# --- Channel 1 ---
mux_set_channel 1i
store_var ch 1i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 2 ---
mux_set_channel 2i
store_var ch 2i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 3 ---
mux_set_channel 3i
store_var ch 3i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 4 ---
mux_set_channel 4i
store_var ch 4i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 5 ---
mux_set_channel 5i
store_var ch 5i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 6 ---
mux_set_channel 6i
store_var ch 6i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 7 ---
mux_set_channel 7i
store_var ch 7i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 8 ---
mux_set_channel 8i
store_var ch 8i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 9 ---
mux_set_channel 9i
store_var ch 9i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 10 ---
mux_set_channel 10i
store_var ch 10i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 11 ---
mux_set_channel 11i
store_var ch 11i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 12 ---
mux_set_channel 12i
store_var ch 12i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 13 ---
mux_set_channel 13i
store_var ch 13i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 14 ---
mux_set_channel 14i
store_var ch 14i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 15 ---
mux_set_channel 15i
store_var ch 15i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# --- Channel 16 ---
mux_set_channel 16i
store_var ch 16i ea
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
on_finished:
cell_off