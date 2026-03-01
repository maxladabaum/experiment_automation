# MUX16: configure all mux GPIO pins as outputs
set_gpio_cfg 0x3FFi 1
var c
var p
var f
var r
var ch
set_pgstat_chan 0
set_pgstat_mode 2
set_max_bandwidth 1600
set_range_minmax da -520m 520m
set_range ba 5m
set_autoranging ba 100n 5m
# --- Channel 1 ---
set_gpio 0x00i
wait 100m  # settle after switching channel
store_var ch 1i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 2 ---
set_gpio 0x11i
wait 100m  # settle after switching channel
store_var ch 2i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 3 ---
set_gpio 0x22i
wait 100m  # settle after switching channel
store_var ch 3i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 4 ---
set_gpio 0x33i
wait 100m  # settle after switching channel
store_var ch 4i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 5 ---
set_gpio 0x44i
wait 100m  # settle after switching channel
store_var ch 5i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 6 ---
set_gpio 0x55i
wait 100m  # settle after switching channel
store_var ch 6i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 7 ---
set_gpio 0x66i
wait 100m  # settle after switching channel
store_var ch 7i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 8 ---
set_gpio 0x77i
wait 100m  # settle after switching channel
store_var ch 8i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 9 ---
set_gpio 0x88i
wait 100m  # settle after switching channel
store_var ch 9i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 10 ---
set_gpio 0x99i
wait 100m  # settle after switching channel
store_var ch 10i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 11 ---
set_gpio 0xAAi
wait 100m  # settle after switching channel
store_var ch 11i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 12 ---
set_gpio 0xBBi
wait 100m  # settle after switching channel
store_var ch 12i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 13 ---
set_gpio 0xCCi
wait 100m  # settle after switching channel
store_var ch 13i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 14 ---
set_gpio 0xDDi
wait 100m  # settle after switching channel
store_var ch 14i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 15 ---
set_gpio 0xEEi
wait 100m  # settle after switching channel
store_var ch 15i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
# --- Channel 16 ---
set_gpio 0xFFi
wait 100m  # settle after switching channel
store_var ch 16i ea
cell_on
meas_loop_swv p c f r -500m 500m 2m 20m 15
	pck_start
		pck_add ch
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
cell_off
on_finished:
cell_off