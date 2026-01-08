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
on_finished:
cell_off