# MUX16: configure all mux GPIO pins as outputs
set_gpio_cfg 0x3FFi 1
var c
var p
var ch
set_pgstat_chan 0
set_pgstat_mode 2
set_max_bandwidth 40
set_range ba 100u
set_autoranging ba 1n 100u
# --- Channel 1 ---
set_gpio 0x00i
wait 100m  # settle after switching channel
store_var ch 1i ea
set_e 0
cell_on
# CV measurement loop
meas_loop_cv p c 0 -500m 500m 2m 100m
	pck_start
		pck_add ch
		pck_add p
		pck_add c
	pck_end
endloop
cell_off
on_finished:
cell_off