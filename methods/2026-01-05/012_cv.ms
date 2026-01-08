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
# --- Channel 2 ---
set_gpio 0x11i
wait 100m  # settle after switching channel
store_var ch 2i ea
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
# --- Channel 3 ---
set_gpio 0x22i
wait 100m  # settle after switching channel
store_var ch 3i ea
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
# --- Channel 4 ---
set_gpio 0x33i
wait 100m  # settle after switching channel
store_var ch 4i ea
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
# --- Channel 5 ---
set_gpio 0x44i
wait 100m  # settle after switching channel
store_var ch 5i ea
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
# --- Channel 6 ---
set_gpio 0x55i
wait 100m  # settle after switching channel
store_var ch 6i ea
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
# --- Channel 7 ---
set_gpio 0x66i
wait 100m  # settle after switching channel
store_var ch 7i ea
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
# --- Channel 8 ---
set_gpio 0x77i
wait 100m  # settle after switching channel
store_var ch 8i ea
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
# --- Channel 9 ---
set_gpio 0x88i
wait 100m  # settle after switching channel
store_var ch 9i ea
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
# --- Channel 10 ---
set_gpio 0x99i
wait 100m  # settle after switching channel
store_var ch 10i ea
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
# --- Channel 11 ---
set_gpio 0xAAi
wait 100m  # settle after switching channel
store_var ch 11i ea
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
# --- Channel 12 ---
set_gpio 0xBBi
wait 100m  # settle after switching channel
store_var ch 12i ea
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
# --- Channel 13 ---
set_gpio 0xCCi
wait 100m  # settle after switching channel
store_var ch 13i ea
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
# --- Channel 14 ---
set_gpio 0xDDi
wait 100m  # settle after switching channel
store_var ch 14i ea
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
# --- Channel 15 ---
set_gpio 0xEEi
wait 100m  # settle after switching channel
store_var ch 15i ea
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
# --- Channel 16 ---
set_gpio 0xFFi
wait 100m  # settle after switching channel
store_var ch 16i ea
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