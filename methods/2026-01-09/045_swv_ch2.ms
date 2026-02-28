e
# MUX16 channel select
set_gpio_cfg 0x3FFi 1
set_gpio 17i
var c
var p
var f
var r
set_pgstat_mode 2
set_max_bandwidth 1600
set_range_minmax da -1036m 236m
set_range ba 5m
set_autoranging ba 100n 5m
cell_on
# Equilibrate at -1000m for 1s
set_e -1000m
wait 1
meas_loop_swv p c f r -1000m 200m 2m 36m 200
	pck_start
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
on_finished:
cell_off