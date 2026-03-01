e
# MUX16 channel select
set_gpio_cfg 0x3FFi 1
set_gpio 0i
var c
var p
var f
var r
set_pgstat_mode 2
set_max_bandwidth 1600
set_range_minmax da -520m 20m
set_range ba 5m
set_autoranging ba 100n 5m
cell_on
# SWV measurement loop for 2 scans
# SWV scan 1/2
meas_loop_swv p c f r -500m 0 2m 20m 100
	pck_start
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
# SWV scan 2/2
meas_loop_swv p c f r -500m 0 2m 20m 100
	pck_start
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
on_finished:
cell_off