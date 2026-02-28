e
set_gpio_cfg 0x3FFi 1
set_gpio 119i
var c
var p
var f
var g
set_pgstat_chan 1
set_pgstat_mode 0
set_pgstat_chan 0
set_pgstat_mode 3
set_max_bandwidth 4k
set_range_minmax da -536m 36m
set_range ba 59n
set_autoranging ba 59n 59n
set_e -500m
cell_on
meas_loop_ca p c -500m 1 1
	pck_start
		pck_add p
		pck_add c
	pck_end
endloop
meas_loop_swv p c f g -500m 0 2m 36m 100
	pck_start
		pck_add p
		pck_add c
		pck_add f
		pck_add g
	pck_end
endloop
on_finished:
	cell_off