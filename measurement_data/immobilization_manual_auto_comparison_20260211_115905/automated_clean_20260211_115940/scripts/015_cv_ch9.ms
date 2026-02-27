e
set_gpio_cfg 0x3FFi 1
set_gpio 136i
var c
var p
var t
set_pgstat_chan 1
set_pgstat_mode 0
set_pgstat_chan 0
set_pgstat_mode 2
set_max_bandwidth 400
set_range_minmax da -400m 1300m
set_range ba 29500n
set_autoranging ba 29500n 29500n
set_e -400m
cell_on
store_var t 0 eb
# CV measurement loop
meas_loop_cv p c -400m -400m 1300m 10m 1000m nscans(50)
	pck_start
		pck_add p
		pck_add c
		pck_add t
	pck_end
	add_var t 9879557n
endloop
on_finished:
	cell_off