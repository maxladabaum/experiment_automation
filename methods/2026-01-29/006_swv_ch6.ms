e
set_gpio_cfg 0x3FFi 1
set_gpio 85i
var c
var p
var f
var r
set_pgstat_chan 1
set_pgstat_mode 0
set_pgstat_chan 0
set_pgstat_mode 3
set_max_bandwidth 4k
set_range_minmax da -586m 36m
set_range ba 59n
set_autoranging ba 59n 59n
cell_on
# Equilibrate at -650m for 1s
set_e -650m
wait 1
meas_loop_swv p c f r -650m -150m 2m 36m 400
	pck_start
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
on_finished:
	cell_off