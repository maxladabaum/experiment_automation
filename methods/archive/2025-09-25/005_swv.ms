e
var c
var p
var f
var r
set_pgstat_mode 2
set_max_bandwidth 1600
set_range_minmax da -436m 136m
set_range ba 5m
set_autoranging ba 100n 5m
cell_on
# SWV measurement loop
meas_loop_swv p c f r -400m 100m 2m 36m 200
	pck_start
	pck_add p
	pck_add c
	pck_add f
	pck_add r
	pck_end
endloop
on_finished:
cell_off