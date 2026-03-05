e
var c
var p
var f
var r
set_pgstat_chan 1
set_pgstat_mode 0
set_pgstat_chan 0
set_pgstat_mode 3
set_max_bandwidth 4k
set_range_minmax da -636m 36m
set_range ba 59n
set_autoranging ba 59n 59n
cell_on
set_e -600m
meas_loop_swv p c f r -600m 0 2m 36m 200
	pck_start
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
on_finished:
	cell_off