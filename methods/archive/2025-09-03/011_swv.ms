e
var c
var p
var f
var r
set_pgstat_mode 2
set_max_bandwidth 1600
set_range_minmax da -520m 520m
set_range ba 5m
set_autoranging ba 100n 5m
cell_on
# Equilibrate at 0 for 5s
set_e 0
wait 5
# SWV measurement loop
meas_loop_swv p c f r -500m 500m 5m 20m 15
	pck_start
	pck_add p
	pck_add c
	pck_add f
	pck_add r
	pck_end
endloop
on_finished:
cell_off