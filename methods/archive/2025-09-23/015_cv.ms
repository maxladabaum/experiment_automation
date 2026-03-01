e
var c
var p
set_pgstat_mode 2
set_max_bandwidth 40
set_range ba 100u
set_autoranging ba 1n 100u
set_e 0
cell_on
# CV measurement loop
meas_loop_cv p c 0 -200m 1.5 10m 1
	pck_start
	pck_add p
	pck_add c
	pck_end
endloop
on_finished:
cell_off