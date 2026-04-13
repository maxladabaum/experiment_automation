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
set_range_minmax da -715m -85m
set_range ba 489n
set_autoranging ba 489n 489n
cell_on
# Equilibrate at -700m for 5s
set_e -700m
wait 5
set_e -700m
meas_loop_swv p c f r -700m -100m 1m 15m 200
	pck_start
		pck_add p
		pck_add c
		pck_add f
		pck_add r
	pck_end
endloop
on_finished:
cell_off