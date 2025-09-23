e
var c
var p
var f
var r
var i
store_var i 0i ja
set_pgstat_chan 0
set_pgstat_mode 2
set_max_bandwidth 40
set_range_minmax da -520m 520m
set_range ba 100u
set_autoranging ba 1n 100u
cell_on
# Measure SWV: E, I, I_fwd, I_rev, E_begin, E_end, E_step, E_amp, freq
meas_loop_swv p c f r -0.5 0.5 0.01 0.02 15
	pck_start
	pck_add p
	pck_add c
	pck_add f
	pck_add r
	pck_end
endloop
on_finished:
cell_off

