e
# MUX16 channel select
set_gpio_cfg 0x3FFi 1
set_gpio 136i
# End MUX16 channel select
var c
var p
var f
var r
var t
set_pgstat_chan 1
set_pgstat_mode 0
set_pgstat_chan 0
set_pgstat_mode 3
set_max_bandwidth 4k
set_range_minmax da -518m 18m
set_range ba 100n
set_autoranging ba 100n 100n
set_e -500m
cell_on
# Equilibrium check at -500m for 1.0s
store_var t 0 eb
meas_loop_ca p c -500m 200m 1
	pck_start
		pck_add t
		pck_add p
		pck_add c
	pck_end
	add_var t 200m
endloop
store_var t 0 eb
set_e -500m
meas_loop_swv p c f r -500m 0 5m 18m 200
	pck_start
		pck_add p
		pck_add c
		pck_add f
		pck_add r
		pck_add t
	pck_end
	add_var t 5m
endloop
on_finished:
cell_off