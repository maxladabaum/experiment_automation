e
# MUX16 channel select
set_gpio_cfg 0x3FFi 1
set_gpio 34i
var c
var p
var f
var r
var t
set_pgstat_chan 1
set_pgstat_mode 0
set_pgstat_chan 0
set_pgstat_mode 3
set_max_bandwidth 8k
set_range_minmax da -736m -64m
set_range ba 59n
set_autoranging ba 59n 59n
set_e -700m
cell_on
# Equilibrium check at -700m for 5s
store_var t 0 eb
meas_loop_ca p c -700m 200m 5
	pck_start
		pck_add t
		pck_add p
		pck_add c
	pck_end
	add_var t 200m
endloop
store_var t 0 eb
set_e -700m
meas_loop_swv p c f r -700m -100m 2m 36m 200
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