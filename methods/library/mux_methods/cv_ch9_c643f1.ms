e
# MUX16 channel select
set_gpio_cfg 0x3FFi 1
set_gpio 136i
var c
var p
set_pgstat_chan 1
set_pgstat_mode 0
set_pgstat_chan 0
set_pgstat_mode 2
set_max_bandwidth 66667m
set_range_minmax da -200m 1100m
set_range ba 59n
set_autoranging ba 59n 59n
set_e -200m
cell_on
# CV measurement loop
meas_loop_cv p c -200m -200m 1100m 10m 1000m nscans(50)
	pck_start
	pck_add p
	pck_add c
	pck_end
endloop
on_finished:
cell_off