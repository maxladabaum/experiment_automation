e
# MUX16 channel select
set_gpio_cfg 0x3FFi 1
set_gpio 17i
var f
var z
var i
set_pgstat_chan 1
set_pgstat_mode 0
set_pgstat_chan 0
set_pgstat_mode 3
cell_on
set_e 0
# Dry alignment EIS sweep
meas_loop_eis f z i 10m 10 10000 13i 0
	pck_start
		pck_add f
		pck_add z
		pck_add i
	pck_end
endloop
on_finished:
cell_off
