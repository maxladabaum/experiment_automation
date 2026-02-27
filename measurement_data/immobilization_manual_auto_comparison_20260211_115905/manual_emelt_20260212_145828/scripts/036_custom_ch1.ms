e
set_gpio_cfg 0x3FFi 1
set_gpio 0i
var c
var p
var t
set_pgstat_chan 1
set_pgstat_mode 0
set_pgstat_chan 0
set_pgstat_mode 3
set_max_bandwidth 400
set_range_minmax da -1 -1
set_range ba 118u
set_autoranging ba 118u 118u
set_e -1
cell_on
store_var t 0 eb
meas_loop_ca p c -1 200m 1
  pck_start
    pck_add t
    pck_add p
    pck_add c
  pck_end
  add_var t 200m
endloop
store_var t 0 eb
meas_loop_ca p c -1 10m 20010m
  pck_start
    pck_add t
    pck_add p
    pck_add c
  pck_end
  add_var t 10m
endloop
on_finished:
  cell_off
