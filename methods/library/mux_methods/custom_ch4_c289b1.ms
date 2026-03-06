e
# MUX16 channel select
set_gpio_cfg 0x3FFi 1
set_gpio 51i
var c
var p
var t
set_pgstat_chan 1
set_pgstat_mode 0
set_pgstat_chan 0
set_pgstat_mode 2
set_max_bandwidth 4
set_range_minmax da -1 -700m
set_range ba 7375n
set_autoranging ba 7375n 7375n
set_e -700m
cell_on
store_var t 0 eb
meas_loop_lsv p c -700m -1 1m 1m
  pck_start
    pck_add p
    pck_add c
    pck_add t
  pck_end
  add_var t 1185547u
endloop
on_finished:
  cell_off