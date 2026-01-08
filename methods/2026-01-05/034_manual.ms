e 
set_gpio_cfg 0x3FFi 1 
set_gpio 0x11i 
var i 
var c 
var p 
store_var i 0i aa 
set_pgstat_chan 0 
set_pgstat_mode 2 
set_max_bandwidth 400 
set_pot_range -1 1 
set_cr 1m 
set_autoranging 10u 1m 
cell_on 
loop i <= 0xFFi 
 set_gpio i 
 set_e -1000m 
 wait 100m 
     meas_loop_lsv p c -1 1 10m 1 
        pck_start 
        pck_add p 
        pck_add c 
        pck_end 
     endloop 
 add_var i 0x11i 
endloop 
on_finished

