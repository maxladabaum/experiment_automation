e
set_gpio_cfg 0x3FFi 1

var i
var h
var p
var c

store_var i 0i aa
store_var h 1i ea

set_pgstat_chan 0
set_pgstat_mode 2
set_max_bandwidth 400
set_pot_range -1 1
set_cr 1m
set_autoranging 10u 1m

cell_on

# i will take 0x00,0x11,...0xFF (16 steps)
loop i <= 0xFFi
    set_gpio i

    # derive channel number from i: h = (i/0x11)+1
    # MethodSCRIPT has limited math; easiest is just increment h each loop:
    # (so we keep h as loop counter and i as gpio word)
    # We'll do that by resetting h before loop and incrementing inside.
    
    set_e -1000m
    wait 100m

    meas_loop_lsv p c -1 1 10m 1
        pck_start
            pck_add h
            pck_add p
            pck_add c
        pck_end
    endloop

    add_var i 0x11i
    add_var h 1i
endloop

on_finished:
cell_off

