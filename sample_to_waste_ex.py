from centris_pure import CentrisPumpPure

PORT_SAMPLE = 1   # set to the sample port you want
PORT_WASTE  = 9   # set to your waste port (or whichever you prefer)

pump = CentrisPumpPure(com_port=8, dev=1, baud=9600)
pump.open()
print("Init:", pump.initialize())

# sanity check: tiny air move
print("Asp 10 µL:", pump.aspirate_ul(10))
print("Disp 10 µL:", pump.dispense_ul(10))

# go to sample, aspirate a small volume
print("Valve->SAMPLE:", pump.valve_to(PORT_SAMPLE))
print("Asp 50 µL sample:", pump.aspirate_ul(50))

# go to waste, dispense
print("Valve->WASTE:", pump.valve_to(PORT_WASTE))
print("Disp 50 µL waste:", pump.dispense_ul(50))

pump.close()
print("Done.")
