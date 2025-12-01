import Pyfemap

import femap

if __name__ == "__main__":
    f = femap.femap()
    f.export_group_to_bdf(205, r"C:\Users\alessandro.vignola\GROUP1.bdf")

