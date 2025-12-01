import pythoncom
import Pyfemap
import sys


"""
This class allows to connect to the current femap session and contains several useful methods
"""

class femap:
    def __init__(self):
        try:
            existObj = pythoncom.connect(Pyfemap.model.CLSID)
            self.app = Pyfemap.model(existObj)
            self.app.feAppMessage(0,"Python connected to femap")
        except:
            sys.exit("femap not open")

    def export_group_to_bdf(self, group_id, filepath):
        """
        Export a FEMAP group into a .bdf (Nastran) file.
        """
        group = self.app.feGroup
        rc = group.Get(group_id)
        if rc == 0:
            self.app.feAppMessage(3, f"Group {group_id} does not exist.")
            raise Exception(f"Group {group_id} does not exist.")

        # Clear the selection
        set_sel = self.app.feSet
        set_sel.clear()

        # Put the group contents into a Set
        rc = set_sel.Put(group.ID)
        if rc == 0:
            raise Exception(f"Failed to put group contents into Set for group {group_id}.")

        # Export as a Nastran bulk data file
        rc = set_sel.feFileWriteNastran(0,filepath)
        if rc == 0:
            self.app.feAppMessage(0, "Group does not exist.")
            raise Exception(f"Failed to export group {group_id} to BDF.")

        self.app.feAppMessage(0, f"Group {group_id} exported to: {filepath}")

