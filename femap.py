import pythoncom
import Pyfemap
from Pyfemap import constants as c
import sys
import os

"""
This class allows to connect to the current femap session and contains several useful methods
"""


class femap:
    def __init__(self):
        try:
            existObj = pythoncom.connect(Pyfemap.model.CLSID)
            self.app = Pyfemap.model(existObj)
            self.app.feAppUndoCheckpoint("Git Push")  # Create checkpoint for FEMAP “undo” to revert to
            self.app.feAppMessage(0, "Python connected to femap")
        except:
            sys.exit("femap not open")

    def select_folder_path(self):
        return

    def export_group_to_bdf(self, filefolder):
        """
        Export a FEMAP group into a .bdf file.
        """
        selection = self.app.feSet
        selection.clear()
        group = self.app.feGroup

        # Put the group contents into the selection
        rc = selection.SelectMultiIDV2(c.FT_GROUP, 1, "Select Groups to Export to BDFs")
        if rc != c.FE_OK:
            self.app.feAppMessage(3, f"Failed to put group contents into Set.")
            raise Exception(f"Failed to put group contents into Set.")

        # Initialize analysis
        analysis = self.app.feAnalysisMgr
        analysis.NasExecSkipStandard = False
        analysis.NasBulkSkipStandard = False
        analysis.NasBulkLargeField = 4
        analysis.ID = self.app.Info_NextID(c.FT_AMGR_DIR)
        analysis.InitAnalysisMgr(c.FAM_NX_NASTRAN, c.FAT_MODES, True)
        analysis.LinkedSolverOption = c.FS_LINKED
        analysis.NasBulkOn = True
        analysis.Active = analysis.ID

        # Exports
        makerun = not os.path.exists(f"{filefolder}\\run.dat")
        while group.NextInSet(selection.ID):
            analysis.NasBulkGroupID = group.ID
            analysis.Put(analysis.ID)
            filename = f"{filefolder}\\{group.title}.bdf"
            self.app.feFileWriteNastran(0, filename)
            self.parse_file(filename, filefolder, makerun)
            makerun = False
            self.add_includes(filefolder,"run",filename)
            self.app.feAppMessage(0, f"Group {group.ID} exported to: {filefolder}")

        self.app.feDelete(c.FT_AMGR_DIR, -analysis.ID)
        self.app.feAppMessage(3, f"All files exported")

    def parse_file(self, filepath,filefolder, firstit):
        with open(filepath, 'r') as file:
            lines = file.readlines()

        INIT = True
        for idx, line in enumerate(lines):
            if "INIT" in line:
                start_run = idx
            if "CORD2S" in line and INIT:
                start_index = idx + 2
                INIT = False
            if "ENDDATA" in line:
                end_index = idx
                break

        if firstit:
            try:
                with open(f"{filefolder}\\run.dat", "w") as runfile:
                    runfile.writelines(lines[start_run:start_index-2]+[lines[end_index]])
            except:
                print("\033[91mWARNING: Could not locate INIT statement,"
                  " check your input file or create INIT file manually.\033[0m")

        try:
            with open(filepath, 'w') as file:
                file.writelines(lines[start_index:end_index])
        except:
            print("\033[91mWARNING: Could not locate INIT statement,"
                  " check your input file or create INIT file manually.\033[0m")

    def add_includes(self, filefolder, runname, groupname):
        with open(f"{filefolder}\\{runname}.dat", 'r') as f:
            lines = f.readlines()

        flag = False
        for line in lines:
            if f"INCLUDE '{groupname}'" in line:
                flag = True

        if len(lines) > 0 and not flag:
            lines.insert(-1, f"INCLUDE '{groupname}'\n")
            with open(f"{filefolder}\\{runname}.dat", 'w') as f:
                f.writelines(lines)







