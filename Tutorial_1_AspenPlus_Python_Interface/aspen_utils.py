# Created on 12 May 2022 by Zihao Wang, zwang@mpi-magdeburg.mpg.de


import os
import time
from win32com.client import GetObject, Dispatch
import numpy as np


class Aspen_Plus_Interface():
    def __init__(self):
        self.Application = Dispatch("Apwn.Document")
        # {"V10.0": "36.0", V11.0": "37.0", V12.0": "38.0"}
        print(self.Application)

    def load_bkp(self, bkp_file, visible_state=1, dialog_state=0):
        """
        Load a process via bkp file
        :param bkp_file: location of the Aspen Plus file
        :param visible_state: Aspen Plus user interface, 0 is invisible and 1 is visible
        :param dialog_state: Aspen Plus dialogs, 0 is not to suppress and 1 is to suppress
        """
        self.Application.InitFromArchive2(os.path.abspath(bkp_file))
        self.Application.Visible = visible_state
        self.Application.SuppressDialogs = dialog_state

    def run_simulation(self):
        # run the process simulation
        self.Application.Engine.Run2()

    def check_run_completion(self, time_limit=60):
        # check whether the simulation completed
        times = 0
        while self.Application.Engine.IsRunning == 1:
            time.sleep(1)
            times += 1
            if times >= time_limit:
                print("Violate time limitation")
                self.Application.Engine.Stop
                break

    def close_bkp(self):
        # close the Aspen Plus file
        self.Application.Quit()

    # Test
    def ExportSummaryFile(self, filename: str) -> None:
        """Saves SummaryFile (.sum) of Aspen Simulation with a given name.

        Args:
            Filename: String which gives the File name.
        """
        self.Application.Export(3, filename)

    def ExportRunMessagesFile(self, filename: str) -> None:
        """Saves Messages, Errors, Warnings and diagnostics from running the Simulation for each run.

        Args:
            Filename: String which gives the File name.
        """
        self.Application.Export(6, filename)

    def SaveAs(self, Filename: str, overwrite: bool = True) -> None:
        """Saves the current Aspen Simulation,(.apw) with a new name with/out overwritting.

        Args:
            Filename: String which gives the File name.
            overwrite: Should file be overwritten when the File already exists? True or False, standard is True
        """
        self.Application.SaveAs(Filename, overwrite)


def KillAspen():
    # kill the Aspen Plus
    WMI = GetObject("winmgmts:")
    for p in WMI.ExecQuery("select * from Win32_Process where Name='AspenPlus.exe'"):
        os.system("taskkill /pid " + str(p.ProcessId))





