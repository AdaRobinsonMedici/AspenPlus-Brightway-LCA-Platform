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

    def re_initialization(self):
        # initial the chemical process in Aspen Plus
        self.Application.Reinit()

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


    def check_convergency(self, last_line_checked):
        # check the simulation convergency by detecting errors in the history file
        runID = self.Application.Tree.FindNode(r"\Data\Results Summary\Run-Status\Output\RUNID").Value
        his_file = rf"C:\Users\Jonas\PycharmProjects\master_thesis\Aspen\AspenPlus-Python-Interface\Automation tests\Test run_base model\{runID}.his"

        with open(his_file, "r") as f:

            for _ in range(last_line_checked):
                f.readline()

            lines = f.readlines()
            error_Solver1 = True
            error_Solver2 = True
            for i, line in enumerate(lines):
                if "$OLVER01" and "*** CONVERGED ***" in line:
                    # Speichere die Nummer der gefundenen Zeile
                    last_line_checked += i + 1
                    error_Solver1 = False
                elif "$OLVER02" and "*** CONVERGED ***" in line:
                    # Speichere die Nummer der gefundenen Zeile
                    last_line_checked += i + 1
                    error_Solver2 = False

            if error_Solver1 == False and error_Solver2 == False:
                print("No errors")
            elif error_Solver1 == True:
                print("Error encountered in Solver 1")
            else:
                print("Error encountered in Solver 2")

        return last_line_checked
            # noERROR_Solver1 = np.any(np.array([("$OLVER01" in line) and ("*** CONVERGED ***" in line) for line in lines]) >= 0)
            # noERROR_Solver2 = np.any(np.array([("$OLVER02" in line) and ("*** CONVERGED ***" in line) for line in lines]) >= 0)
            #
            # if noERROR_Solver1 == True and noERROR_Solver2 == True:
            #     return "No error"
            # else:
            #     return "Error encountered"



    def close_bkp(self):
        # close the Aspen Plus file
        self.Application.Quit()

    def collect_stream(self):
        # colloct all streams involved in the process
        streams = []
        node = self.Application.Tree.FindNode(r"\Data\Streams")
        for item in node.Elements:
            streams.append(item.Name)
        return tuple(streams)

    def collect_block(self):
        # colloct all blocks involved in the process
        blocks = []
        node = self.Application.Tree.FindNode(r"\Data\Blocks")
        for item in node.Elements:
            blocks.append(item.Name)
        return tuple(blocks)

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


def SequenceWithEndPoint(start, stop, step):
    # generate evenly spaced values containing the end point
    return np.arange(start, stop + step, step)


def ListValue2Str(alist):
    # convert a list of mixed variables into string formats
    return list(map(str, alist))


