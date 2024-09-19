import os
import win32com.client as com
import sys

sys.path.append(os.path.dirname(__file__))

import get_inputs as gi


def main(path, runID, demandVer, outputs):
    Visum = com.Dispatch("Visum.Visum.240")
    Visum.LoadVersion(os.path.join(path, f"demand\\{demandVer}_Demand.ver"))

    Visum.Procedures.OpenXmlWithOptions(os.path.join(path, "results\\PSeq.xml"), True, True, 0)

    Visum.Procedures.Operations.ItemByKey(4).SaveVersionParameters.SetAttValue("FileName", f"{runID}_{demandVer}_Assigned.ver")
    Visum.SaveVersion(os.path.join(path, f"results\\{runID}_{demandVer}.ver"))

    if not outputs:
        Visum.Procedures.Operations.ItemByKey(5).SetAttValue('ACTIVE', 0)

    Visum.Procedures.Execute()



if __name__ == "__main__":

    path = os.path.dirname(__file__)
    input_path = os.path.join(path, "input\\inputs.csv")

    runAssignment = gi.readAssignmentInputs(input_path)

    if runAssignment[0] == "TRUE":
        runID = runAssignment[1]
        demandVer = runAssignment[2]
        outputs = runAssignment[3]

        main(path, runID, demandVer, outputs)