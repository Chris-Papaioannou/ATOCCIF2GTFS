import os
import win32com.client as com
import sys
import shutil
import traceback

sys.path.append(os.path.dirname(__file__))

import get_inputs as gi

import logging

logging.basicConfig(
    filename="ModelBuilder.log",
    encoding="utf-8",
    filemode="a",
    format="{asctime} - {levelname} - {message}",
    style="{",
    datefmt="%Y-%m-%d %H:%M",
    level=logging.INFO # Change to logging.DEBUG for more details
)

def main(path, runID, demandVer, outputs):
    try:
        shutil.copy(os.path.join(path, f"demand\\{demandVer}_Demand.ver"), os.path.join(path, f"results\\{runID}_{demandVer}.ver"))
        
        Visum = com.Dispatch("Visum.Visum.240")
        Visum.LoadVersion(os.path.join(path, f"results\\{runID}_{demandVer}.ver"))

        Visum.Procedures.OpenXmlWithOptions(os.path.join(path, "results\\PSeq.xml"), True, True, 0)

        Visum.Procedures.Operations.ItemByKey(4).SaveVersionParameters.SetAttValue("FileName", f"{path}\\results\\{runID}_{demandVer}_Assigned.ver")
        Visum.SaveVersion(os.path.join(path, f"results\\{runID}_{demandVer}.ver"))

        if not outputs:
            Visum.Procedures.Operations.ItemByKey(5).SetAttValue('ACTIVE', 0)

        Visum.Procedures.Execute()
    except:
        logging.error(traceback.format_exc())



if __name__ == "__main__":

    path = os.path.dirname(__file__)
    input_path = os.path.join(path, "input\\inputs.csv")

    runAssignment = gi.readAssignmentInputs(input_path)

    logging.info(f'Assignment run: {runAssignment[0]}')

    if bool(runAssignment[0]):
        runID = runAssignment[1]
        demandVer = runAssignment[2]
        outputs = bool(runAssignment[3])

        logging.info(f'RunID: {runAssignment[1]}')
        logging.info(f'Demand version file: {runAssignment[2]}')
        logging.info(f'Outputs run: {runAssignment[3]}')

        main(path, runID, demandVer, outputs)