import comtypes.client
import os.path


def attachToETABSInstance():
	# attach to running ETABS instance
	try:
		ETABS = comtypes.client.GetActiveObject('CSI.ETABS.API.ETABSObject')
	except(OSError, comtypes.COMError):
		print('ETABS not running!')

	# init model object for running commands
	Model = ETABS.SapModel

	return Model


def attachToSAP2000Instance():
	# attach to running ETABS instance
	try:
		ETABS = comtypes.client.GetActiveObject('CSI.SAP2000.API.ETABSObject')
	except(OSError, comtypes.COMError):
		print('SAP2000 not running!')

	# init model object for running commands
	Model = ETABS.SapModel

	return Model


def _openETABSModel(modelPath, programPath):
	if not os.path.exists(programPath):
		raise Exception('Program file path does not exist')
	if not os.path.exists(modelPath):
		raise Exception('Model path does not exist')

	helper = comtypes.client.CreateObject('ETABSv1.Helper')
	helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)

	try:
		SAPObject = helper.CreateObject(programPath)
	except (OSError, comtypes.COMError):
		print('Cannot start a new instance of the program from ' + programPath)

	SAPObject.ApplicationStart()
	Model = SAPObject.SapModel
	Model.InitializeNewModel()
	execute(Model.File.OpenFile(modelPath))
	return Model


def openETABS18Model(modelPath, programPath=None):
	if (programPath == None):
		programPath = "c:\Program Files\Computers and Structures\ETABS 18\ETABS.exe"

	return _openETABSModel(modelPath, programPath)


def openETABS19Model(modelPath, programPath=None):
	if (programPath == None):
		programPath = "c:\Program Files\Computers and Structures\ETABS 19\ETABS.exe"

	return _openETABSModel(modelPath, programPath)


def getUnitsEnum(unitName):
	Units = {
            'lb_in_F': 1,
            'lb_ft_F': 2,
            'kip_in_F': 3,
            'kip_ft_F': 4,
            'kN_mm_C': 5,
            'kN_m_C': 6,
            'kgf_mm_C': 7,
            'kgf_m_C': 8,
            'N_mm_C': 9,
            'N_m_C': 10,
            'Ton_mm_C': 11,
            'Ton_m_C': 12,
            'kN_cm_C': 13,
            'kgf_cm_C': 14,
            'N_cm_C': 15,
            'Ton_cm_C': 16
        }

	return Units[unitName]


def execute(func, msg=None):
    ret = func
    success = False
    returnedValues = None
    if (isinstance(ret, list)):
        success = (ret[-1] == 0)
        returnedValues = ret[:-1]
    else:
        success = (ret == 0)
    if success is True:
        return returnedValues
    else:
        if (msg == None):
            msg = 'API returned none-zero result'
        raise Exception(msg)
