import pandas as pd


def readNetworkInputs(input_path):
    df = pd.read_csv(input_path, header=None, names=['variable', 'value'])
    df.set_index('variable', inplace=True)

    buildNetworkBool = df.at['BuildNetwork', 'value']
    shapes_path = df.at['NetworkShapes', 'value']
    tiploc_path = df.at['TIPLOCs', 'value']
    BPLAN_path = df.at['BPLAN', 'value']
    ELR_path = df.at['ELR', 'value']
    merge_path = df.at['MergeStopsPath', 'value']
    tsys_path = df.at['TransportSystems', 'value']
    xfer_path = df.at['TransferLinks', 'value']

    buildNetwork = [buildNetworkBool, shapes_path, tiploc_path, BPLAN_path, ELR_path, merge_path, tsys_path, xfer_path]

    return buildNetwork


def readMergeInputs(input_path):
    df = pd.read_csv(input_path, header=None, names=['variable', 'value'])
    df.set_index('variable', inplace=True)

    mergeStopsBool = df.at['MergeStops', 'value']
    merge_path = df.at['MergeStopsPath', 'value']

    mergeStops = [mergeStopsBool, merge_path]

    return mergeStops

def readTimetableInputs(input_path):
    df = pd.read_csv(input_path, header=None, names=['variable', 'value'])
    df.set_index('variable', inplace=True)

    importTimetableBool = df.at['ImportTimetable', 'value']
    timetable_path = df.at['TimetablePath', 'value']

    importTimetable = [importTimetableBool, timetable_path]

    return importTimetable


def readVerInputs(input_path):
    df = pd.read_csv(input_path, header=None, names=['variable', 'value'])
    df.set_index('variable', inplace=True)

    createVersBool = df.at['CreateVers', 'value']
    ver1 = df.at['Ver1', 'value']
    ver2 = df.at['Ver2', 'value']
    ver3 = df.at['Ver3', 'value']
    ver4 = df.at['Ver4', 'value']
    ver5 = df.at['Ver5', 'value']

    # verx should be of the form {TSys:[TsysCode1, TsysCode2, ...], StartDate:dd/mm/yyyy, EndDate:dd/mm/yyyy}

    createVers = [createVersBool]

    for v in [ver1, ver2, ver3, ver4, ver5]:
        if v != "":
            createVers.append(v)

    return createVers



def read_inputs(input_path):
    df = pd.read_csv(input_path, header=None, names=['variable', 'value'])

    df.set_index('variable', inplace=True)

    workingPath = df.at['WorkingFolder', 'value']
    buildNetworkBool = df.at['BuildNetwork', 'value']
    shapes_path = df.at['NetworkShapes', 'value']
    tiploc_path = df.at['TIPLOCs', 'value']
    BPLAN_path = df.at['BPLAN', 'value']
    ELR_path = df.at['ELR', 'value']
    merge_path = df.at['MergeStopsPath', 'value']
    tsys_path = df.at['TransportSystems', 'value']

    buildNetwork = [buildNetworkBool, shapes_path, tiploc_path, BPLAN_path, ELR_path, merge_path, tsys_path]

    importTimetableBool = df.at['ImportTimetable', 'value']
    timetable_path = df.at['TimetablePath', 'value']

    importTimetable = [importTimetableBool, timetable_path]

    mergeStopsBool = df.at['MergeStops', 'value']

    mergeStops = [mergeStopsBool, merge_path]

    createVersBool = df.at['CreateVers', 'value']
    ver1 = df.at['Ver1', 'value']
    ver2 = df.at['Ver2', 'value']
    ver3 = df.at['Ver3', 'value']
    ver4 = df.at['Ver4', 'value']
    ver5 = df.at['Ver5', 'value']

    # verx should be of the form {TSys:[TsysCode1, TsysCode2, ...], StartDate:dd/mm/yyyy, EndDate:dd/mm/yyyy}

    createVers = [createVersBool]

    for v in [ver1, ver2, ver3, ver4, ver5]:
        if v != "":
            createVers.append(v)

    return workingPath, buildNetwork, importTimetable, mergeStops, createVers


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
import wx


class Form(wx.Panel):
    ''' The Form class is a wx.Panel that creates a bunch of controls
        and handlers for callbacks. Doing the layout of the controls is 
        the responsibility of subclasses (by means of the doLayout()
        method). '''

    def __init__(self, *args, **kwargs):
        super(Form, self).__init__(*args, **kwargs)
        self.referrers = ['friends', 'advertising', 'websearch', 'yellowpages']
        self.colors = ['blue', 'red', 'yellow', 'orange', 'green', 'purple',
                       'navy blue', 'black', 'gray']
        self.createControls()
        self.bindEvents()
        self.doLayout()

    def createControls(self):
        self.logger = wx.TextCtrl(self, style=wx.TE_MULTILINE|wx.TE_READONLY)
        self.saveButton = wx.Button(self, label="Save")
        self.nameLabel = wx.StaticText(self, label="Your name:")
        self.nameTextCtrl = wx.TextCtrl(self, value="Enter here your name")
        self.referrerLabel = wx.StaticText(self, 
            label="How did you hear from us?")
        self.referrerComboBox = wx.ComboBox(self, choices=self.referrers, 
            style=wx.CB_DROPDOWN)
        self.insuranceCheckBox = wx.CheckBox(self, 
            label="Do you want Insured Shipment?")
        self.colorRadioBox = wx.RadioBox(self, 
            label="What color would you like?", 
            choices=self.colors, majorDimension=3, style=wx.RA_SPECIFY_COLS)

    def bindEvents(self):
        for control, event, handler in \
            [(self.saveButton, wx.EVT_BUTTON, self.onSave),
             (self.nameTextCtrl, wx.EVT_TEXT, self.onNameEntered),
             (self.nameTextCtrl, wx.EVT_CHAR, self.onNameChanged),
             (self.referrerComboBox, wx.EVT_COMBOBOX, self.onReferrerEntered),
             (self.referrerComboBox, wx.EVT_TEXT, self.onReferrerEntered),
             (self.insuranceCheckBox, wx.EVT_CHECKBOX, self.onInsuranceChanged),
             (self.colorRadioBox, wx.EVT_RADIOBOX, self.onColorchanged)]:
            control.Bind(event, handler)

    def doLayout(self):
        ''' Layout the controls that were created by createControls(). 
            Form.doLayout() will raise a NotImplementedError because it 
            is the responsibility of subclasses to layout the controls. '''
        raise NotImplementedError 

    # Callback methods:

    def onColorchanged(self, event):
        self.__log('User wants color: %s'%self.colors[event.GetInt()])

    def onReferrerEntered(self, event):
        self.__log('User entered referrer: %s'%event.GetString())

    def onSave(self,event):
        self.__log('User clicked on button with id %d'%event.GetId())

    def onNameEntered(self, event):
        self.__log('User entered name: %s'%event.GetString())

    def onNameChanged(self, event):
        self.__log('User typed character: %d'%event.GetKeyCode())
        event.Skip()

    def onInsuranceChanged(self, event):
        self.__log('User wants insurance: %s'%event.IsChecked())

    # Helper method(s):

    def __log(self, message):
        ''' Private method to append a string to the logger text
            control. '''
        self.logger.AppendText('%s\n'%message)


class FormWithAbsolutePositioning(Form):
    def doLayout(self):
        ''' Layout the controls by means of absolute positioning. '''
        for control, x, y, width, height in \
                [(self.logger, 300, 20, 200, 300),
                 (self.nameLabel, 20, 60, -1, -1),
                 (self.nameTextCtrl, 150, 60, 150, -1),
                 (self.referrerLabel, 20, 90, -1, -1),
                 (self.referrerComboBox, 150, 90, 95, -1),
                 (self.insuranceCheckBox, 20, 180, -1, -1),
                 (self.colorRadioBox, 20, 210, -1, -1),
                 (self.saveButton, 200, 300, -1, -1)]:
            control.SetDimensions(x=x, y=y, width=width, height=height)


class FormWithSizer(Form):
    def doLayout(self):
        ''' Layout the controls by means of sizers. '''

        # A horizontal BoxSizer will contain the GridSizer (on the left)
        # and the logger text control (on the right):
        boxSizer = wx.BoxSizer(orient=wx.HORIZONTAL)
        # A GridSizer will contain the other controls:
        gridSizer = wx.FlexGridSizer(rows=5, cols=2, vgap=10, hgap=10)

        # Prepare some reusable arguments for calling sizer.Add():
        expandOption = dict(flag=wx.EXPAND)
        noOptions = dict()
        emptySpace = ((0, 0), noOptions)
    
        # Add the controls to the sizers:
        for control, options in \
                [(self.nameLabel, noOptions),
                 (self.nameTextCtrl, expandOption),
                 (self.referrerLabel, noOptions),
                 (self.referrerComboBox, expandOption),
                  emptySpace,
                 (self.insuranceCheckBox, noOptions),
                  emptySpace,
                 (self.colorRadioBox, noOptions),
                  emptySpace, 
                 (self.saveButton, dict(flag=wx.ALIGN_CENTER))]:
            gridSizer.Add(control, **options)

        for control, options in \
                [(gridSizer, dict(border=5, flag=wx.ALL)),
                 (self.logger, dict(border=5, flag=wx.ALL|wx.EXPAND, 
                    proportion=1))]:
            boxSizer.Add(control, **options)

        self.SetSizerAndFit(boxSizer)


class FrameWithForms(wx.Frame):
    def __init__(self, *args, **kwargs):
        super(FrameWithForms, self).__init__(*args, **kwargs)
        notebook = wx.Notebook(self)
        form1 = FormWithAbsolutePositioning(notebook)
        form2 = FormWithSizer(notebook)
        notebook.AddPage(form1, 'Absolute Positioning')
        notebook.AddPage(form2, 'Sizers')
        # We just set the frame to the right size manually. This is feasible
        # for the frame since the frame contains just one component. If the
        # frame had contained more than one component, we would use sizers
        # of course, as demonstrated in the FormWithSizer class above.
        self.SetClientSize(notebook.GetBestSize()) 





if __name__ == '__main__':
    app = wx.App(0)
    frame = FrameWithForms(None, title='Demo with Notebook')
    frame.Show()
    app.MainLoop()

