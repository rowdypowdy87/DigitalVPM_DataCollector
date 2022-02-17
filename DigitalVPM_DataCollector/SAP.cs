using System;
using System.Diagnostics;
using SAPFEWSELib;
using SapROTWr;
using System.Data;
using System.ComponentModel;

/*
    LIBRARY TO AUTOMATE SAP WITH C# USING SAP SCRIPTING ENGINE
    WRITTEN BY CLINTON WHEELER - STRICTLY FOR USE BY UGL PTY LTD
    WRITTEN 2021
 */

// Structure to hold SAP mobility info
public struct MobilityTask
{
    public string Name { set; get; }
    public string CEL { set; get; }
    public string Group { set; get; }
    public string Counter { set; get; }
    public string Workcentre { set; get; }
    public string WBS { get; set; }
    public string Module { get; set; }
    public string LongText { set; get; }
    public string SalesDoc { get; set; }
    public string SalesDocItem { get; set; }
    public string Equipment1 { set; get; }
    public string Equipment2 { set; get; }
    public string Equipment3 { set; get; }
    public string Equipment4 { set; get; }
    public string Equipment5 { set; get; }
    public string Equipment6 { set; get; }
    public string Equipment7 { set; get; }
    public string Equipment8 { set; get; }
    public string Equipment9 { set; get; }
    public string Equipment10 { set; get; }
    public string FolderPath { set; get; }
    public bool WarrantyClaim { get; set; }
};

public struct QuoteInformation
{
    public string Equipment { get; set; }
    public string QuoteTemplate { get; set; }
    public string PartsTemplate { get; set; }
    public string MandPartsTemplate { get; set; }
};

public struct MobilityMeasurement
{
    public string Number { get; set; }
    public string Position { get; set; }
    public string Description { get; set; }
    public string CharCode { get; set; }
    public string CodeGroup { get; set; }
    public string Decimals { get; set; }
    public string TargetValue { get; set; }
    public string IsValueCodeSuff { get; set; }
    public string IsCounter { get; set; }
    public string LowerLimit { get; set; }
    public string UpperLimit { get; set; }
    public string TargetText { get; set; }

};

public struct MobilityEquipment
{
    public string SerialNumber { set; get; }
    public string EquipmentNumber { set; get; }
    public string Description { set; get; }
    public string ObjectType { set; get; }
    public string CatProfile { set; get; }
    public string SalesOrg { set; get; }
    public string DistChan { set; get; }
    public string Division { set; get; }
    public string EntryListName { set; get; }
    public string FunctionLoc { set; get; }
    public string ZAWA { set; get; }
    public string ZDI1 { set; get; }
    public string ModelNumber { set; get; }
    public string TemplateName { set; get; }
    public string TemplateNumber { set; get; }
    public bool UpdateToTemplate { set; get; }
};

public struct MobilityServiceOrder
{
    public string Description { set; get; }
    public string NotificationType { set; get; }
    public string BasicStartDate { set; get; }
    public string BasicEndDate { set; get; }
    public string ActualStartDate { set; get; }
    public string ActualEndDate { set; get; }
    public string OrderType { set; get; }
    public string PurchaseOrder { set; get; }
    public string PurchaseOrderDate { set; get; }
    public string SoldToParty { set; get; }
    public string CustomerName { set; get; }
    public string Priority { set; get; }
    public string ActivityType { set; get; }
    public string SalesGroup { set; get; }
    public string SalesOffice { set; get; }
    public string ExternalReference { set; get; }
    public string Plant { set; get; }
    public string PlannerGroup { set; get; }
    public string Location { set; get; }
    public bool   CreateWithoutEq { set; get; }
};

namespace DigitalVPM_DataCollector
{
    public class AUTOSAP
    {
        // Internal class control objects
        public static GuiSession       SapSession    { get; set; }
        public static GuiMainWindow    SapWindow     { get; set; }
        public static GuiApplication   SapApp        { get; set; }
        public static GuiConnection    SapConnection { get; set; }
        public static bool             SapConnected  { get; set; }

        // Find any open SAP session, return false if not open
        public bool GetSession()
        {
            // Create new instance of ROT wrapper to attach to SAPGUI
            CSapROTWrapper sapROT = new CSapROTWrapper();

            // Attach to GUI
            object sapGui = sapROT.GetROTEntry("SAPGUI");

            if (sapGui == null) return false;

            // Find & attach to scripting engine
            try
            {
                SapApp = (sapGui.GetType().InvokeMember("GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, sapGui, null) as GuiApplication);
            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
                return false;
            }

            // Find connection to scripting engine
            for (int i = 0; i< SapApp.Children.Count; i++)
            {
                SapConnection = (SapApp.Children.ElementAt(i) as GuiConnection);
            }

            // Verify the connection
            if (SapConnection == null)
            {
                SapConnected = false;
                
                return false;
            }

            // Get SAP session
            SapSession = (GuiSession)SapConnection.Children.ElementAt(0);

            // Verify the sesssion
            if (SapSession == null)
            {
                SapConnected = false;
                return false;
            }

            // Get SAP window
            SapWindow = (GuiMainWindow)SapSession.Children.ElementAt(0);

            // Success
            SapConnected = true;
            return true;
        }

        // Active window
        public GuiFrameWindow GetActiveWindow()
        {
            return SapSession.ActiveWindow;
        }

        // Get connected
        public bool CheckConnected()
        {
            return SapConnected;
        }

        // Get ID
        public dynamic GetFormById(string FormID)
        {
            if (SapConnected)
            {
                return SapSession.FindById(FormID);
            }
            else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get by name
        public dynamic GetFormByName(string FormName, string FormType)
        {
            if (SapConnected)
            {
                return SapWindow.FindByName(FormName, FormType);
            }
            else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get shell
        public GuiShell GetShell(string ShellName)
        {
            GuiShell Shell;

            if (SapConnected)
            {
                try
                {
                    Shell = (GuiShell)SapWindow.FindByName(ShellName, "GuiShell");
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return Shell;
            }
            else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Start a SAP transaction
        public void StartTransaction(string transCode)
        {
            if (SapConnected)
            { 
                SapSession.StartTransaction(transCode);

                string statusTxt = ((GuiStatusbar)SapWindow.FindById("sbar")).Text;

                if (statusTxt.Contains("You are not authorized to use transaction"))
                {
                    Console.WriteLine($"You are not authorized to use transaction {transCode}.");
                    return;
                }
            }
                else
            {
                Console.WriteLine("Please ensure SAP is running");
                return;
            }
        }

        // Change store location
        public bool ChangeStore(string NewStoreLocation, string NewSpecialStock, BackgroundWorker Parent)
        {
            string                 PrevLoc, PrevStock                  = "";
            GuiComponentCollection StoreLocations, SpecialStock        = null;
            GuiCTextField          CurrentLoc, CurrentStock, Check     = null;
            GuiTableControl        ComponentTable                      = null;
            int VisibleRow, RowJump;
            double Progress = 0.0;

            VisibleRow = 0;
            RowJump = 0;


            if (SapConnected) {

                for (int i = 0; i < 9999; i++) 
                {
                    // Re init collection
                    ComponentTable = ((GuiTableControl)SapWindow.FindByName("SAPLCOMKTCTRL_3020", "GuiTableControl"));
                    StoreLocations = ComponentTable.FindAllByName("RESBD-LGORT", "GuiCTextField");
                    SpecialStock   = ComponentTable.FindAllByName("RESBD-SOBKZ_D", "GuiCTextField");

                    // Try to change
                    if (StoreLocations.Item(VisibleRow).Type == "GuiCTextField") 
                    {
                        // Get field collections
                        CurrentLoc   = ((GuiCTextField)StoreLocations.Item(VisibleRow));
                        CurrentStock = ((GuiCTextField)SpecialStock.Item(VisibleRow));
                        Check        = ((GuiCTextField)ComponentTable.FindAllByName("RESBD-MATNR", "GuiCTextField").Item(VisibleRow));

                        // Exit on blank
                        if (Check.Text == "") break;

                        // Get current value in case of error
                        PrevLoc   = CurrentLoc.Text;
                        PrevStock = CurrentStock.Text;

                        // Change to updated value
                        CurrentLoc.Text     = NewStoreLocation;
                        CurrentStock.Text   = NewSpecialStock;

                        // Revert if failed
                        if (!ClearErrors(5, false)){
                            // Re-connect to fields
                            ComponentTable          = ((GuiTableControl)SapWindow.FindByName("SAPLCOMKTCTRL_3020", "GuiTableControl"));
                            StoreLocations          = ComponentTable.FindAllByName("RESBD-LGORT", "GuiCTextField");
                            SpecialStock            = ComponentTable.FindAllByName("RESBD-SOBKZ_D", "GuiCTextField");
                            CurrentLoc              = ((GuiCTextField)StoreLocations.Item(VisibleRow));
                            CurrentStock            = ((GuiCTextField)SpecialStock.Item(VisibleRow)); 
                            CurrentLoc.Text         = PrevLoc;
                            CurrentStock.Text       = PrevStock;

                            ClearErrors(5, false);
                        }

                        if(Parent.CancellationPending) { return false; }

                        // Re-init table
                        ComponentTable = ((GuiTableControl)SapWindow.FindByName("SAPLCOMKTCTRL_3020", "GuiTableControl"));

                        RowJump = StoreLocations.Count - 1;

                        // Update progress
                        Progress = Convert.ToDouble(VisibleRow) / Convert.ToDouble(RowJump);

                        VisibleRow++;

                        // Move scroll bar if we go out of range
                        if (VisibleRow > RowJump){
                            
                            ComponentTable.VerticalScrollbar.Position += RowJump;
                            VisibleRow = 0;
                            
                        }
                    }
                }

                return true;
                
            }
            else {
                Console.WriteLine("Please ensure SAP is running");
                return false;
            }
        }

        // Window
        public GuiMainWindow GetWindow() { return SapWindow; }

        // Application
        public GuiApplication GetApp() { return SapApp; }

        // Info
        public GuiSessionInfo GetSessionInfo() { return SapSession.Info; }

        // Sessionob
        public GuiSession GetSessionObj() { return SapSession; }

        // Close trans
        public void EndTransaction() { SapSession.EndTransaction(); }

        // Get all fields by name & type
        public GuiComponentCollection GetAllByName(string fieldName, string fieldType)
        {
            GuiComponentCollection fieldComponentCollection;

            if (SapConnected)
            {
                try
                {
                    fieldComponentCollection = SapWindow.FindAllByName(fieldName, fieldType);
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return fieldComponentCollection;
            }
                else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get GuiComboBox by name
        public GuiComboBox GetComboBox(string fieldName)
        {
            GuiComboBox comboBox;

            if (SapConnected)
            {
                // Try to find by name and catch any exceptions
                try
                {
                    comboBox = (GuiComboBox)SapWindow.FindByName(fieldName, "GuiComboBox");
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return comboBox;
            }
            else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get CTextfield by ID
        public GuiCTextField GetCTextFieldByID(string Address)
        {
            GuiCTextField textField;

            if (SapConnected) {
                // Try to find by name and catch any exceptions
                try {
                    textField = (GuiCTextField)SapWindow.FindById(Address);
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return textField;
            }
            else {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get Textfield by ID
        public GuiTextField GetTextFieldByID(string Address) {
            GuiTextField textField;

            if (SapConnected) {
                // Try to find by name and catch any exceptions
                try {
                    textField = (GuiTextField)SapWindow.FindById(Address);
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return textField;
            }
            else {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get Button by ID
        public GuiButton GetButtonByID(string Address) {
            GuiButton textField;

            if (SapConnected) {
                // Try to find by name and catch any exceptions
                try {
                    textField = (GuiButton)SapWindow.FindById(Address);
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return textField;
            }
            else {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get CTextField by name
        public GuiCTextField GetCTextField(string fieldName)
        {
            GuiCTextField textField;

            if (SapConnected)
            {
                // Try to find by name and catch any exceptions
                try
                {
                    textField = (GuiCTextField)SapWindow.FindByName(fieldName, "GuiCTextField");
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return textField;
            }
                else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get TextField by name
        public GuiTextField GetTextField(string fieldName)
        {
            GuiTextField textField;

            if (SapConnected)
            {
                // Try to find by name and catch any exceptions
                try
                {
                    textField = (GuiTextField)SapWindow.FindByName(fieldName, "GuiTextField");
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return textField;
            }
                else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get GuiButton
        public GuiButton GetButton(string fieldName)
        {
            GuiButton button;

            if (SapConnected)
            {
                // Try to find by name and catch any exceptions
                try
                {
                    button = (GuiButton)SapWindow.FindByName(fieldName, "GuiButton");
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return button;
            }
                else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get GuiTable
        public GuiTableControl GetTable(string fieldName)
        {
            GuiTableControl table;

            if (SapConnected)
            {
                // Try to find by name and catch any exceptions
                try
                {
                    table = (GuiTableControl)SapWindow.FindByName(fieldName, "GuiTableControl");
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return table;
            }
                else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get GuiGridView
        public GuiGridView GetGridView(string fieldName)
        {
            GuiGridView gridView;

            if (SapConnected)
            {
                // Try to find by name and catch any exceptions
                try
                {
                    gridView = (GuiGridView)SapWindow.FindByName(fieldName, "GuiGridView");
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return gridView;
            }
                else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get GuiCheckBox
        public GuiCheckBox GetCheckBox(string fieldName)
        {
            GuiCheckBox checkBox;

            if (SapConnected)
            {
                // Try to find by name and catch any exceptions
                try
                {
                    checkBox = (GuiCheckBox)SapWindow.FindByName(fieldName, "GuiCheckBox");
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return checkBox;
            }
                else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Get GuiRadioButton
        public GuiRadioButton GetRadioButton(string fieldName)
        {
            GuiRadioButton radioButton;

            // Try to find by name and catch any exceptions
            if (SapConnected)
            {
                try
                {
                    radioButton = (GuiRadioButton)SapWindow.FindByName(fieldName, "GuiRadioButton");
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return radioButton;
            }
            else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }

        }

        // Get GuiTab
        public GuiTab GetTab(string fieldName)
        {
            GuiTab tab;

            if (SapConnected)
            {
                // Try to find by name and catch any exceptions
                try
                {
                    tab = (GuiTab)SapWindow.FindByName(fieldName, "GuiTab");
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
                return tab;
            }
                else
            {
                Console.WriteLine("Please ensure SAP is running");
                return null;
            }
        }

        // Send key stroke
        public void SendVKey(int keyId)
        {

            if (SapConnected)
            {
                try
                {
                    SapWindow.SendVKey(keyId);
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }
            }
            else
            {
                Console.WriteLine("Please ensure SAP is running");
                return;
            }

        }

        public bool ClearErrors(int RetryAmount, bool ShowError)
        {
            int Retrys = 0;

            if (SapConnected)
            {
                SendVKey(0);

                while (SapSession.ActiveWindow.Type != "GuiMainWindow")
                {
                    string WindowText = SapSession.ActiveWindow.Text;

                    if (WindowText.Contains("Priority Change"))
                    {
                        GuiButton button = (GuiButton)SapSession.ActiveWindow.FindByName("BUTTON_2", "GuiButton");
                        button.Press();
                    }

                    if (WindowText.Contains("Delete partner(s)"))
                    {
                        GuiButton button = (GuiButton)SapSession.ActiveWindow.FindByName("POP-OPTION1", "GuiButton");
                        button.Press();
                    }

                    SendVKey(0);
                    SendVKey(0);

                    Retrys++;

                    // Check how many times
                    if (Retrys > RetryAmount)
                    {
                        if(ShowError) Console.WriteLine("Could not clear popup error/warning");
                        break;
                    }
                }

                // Check status bar
                GuiStatusbar status = (GuiStatusbar)SapWindow.FindById("sbar");

                // Reset retry
                Retrys = 0;

                while (status.Text != "")
                {
                    SendVKey(0);
                    Retrys++;

                    // Check how many times
                    if (Retrys > RetryAmount)
                    {
                        if (ShowError) Console.WriteLine("Could not clear status bar error/warning");
                        return false;
                    }
                }
            }
            else
            {
                if (ShowError) Console.WriteLine("Please ensure SAP is running");
                return false;
            }

            return true;

        }

        // Set variant
        public string SetVariant(string VarName)
        {
            if (SapConnected)
            {
                GetButton("btn[17]").Press();

                ((GuiTextField)SapSession.FindById("wnd[1]/usr/txtV-LOW")).Text = VarName;
                ((GuiTextField)SapSession.FindById("wnd[1]/usr/ctxtENVIR-LOW")).Text = "";
                ((GuiTextField)SapSession.FindById("wnd[1]/usr/txtENAME-LOW")).Text = "";
                ((GuiTextField)SapSession.FindById("wnd[1]/usr/txtAENAME-LOW")).Text = "";
                ((GuiTextField)SapSession.FindById("wnd[1]/usr/txtMLANGU-LOW")).Text = "";

                SendVKey(8);

                return "";

            }
                else
            {
                Console.WriteLine("Please ensure SAP is running");
                return "SAP is not connected";
            }
        }

        // Create measurement point
        public string CreateMeasurementPoint(MobilityMeasurement NewMeasurement, string EquipmentNumber)
        {
            if (SapConnected)
            {
                SapSession.StartTransaction("IK01");

                string temp = GetCTextField("RIMR0-MPOTY").Text;

                GetCTextField("RIMR0-MPOTY").Text = "IEQ";
                if (temp != "IEQ") { SendVKey(0); }
                GetCTextField("EQUI-EQUNR").Text = EquipmentNumber;
                GetCTextField("IMPT-MPTYP").Text = "Z";

                SendVKey(0);

                // Check results
                switch (GetSessionInfo().ScreenNumber)
                {
                    // Equipment locked
                    case 1110:
                        return "Could not create measurement point";

                    // All good
                    default:
                        GetTextField("IMPT-PSORT").Text  = NewMeasurement.Position;
                        GetCTextField("IMPT-ATNAM").Text = NewMeasurement.CharCode;
                        GetCTextField("IMPT-CODGR").Text = NewMeasurement.CodeGroup;

                        ClearErrors(30, true);

                        GetTextField("IMPT-PTTXT").Text  = NewMeasurement.Description;
                        GetTextField("IMPT-DECIM").Text  = NewMeasurement.Decimals;
                        GetTextField("RIMR0-DESIC").Text = NewMeasurement.TargetValue;
                        GetTextField("IMPT-DSTXT").Text  = NewMeasurement.TargetText;

                        ClearErrors(30, true);

                        // Check boxes
                        if (NewMeasurement.IsValueCodeSuff == "X") GetCheckBox("IMPT-INDCT").Selected = true;
                        if (NewMeasurement.IsCounter == "X")       GetCheckBox("IMPT-CDSUF").Selected = true;

                        // Goto upper and lower limits
                        GetButton("DETAIL").Press();
                        ((GuiTextField)SapSession.FindById("wnd[1]/usr/txtRIMR0-MRMAC")).Text = NewMeasurement.UpperLimit;
                        ((GuiTextField)SapSession.FindById("wnd[1]/usr/txtRIMR0-MRMIC")).Text = NewMeasurement.LowerLimit;
                        ((GuiButton)SapSession.FindById("wnd[1]/tbar[0]/btn[0]")).Press();

                        // Save
                        GetButton("btn[11]").Press();

                        break;

                }

                return "";
            }
            else
            {
                Console.WriteLine("Please ensure SAP is running");
                return "SAP is not connected";
            }

        }

        // Change measurement point
        public string ChangeMeasurementPoint(MobilityMeasurement ExistingMeasurement, string Number)
        {
            if (SapConnected)
            {
                StartTransaction("IK02");
                GetCTextField("IMPT-POINT").Text = Number;
                SendVKey(0);
                // Clear decimals and description
                GetTextField("IMPT-DECIM").Text = "";
                GetTextField("RIMR0-DESIC").Text = "";
                SendVKey(0);
                SendVKey(0);
                // Clear upper and lower limits
                GetButton("DETAIL").Press();
                ((GuiTextField)SapSession.FindById("wnd[1]/usr/txtRIMR0-MRMAC")).Text = "";
                ((GuiTextField)SapSession.FindById("wnd[1]/usr/txtRIMR0-MRMIC")).Text = "";
                ((GuiButton)SapSession.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                // Save
                GetButton("btn[11]").Press();

                // Enter data
                StartTransaction("IK02");
                GetCTextField("IMPT-POINT").Text = Number;
                SendVKey(0);

                // Update position
                if(ExistingMeasurement.Position != GetTextField("IMPT-PSORT").Text){
                    GetTextField("IMPT-PSORT").Text = ExistingMeasurement.Position;
                    SendVKey(0);
                    SendVKey(0);
                }

                GetTextField("IMPT-PTTXT").Text = ExistingMeasurement.Description;
                GetTextField("IMPT-DECIM").Text = ExistingMeasurement.Decimals;
                GetCTextField("IMPT-CODGR").Text = ExistingMeasurement.CodeGroup;
                GetTextField("RIMR0-DESIC").Text = ExistingMeasurement.TargetValue;
                SendVKey(0);
                GetTextField("IMPT-DSTXT").Text = ExistingMeasurement.TargetText;
                if (ExistingMeasurement.IsCounter == "X") GetCheckBox("IMPT-INDCT").Selected = true;
                if (ExistingMeasurement.IsValueCodeSuff == "X") GetCheckBox("IMPT-CDSUF").Selected = true;

                // Upper lower limit
                GetButton("DETAIL").Press();
                ((GuiTextField)SapSession.FindById("wnd[1]/usr/txtRIMR0-MRMAC")).Text = ExistingMeasurement.UpperLimit;
                ((GuiTextField)SapSession.FindById("wnd[1]/usr/txtRIMR0-MRMIC")).Text = ExistingMeasurement.LowerLimit;
                ((GuiButton)SapSession.FindById("wnd[1]/tbar[0]/btn[0]")).Press();
                // Save
                GetButton("btn[11]").Press();

            }
            else
            {
                Console.WriteLine("Please ensure SAP is running");
                return "SAP is not connected";
            }

            return "";
        }

        // De-active measurement
        public void DeactiveMeasurement(string Number) 
        {
            if (SapConnected) 
            {
                SapSession.StartTransaction("IK02");
                GetCTextField("IMPT-POINT").Text = Number;
                SendVKey(0);
                string Title = GetTextField("IMPT-PTTXT").Text;

                if (Title.Length > 30) {
                    GetTextField("IMPT-PTTXT").Text = Title.Substring(0, 30) + "_OLD";
                } else {
                    GetTextField("IMPT-PTTXT").Text = Title + "_OLD";
                }

                ((GuiMenu)GetFormById("wnd[0]/mbar/menu[0]/menu[7]/menu[0]/menu[0]")).Select();
                GetButton("btn[11]").Press();
            }
        }

   

        public string CreateNewEquipment(string TemplateEq, string SerialNumber, string MaterialNumber, string FuncLoc)
        {
            if (SapConnected)
            {
                StartTransaction("IE01");

                GetCTextField("RM63E-EQUNR").Text = "";
                GetCTextField("RM63E-EQTYP").Text = "S";
                GetCTextField("RM63E-REFEQ").Text = TemplateEq;
                GetCTextField("RM63E-REFMA").Text = "";
                SendVKey(0);
                ((GuiButton)SapSession.FindById("wnd[1]/tbar[0]/btn[30]")).Press();
                ((GuiButton)SapSession.FindById("wnd[1]/tbar[0]/btn[30]")).Press();
                ((GuiCheckBox)SapSession.FindById("wnd[1]/usr/chkRILOCP-CPMPT")).Selected = true;
                ((GuiCheckBox)SapSession.FindById("wnd[1]/usr/chkRILOCP-CPMPX")).Selected = true;
                ((GuiCheckBox)SapSession.FindById("wnd[1]/usr/chkRILOCP-CPMPC")).Selected = true;
                ((GuiCheckBox)SapSession.FindById("wnd[1]/usr/chkRILOCP-CPMPD")).Selected = true;
                ((GuiCheckBox)SapSession.FindById("wnd[1]/usr/chkRILOCP-CPKON")).Selected = true;
                ((GuiCheckBox)SapSession.FindById("wnd[1]/usr/chkRILOCP-J_3GCPBAU")).Selected = true;
                ((GuiButton)SapSession.FindById("wnd[1]/tbar[0]/btn[7]")).Press();

                GetTab(@"T\07").Select();
                GetTextField("ITOB-SERNR").Text = SerialNumber;
                GetCTextField("ITOB-MATNR").Text = MaterialNumber;

                GetButton("btn[11]").Press();
                EndTransaction();

                StartTransaction("IE02");

                return GetCTextField("RM63E-EQUNR").Text;

            }
            else
            {
                Console.WriteLine("Please ensure SAP is running");
                return "SAP is not connected";
            }
        }

        public bool CheckMaterialSerial(string serial, string material)
        {
            if (SapConnected)
            {
                SapSession.StartTransaction("IQ02");

                ((GuiCTextField)SapWindow.FindByName("RISA0-MATNR", "GuiCTextField")).Text = material;
                ((GuiCTextField)SapWindow.FindByName("RISA0-SERNR", "GuiCTextField")).Text = serial;

                SapWindow.SendVKey(0);

                if (SapSession.Info.ScreenNumber != 1000)
                {
                    if (((GuiCTextField)SapWindow.FindByName("ITOBATTR-EQUNR","GuiCTextField")).Text == "")
                    {
                        Console.WriteLine("This equipment serial number is linked to a material number in SAP but no equipment, please convert the material serial number to an equipment number to continue");
                        return false;
                    }
                    return true; 
                }

                return true;
            }
            else
            {
                Console.WriteLine("Please ensure SAP is running");
                return false;
            }
        }

        public DataTable GetMeasurementPoints(string InputEquipment)
        {
            ExcelDataTables TableManager    = new ExcelDataTables();

            if (SapConnected)
            {
                StartTransaction("IK07");
                SetVariant("/FSDS-25-123");
                GetCTextField("EQUNR-LOW").Text = InputEquipment;
                SendVKey(8);

                switch (GetSessionInfo().ScreenNumber) {
                    // No measurements
                    case 1000: return null;

                    // Single point
                    case 5110:
                        return null;

                    // List view
                    default:

                        // Build unique filename
                        string FilePath = $"{Environment.GetEnvironmentVariable("TEMP")}\\";
                        string FileName = $"IK08DUMP{InputEquipment}{DateTime.Now.Year}{DateTime.Now.Month}{DateTime.Now.Day}{DateTime.Now.Hour}{DateTime.Now.Minute}{DateTime.Now.Second}.txt";

                        GuiMenu Menu = (GuiMenu)SapWindow.FindById("wnd[0]/mbar/menu[0]/menu[10]/menu[2]");

                        Menu.Select();

                        // Text file format popup
                        GuiModalWindow Popup = (GuiModalWindow)SapSession.FindById("wnd[1]");
                        ((GuiRadioButton)Popup.FindById("usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]")).Select();
                        ((GuiButton)Popup.FindById("tbar[0]/btn[0]")).Press();

                        // Filename popup
                        Popup = (GuiModalWindow)SapSession.FindById("wnd[1]");
                        ((GuiCTextField)Popup.FindByName("DY_PATH", "GuiCTextField")).Text      = FilePath;
                        ((GuiCTextField)Popup.FindByName("DY_FILENAME", "GuiCTextField")).Text  = FileName;
                        ((GuiButton)Popup.FindById("tbar[0]/btn[11]")).Press();

                        return TableManager.ControlTexttoDataTable($"{FilePath}{FileName}");

                }
            }
            else
            {
                return null;
            }
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        // End session
        public void EndSession()
        {
            SapConnection.CloseConnection();
            SapConnection   = null;
            SapApp          = null;
            SapWindow       = null;
            SapConnected    = false;
        }
    }
}
