using ClosedXML.Excel;
using ExportToExcel.RMSServiceReference;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExportToExcel
{
    public partial class Form1 : Form
    {
        //add the service reference 
        PublicServiceClient client = new PublicServiceClient();

        private string _Token;
        private PublicServiceClient _PublicServiceClient;

        string _fileName="";
        string _timeStamp="";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Get_Connection();
        }

        public void Get_Connection()
        {

            try
            {
                label1.Text = "";

                ConnectionInfoRequest connectionInfoRequest = new ConnectionInfoRequest();
                ConnectionInfoResponse connectionInfoResponse = new ConnectionInfoResponse();


                connectionInfoRequest.AgentId = 15;
                connectionInfoRequest.AgentPassword = "1h&29$vk449f8";
                connectionInfoRequest.RMSClientNo = 3038;
                connectionInfoRequest.ClientPassword = "X!b69rN*";

                //pass conInfo and get info and store to a variable 
                connectionInfoResponse = client.GetConnectionInfo(connectionInfoRequest);


                // 'The secret token for this Client/Agent combination. This should be refreshed every hour
                _Token = connectionInfoResponse.Token;

                // Create an instance of the public service client with the correct endpoint URL for this client
                _PublicServiceClient = GetPublicWebServiceClientInstance(connectionInfoResponse.WebserviceURL);


                label1.Text = client.TestCall();
            }

            catch (Exception ex)
            {
                label1.Text = "Error:" + ex.Message;
            }
        }

        private PublicServiceClient GetPublicWebServiceClientInstance(string sWebserviceURL)
        {
            BasicHttpBinding oBinding;
            EndpointAddress oEndpointAddress;

            oEndpointAddress = new EndpointAddress(sWebserviceURL);

            oBinding = new BasicHttpBinding();
            {
                var withBlock = oBinding;
                withBlock.Security.Mode = BasicHttpSecurityMode.Transport;
                withBlock.MaxReceivedMessageSize = 20000000;
                withBlock.MaxBufferSize = 20000000;
                withBlock.ReaderQuotas.MaxArrayLength = 20000000;

                withBlock.ReceiveTimeout = new TimeSpan(1, 0, 0, 0);
                withBlock.SendTimeout = new TimeSpan(1, 0, 0, 0);
            }

            return new PublicServiceClient(oBinding, oEndpointAddress);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DataSource = null;//clear datagridview1

                label1.Text = "";

                if (_fileName != "" || _timeStamp != "") {
                    _fileName = "";
                    _timeStamp = "";
                }

                string text = DateTime.Now.ToString();
               
                _timeStamp = text.Replace("/", "_").Replace(":", "_");
                _fileName = "ReservationsDetailsOn_";

                ResRequest oRequest = new ResRequest();
                ResResult oResponse = new ResResult();

                // Build the request to get some reservations
                // All of the filters will be applied together
                {
                    var withBlock = oRequest;
                    withBlock.ResIdFrom = 0;
                    withBlock.ResIdTo = 10000;
                    withBlock.ListOfPropertyIds = new int[] { 1, 2, 3, 4 };
                   

                    // Specify some optional data to be populated
                    withBlock.ResOptionalFieldList = new OptionalFieldsRes();
                    {
                        var withBlock1 = withBlock.ResOptionalFieldList;
                        withBlock1.Company = true;
                        withBlock1.AccountBalance = true;
                        
                    }
                }

                // Get the data from the server
                oResponse = _PublicServiceClient.GetListOfReservations(_Token, oRequest);

                dataGridView1.DataSource = oResponse.ListOfRes;
                var FetchedData = oResponse.ListOfRes;

                if (oResponse != null && oResponse.ListOfRes != null) { 
                    if (FetchedData.Count() == 0) {
                        label1.Text = "Get Data Success: No reservations found!";
                    }
                    else {
                        label1.Text = "Get Data Success:" + FetchedData.Count() + " reservations found!";
                    }
                }
                else
                    label1.Text = " No Reservations found.";
            }
            catch (Exception ex)
            {
                label1.Text = "Error: " + ex.Message;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
           // Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xls)|*.xls";
            sfd.FileName = _fileName + _timeStamp+".xls";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                // Copy DataGridView results to clipboard
                copyAlltoClipboard();

                object misValue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Application xlexcel = new Microsoft.Office.Interop.Excel.Application();

                xlexcel.DisplayAlerts = false; // Without this you will get two confirm overwrite prompts
                Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                // Format column D as text before pasting results, this was required for my data
                Excel.Range rng = xlWorkSheet.get_Range("D:D").Cells;
                rng.NumberFormat = "@";

                // Paste clipboard results to worksheet range
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                // For some reason column A is always blank in the worksheet. ¯\_(ツ)_/¯
                // Delete blank column A and select cell A1
                Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                delRng.Delete(Type.Missing);
                xlWorkSheet.get_Range("A1").Select();

                // Save the excel file under the captured location from the SaveFileDialog
                xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlexcel.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);

                // Clear Clipboard and DataGridView selection
                Clipboard.Clear();
                dataGridView1.ClearSelection();

                // Open the newly saved excel file
                if (File.Exists(sfd.FileName))
                    System.Diagnostics.Process.Start(sfd.FileName);

                dataGridView1.DataSource = null;//clear datagridview1
                label1.Text = "Cleared all data from this view!";
            }
        }

        private void copyAlltoClipboard()
        {
            
            dataGridView1.RowHeadersVisible = true;
            dataGridView1.SelectAll();
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DataSource = null;//clear datagridview1

                label1.Text = "";
                if (_fileName != "" || _timeStamp != "")
                {
                    _fileName = "";
                    _timeStamp = "";
                }

                string text = DateTime.Now.ToString();

                _timeStamp = text.Replace("/", "_").Replace(":", "_");
                _fileName = "ClientsDetailsOn_";

                ClientRequest oRequest = new ClientRequest();
                ClientResult oResponse = new ClientResult();

                // Build the request to get some clients
                // All of the filters will be applied together
                {
                    var withBlock = oRequest;
                    withBlock.ClientIdFrom = 0;
                    withBlock.ClientIdTo = 10000;
                    withBlock.ListOfPropertyIds = new int[] { 1, 2, 3, 4 };

                    // Specify some optional data to be populated
                    withBlock.ClientOptionalFieldList = new OptionalFieldsClient();
                    {
                        var withBlock1 = withBlock.ClientOptionalFieldList;
                        withBlock1.Company = true;
                        withBlock1.AccountBalance = true;
                    }
                }

                // Get the data from the server
                oResponse = _PublicServiceClient.GetListOfClients(_Token, oRequest);
                dataGridView1.DataSource = oResponse.ListOfClients;

                label1.Text = "Get Data Success:";

                if (oResponse != null && oResponse.ListOfClients != null) {

                    if (oResponse.ListOfClients.Count() == 0) {
                        label1.Text += " " + "No clients found!";
                    }
                    else {
                        label1.Text += " " + oResponse.ListOfClients.Count() + " clients found.";
                    }
                   
                }
                  
                else
                    label1.Text += " No clients found.";
            }
            catch (Exception ex)
            {
                label1.Text = "Error: " + ex.Message;
            }
        }
    }
}
