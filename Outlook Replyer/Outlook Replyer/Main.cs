using System;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.IO;    // for StreamReader
using System.Collections;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Data;
using System.Data.OleDb;
using System.Net.Mail;
using System.Linq;

namespace Outlook_Replyer
{
    public partial class Main : Form
    {
        
        DataTable dtMails = new DataTable();
        DataTable dt = new DataTable();
        Outlook.Application app;
        Outlook.NameSpace nameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Explorer explorer;
        internal static string pwd;

        private static string CriaJson(string description, string summary, string Category, string name, string analista, string subCategory, string subCategory2, string subCategory3, string subCategory4)
        {
            string postData =
               "{\"fields\":{\"issuetype\":{\"name\":\"Service Desk\"},\"project\":{\"key\":\"SC\"},\"summary\":\"" + summary + "\",\"description\":\"" + description + "\",\"customfield_10601\":{\"value\":\"No\"},\"customfield_11005\":{\"value\":\"No\"},\"customfield_10600\":{\"name\":\"" + name + "\"},\"customfield_10700\":\"" + Category + "\",\"customfield_10701\":\"" + subCategory + "\",\"customfield_10702\":\"" + subCategory2 + "\",\"customfield_10703\":\"" + subCategory3 + "\",\"customfield_10902\":\"" + subCategory4 + "\",\"assignee\":{\"name\":\"" + analista + "\"}}}";
            return postData;
        }

        private static string SendRequest(string postdata, string url)// retorna a resposta da request
        {

            try
            {
                // Create the Web request object.
                byte[] byteArray = Encoding.UTF8.GetBytes(postdata); // convert to byte input
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url); //coment
                req.Method = "POST";
                req.ContentType = "application/json";
                req.ContentLength = byteArray.Length;
                string mergedCredentials = string.Format("{0}:{1}", "lpicazio", pwd);
                byte[] byteCredentials = UTF8Encoding.UTF8.GetBytes(mergedCredentials);
                string base64Credentials = Convert.ToBase64String(byteCredentials);
                req.Headers.Add("Authorization", "Basic " + base64Credentials);



                // Get the request stream.
                Stream dataStream = req.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);// Write the data to the request stream.
                dataStream.Close();// Close the Stream object.


                HttpWebResponse response = req.GetResponse() as HttpWebResponse;
                StreamReader streamReader = new StreamReader(response.GetResponseStream());
                string str = streamReader.ReadToEnd();
                Console.WriteLine(str);
                response.Close();
                return str;


            }
            catch (WebException e)
            {
                // Display any errors. In particular, display any protocol-related error.
                if (e.Status == WebExceptionStatus.ProtocolError)

                {
                    HttpWebResponse hresp = (HttpWebResponse)e.Response;
                    Console.WriteLine("\nAuthentication Failed, " + hresp.StatusCode);
                    Console.WriteLine("Status Code: " + (int)hresp.StatusCode);
                    Console.WriteLine("Status Description: " + hresp.StatusDescription);
                    return "Error";
                }
                
                Console.WriteLine("Caught Exception: " + e.Message);

                Console.WriteLine("Stack: " + e.StackTrace);
                return "Error 2"; 
            }


        }

        private static void GetRequest()
        {
            string html = string.Empty;
            string url = @"http://jiracorebr.csintra.net/rest/api/2/issue/SC-33291";
            string mergedCredentials = string.Format("{0}:{1}", "lpicazio", pwd);
            byte[] byteCredentials = UTF8Encoding.UTF8.GetBytes(mergedCredentials);
            string base64Credentials = Convert.ToBase64String(byteCredentials);
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Headers.Add("Authorization", "Basic " + base64Credentials);


            using (HttpWebResponse response = (HttpWebResponse)req.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                html = reader.ReadToEnd();
            }
            Console.WriteLine(html);
        }



        private static void displayPageContent(Stream ReceiveStream)
        {

            // Create an ASCII encoding object.

            Encoding ASCII = Encoding.ASCII;

            // Define the byte array to temporarily hold the current read bytes.


            Byte[] read = new Byte[512];
            Console.WriteLine("\r\nPage Content...\r\n");
            // Read the page content and display it on the console.
            // Read the first 512 bytes.
            int bytes = ReceiveStream.Read(read, 0, 512);
            while (bytes > 0)
            {
                Console.Write(ASCII.GetString(read, 0, bytes));
                bytes = ReceiveStream.Read(read, 0, 512);
            }
            Console.WriteLine("");
        }



        public Main()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                if (textBox1.Text == "")
                {
                    MessageBox.Show("Campo summary obrigatório");
                    return;
                }
                    
                Outlook.MailItem mail = explorer.Selection[1];
                Outlook.MailItem resp = mail.ReplyAll();

                string ReqBy = RequestBy(mail);

                string text = cortaBody(resp.Body);

                string description = text;
                string summary = textBox1.Text;
                string Analista = selectAssigne(ReqBy);
                if (Analista == "atd") Analista = "lpicazio";
                string Category = "Incident";
                string SubCategory = "SW - Third Party Systems";
                string SubCategory2 = "Others";
                string SubCategory3 = "N/A";
                string SubCategory4 = "N/A";


                string json = criaJson(description, summary, Category, ReqBy, Analista, SubCategory, SubCategory2, SubCategory3, SubCategory4);
                string url = "http://jiracorebr.csintra.net/rest/api/2/issue";
                string respReq = SendRequest(json, url);

                string nTicket = respReq.Substring(respReq.IndexOf("SC-"), 8); // Tira o SC do ticket
                string sc = "SC-";
                nTicket = nTicket.TrimStart(sc.ToCharArray());

                string texto = "O ticket <b>" + nTicket + "</b> foi aberto para um de nossos analistas verificar.";
                resp.HTMLBody = geraHTML(resp.HTMLBody, texto);
                resp.Display();

            }
            if (radioButton2.Checked == true)
            {
                Outlook.MailItem mail = explorer.Selection[1]; //sim
                string summary = "Shared Drive - Acesso";
                Outlook.MailItem resp = mail.ReplyAll(); //sim
                string description = cortaBody(resp.Body);

                string reqBy = RequestBy(mail);

                string Analista = selectAssigne(reqBy);

                string Category = "Access Request";
                string SubCategory = "Shared Drive";
                string SubCategory2 = "N/A";
                string SubCategory3 = "N/A";
                string SubCategory4 = "N/A";

                string json = criaJson(description, summary, Category, reqBy, Analista, SubCategory, SubCategory2, SubCategory3, SubCategory4);
                string url = "http://jiracorebr.csintra.net/rest/api/2/issue";

                string respReq = SendRequest(json, url);
                string nTicket = respReq.Substring(respReq.IndexOf("SC-"), 8);
                string sc = "SC-";
                nTicket = nTicket.TrimStart(sc.ToCharArray());
                string texto = "O ticket <b>" + nTicket + " </b>foi aberto para um de nossos analistas verificar o acesso, assim que aprovado.<b>LEMBRAR DE COLOCAR APROVADOR</b>";
                resp.HTMLBody = geraHTML(resp.HTMLBody, texto);
                resp.Display();
            }
            if (radioButton3.Checked == true)
            {
                Outlook.MailItem mail = explorer.Selection[1];
                string summary = "File Record  - Gravação";
                Outlook.MailItem resp = mail.ReplyAll();
                string description = cortaBody(resp.Body);

                string name = RequestBy(mail);

                string Analista = selectAssigne(name);

                string Category = "RFS (Request for Service)";
                string SubCategory = "File Record";
                string SubCategory2 = "File Outing - Other";
                string SubCategory3 = "N/A";
                string SubCategory4 = "N/A";

                string json = criaJson(description, summary, Category, name, Analista, SubCategory, SubCategory2, SubCategory3, SubCategory4);
                string url = "http://jiracorebr.csintra.net/rest/api/2/issue";
                string respReq = SendRequest(json, url);
                string nTicket = respReq.Substring(respReq.IndexOf("SC-"), 8);
                string sc = "SC-";
                nTicket = nTicket.TrimStart(sc.ToCharArray());
                string texto = "O ticket<b> " + nTicket + "</b> foi aberto para um de nossos analistas verificar a gravação.";
                resp.HTMLBody = geraHTML(resp.HTMLBody, texto);
                resp.Display();
            }
            if (radioButton4.Checked == true)
            {
                Comment();
                
            }
            if(radioButton5.Checked == true)
            {
                Outlook.MailItem mail = explorer.Selection[1];
                string comment = mail.Body.Substring(0, mail.Body.IndexOf("From:", StringComparison.CurrentCulture));
                int start = mail.Body.IndexOf("Ticket", StringComparison.CurrentCultureIgnoreCase);
                string ticket = mail.Body.Substring(start, 12);
                string nticket = ticket.Substring(6, 6).Trim();
                Outlook.MailItem resp = mail.ReplyAll();
                string text = resp.Body.Replace("\"", "");
                text = text.Replace("\r", "");
                text = text.Replace("\\", "");
                text = text.Replace("\n", "\\n");
                text = text.Replace("\'", "");
                text = text.Replace("\t", "");
                text = text.Substring(text.IndexOf("From"));
                text = text.Substring(0, text.IndexOf("From", 6));
                string json = "{\"update\":{\"comment\":[{\"add\":{\"body\":\""+text+"\"}}]},\"transition\":{\"id\":\"31\"}}";
                SendRequest(json, "http://jiracorebr.csintra.net/rest/api/2/issue/SC-" + nticket + "/transitions");
                string texto = "Ticket <b>" + nticket + "</b> atualizado com a aprovação.";
                resp.HTMLBody = geraHTML(resp.HTMLBody, texto);
                resp.Display();
            }
            textBox1.Clear();

        }

        private string RequestBy(Outlook.MailItem mail)
        {
            if (radioButton8.Checked == true) return mail.Sender.GetExchangeUser().Alias;
            return textBox3.Text.Substring(0, textBox3.Text.IndexOf('-'));
        }

        private string selectAssigne(string name)
        {
            if (radioButton6.Checked == true)
            {
                return AnalistaDe(name, 1);
                
            }
            if (radioButton7.Checked == true)
            {
                return AnalistaDe(name, 2);
                
            }
            return textBox2.Text.Substring(0, textBox2.Text.IndexOf('-'));
        }

        private void Comment()
        {
            Outlook.MailItem mail = explorer.Selection[1];
            string comment = mail.Body.Substring(0, mail.Body.IndexOf("From:", StringComparison.CurrentCulture));
            int start = mail.Body.IndexOf("Ticket", StringComparison.CurrentCultureIgnoreCase);
            string ticket = mail.Body.Substring(start, 12);
            string nticket = ticket.Substring(6, 6).Trim();
            Outlook.MailItem resp = mail.ReplyAll();
            string text = cortaComment( resp.Body);
            string json = "{\"body\":\"" + text + "\"}";
            SendRequest(json, "http://jiracorebr.csintra.net/rest/api/2/issue/SC-" + nticket + "/comment");
            string texto = "Seu comentario foi adicionado ao ticket <b>" + nticket + "</b> para o analista responsável verificar e te dar um retorno sobre o proceso.";
            resp.HTMLBody = geraHTML(resp.HTMLBody, texto);
            resp.Display();
            
        }

        private string cortaComment(string text) //pega apenas ultimo e-mail de todo o body e deixa no formato para ser colocado no ticket
        {
            text = cortaBody(text);
            return text.Substring(0, text.IndexOf("From", 6));// método que deixa no formato
        }

        private string cortaBody(string text) // método que deixa no formato para envio no ticket
        {
            return text.Replace("\"", "").
            Replace("\r", "").
            Replace("\\", "").
            Replace("\n", "\\n").
            Replace("\'", "").
            Replace("\t", "").
            Substring(text.IndexOf("From"));
        }

        private string geraHTML(string html, string texto) // Gera body do e-mail para ser feito a resposta.
        {
            return html.Substring(0, html.IndexOf("<div class=WordSection1>")) + 
                @"<p class=MsoNormal>
			<span lang=PT-BR style='color:#1F497D;mso-ansi-language:PT-BR'>
				" + hora() + @",
				<o:p></o:p>
			</span>
		</p>
		<p class=MsoNormal>
			<span lang=PT-BR style='color:#1F497D;mso-ansi-language:PT-BR'>
				<o:p>&nbsp;</o:p>
			</span>
		</p>
		<p class=MsoNormal>
			<span lang=PT-BR style='color:#1F497D;mso-ansi-language:PT-BR'>
				"+texto+ @"
				<o:p></o:p>
			</span>
		</p>
        <p class=MsoNormal>
			<span lang=PT-BR style='color:#1F497D;mso-ansi-language:PT-BR'>
				<o:p>&nbsp;</o:p>
			</span>
		</p>
		" + html.Substring(html.IndexOf("<div>"), html.Length - html.Substring(0, html.IndexOf("<div>")).Length);
        }

        private string criaJson(string description, string summary, string Category, string name, string analista, string subCategory, string subCategory2, string subCategory3, string subCategory4)
        {
            string postData =
                "{\"fields\":{\"issuetype\":{\"name\":\"Service Desk\"},\"project\":{\"key\":\"SC\"},\"summary\":\"" + summary + "\",\"description\":\"" + description + "\",\"customfield_10601\":{\"value\":\"No\"},\"customfield_10600\":{\"name\":\"" + name + "\"},\"customfield_10700\":\"" + Category + "\",\"customfield_10701\":\"" + subCategory + "\",\"customfield_10702\":\"" + subCategory2 + "\",\"customfield_10703\":\"" + subCategory3 + "\",\"customfield_10902\":\"" + subCategory4 + "\",\"customfield_10704\": \"Service Desk\",\"assignee\":{\"name\":\"" + analista + "\"}}}";
            return postData;
        }
        private string hora()
        {
            int h = DateTime.Now.Hour;
            if (h < 12) return "Bom dia";
            if (h >= 12 && h < 18) return "Boa Tarde";
            if (h >= 18) return "Boa Noite";
            return "Prezados";
        }

        private string AnalistaDe(string name, int i)
        {
            try
            {
                var drow = dt.Select("LOGIN =" + "'" + name + "'");
                if ((string)drow[0][i] == "ATENDIMENTO") return "atd";
                var drow2 = dt.Select("Name =" + "'" + drow[0][i].ToString() + "'");
                return drow2[0][3].ToString();
            }
            catch ( Exception e)
            {
                return "Unassigned";
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

            app = new Outlook.Application();
            nameSpace = app.GetNamespace("MAPI");
            explorer = app.ActiveExplorer();
            explorer.SelectionChange += Explorer_SelectionChange;
            dtMails.Columns.Add("From", typeof(string));
            dtMails.Columns.Add("Subject", typeof(string));
            dtMails.Columns.Add("Received", typeof(DateTime));
            dtMails.Columns.Add("ID", typeof(int));

            inbox = nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            
            //int i = 1;

            //foreach (var item in emails)
            //{
            //    try
            //    {
            //        Outlook.MailItem mail = (Outlook.MailItem)item;
            //        DataRow dr = dtMails.NewRow();
            //        dr["From"] = mail.SenderName;
            //        dr["Subject"] = mail.Subject;
            //        dr["Received"] = mail.ReceivedTime;
            //        dr["ID"] = i;
            //        dtMails.Rows.Add(dr);
            //        i++;
            //    }
            //    catch (Exception y)
            //    {
            //        i++;
            //    }

            //}

            //-------------------------------------abre planilha e mostra analista-----------------------------//
            string path = @"\\csao11p20011c\IT\PCHardw\INFRA\Controles\Levantamentos\Controle_de_Hardware_e_Usuários.xlsx";
            string excelConnStr = string.Empty;
            OleDbCommand excelCommand = new OleDbCommand();
            OleDbDataAdapter excelDataAdapter = new OleDbDataAdapter();


            excelConnStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path + ";Extended Properties=Excel 12.0";


            OleDbConnection excelConn = new OleDbConnection(excelConnStr);
            excelConn.Open();

            excelCommand = new OleDbCommand("SELECT * FROM [Usuários$]", excelConn);
            excelDataAdapter.SelectCommand = excelCommand;
            excelDataAdapter.Fill(dt);

            //-------------------------------------abre planilha e mostra analista-----------------------------//
            string[] postSource = dt.AsEnumerable().Select(r => r.Field<string>("LOGIN")).ToArray();
            for( int i = 0; i<postSource.Length; i++)
            {
                if (postSource[i] == null) postSource[i] = "-";
                postSource[i] = postSource[i] + " - " + dt.Rows[i].ItemArray[0];
            }
            
            var source = new AutoCompleteStringCollection();
            source.AddRange(postSource);
            textBox2.AutoCompleteCustomSource = source;
            textBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox2.AutoCompleteSource = AutoCompleteSource.CustomSource;
            textBox3.AutoCompleteCustomSource = source;
            textBox3.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox3.AutoCompleteSource = AutoCompleteSource.CustomSource;

            radioButton1.Checked = true;
            radioButton6.Checked = true;
            radioButton8.Checked = true;
        }

       

        private void Explorer_SelectionChange()
        {
            
            try
            {
                if (explorer.Selection[1] is Outlook.MailItem)
                {
                    Outlook.MailItem mail = explorer.Selection[1];


                    if (mail.Sender.Address.Contains("cn=") && !mail.SenderName.StartsWith("DD") && !mail.SenderName.StartsWith ("GG"))
                    {
                        try
                        {
                            if (label1.InvokeRequired)
                            {
                                textBox1.Invoke(new MethodInvoker(delegate { label1.Text = AnalistaDe(mail.Sender.GetExchangeUser().Alias, 1); }));
                            }
                            if (label2.InvokeRequired)
                            {
                                label2.Invoke(new MethodInvoker(delegate { label2.Text = AnalistaDe(mail.Sender.GetExchangeUser().Alias, 2); }));
                            }
                            if (label4.InvokeRequired)
                            {
                                label2.Invoke(new MethodInvoker(delegate { label4.Text = mail.Sender.GetExchangeUser().Alias; }));
                            }
                        }
                        catch (Exception t)
                        {
                            if (label1.InvokeRequired)
                            {
                                label1.Invoke(new MethodInvoker(delegate { label1.Text = "Unassigned"; }));
                            }
                            if (label2.InvokeRequired)
                            {
                                label2.Invoke(new MethodInvoker(delegate { label2.Text = "Unassigned"; }));
                            }
                            if (label4.InvokeRequired)
                            {
                                label2.Invoke(new MethodInvoker(delegate { label4.Text = mail.Sender.GetExchangeUser().Alias; }));
                            }
                        }
                    }
                }
            }
            catch ( Exception x)
            {
                if (label1.InvokeRequired)
                {
                    label1.Invoke(new MethodInvoker(delegate { label1.Text = "Unassigned"; }));
                }
                if (label2.InvokeRequired)
                {
                    label2.Invoke(new MethodInvoker(delegate { label2.Text = "Unassigned"; }));
                }
                
            }
            
            
            
        }

     
        private void textBox2_MouseClick(object sender, MouseEventArgs e)
        {
            radioButton6.Checked = false;
            radioButton7.Checked = false;
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            radioButton8.Checked = false;
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            textBox2.Clear();
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            textBox2.Clear();
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            textBox3.Clear();
        }

        private void optionsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Options options = new Options();
            options.ShowDialog();
        }
    }
}
