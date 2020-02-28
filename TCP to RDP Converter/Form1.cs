using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RDPCOMAPILib;
using AxMSTSCLib;
using System.Runtime.InteropServices;
using System.Data;
using System.Data.SqlClient;
using System.Data.Odbc;
using System.Management;
using System.Diagnostics;
using Microsoft.VisualBasic.Devices;
using System.IO;
using Microsoft.Win32;
//using System.Windows.Forms;

namespace TCP_to_RDP_Converter
{
    public partial class Form1 : Form
    {
        OdbcConnection conexao = new OdbcConnection(System.Configuration.ConfigurationManager.ConnectionStrings["programas1"].ConnectionString.ToString());
        // informacaoPc = new System.Windows.Forms.SystemInformation();
      ManagementObjectSearcher infoProce = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Processor");
        string nomeProcessador,memoriaPc,NomeWindows,macAddress,IP,criaRegistro,idM,tamanhoDIsco,UsadoDisco,tamanhoDiscoLivre,quantSoft;
        public static RDPSession currentSession = null;

        class config
        {
            public string idmaquina { get; set; }
            public string nome { get; set; }
            public string macAdress { get; set; }
           

        }
        class softw
        {
          
            public string nome { get; set; }
           


        }

      

        public Form1()
        {
            InitializeComponent();
          //  Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
          //  key.SetValue("CONEXÃO REMOTA INDUSCABOS", Application.ExecutablePath.ToString());
            carregaNomeProcessador();
            carregaQtdMemoria();
            carregaVersaoWindows();
            carregaMacAddressIp();
            carregaEspacoDisco();
            verificaSeExisteMaquina(conexao, System.Windows.Forms.SystemInformation.ComputerName.ToString(), macAddress);
            quantidadeDeSoftware(macAddress, System.Windows.Forms.SystemInformation.ComputerName.ToString(), idM);
            listaDeSoftware();
            if (criaRegistro == "NAO")
            {
                conectarSession();
                atualizarChaveAcesso(conexao,textConnectionString.Text.TrimEnd(), System.Windows.Forms.SystemInformation.ComputerName.ToString(), nomeProcessador, memoriaPc, NomeWindows, IP, macAddress,tamanhoDIsco,UsadoDisco, idM);
                escreverNoArquivoConfig();
                InserirListaSoft(dataGridView1, idM, macAddress, IP, System.Windows.Forms.SystemInformation.ComputerName.ToString());

            }
            else
            {
                conectarSession();
                cadastraNaTabelaConexao(conexao, textConnectionString.Text.TrimEnd(), System.Windows.Forms.SystemInformation.ComputerName.ToString(), nomeProcessador, memoriaPc, NomeWindows, IP, macAddress,tamanhoDIsco,UsadoDisco);
                verificaSeExisteMaquina(conexao, System.Windows.Forms.SystemInformation.ComputerName.ToString(), macAddress);
           
                escreverNoArquivoConfig();
                 InserirListaSoft(dataGridView1,idM,macAddress,IP,System.Windows.Forms.SystemInformation.ComputerName.ToString());

            }

          //  loadArquivoConfig(dataGridView1);

            this.WindowState = FormWindowState.Minimized;
            timer1.Start();
        }

        public static void createSession()
        {
            currentSession = new RDPSession();
        }
        public void carregaNomeProcessador()
        {

            foreach (ManagementObject mo in infoProce.Get())
                nomeProcessador =  mo["Name"].ToString();
        }
        public void carregaQtdMemoria() {

            ManagementObjectSearcher s4 = new ManagementObjectSearcher("SELECT Capacity FROM Win32_PhysicalMemory");

            foreach (ManagementObject mo in s4.Get())
                memoriaPc = Convert.ToString(mo["Capacity"]);

            double memoriaRam = Convert.ToDouble(memoriaPc);
            double memoriaRamTotal = 0;

            memoriaRamTotal = memoriaRam / 1024 / 1024;
            memoriaPc = Convert.ToString(memoriaRamTotal);
        }

        public void carregaVersaoWindows() {
        //    NomeWindows = System.Environment.OSVersion.ToString();
            NomeWindows = new ComputerInfo().OSFullName.ToString();
           // NomeWindows = "[ " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() + " ] " + " [ " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Name.ToString() +" ]";
        
        
        }

        public void carregaMacAddressIp()
         { 
            ManagementObjectSearcher ObjMOS = new ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'"); 
            ManagementObjectCollection ObjMOC = ObjMOS.Get(); 
 
            foreach (ManagementObject mo in ObjMOC) 
            { 
                string[] addresses = (string[])mo["IPAddress"]; 
                IP = addresses[0]; 
                macAddress = addresses[1]; 
 
            } 
        }

        public void carregaEspacoDisco() { 
        
            ConnectionOptions opt = new ConnectionOptions();
			ObjectQuery oQuery = new ObjectQuery("SELECT Size, FreeSpace, Name, FileSystem FROM Win32_LogicalDisk WHERE DriveType = 3");
			ManagementScope scope = new ManagementScope("\\\\localhost\\root\\cimv2", opt);
			
			ManagementObjectSearcher moSearcher = new ManagementObjectSearcher(scope, oQuery);
			ManagementObjectCollection collection = moSearcher.Get();
			foreach (ManagementObject res in collection)
			{
                if (res["Name"].ToString() == "C:")
                {
                    decimal size = Convert.ToDecimal(res["Size"]) / 1024 / 1024 / 1024;
                    decimal freeSpace = Convert.ToDecimal(res["FreeSpace"]) / 1024 / 1024 / 1024;
                    string unidade = res["Name"].ToString();
                    decimal tamanho = Decimal.Round(size, 2);
                    decimal livre = Decimal.Round(freeSpace, 2);
                    decimal usado = Decimal.Round(size - freeSpace, 2);
                    decimal livrepercent = Decimal.Round(usado / size, 2) * 100;

                    tamanhoDIsco = tamanho.ToString();
                    UsadoDisco = usado.ToString();
                    tamanhoDiscoLivre = livre.ToString();
                    tamanhoDIsco ="UNIDADE " +res["Name"].ToString() + " TAMANHO: [ " + tamanhoDIsco + " ] LIVRE PARA USO [ " + tamanhoDiscoLivre + " ]";

                }


           }
        
        }


        public void listaDeSoftware() {

            List<softw> listaNome = new List<softw>();
              try
              {

                  //string registry_key = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
                  //using (Microsoft.Win32.RegistryKey key = Registry.LocalMachine.OpenSubKey(registry_key))
                  //{
                  //    foreach (string subkey_name in key.GetSubKeyNames())
                  //    {
                  //        using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                  //        {
                  //            listaNome.Add( subkey.GetValue("DisplayName").ToString());
                  //            //  Console.WriteLine(subkey.GetValue("DisplayName"));
                  //        }
                  //    }
                  //}
                  string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
                  using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(uninstallKey))
                  {
                      foreach (string skName in rk.GetSubKeyNames())
                      {
                          using (RegistryKey sk = rk.OpenSubKey(skName))
                          {
                              try
                              {

                                  listaNome.Add(new softw() { nome = sk.GetValue("DisplayName").ToString() });
                        
                              }
                              catch (Exception ex)
                              { }
                          }
                      }
                  }

              }
              catch
              {

              }
          
            dataGridView1.DataSource = listaNome;
        }



        public static void Connect(RDPSession session)
        {
            session.OnAttendeeConnected += Incoming;
            session.Open();
            
         
        }

        public static void Disconnect(RDPSession session)
        {
            session.Close();
        }

        public static string getConnectionString(RDPSession session, String authString, 
            string group, string password, int clientLimit)
        {
            IRDPSRAPIInvitation invitation =
                session.Invitations.CreateInvitation
                (authString, group, password, clientLimit);
                        return invitation.ConnectionString;
        }

        private static void Incoming(object Guest)
        {
            IRDPSRAPIAttendee MyGuest = (IRDPSRAPIAttendee)Guest;
            MyGuest.ControlLevel = CTRL_LEVEL.CTRL_LEVEL_INTERACTIVE;
        }

      
       

        private void button1_Click(object sender, EventArgs e)
        {
            
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
          
        }

        public void cadastraNaTabelaConexao(OdbcConnection conexao, string caminho,string nomeMaquina,string processador,string memoria,string windows,string ip,string mac,string tamanhoHd,string usadoHd)
        {
            if (conexao.State == ConnectionState.Closed) //Validar a conexão
            {
                conexao.Open();
            }

            try
            {
                string sql = "insert into conexaoRemota (string,maquina,dtcad,processador,memoria,windows,ip,macAddress,tamanhoHD,usadoHD) values ('" + caminho + "','" + nomeMaquina + "',getdate(),'"+processador+"','"+memoria+"','"+windows+"','"+ip+"','"+mac+"','"+tamanhoHd+"','"+usadoHd+"')";

                OdbcCommand cmd1 = new OdbcCommand(sql, conexao);

                cmd1.ExecuteNonQuery();
            }
            catch { conexao.Close(); }

            conexao.Close();
        }


        public void loadArquivoConfig(DataGridView grid)
        {
            //grid.Columns[0].Name = "IDMAQUINA";
            //grid.Columns[1].Name = "NOME";
            //grid.Columns[2].Name = "MACADRESS";
         //   grid.Columns[3].Name = "nomesetor";



            List<config> lista = new List<config>();
            try
            {
                System.IO.StreamReader arquivo = new System.IO.StreamReader(System.AppDomain.CurrentDomain.BaseDirectory.ToString()+@"\config.txt");
                string linha = "";
                while (true)
                {
                    linha = arquivo.ReadLine();
                    if (linha != null)
                    {
                        string[] DadosColetados = linha.Split(',');                                                      //, Setor = DadosColetados[3]
                        lista.Add(new config { idmaquina = DadosColetados[0], nome = DadosColetados[1], macAdress = DadosColetados[2] });
                    }
                    else
                        break;
                }
                grid.DataSource = lista;
                arquivo.Close();
            }

            catch (System.Exception)
            {
                MessageBox.Show("PROBLEMA NO ARQUIVO CONFIG.");
                Close();
            }

        }

        public void escreverNoArquivoConfig() {

            string[] lines = { idM+","+ System.Windows.Forms.SystemInformation.ComputerName.ToString()+","+ macAddress+"" };
            string caminho = System.AppDomain.CurrentDomain.BaseDirectory.ToString()+@"\config.txt";
           using (System.IO.StreamWriter file = 
            new System.IO.StreamWriter(caminho))
                {
                    foreach (string line in lines)
                    {
               
                            file.WriteLine(line);
                
                    }
                }


        }


        public void verificaSeExisteMaquina(OdbcConnection conexao, string maquina,string macAdrr)
        {


            string sql = "select top(1) * from conexaoRemota where macAddress ='"+macAdrr+"' and maquina ='"+maquina+"' and delet ='' order by idM desc";
            if (conexao.State == ConnectionState.Closed) //Validar a conexão
            {
                conexao.Open();
            }

            OdbcDataAdapter ADAP = new OdbcDataAdapter(sql, conexao);
            DataSet DS = new DataSet();
            ADAP.Fill(DS, "past11");

            if (DS.Tables["past11"].Rows.Count > 0)
            {

                criaRegistro = "NAO";
                idM = DS.Tables["past11"].Rows[0]["idM"].ToString();

            }
            else {
                criaRegistro = "SIM";
            }


            conexao.Close();


        }

        public void conectarSession() {

            createSession();
            Connect(currentSession);
            textConnectionString.Text = getConnectionString(currentSession,
                "Conect", "INDUSCABOS", "", 5);
             this.WindowState = FormWindowState.Minimized;
            lbDes.Text = "CONECTADO.";
        }

        public void desconectarSession() {
            Disconnect(currentSession);
            lbDes.Text = "SOLICITAR ACESSO.";
        }


        public void atualizarChaveAcesso(OdbcConnection conexao,string caminho,string nomeMaquina,string processador,string memoria,string windows,string ip,string mac,string tamanhoHd,string usadoHd, string idM)
        {
            if (conexao.State == ConnectionState.Closed) //Validar a conexão
            {
                conexao.Open();
            }

            try
            {
                string sql = "update conexaoRemota set string='"+caminho+"',maquina='"+nomeMaquina+"',processador='"+processador+"',memoria='"+memoria+"',windows='"+windows+"',ip='"+ip+"',macAddress='"+mac+"',tamanhoHD='"+tamanhoHd+"',usadoHD='"+usadoHd+"',dtcad =GETDATE() where idM = '" + idM.TrimEnd() + "'";

                OdbcCommand cmd1 = new OdbcCommand(sql, conexao);

                cmd1.ExecuteNonQuery();
            }
            catch { conexao.Close(); }

            conexao.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.WindowsShutDown)
            {
                e.Cancel = false;
                this.Hide();

            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            // isto é para quando carregas no "X" ele não fechar mas sim esconder
            e.Cancel = true;
            // aqui minimizo e escondo
            this.WindowState = FormWindowState.Minimized;
            Hide();
        }

        private void notifyIcon1_BalloonTipClosed(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        private void InserirListaSoft(DataGridView grid, string idMaquina,string macAddres,string ip,string nomeMaquina)
        {
            string sql = "INSERT INTO listaSoftMaquina (idM, macAddress, ip,programa,dtcad,nomeMaquina) " +
                      "VALUES (?, ?, ?, ?,GETDATE(),?)";

            OdbcCommand cmd = new OdbcCommand(sql, conexao);
            if (conexao.State == ConnectionState.Closed)
            {
                conexao.Open();
            }

            if (grid.Rows.Count > Convert.ToInt64(quantSoft))
            {

                string sql1 = "delete from listaSoftMaquina where macAddress ='" + macAddres + "' and nomeMaquina ='" + nomeMaquina + "' and idM='" + idMaquina + "' and delet =''";

                OdbcCommand cmd1 = new OdbcCommand(sql1, conexao);
                cmd1.ExecuteNonQuery();


                for (int i = 0; i < grid.Rows.Count; i++)
                {

                    cmd.Parameters.Clear();

                    cmd.Parameters.AddWithValue("@idM", idMaquina);
                    cmd.Parameters.AddWithValue("@macAddress", macAddres);
                    cmd.Parameters.AddWithValue("@ip", ip);
                    cmd.Parameters.AddWithValue("@programa", grid.Rows[i].Cells[0].Value.ToString());
                    cmd.Parameters.AddWithValue("@nomeMaquina", nomeMaquina);

                    cmd.ExecuteNonQuery();

                }


            }
           
            conexao.Close();
        }


        public void quantidadeDeSoftware(string macAdrr, string maquina,string idMaquin)
        {

            string sql = "select count(*) as c from listaSoftMaquina where macAddress ='" + macAdrr + "' and nomeMaquina ='" + maquina + "' and idM='"+idMaquin+"' and delet =''";
            if (conexao.State == ConnectionState.Closed) //Validar a conexão
            {
                conexao.Open();
            }

            OdbcDataAdapter ADAP = new OdbcDataAdapter(sql, conexao);
            DataSet DS = new DataSet();
            ADAP.Fill(DS, "past11");

           
            quantSoft = DS.Tables["past11"].Rows[0]["c"].ToString();

           

            conexao.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Hide();
            timer1.Stop();
        }




    }
}
