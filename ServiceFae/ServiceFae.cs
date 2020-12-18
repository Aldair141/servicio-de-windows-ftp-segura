using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Office.Interop.Excel;
using ServiceFae.Models;
using ServiceFae.Repository;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using WinSCP;

namespace ServiceFae
{
    public partial class ServiceFae : ServiceBase
    {
        System.Timers.Timer tm;

        public ServiceFae()
        {
            InitializeComponent();

            eventosSistema = new EventLog();
            if (!EventLog.SourceExists("ServiceFae"))
            {
                EventLog.CreateEventSource("ServiceFae", "Application");
            }

            eventosSistema.Source = "ServiceFae";
            eventosSistema.Log = "Application";
        }

        protected override void OnStart(string[] args)
        {
            tm = new System.Timers.Timer();
            tm.Interval = 30000;
            tm.Elapsed += new System.Timers.ElapsedEventHandler(GenerarDocumento);
            tm.Start();

            eventosSistema.WriteEntry($"Iniciando servicio de respuesta ({ ConfigurationManager.AppSettings["hora"].ToString() }).");
            WriteToFile($"Iniciando servicio de respuesta ({ ConfigurationManager.AppSettings["hora"].ToString() }).");
        }

        protected override void OnStop()
        {
            eventosSistema.WriteEntry("Servicio de respuesta detenido.");
        }

        private void GenerarDocumento(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (DateTime.Now.ToString("HH:mm") == ConfigurationManager.AppSettings["hora"].ToString())
            {
                DateTime dt = DateTime.Now;
                string momenclatura = dt.ToString("yyyyMMdd");

                string nomarchivo = $"stock_"+ momenclatura + ".txt";
                DateTime fecha = DateTime.Now;

                try
                {//Subiendo el archivo a un servidor FTP
                    SessionOptions sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Ftp,
                        FtpSecure = FtpSecure.Explicit,
                        HostName = "2c94d577fa.nxcli.net",
                        PortNumber = 21,
                        UserName = "ftp@pieers.com",
                        Password = "HeartsInmostJohnsBrad",
                        GiveUpSecurityAndAcceptAnyTlsHostCertificate = true
                    };

                    using (Session session = new Session())
                    {
                        session.Open(sessionOptions);

                        StockRepository stockRepository = new StockRepository();
                        IEnumerable<Tb_Wb_StockBE> listaStock = stockRepository.GetAll();

                        string[] array;

                        if (listaStock != null)
                        {
                            List<Tb_Wb_StockBE> lista = listaStock as List<Tb_Wb_StockBE>;

                            array = new string[lista.Count + 1];

                            array[0] = "sku,qty,is_in_stock";

                            for (int i = 0; i < lista.Count; i++)
                            {
                                //array[i] = $"{ lista[i].Productid },{ lista[i].StockDisponibleWeb },{ (Convert.ToInt32(lista[i].Stock) == 0 ? "0" : "1") }";
                                array[i + 1] = $"{ lista[i].Productid },{ lista[i].StockDisponibleWeb },{ (Convert.ToInt32(lista[i].StockDisponibleWeb) == 0 ? "0" : "1") }";
                            }
                        }
                        else
                        {
                            array = new string [] { "sku,qty,is_in_stock" };
                        }

                        File.WriteAllLines("C:\\" + nomarchivo, array);

                        //Lo subimos al FTP
                        TransferOptions transferOptions = new TransferOptions();
                        transferOptions.TransferMode = TransferMode.Binary;

                        TransferOperationResult transferResult;
                        transferResult = session.PutFiles("C:\\" + nomarchivo, "/pieers.com/html/var/import/", false, transferOptions);transferResult.Check();

                        if (transferResult.IsSuccess)
                        {
                            string mensaje = String.Format("El documento se subió al Servidor FTP EL {0} a las {1}.", fecha.ToString("dd/MM/yyyy"), fecha.ToString("H:mm"));

                            eventosSistema.WriteEntry(mensaje);
                            WriteToFile(mensaje);

                            try
                            {
                                //Enviar correo
                                MailMessage oMail = new MailMessage();
                                SmtpClient oSmtp = new SmtpClient("smtp.gmail.com");
                                oMail.From = new MailAddress("juanaldairdeveloper@gmail.com", "ARCHIVO DE STOCK", Encoding.UTF8);
                                oMail.Subject = "Envío del archivo de stock al servidor FTP";
                                oMail.Body = mensaje;
                                oMail.To.Add("sistemas4@pieers.com");
                                oMail.To.Add("analistaventadirecta@pieers.com");
                                oMail.To.Add("Coordinadora.marketing@pieers.com");
                                oMail.To.Add("henrycarloman@hotmail.com");

                                oSmtp.Port = 587;
                                oSmtp.Credentials = new NetworkCredential("juanaldairdeveloper@gmail.com", "S0p0rt3123");
                                oSmtp.EnableSsl = true;

                                oSmtp.Send(oMail);
                            }
                            catch (Exception ex)
                            {
                                eventosSistema.WriteEntry("Error al enviar correo: " + ex.Message);
                                WriteToFile("Error al enviar correo: " + ex.Message);
                            }
                        }
                        else
                        {
                            eventosSistema.WriteEntry("No se pudo subir el documento al Servidor FTP.");
                            WriteToFile("No se pudo subir el documento al Servidor FTP.");
                        }

                        session.Close();
                        session.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    eventosSistema.WriteEntry("Hubo un error interno: " + ex.Message);
                    WriteToFile("Hubo un error interno: " + ex.Message);
                }
                finally
                {
                    bool borra = Convert.ToBoolean(ConfigurationManager.AppSettings["optBorra"]);
                    //Borra el archivo
                    if (borra)
                    {
                        if (File.Exists("C:\\" + nomarchivo))
                            File.Delete("C:\\" + nomarchivo);
                    }
                }
            }
        }

        private void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
    }
}