using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL;
using WebScrapingImpressoras.Entity;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Net.Mail;
using System.Net;

namespace WebScrapingImpressoras
{
    class Program
    {
        static void Main(string[] args)
        {

            IWebDriver driver = new ChromeDriver();


            //Faz acesso nas paginas de cada impressora e coleta os dados para gravar no banco de dados.
            RelatorioMFP432(driver);
            RelatorioMFP479(driver);
            RelatorioMFPE52645(driver);

            driver.Quit();
            EnviarEmail();

        }

        private static void RelatorioMFP432(IWebDriver driver)
        {
            int refresh = 0;
            bool sucess = false;
            List<ImpressorasInfo> ImpressorasMFP432 = ObterImpressoras("MFP432");
            List<ObjImpressoraInfo> ListaMFP432 = new List<ObjImpressoraInfo>();

            Console.WriteLine("Abrindo Navegador...");


            foreach (ImpressorasInfo print in ImpressorasMFP432)
            {
                refresh = 0;
                sucess = false;
                while (refresh < 5 && !sucess)
                {

                    Console.WriteLine("Carregando pagina...");

                    try
                    {
                        driver.Navigate().GoToUrl(print.Endereco);
                        driver.FindElement(By.XPath("/html/body/div/div[2]/button[3]")).Click();
                        driver.FindElement(By.XPath("/html/body/div/div[3]/p[2]/a")).Click();
                        Console.WriteLine("Procurando dados...");

                        var toner = 999;
                        while (toner == 999)
                        {
                            //Tenta encontrar o elemento na página, caso não achar, atualiza
                            try
                            {
                                Thread.Sleep(5000);
                                ObjImpressoraInfo MFP432 = new ObjImpressoraInfo();
                                //Tenta pegar o elemento, se não pegar, da um sleep e tenta novamente, se não atualiza a pagina.
                                try
                                {
                                    var pToner = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div/div/div/div/form/fieldset[1]/div/div/div[1]/div/div/div/div[1]/div/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[2]/div/div/table/tbody/tr/td[2]/div/div/div[2]")).Text.Replace("%", "").Replace("--", "0");
                                    toner = int.Parse(pToner);
                                    var pIP = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div/div/div/div/form/div/div/div/fieldset/div/div/div/div/div/div/div[3]/div/div/div/div[4]/div[2]")).Text;
                                    var pUnidImg = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div/div/div/div/form/fieldset[1]/div/div/div[1]/div/div/div/div[3]/div/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[2]/div/div/table/tbody/tr/td[2]/div/div/div[2]")).Text.Replace("%", "").Replace("--", "0");

                                    Console.WriteLine("Armazenando dados...");
                                    MFP432.Toner = int.Parse(pToner);
                                    MFP432.IP = pIP;
                                    MFP432.UnidImg = int.Parse(pUnidImg);
                                    MFP432.DataRegistro = DateTime.Now;
                                    MFP432.Turno = VerificaTurno(DateTime.Now);
                                    MFP432.Toner_Preto = 0;
                                    MFP432.Toner_Cyan = 0;
                                    MFP432.Toner_Magenta = 0;
                                    MFP432.Toner_Amarelo = 0;

                                    ListaMFP432.Add(MFP432);
                                    sucess = true;

                                }//Tenta dnv após 10miliseg.
                                catch (Exception)
                                {
                                    Thread.Sleep(10000);

                                    var pToner = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div/div/div/div/form/fieldset[1]/div/div/div[1]/div/div/div/div[1]/div/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[2]/div/div/table/tbody/tr/td[2]/div/div/div[2]")).Text.Replace("%", "");
                                    toner = int.Parse(pToner);
                                    var pIP = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div/div/div/div/form/div/div/div/fieldset/div/div/div/div/div/div/div[3]/div/div/div/div[4]/div[2]")).Text;
                                    var pUnidImg = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div/div[2]/div[1]/div/div/div/div/div/form/fieldset[1]/div/div/div[1]/div/div/div/div[3]/div/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[2]/div/div/table/tbody/tr/td[2]/div/div/div[2]")).Text.Replace("%", "");

                                    Console.WriteLine("Armazenando dados...");
                                    MFP432.Toner = int.Parse(pToner);
                                    MFP432.IP = pIP;
                                    MFP432.UnidImg = int.Parse(pUnidImg);
                                    MFP432.DataRegistro = DateTime.Now;
                                    MFP432.Turno = VerificaTurno(DateTime.Now);
                                    MFP432.Toner_Preto = 0;
                                    MFP432.Toner_Cyan = 0;
                                    MFP432.Toner_Magenta = 0;
                                    MFP432.Toner_Amarelo = 0;

                                    ListaMFP432.Add(MFP432);
                                    sucess = true;
                                }

                            }//Atualiza pagina.
                            catch (Exception)
                            {
                                Console.WriteLine("Atualizando Pagina...");
                                driver.Navigate().Refresh();
                                toner = 999;
                                refresh++;
                                sucess = false;
                            }
                        }
                    }//NÃO CONSEGUE ABRIR A PAGINA, TENTA POR 5 VEZES E CAI NO IF DA IMPRESSORA OFFLINE.
                    catch (Exception)
                    {
                        Console.WriteLine("Atualizando Pagina...");
                        driver.Navigate().Refresh();
                        refresh++;
                        sucess = false;
                    }
                }
                if (!sucess)
                {
                    Console.WriteLine("Erro: Impressora Offline!");
                    List<ObjImpressoraInfo> RegistrosAnterior = ObterImpressorasUltimoRegistro("MFP432");

                    if (RegistrosAnterior.Count != 0)
                    {
                        ObjImpressoraInfo ImpressoraAnterior = new ObjImpressoraInfo();
                        ImpressoraAnterior = RegistrosAnterior.Find(x => x.IP == print.IP);
                        ImpressoraAnterior.Turno = VerificaTurno(DateTime.Now);
                        ImpressoraAnterior.DataRegistro = DateTime.Now;
                        ListaMFP432.Add(ImpressoraAnterior);
                    }
                }


            }
            //Salvar no Banco
            Console.WriteLine("Salvando dados na base de dados...");
            PopularTabelaImpressoras(ListaMFP432, "MFP432");
        }

        private static void RelatorioMFP479(IWebDriver driver)
        {
            bool sucess = false;
            int refresh = 0;
            List<ImpressorasInfo> ImpressorasMFP479 = ObterImpressoras("MFP479");
            List<ObjImpressoraInfo> ListaMFP479 = new List<ObjImpressoraInfo>();

            Console.WriteLine("Abrindo Navegador...");


            foreach (ImpressorasInfo print in ImpressorasMFP479)
            {
                refresh = 0;
                sucess = false;
                while (refresh < 5 && !sucess)
                {
                    try
                    {
                        Console.WriteLine("Carregando pagina...");
                        driver.Navigate().GoToUrl(print.Endereco);

                        driver.FindElement(By.XPath("/html/body/div/div[2]/button[3]")).Click();
                        driver.FindElement(By.XPath("/html/body/div/div[3]/p[2]/a")).Click();
                        Console.WriteLine("Procurando dados...");

                        var toner = 999;
                        while (refresh < 5 && toner == 999)
                        {
                            //Tenta encontrar o elemento na página, caso não achar, atualiza
                            try
                            {
                                ObjImpressoraInfo MFP479 = new ObjImpressoraInfo();
                                //Tenta pegar o elemento, se não pegar, da um sleep e tenta novamente, se não atualiza a pagina.
                                try
                                {
                                    Thread.Sleep(5000);
                                    //var pTonerPreto = driver.FindElement(By.XPath("/html/body/div[1]/div[5]/div[1]/div[2]/form/div[1]/div[2]/div[1]/div/table/tbody/tr[8]/td[2]")).Text;//.Replace(">", "").Replace("*", "").Replace("†", "");
                                    var pTonerPreto = driver.FindElement(By.CssSelector("#appConsumable-inkCart-tbl-Tbl > tbody > tr:nth-child(8) > td:nth-child(2)")).Text.Replace(">", "").Replace("*", "").Replace("†", "").Replace("¶", "").Replace("--", "0");
                                    toner = int.Parse(pTonerPreto);
                                    var pTonerCyan = driver.FindElement(By.CssSelector("#appConsumable-inkCart-tbl-Tbl > tbody > tr:nth-child(8) > td:nth-child(3)")).Text.Replace(">", "").Replace("*", "").Replace("†", "").Replace("¶", "").Replace("--", "0");
                                    toner = int.Parse(pTonerCyan);
                                    var pTonerMagenta = driver.FindElement(By.CssSelector("#appConsumable-inkCart-tbl-Tbl > tbody > tr:nth-child(8) > td:nth-child(4)")).Text.Replace(">", "").Replace("*", "").Replace("†", "").Replace("¶", "").Replace("--", "0");
                                    toner = int.Parse(pTonerMagenta);
                                    var pTonerAmarelo = driver.FindElement(By.CssSelector("#appConsumable-inkCart-tbl-Tbl > tbody > tr:nth-child(8) > td:nth-child(5)")).Text.Replace(">", "").Replace("*", "").Replace("†", "").Replace("¶", "").Replace("--", "0");
                                    toner = int.Parse(pTonerAmarelo);


                                    Console.WriteLine("Armazenando dados...");
                                    MFP479.Toner = 0;
                                    MFP479.IP = print.IP;
                                    MFP479.UnidImg = 0;
                                    MFP479.DataRegistro = DateTime.Now;
                                    MFP479.Turno = VerificaTurno(DateTime.Now);
                                    MFP479.Toner_Preto = int.Parse(pTonerPreto);
                                    MFP479.Toner_Cyan = int.Parse(pTonerCyan);
                                    MFP479.Toner_Magenta = int.Parse(pTonerMagenta);
                                    MFP479.Toner_Amarelo = int.Parse(pTonerAmarelo);

                                    ListaMFP479.Add(MFP479);
                                    sucess = true;
                                }//Tenta dnv após 15000 miliseconds.
                                catch (Exception)
                                {
                                    Thread.Sleep(15000);

                                    //var pTonerPreto = driver.FindElement(By.XPath("/html/body/div[1]/div[5]/div[1]/div[2]/form/div[1]/div[2]/div[1]/div/table/tbody/tr[8]/td[2]")).Text;//.Replace(">", "").Replace("*", "").Replace("†", "");
                                    var pTonerPreto = driver.FindElement(By.CssSelector("#appConsumable-inkCart-tbl-Tbl > tbody > tr:nth-child(8) > td:nth-child(2)")).Text.Replace(">", "").Replace("*", "").Replace("†", "").Replace("¶", "").Replace("--", "0");
                                    pTonerPreto.Substring(0, 6);
                                    toner = int.Parse(pTonerPreto);
                                    var pTonerCyan = driver.FindElement(By.CssSelector("#appConsumable-inkCart-tbl-Tbl > tbody > tr:nth-child(8) > td:nth-child(3)")).Text.Replace(">", "").Replace("*", "").Replace("†", "").Replace("¶", "").Replace("--", "0");
                                    toner = int.Parse(pTonerCyan);
                                    var pTonerMagenta = driver.FindElement(By.CssSelector("#appConsumable-inkCart-tbl-Tbl > tbody > tr:nth-child(8) > td:nth-child(4)")).Text.Replace(">", "").Replace("*", "").Replace("†", "").Replace("¶", "").Replace("--", "0");
                                    toner = int.Parse(pTonerMagenta);
                                    var pTonerAmarelo = driver.FindElement(By.CssSelector("#appConsumable-inkCart-tbl-Tbl > tbody > tr:nth-child(8) > td:nth-child(5)")).Text.Replace(">", "").Replace("*", "").Replace("†", "").Replace("¶", "").Replace("--", "0");
                                    toner = int.Parse(pTonerAmarelo);


                                    Console.WriteLine("Armazenando dados...");
                                    MFP479.Toner = 0;
                                    MFP479.IP = print.IP;
                                    MFP479.UnidImg = 0;
                                    MFP479.DataRegistro = DateTime.Now;
                                    MFP479.Turno = VerificaTurno(DateTime.Now);
                                    MFP479.Toner_Preto = int.Parse(pTonerPreto);
                                    MFP479.Toner_Cyan = int.Parse(pTonerCyan);
                                    MFP479.Toner_Magenta = int.Parse(pTonerMagenta);
                                    MFP479.Toner_Amarelo = int.Parse(pTonerAmarelo);

                                    ListaMFP479.Add(MFP479);
                                    sucess = true;
                                }

                            }//Atualiza pagina.
                            catch (Exception)
                            {
                                Console.WriteLine("Atualizando Pagina...");

                                driver.Navigate().Refresh();
                                toner = 999;
                                refresh++;
                                sucess = false;
                            }
                        }
                    }//NÃO CONSEGUE ABRIR A PAGINA, TENTA POR 5 VEZES E CAI NO IF DA IMPRESSORA OFFLINE.
                    catch (Exception)
                    {
                        Console.WriteLine("Atualizando Pagina...");
                        driver.Navigate().Refresh();
                        refresh++;
                        sucess = false;
                    }
                }
                if (!sucess)
                {
                    Console.WriteLine("Erro: Impressora Offline!");
                    List<ObjImpressoraInfo> RegistrosAnterior = ObterImpressorasUltimoRegistro("MFP479");
                    if (RegistrosAnterior.Count != 0)
                    {
                        ObjImpressoraInfo ImpressoraAnterior = new ObjImpressoraInfo();
                        ImpressoraAnterior = RegistrosAnterior.Find(x => x.IP == print.IP);
                        ImpressoraAnterior.Turno = VerificaTurno(DateTime.Now);
                        ImpressoraAnterior.DataRegistro = DateTime.Now;
                        ListaMFP479.Add(ImpressoraAnterior);
                    }
                }


            }
            //Salvar no Banco
            Console.WriteLine("Salvando dados na base de dados...");
            PopularTabelaImpressoras(ListaMFP479, "MFP479");
        }

        private static void RelatorioMFPE52645(IWebDriver driver)
        {
            bool sucess = false;
            int refresh = 0;
            List<ImpressorasInfo> ImpressorasMFPE52645 = ObterImpressoras("MFPE52645");
            List<ObjImpressoraInfo> ListaMFPE52645 = new List<ObjImpressoraInfo>();

            Console.WriteLine("Abrindo Navegador...");


            foreach (ImpressorasInfo print in ImpressorasMFPE52645)
            {
                refresh = 0;
                sucess = false;
                while (refresh < 5 && !sucess)
                {
                    try
                    {

                        Console.WriteLine("Carregando pagina...");
                        driver.Navigate().GoToUrl(print.Endereco);

                        driver.FindElement(By.XPath("/html/body/div/div[2]/button[3]")).Click();
                        driver.FindElement(By.XPath("/html/body/div/div[3]/p[2]/a")).Click();
                        Console.WriteLine("Procurando dados...");

                        var toner = 999;
                        while (refresh < 5 && toner == 999)
                        {
                            //Tenta encontrar o elemento na página, caso não achar, atualiza
                            try
                            {
                                Thread.Sleep(5000);
                                ObjImpressoraInfo MFPE52645 = new ObjImpressoraInfo();
                                //Tenta pegar o elemento, se não pegar, da um sleep e tenta novamente, se não atualiza a pagina.
                                try
                                {
                                    var pToner = driver.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[1]/span[2]")).Text.Replace("%", "").Replace("*", "").Replace("<", "").Replace("--", "0");
                                    toner = int.Parse(pToner);
                                    var pIP = driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div/p[2]")).Text;
                                    var pUnidImg = driver.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[2]/span[2]")).Text.Replace("%", "").Replace("*", "").Replace("<", "").Replace("--", "0");

                                    Console.WriteLine("Armazenando dados...");
                                    MFPE52645.Toner = int.Parse(pToner);
                                    MFPE52645.IP = pIP;
                                    MFPE52645.UnidImg = int.Parse(pUnidImg);
                                    MFPE52645.DataRegistro = DateTime.Now;
                                    MFPE52645.Turno = VerificaTurno(DateTime.Now);
                                    MFPE52645.Toner_Preto = 0;
                                    MFPE52645.Toner_Cyan = 0;
                                    MFPE52645.Toner_Magenta = 0;
                                    MFPE52645.Toner_Amarelo = 0;

                                    ListaMFPE52645.Add(MFPE52645);
                                    sucess = true;

                                }//Tenta dnv após 10seg.
                                catch (Exception)
                                {
                                    Thread.Sleep(10000);

                                    var pToner = driver.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[1]/span[2]")).Text.Replace("%", "").Replace("*", "").Replace("<", "");
                                    toner = int.Parse(pToner);
                                    var pIP = driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div/p[2]")).Text;
                                    var pUnidImg = driver.FindElement(By.XPath("/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[2]/span[2]")).Text.Replace("%", "").Replace("*", "").Replace("<", "");

                                    Console.WriteLine("Armazenando dados...");
                                    MFPE52645.Toner = int.Parse(pToner);
                                    MFPE52645.IP = pIP;
                                    MFPE52645.UnidImg = int.Parse(pUnidImg);
                                    MFPE52645.DataRegistro = DateTime.Now;
                                    MFPE52645.Turno = VerificaTurno(DateTime.Now);
                                    MFPE52645.Toner_Preto = 0;
                                    MFPE52645.Toner_Cyan = 0;
                                    MFPE52645.Toner_Magenta = 0;
                                    MFPE52645.Toner_Amarelo = 0;

                                    ListaMFPE52645.Add(MFPE52645);
                                    sucess = true;
                                }

                            }//Atualiza pagina.
                            catch (Exception)
                            {
                                Console.WriteLine("Atualizando Pagina...");
                                driver.Navigate().Refresh();
                                toner = 999;
                                refresh++;
                                sucess = false;
                            }

                        }
                    }//NÃO CONSEGUE ABRIR A PAGINA, TENTA POR 5 VEZES E CAI NO IF DA IMPRESSORA OFFLINE.
                    catch (Exception)
                    {
                        Console.WriteLine("Atualizando Pagina...");
                        driver.Navigate().Refresh();
                        refresh++;
                        sucess = false;
                    }

                }
                if (!sucess)
                {
                    Console.WriteLine("Erro: Impressora Offline!");
                    List<ObjImpressoraInfo> RegistrosAnterior = ObterImpressorasUltimoRegistro("MFPE52645");
                    if (RegistrosAnterior.Count != 0)
                    {
                        ObjImpressoraInfo ImpressoraAnterior = new ObjImpressoraInfo();
                        ImpressoraAnterior = RegistrosAnterior.Find(x => x.IP == print.IP);
                        ImpressoraAnterior.Turno = VerificaTurno(DateTime.Now);
                        ImpressoraAnterior.DataRegistro = DateTime.Now;
                        ListaMFPE52645.Add(ImpressoraAnterior);
                    }

                }


            }
            //Salvar no Banco
            Console.WriteLine("Salvando dados na base de dados...");
            PopularTabelaImpressoras(ListaMFPE52645, "MFPE52645");
        }

        private static int PegaTurnoAnterior(int turno)
        {
            if (turno == 1)
                return 3;
            else if (turno == 2)
                return 1;
            else
                return 2;
        }

        private static List<TonerInfo> ObterTonersEstoque()
        {
            AcessoBancoDados acessoBancoDados = new AcessoBancoDados();

            try
            {
                List<TonerInfo> Toners = new List<TonerInfo>();
                acessoBancoDados.LimparParametros();

                //Carregar DataTable
                DataTable dt = acessoBancoDados.ExecutarConsulta(CommandType.StoredProcedure, "usp_ObterTonersEstoque");

                if (dt.Rows.Count > 0)
                {
                    Toners = acessoBancoDados.ConverteParaLista<TonerInfo>(dt);
                }
                return Toners;
            }
            catch (Exception ex)
            {
                throw new Exception("Não foi possivel obter lista de impressoras da data anterior. Detalhes: " + ex.Message);
            }
        }
        private static void EnviarEmail()
        {
            try
            {
                //usp_ObterRelatorioImpressoras --> Gera o HTML dos dados das impressoras coletados do mesmo turno que é chamado o metodo
                string Tabela432 = ObterRelatorioImpressoras("MFP432", VerificaTurno(DateTime.Now));
                string Tabela479 = ObterRelatorioImpressoras("MFP479", VerificaTurno(DateTime.Now));
                string TabelaE52645 = ObterRelatorioImpressoras("MFPE52645", VerificaTurno(DateTime.Now));
                string Turno = VerificaTurnoString(DateTime.Now);
                List<TonerInfo> TonersEstoque = ObterTonersEstoque();
                int Toner432Estoque = TonersEstoque.Find(x => x.Modelo == "MFP432").TonerUnico;
                int TonerE52645Estoque = TonersEstoque.Find(x => x.Modelo == "MFPE52645").TonerUnico;
                int Toner479PretoEstoque = TonersEstoque.Find(x => x.Modelo == "MFP479").ColorBlack;
                int Toner479CyanEstoque = TonersEstoque.Find(x => x.Modelo == "MFP479").ColorCiano;
                int Toner479MagentaEstoque = TonersEstoque.Find(x => x.Modelo == "MFP479").ColorMagenta;
                int Toner479AmareloEstoque = TonersEstoque.Find(x => x.Modelo == "MFP479").ColorYellow;

                string[] emails = { "kelvin.rocha@plural.com.br", "suporte@plural.com.br" };

                foreach (string destino in emails)
                {

                    MailMessage email = new MailMessage("Sistemas.plural@plural.com.br", destino);

                    email.Subject = Turno;

                    email.IsBodyHtml = true;
                    email.Body = "<html><head><style type=\"text/css\">caption{font-weight: bold;font-size: 20px;margin-bottom: 5px;}table {padding: 0;border-spacing: 0; border-collapse: collapse;}thead{background: #063690; border: 1px solid #ddd;}th {padding: 5px;font-weight: bold;border: 1px solid #000;color: #fff;}tr{padding: 0;}td{padding: 5px; border: 1px solid #000; margin:0; text-align:center;}</style></head><body><h3 style = \"margin-left:140px\">" + Turno + "</h3><p> Relatório gerado em: <strong>" + DateTime.Now.ToString("dd/MM/yyyy") + "</strong></p><p> Segue Relatórios de todas as impressoras e estoque de suprimentos deste turno:</p><p></p>" + Tabela432 + "<p></p><p> Toners MFP 432FDN em estoque: <strong>" + Toner432Estoque + "</strong></p><p></p><p>" + TabelaE52645 + "</p><p> Toners E52645 em estoque: <strong>" + TonerE52645Estoque + "</strong></p><p></p>" + Tabela479 + "<p></p><p><br> Toner M479<strong> Preto</strong> em estoque: <strong> " + Toner479PretoEstoque + " </strong></br><br> Toner M479<strong> Ciano</strong> em estoque: <strong> " + Toner479CyanEstoque + " </strong></br><br> Toner M479<strong> Magenta</strong> em estoque: <strong> " + Toner479MagentaEstoque + " </strong></br><br> Toner M479<strong> Amarelo</strong> em estoque: <strong> " + Toner479AmareloEstoque + " </strong></br></p><p></p><p> Atualização do Estoque de Toners via sistema CETI.</p><p> Abra a planilha de Controle de Estoque Simpress <a href = \"\\\\srvsao040\\Departamentos\\TI\\Suporte\\Estoque Simpress\\Estoque (Simpress).xlsx\"target = \"_blank\"> Clicando aqui </a></p><p> Suporte TI - (11) 4152 - 9518 / 9821 </p></body></html>";

                    email.SubjectEncoding = Encoding.GetEncoding("UTF-8");
                    email.BodyEncoding = Encoding.GetEncoding("UTF-8");

                    SmtpClient smtpClient = new SmtpClient();
                    //smtpClient.Host = "email.plural.com.br";
                    //smtpClient.Port = 587;
                    smtpClient.Host = "smtp-relay.gmail.com";
                    smtpClient.Port = 25;

                    //smtpClient.UseDefaultCredentials = false;
                    smtpClient.Credentials = new NetworkCredential("Sistemas.plural", "asdf321!@#");
                    //smtpClient.EnableSsl = true;

                    smtpClient.Send(email);
                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        private static string ObterRelatorioImpressoras(string pModelo, int pTurno)
        {
            AcessoBancoDados acessoBancoDados = new AcessoBancoDados();

            try
            {

                acessoBancoDados.LimparParametros();
                acessoBancoDados.AdicionarParametros("@MODELO", pModelo);
                acessoBancoDados.AdicionarParametros("@TURNO", pTurno);

                string Relatorio = acessoBancoDados.ExecutarManipulacao(CommandType.StoredProcedure, "usp_ObterRelatorioImpressoras").ToString();
                return Relatorio;

            }
            catch (Exception ex)
            {
                throw new Exception("Não foi possivel obter dados das impressoras. Detalhes: " + ex.Message);
            }

        }

        public static void PopularTabelaImpressoras(List<ObjImpressoraInfo> objImpressoras, string pModelo)
        {
            AcessoBancoDados acessoBancoDados = new AcessoBancoDados();

            try
            {
                foreach (ObjImpressoraInfo impressora in objImpressoras)
                {
                    acessoBancoDados.LimparParametros();
                    acessoBancoDados.AdicionarParametros("@MODELO", pModelo);
                    acessoBancoDados.AdicionarParametros("@IP", impressora.IP);
                    acessoBancoDados.AdicionarParametros("@Toner", impressora.Toner);
                    acessoBancoDados.AdicionarParametros("@UnidImg", impressora.UnidImg);
                    acessoBancoDados.AdicionarParametros("@DataRegistro", impressora.DataRegistro);
                    acessoBancoDados.AdicionarParametros("@Turno", impressora.Turno);
                    acessoBancoDados.AdicionarParametros("@Toner_Preto", impressora.Toner_Preto);
                    acessoBancoDados.AdicionarParametros("@Toner_Cyan", impressora.Toner_Cyan);
                    acessoBancoDados.AdicionarParametros("@Toner_Magenta", impressora.Toner_Magenta);
                    acessoBancoDados.AdicionarParametros("@Toner_Amarelo", impressora.Toner_Amarelo);

                    acessoBancoDados.ExecutarManipulacao(CommandType.StoredProcedure, "usp_InserirDadosImpressora");
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Não foi possivel inserir dados das impressoras. Detalhes: " + ex.Message);
            }
        }

        public static List<ObjImpressoraInfo> ObterImpressorasUltimoRegistro(string pModelo)
        {
            AcessoBancoDados acessoBancoDados = new AcessoBancoDados();

            try
            {
                List<ObjImpressoraInfo> objImpressoras = new List<ObjImpressoraInfo>();
                acessoBancoDados.LimparParametros();
                acessoBancoDados.AdicionarParametros("@MODELO", pModelo);

                //Carregar DataTable
                DataTable dt = acessoBancoDados.ExecutarConsulta(CommandType.StoredProcedure, "usp_ObterImpressorasPorData");

                if (dt.Rows.Count > 0)
                {
                    objImpressoras = acessoBancoDados.ConverteParaLista<ObjImpressoraInfo>(dt);
                }
                return objImpressoras;
            }
            catch (Exception ex)
            {
                throw new Exception("Não foi possivel obter lista de impressoras da data anterior. Detalhes: " + ex.Message);
            }
        }

        //public static List<ObjImpressoraInfo> ObterImpressorasUltimoRegistro(string pModelo)
        //{
        //    AcessoBancoDados acessoBancoDados = new AcessoBancoDados();

        //    try
        //    {
        //        List<ObjImpressoraInfo> objImpressoras = new List<ObjImpressoraInfo>();
        //        acessoBancoDados.LimparParametros();
        //        acessoBancoDados.AdicionarParametros("@MODELO", pModelo);

        //        //Carregar DataTable
        //        DataTable dt = acessoBancoDados.ExecutarConsulta(CommandType.StoredProcedure, "usp_ObterImpressorasUltimoRegistro");

        //        if (dt.Rows.Count > 0)
        //        {
        //            objImpressoras = acessoBancoDados.ConverteParaLista<ObjImpressoraInfo>(dt);
        //        }
        //        return objImpressoras;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception("Não foi possivel obter lista de impressoras do ultimo registro. Detalhes: " + ex.Message);
        //    }
        //}

        public static List<ImpressorasInfo> ObterImpressoras(string pModelo)
        {
            AcessoBancoDados acessoBancoDados = new AcessoBancoDados();

            try
            {
                List<ImpressorasInfo> objImpressoras = new List<ImpressorasInfo>();
                acessoBancoDados.LimparParametros();
                acessoBancoDados.AdicionarParametros("@MODELO", pModelo);

                //Carregar DataTable
                DataTable dt = acessoBancoDados.ExecutarConsulta(CommandType.StoredProcedure, "usp_ObterImpressoras");

                if (dt.Rows.Count > 0)
                {
                    objImpressoras = acessoBancoDados.ConverteParaLista<ImpressorasInfo>(dt);
                }
                return objImpressoras;
            }
            catch (Exception ex)
            {
                throw new Exception("Não foi possivel obter lista de impressoras. Detalhes: " + ex.Message);
            }
        }

        private static int VerificaTurno(DateTime date)
        {
            if (date.Hour >= 5 && date.Hour <= 7)
                return 1;
            else if (date.Hour >= 13 && date.Hour <= 15)
                return 2;
            else if (date.Hour >= 21 && date.Hour <= 23)
                return 3;
            else
                return 0;

        }

        private static string VerificaTurnoString(DateTime date)
        {
            if (date.Hour >= 5 && date.Hour <= 7)
            {
                string Turno = "Primeiro Turno";
                return Turno;
            }
            else if (date.Hour >= 13 && date.Hour <= 15)
            {
                string Turno = "Segundo Turno";
                return Turno;
            }
            else if (date.Hour >= 21 && date.Hour <= 23)
            {
                string Turno = "Terceiro Turno";
                return Turno;
            }
            else
            {
                string Turno = "Fora de Turno";
                return Turno;
            }


        }
    }
}
